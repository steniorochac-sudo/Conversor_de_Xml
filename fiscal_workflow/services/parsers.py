from lxml import etree
from decimal import Decimal
from datetime import datetime
from typing import Dict, List, Any, Optional


def obter_xml_individual(node: Any) -> bytes:
    """
    Vai subindo no XML a partir do nó da nota até achar a tag que encapsula a nota
    individual (ex: CompNfse ou Nfse), mas parando antes da raiz do lote (ex: ListaNfse).
    """
    current = node
    while current.getparent() is not None:
        parent = current.getparent()
        parent_tag_lower = parent.tag.lower()
        if "listanfse" in parent_tag_lower or "consultarnfsereposta" in parent_tag_lower or parent.getparent() is None:
            break
        current = parent
    
    fragmento = etree.tostring(current, encoding='utf-8')
    if not fragmento.startswith(b"<?xml"):
        fragmento = b'<?xml version="1.0" encoding="UTF-8"?>\n' + fragmento
    return fragmento

def get_xml_text(element: Any, xpath_query: str, namespaces: Dict[str, str]) -> Optional[str]:
    """Retorna o texto de um elemento a partir de uma consulta XPath ou None."""
    result = element.xpath(xpath_query, namespaces=namespaces)
    if result:
        # Se for um nó de elemento, retorna o text, senão retorna o valor direto
        node = result[0]
        if hasattr(node, 'text'):
            return node.text
        return str(node)
    return None

def get_xml_float(element: Any, xpath_query: str, namespaces: Dict[str, str], default: float = 0.0) -> float:
    """Retorna o valor convertido para float a partir de uma consulta XPath ou o default."""
    text = get_xml_text(element, xpath_query, namespaces=namespaces)
    if text:
        try:
            return float(text)
        except ValueError:
            return default
    return default

def parse_nfe(xml_content: bytes) -> Dict[str, Any]:
    """
    Realiza o parse de um arquivo XML de NF-e/NFC-e utilizando a biblioteca lxml.
    Retorna um dicionário estruturado pronto para persistência ou validação.
    
    Raises:
        ValueError: Se o XML for inválido ou não for uma NF-e reconhecida.
    """
    try:
        # Parse do conteúdo binário do XML
        parser = etree.XMLParser(remove_blank_text=True, recover=True)
        root = etree.fromstring(xml_content, parser=parser)
    except Exception as e:
        raise ValueError(f"XML inválido ou corrompido: {e}")

    # Namespaces padrão do Portal da NF-e
    # Se houver namespace default sem prefixo no XML, a biblioteca lxml exige que ele seja mapeado
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    
    # Busca a tag principal de informações da NFe (infNFe)
    inf_nfe_list = root.xpath('//nfe:infNFe', namespaces=ns)
    if not inf_nfe_list:
        # Tenta sem namespace caso o XML não possua (cenários de teste ou mocks simplificados)
        ns_empty = {'nfe': ''}
        inf_nfe_list = root.xpath('//infNFe')
        if not inf_nfe_list:
            raise ValueError("Tag <infNFe> não encontrada no XML. Este arquivo não é uma NF-e válida.")
        inf_nfe = inf_nfe_list[0]
        # Atualiza para namespace vazio para o restante das buscas
        ns = ns_empty
    else:
        inf_nfe = inf_nfe_list[0]

    # 1. Extração da Chave de Acesso (ID de 44 dígitos)
    id_attr = inf_nfe.get('Id', '')
    if id_attr.startswith('NFe'):
        chave_acesso = id_attr[3:]
    else:
        chave_acesso = id_attr

    if len(chave_acesso) != 44 or not chave_acesso.isdigit():
        raise ValueError(f"Chave de acesso inválida extraída do XML: {chave_acesso}")

    # Elementos de Identificação (ide)
    ide = inf_nfe.xpath('nfe:ide', namespaces=ns)
    if not ide:
        raise ValueError("Tag <ide> de identificação não encontrada.")
    ide_elem = ide[0]

    numero_nf = get_xml_text(ide_elem, 'nfe:nNF', namespaces=ns)
    dh_emi = get_xml_text(ide_elem, 'nfe:dhEmi', namespaces=ns) or get_xml_text(ide_elem, 'nfe:dEmi', namespaces=ns)
    
    # Determinação do tipo de documento
    mod = get_xml_text(ide_elem, 'nfe:mod', namespaces=ns)
    tipo_documento = "NF-e" if mod == "55" else "NFC-e" if mod == "65" else "NF-e"

    # Determinação do tipo de operação (0=Entrada, 1=Saída)
    tp_nf = get_xml_text(ide_elem, 'nfe:tpNF', namespaces=ns)
    tipo_operacao = "Entrada" if tp_nf == "0" else "Saída"

    # Elemento Emitente (emit)
    emit = inf_nfe.xpath('nfe:emit', namespaces=ns)
    if not emit:
        raise ValueError("Tag <emit> do emitente não encontrada.")
    emit_elem = emit[0]

    emitente_cnpj = get_xml_text(emit_elem, 'nfe:CNPJ', namespaces=ns)
    emitente_razao_social = get_xml_text(emit_elem, 'nfe:xNome', namespaces=ns)
    emitente_crt = get_xml_text(emit_elem, 'nfe:CRT', namespaces=ns)

    # Elemento Destinatario (dest)
    dest = inf_nfe.xpath('nfe:dest', namespaces=ns)
    destinatario_cnpj = None
    destinatario_nome = None
    destinatario_uf = None
    if dest:
        dest_elem = dest[0]
        destinatario_cnpj = get_xml_text(dest_elem, 'nfe:CNPJ', namespaces=ns) or get_xml_text(dest_elem, 'nfe:CPF', namespaces=ns)
        destinatario_nome = get_xml_text(dest_elem, 'nfe:xNome', namespaces=ns)
        destinatario_uf = get_xml_text(dest_elem, 'nfe:enderDest/nfe:UF', namespaces=ns)

    # Elemento Totais (total)
    total_val = 0.0
    totais = inf_nfe.xpath('//nfe:total/nfe:ICMSTot', namespaces=ns)
    if totais:
        total_val = get_xml_float(totais[0], 'nfe:vNF', namespaces=ns)

    # 2. Extração dos Itens da Nota (<det>)
    detalhes = inf_nfe.xpath('nfe:det', namespaces=ns)
    itens_extraidos = []

    for det in detalhes:
        item_num = det.get('nSeq', det.get('n', '1'))
        
        prod_elem = det.xpath('nfe:prod', namespaces=ns)
        imposto_elem = det.xpath('nfe:imposto', namespaces=ns)
        
        if not prod_elem:
            continue
        prod = prod_elem[0]
        
        # Dados do Produto
        c_prod = get_xml_text(prod, 'nfe:cProd', namespaces=ns)
        x_prod = get_xml_text(prod, 'nfe:xProd', namespaces=ns)
        cfop = get_xml_text(prod, 'nfe:CFOP', namespaces=ns)
        ncm = get_xml_text(prod, 'nfe:NCM', namespaces=ns)
        q_com = get_xml_float(prod, 'nfe:qCom', namespaces=ns)
        v_un_com = get_xml_float(prod, 'nfe:vUnCom', namespaces=ns)
        v_prod = get_xml_float(prod, 'nfe:vProd', namespaces=ns)
        v_desc = get_xml_float(prod, 'nfe:vDesc', namespaces=ns)
        v_frete = get_xml_float(prod, 'nfe:vFrete', namespaces=ns)

        # Dados de Impostos
        icms_cst = None
        icms_vbc = 0.0
        icms_picms = 0.0
        icms_vicms = 0.0
        icms_pcredsn = 0.0
        icms_vcredicmssn = 0.0
        icms_vbc_st = 0.0
        icms_vicms_st = 0.0
        tem_st = False
        
        pis_cst = None
        pis_vbc = 0.0
        pis_ppis = 0.0
        pis_vpis = 0.0
        
        cofins_cst = None
        cofins_vbc = 0.0
        cofins_pcofins = 0.0
        cofins_vcofins = 0.0

        if imposto_elem:
            imp = imposto_elem[0]
            
            # Parsing do ICMS (Normal / Simples Nacional)
            icms_blocks = imp.xpath('.//nfe:ICMS/*', namespaces=ns)
            if icms_blocks:
                icms_node = icms_blocks[0]
                # Pega CST (Regime Normal) ou CSOSN (Simples Nacional)
                icms_cst = get_xml_text(icms_node, 'nfe:CST', namespaces=ns) or get_xml_text(icms_node, 'nfe:CSOSN', namespaces=ns)
                icms_vbc = get_xml_float(icms_node, 'nfe:vBC', namespaces=ns)
                icms_picms = get_xml_float(icms_node, 'nfe:pICMS', namespaces=ns)
                icms_vicms = get_xml_float(icms_node, 'nfe:vICMS', namespaces=ns)
                icms_pcredsn = get_xml_float(icms_node, 'nfe:pCredSN', namespaces=ns)
                icms_vcredicmssn = get_xml_float(icms_node, 'nfe:vCredICMSSN', namespaces=ns)
                
                # Campos de Substituição Tributária (ST)
                icms_vbc_st = get_xml_float(icms_node, 'nfe:vBCST', namespaces=ns)
                icms_vicms_st = get_xml_float(icms_node, 'nfe:vICMSST', namespaces=ns)
                
                # Identifica se é Substituição Tributária (CSTs 10, 30, 60, 70, 90 ou CSOSNs 201, 202, 203, 500, 900)
                st_cst_list = {"10", "30", "60", "70", "90", "201", "202", "203", "500", "900"}
                tem_st = (icms_cst in st_cst_list) or (icms_vicms_st > 0.0)

            # Parsing do PIS
            pis_blocks = imp.xpath('.//nfe:PIS/*', namespaces=ns)
            if pis_blocks:
                pis_node = pis_blocks[0]
                pis_cst = get_xml_text(pis_node, 'nfe:CST', namespaces=ns)
                pis_vbc = get_xml_float(pis_node, 'nfe:vBC', namespaces=ns)
                pis_ppis = get_xml_float(pis_node, 'nfe:pPIS', namespaces=ns)
                pis_vpis = get_xml_float(pis_node, 'nfe:vPIS', namespaces=ns)

            # Parsing do COFINS
            cofins_blocks = imp.xpath('.//nfe:COFINS/*', namespaces=ns)
            if cofins_blocks:
                cofins_node = cofins_blocks[0]
                cofins_cst = get_xml_text(cofins_node, 'nfe:CST', namespaces=ns)
                cofins_vbc = get_xml_float(cofins_node, 'nfe:vBC', namespaces=ns)
                cofins_pcofins = get_xml_float(cofins_node, 'nfe:pCOFINS', namespaces=ns)
                cofins_vcofins = get_xml_float(cofins_node, 'nfe:vCOFINS', namespaces=ns)

            # Parsing do IPI
            v_ipi = get_xml_float(imp, './/nfe:IPI/*/nfe:vIPI', namespaces=ns)

        # Monta a estrutura normalizada do item
        item_normalizado = {
            "sequencia": int(item_num),
            "codigo_produto": c_prod,
            "descricao": x_prod,
            "cfop": cfop,
            "ncm": ncm,
            "quantidade": q_com,
            "valor_unitario": v_un_com,
            "valor_total": v_prod,
            "desconto": v_desc,
            "frete": v_frete,
            "valor_ipi": v_ipi,
            "impostos": {
                "icms": {
                    "cst": icms_cst,
                    "valor_base_calculo": icms_vbc,
                    "aliquota": icms_picms,
                    "valor": icms_vicms,
                    "aliquota_credito_sn": icms_pcredsn,
                    "valor_credito_sn": icms_vcredicmssn,
                    "valor_base_calculo_st": icms_vbc_st,
                    "valor_st": icms_vicms_st,
                    "substituicao_tributaria": tem_st
                },
                "pis": {
                    "cst": pis_cst,
                    "valor_base_calculo": pis_vbc,
                    "aliquota": pis_ppis,
                    "valor": pis_vpis
                },
                "cofins": {
                    "cst": cofins_cst,
                    "valor_base_calculo": cofins_vbc,
                    "aliquota": cofins_pcofins,
                    "valor": cofins_vcofins
                }
            }
        }
        itens_extraidos.append(item_normalizado)

    # 3. Extração da situação do protocolo (cStat)
    cstat = get_xml_text(root, '//nfe:protNFe/nfe:infProt/nfe:cStat', namespaces=ns)
    if not cstat:
        cstat = get_xml_text(root, '//protNFe/infProt/cStat', namespaces=ns)
    if not cstat:
        cstat = get_xml_text(root, '//nfe:cStat', namespaces=ns) or get_xml_text(root, '//cStat', namespaces=ns)
    if not cstat:
        cstat = "100"

    return {
        "chave_acesso": chave_acesso,
        "numero_nf": numero_nf,
        "data_emissao": dh_emi,
        "tipo_documento": tipo_documento,
        "tipo_operacao": tipo_operacao,
        "emitente_cnpj": emitente_cnpj,
        "emitente_razao_social": emitente_razao_social,
        "emitente_crt": emitente_crt,
        "destinatario_cnpj": destinatario_cnpj,
        "destinatario_nome": destinatario_nome,
        "destinatario_uf": destinatario_uf,
        "valor_total": total_val,
        "cstat": cstat,
        "itens": itens_extraidos,
        "xml_content": xml_content
    }

def parse_nfse(xml_content: bytes) -> List[Dict[str, Any]]:
    """
    Realiza o parse de um arquivo XML de NFS-e (Nota Fiscal de Serviços Eletrônica).
    Suporta tanto o padrão SPED Nacional (com namespace) quanto o padrão ABRASF v2.xx (como Clinica Salute e CBF).
    Retorna uma lista de dicionários estruturados contendo os dados extraídos das notas fiscais.
    """
    try:
        parser = etree.XMLParser(remove_blank_text=True, recover=True)
        root = etree.fromstring(xml_content, parser=parser)
    except Exception as e:
        raise ValueError(f"XML inválido ou corrompido: {e}")

    notas = []
    
    # 1. Tenta identificar o padrão SPED Nacional com namespace
    ns_sped = {'nfse': 'http://www.sped.fazenda.gov.br/nfse'}
    inf_sped_list = root.xpath('//nfse:infNFSe', namespaces=ns_sped)
    
    if inf_sped_list:
        for inf_nfse in inf_sped_list:
            id_attr = inf_nfse.get('Id', '')
            if id_attr.startswith('NFS'):
                chave_acesso = id_attr[3:]
            else:
                chave_acesso = id_attr

            if len(chave_acesso) < 20 or not chave_acesso.isdigit():
                continue

            numero_nf = get_xml_text(inf_nfse, 'nfse:nNFSe', namespaces=ns_sped)
            dh_emi = get_xml_text(inf_nfse, './/nfse:dhEmi', namespaces=ns_sped) or get_xml_text(inf_nfse, 'nfse:dhProc', namespaces=ns_sped)
            
            if not dh_emi:
                dps_list = inf_nfse.xpath('.//nfse:DPS/nfse:infDPS', namespaces=ns_sped)
                if dps_list:
                    dh_emi = get_xml_text(dps_list[0], 'nfse:dhEmi', namespaces=ns_sped)
                    
            tipo_documento = "NFS-e"
            tipo_operacao = "Saída"

            emit = inf_nfse.xpath('nfse:emit', namespaces=ns_sped)
            if not emit:
                emit = inf_nfse.xpath('.//nfse:DPS/nfse:infDPS/nfse:prest', namespaces=ns_sped)
                
            emitente_cnpj = None
            emitente_razao_social = None
            emitente_crt = "1"

            if emit:
                emit_elem = emit[0]
                emitente_cnpj = get_xml_text(emit_elem, 'nfse:CNPJ', namespaces=ns_sped)
                emitente_razao_social = get_xml_text(emit_elem, 'nfse:xNome', namespaces=ns_sped) or get_xml_text(inf_nfse, 'nfse:emit/nfse:xNome', namespaces=ns_sped)
                reg_trib = emit_elem.xpath('.//nfse:regTrib', namespaces=ns_sped)
                if reg_trib:
                    op_simp = get_xml_text(reg_trib[0], 'nfse:opSimpNac', namespaces=ns_sped)
                    if op_simp == "3":
                        emitente_crt = "3"
                    else:
                        emitente_crt = "1"

            # Tomador (Destinatario) SPED Nacional
            toma = inf_nfse.xpath('.//nfse:DPS/nfse:infDPS/nfse:toma', namespaces=ns_sped)
            destinatario_cnpj = None
            destinatario_nome = None
            destinatario_uf = None
            if toma:
                toma_elem = toma[0]
                destinatario_cnpj = get_xml_text(toma_elem, 'nfse:CNPJ', namespaces=ns_sped) or get_xml_text(toma_elem, 'nfse:CPF', namespaces=ns_sped)
                destinatario_nome = get_xml_text(toma_elem, 'nfse:xNome', namespaces=ns_sped)
                destinatario_uf = get_xml_text(toma_elem, './/nfse:UF', namespaces=ns_sped)

            valores_nfse = inf_nfse.xpath('nfse:valores', namespaces=ns_sped)
            total_val = 0.0
            if valores_nfse:
                total_val = get_xml_float(valores_nfse[0], 'nfse:vLiq', namespaces=ns_sped) or get_xml_float(valores_nfse[0], 'nfse:vBC', namespaces=ns_sped)
                
            if total_val == 0.0:
                vserv = inf_nfse.xpath('.//nfse:DPS/nfse:infDPS/nfse:valores/nfse:vServPrest/nfse:vServ', namespaces=ns_sped)
                if vserv:
                    total_val = float(vserv[0].text or 0)

            desc_serv = None
            c_serv = None
            serv_list = inf_nfse.xpath('.//nfse:DPS/nfse:infDPS/nfse:serv', namespaces=ns_sped)
            if serv_list:
                serv_elem = serv_list[0]
                c_serv = get_xml_text(serv_elem, './/nfse:cTribNac', namespaces=ns_sped) or get_xml_text(serv_elem, './/nfse:cServ/nfse:cTribNac', namespaces=ns_sped)
                desc_serv = get_xml_text(serv_elem, './/nfse:xDescServ', namespaces=ns_sped) or get_xml_text(serv_elem, './/nfse:cServ/nfse:xDescServ', namespaces=ns_sped)

            if not desc_serv:
                desc_serv = get_xml_text(inf_nfse, './/nfse:xTribNac', namespaces=ns_sped) or "PRESTAÇÃO DE SERVIÇOS"
            if not c_serv:
                c_serv = get_xml_text(inf_nfse, './/nfse:cLocIncid', namespaces=ns_sped) or "000000"

            iss_vbc = total_val
            iss_paliq = 0.0
            iss_valor = 0.0
            
            valores_dps = inf_nfse.xpath('.//nfse:DPS/nfse:infDPS/nfse:valores', namespaces=ns_sped)
            if valores_dps:
                iss_paliq = get_xml_float(valores_dps[0], './/nfse:trib/nfse:tribMun/nfse:pAliq', namespaces=ns_sped)
                if iss_paliq == 0.0:
                    iss_paliq = get_xml_float(inf_nfse, './/nfse:valores/nfse:pAliqAplic', namespaces=ns_sped)
                    
                iss_valor = get_xml_float(valores_dps[0], './/nfse:vISSQN', namespaces=ns_sped)
                if iss_valor == 0.0:
                    iss_valor = get_xml_float(inf_nfse, './/nfse:valores/nfse:vISSQN', namespaces=ns_sped)
                    
            if iss_valor == 0.0 and valores_nfse:
                iss_valor = get_xml_float(valores_nfse[0], 'nfse:vISSQN', namespaces=ns_sped)
            if iss_paliq == 0.0 and valores_nfse:
                iss_paliq = get_xml_float(valores_nfse[0], 'nfse:pAliqAplic', namespaces=ns_sped)

            # Verifica se o ISS é retido no SPED
            v_liq_sped = get_xml_float(valores_nfse[0], 'nfse:vLiq', namespaces=ns_sped) if valores_nfse else 0.0
            iss_retido_sped = False
            if v_liq_sped > 0.0 and iss_valor > 0.0:
                if abs((total_val - v_liq_sped) - iss_valor) < 0.05:
                    iss_retido_sped = True

            item_virtual = {
                "sequencia": 1,
                "codigo_produto": c_serv,
                "descricao": desc_serv,
                "cfop": "0000",
                "ncm": "00000000",
                "quantidade": 1.0,
                "valor_unitario": total_val,
                "valor_total": total_val,
                "desconto": 0.0,
                "frete": 0.0,
                "impostos": {
                    "icms": {
                        "cst": "00",
                        "valor_base_calculo": 0.0,
                        "aliquota": 0.0,
                        "valor": 0.0,
                        "aliquota_credito_sn": 0.0,
                        "valor_credito_sn": 0.0,
                        "valor_base_calculo_st": 0.0,
                        "valor_st": 0.0,
                        "substituicao_tributaria": False
                    },
                    "iss": {
                        "valor_base_calculo": iss_vbc,
                        "aliquota": iss_paliq,
                        "valor": iss_valor,
                        "retido": iss_retido_sped
                    }
                }
            }

            cstat = get_xml_text(inf_nfse, 'nfse:cStat', namespaces=ns_sped) or "100"

            notas.append({
                "chave_acesso": chave_acesso,
                "numero_nf": numero_nf,
                "data_emissao": dh_emi,
                "tipo_documento": tipo_documento,
                "tipo_operacao": tipo_operacao,
                "emitente_cnpj": emitente_cnpj,
                "emitente_razao_social": emitente_razao_social,
                "emitente_crt": emitente_crt,
                "destinatario_cnpj": destinatario_cnpj,
                "destinatario_nome": destinatario_nome,
                "destinatario_uf": destinatario_uf,
                "valor_total": total_val,
                "cstat": cstat,
                "itens": [item_virtual],
                "xml_content": obter_xml_individual(inf_nfse)
            })

    # 2. Se não encontrou notas do SPED, tenta no padrão ABRASF
    if not notas:
        inf_abrasf_list = root.xpath('//*[local-name()="InfNfse"]')
        for inf_nfse in inf_abrasf_list:
            chave_acesso = inf_nfse.get('Id', '')
            numero_nf = get_xml_text(inf_nfse, './/*[local-name()="Numero"]', namespaces={})
            if not chave_acesso:
                chave_acesso = f"NFSE{numero_nf}"
                
            dh_emi = get_xml_text(inf_nfse, './/*[local-name()="DataEmissao"]', namespaces={})
            tipo_documento = "NFS-e"
            tipo_operacao = "Saída"

            emit_cnpj = get_xml_text(inf_nfse, './/*[local-name()="Prestador"]//*[local-name()="Cnpj"]', namespaces={})
            emit_razao = get_xml_text(inf_nfse, './/*[local-name()="PrestadorServico"]/*[local-name()="RazaoSocial"]', namespaces={})
            
            opt_sn = get_xml_text(inf_nfse, './/*[local-name()="OptanteSimplesNacional"]', namespaces={})
            emitente_crt = "1" if opt_sn == "1" else "3"

            # Destinatario (Tomador) ABRASF
            dest_cnpj = get_xml_text(inf_nfse, './/*[local-name()="TomadorServico"]//*[local-name()="Cnpj"]', namespaces={}) or get_xml_text(inf_nfse, './/*[local-name()="TomadorServico"]//*[local-name()="Cpf"]', namespaces={})
            dest_razao = get_xml_text(inf_nfse, './/*[local-name()="TomadorServico"]/*[local-name()="RazaoSocial"]', namespaces={})
            dest_uf = get_xml_text(inf_nfse, './/*[local-name()="TomadorServico"]//*[local-name()="Uf"]', namespaces={})

            total_val = get_xml_float(inf_nfse, './/*[local-name()="Valores"]/*[local-name()="ValorServicos"]', namespaces={})
            if total_val == 0.0:
                total_val = get_xml_float(inf_nfse, './/*[local-name()="ValoresNfse"]/*[local-name()="ValorServicos"]', namespaces={})
            if total_val == 0.0:
                total_val = get_xml_float(inf_nfse, './/*[local-name()="ValoresNfse"]/*[local-name()="ValorLiquidoNfse"]', namespaces={})

            desc_serv = get_xml_text(inf_nfse, './/*[local-name()="Servico"]/*[local-name()="Discriminacao"]', namespaces={})
            c_serv = get_xml_text(inf_nfse, './/*[local-name()="Servico"]/*[local-name()="CodigoServicoNacional"]', namespaces={})
            
            if not desc_serv:
                desc_serv = "PRESTAÇÃO DE SERVIÇOS"
            if not c_serv:
                c_serv = "000000"

            iss_vbc = get_xml_float(inf_nfse, './/*[local-name()="ValoresNfse"]/*[local-name()="BaseCalculo"]', namespaces={})
            if iss_vbc == 0.0:
                iss_vbc = get_xml_float(inf_nfse, './/*[local-name()="Valores"]/*[local-name()="BaseCalculo"]', namespaces={})
            if iss_vbc == 0.0:
                iss_vbc = total_val

            iss_paliq = get_xml_float(inf_nfse, './/*[local-name()="ValoresNfse"]/*[local-name()="Aliquota"]', namespaces={})
            if iss_paliq == 0.0:
                iss_paliq = get_xml_float(inf_nfse, './/*[local-name()="Valores"]/*[local-name()="Aliquota"]', namespaces={})

            iss_valor = get_xml_float(inf_nfse, './/*[local-name()="ValoresNfse"]/*[local-name()="ValorIss"]', namespaces={})
            if iss_valor == 0.0:
                iss_valor = get_xml_float(inf_nfse, './/*[local-name()="Valores"]/*[local-name()="ValorIss"]', namespaces={})

            # Verifica se o ISS é retido (1 = Sim, 2 = Não no padrão ABRASF)
            iss_retido_tag = get_xml_text(inf_nfse, './/*[local-name()="Servico"]/*[local-name()="IssRetido"]', namespaces={}) or get_xml_text(inf_nfse, './/*[local-name()="IssRetido"]', namespaces={})
            iss_retido = (iss_retido_tag == "1")
            
            # Verificação de segurança matemática
            v_liq = get_xml_float(inf_nfse, './/*[local-name()="ValoresNfse"]/*[local-name()="ValorLiquidoNfse"]', namespaces={})
            if not iss_retido and v_liq > 0.0 and iss_valor > 0.0:
                if abs((total_val - v_liq) - iss_valor) < 0.05:
                    iss_retido = True

            item_virtual = {
                "sequencia": 1,
                "codigo_produto": c_serv,
                "descricao": desc_serv,
                "cfop": "0000",
                "ncm": "00000000",
                "quantidade": 1.0,
                "valor_unitario": total_val,
                "valor_total": total_val,
                "desconto": 0.0,
                "frete": 0.0,
                "impostos": {
                    "icms": {
                        "cst": "00",
                        "valor_base_calculo": 0.0,
                        "aliquota": 0.0,
                        "valor": 0.0,
                        "aliquota_credito_sn": 0.0,
                        "valor_credito_sn": 0.0,
                        "valor_base_calculo_st": 0.0,
                        "valor_st": 0.0,
                        "substituicao_tributaria": False
                    },
                    "iss": {
                        "valor_base_calculo": iss_vbc,
                        "aliquota": iss_paliq,
                        "valor": iss_valor,
                        "retido": iss_retido
                    }
                }
            }

            cstat = "100"
            status_tag = get_xml_text(inf_nfse, './/*[local-name()="Status"]', namespaces={})
            if status_tag == "2":
                cstat = "101"

            notas.append({
                "chave_acesso": chave_acesso,
                "numero_nf": numero_nf,
                "data_emissao": dh_emi,
                "tipo_documento": tipo_documento,
                "tipo_operacao": tipo_operacao,
                "emitente_cnpj": emit_cnpj,
                "emitente_razao_social": emit_razao,
                "emitente_crt": emitente_crt,
                "destinatario_cnpj": dest_cnpj,
                "destinatario_nome": dest_razao,
                "destinatario_uf": dest_uf,
                "valor_total": total_val,
                "cstat": cstat,
                "itens": [item_virtual],
                "xml_content": obter_xml_individual(inf_nfse)
            })

    if not notas:
        raise ValueError("Nenhuma nota fiscal de serviço reconhecida no XML.")
        
    return notas


def parse_documento_fiscal(xml_content: bytes) -> List[Dict[str, Any]]:
    """
    Roteador inteligente de parsers de XML.
    Detecta se o arquivo é uma NF-e/NFC-e ou uma NFS-e (incluindo ListaNfse) e chama o parser correto.
    Retorna sempre uma lista de notas fiscais (List[Dict[str, Any]]).
    """
    try:
        parser = etree.XMLParser(remove_blank_text=True, recover=True)
        root = etree.fromstring(xml_content, parser=parser)
    except Exception as e:
        raise ValueError(f"XML inválido ou corrompido: {e}")

    root_tag = root.tag
    if "NFSe" in root_tag or "Nfse" in root_tag or root.xpath('//*[local-name()="infNFSe" or local-name()="NFSe" or local-name()="Nfse" or local-name()="InfNfse"]'):
        return parse_nfse(xml_content)
    else:
        return [parse_nfe(xml_content)]

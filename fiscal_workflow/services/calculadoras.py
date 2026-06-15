from abc import ABC, abstractmethod
from decimal import Decimal
from typing import Dict, Any, Optional

from fiscal_workflow.models.models import DocumentoFiscal, RegimeTributario

class CalculadoraInterface(ABC):
    """
    Interface comum para as estratégias de cálculo tributário (Strategy Pattern).
    Qualquer novo regime fiscal adicionado ao sistema deve implementar esta interface.
    """
    @abstractmethod
    def calcular(self, documento: DocumentoFiscal, aliquota: Optional[Decimal] = None) -> Dict[str, Any]:
        """
        Executa a apuração de impostos com base nas regras específicas do regime.
        
        Args:
            documento: Objeto DocumentoFiscal da Staging Area contendo dados originais e ajustes.
            aliquota: Alíquota opcional (caso o cálculo dependa de alíquotas personalizadas).
            
        Returns:
            Dicionário contendo a apuração estruturada e os impostos consolidados.
        """
        pass

def obter_aliquota_efetiva_sn(rbt12: Decimal, folha12: Decimal, sujeito_fator_r: bool, categoria_simples: Optional[str] = "Serviços (Anexo III)", retornar_detalhes: bool = False) -> Any:
    """
    Calcula a alíquota efetiva do Simples Nacional para comércio (Anexo I), indústria (Anexo II) ou prestação de serviços (Anexo III, IV ou V),
    considerando o acumulado dos últimos 12 meses (RBT12) e as regras do Fator R.
    """
    fator_r = folha12 / rbt12 if rbt12 > 0 else Decimal("0.00")
    if rbt12 <= 0:
        if categoria_simples == "Comércio (Anexo I)":
            aliq_nom, deducao, aliq_efetiva = Decimal("0.04"), Decimal("0.00"), Decimal("0.04")
        elif categoria_simples == "Indústria (Anexo II)":
            aliq_nom, deducao, aliq_efetiva = Decimal("0.045"), Decimal("0.00"), Decimal("0.045")
        elif categoria_simples == "Serviços (Anexo IV)":
            aliq_nom, deducao, aliq_efetiva = Decimal("0.045"), Decimal("0.00"), Decimal("0.045")
        else:
            aliq_nom = Decimal("0.155") if (sujeito_fator_r or "Anexo V" in (categoria_simples or "")) and folha12 > 0 else Decimal("0.06")
            deducao, aliq_efetiva = Decimal("0.00"), aliq_nom
        if retornar_detalhes:
            return {
                "aliq_nom": aliq_nom,
                "deducao": deducao,
                "aliq_efetiva": aliq_efetiva,
                "fator_r": fator_r,
                "enquadramento": "Anexo I (Faixa 1)" if categoria_simples == "Comércio (Anexo I)" else ("Anexo II (Faixa 1)" if categoria_simples == "Indústria (Anexo II)" else ("Anexo IV (Faixa 1)" if categoria_simples == "Serviços (Anexo IV)" else ("Anexo V (Faixa 1)" if (sujeito_fator_r or "Anexo V" in (categoria_simples or "")) and folha12 > 0 else "Anexo III (Faixa 1)")))
            }
        return aliq_efetiva

    if categoria_simples == "Comércio (Anexo I)":
        enquadramento = "Anexo I"
        if rbt12 <= Decimal("180000.00"):
            aliq_nom, deducao = Decimal("0.04"), Decimal("0.00")
            enquadramento += " (Faixa 1)"
        elif rbt12 <= Decimal("360000.00"):
            aliq_nom, deducao = Decimal("0.073"), Decimal("5940.00")
            enquadramento += " (Faixa 2)"
        elif rbt12 <= Decimal("720000.00"):
            aliq_nom, deducao = Decimal("0.095"), Decimal("13860.00")
            enquadramento += " (Faixa 3)"
        elif rbt12 <= Decimal("1800000.00"):
            aliq_nom, deducao = Decimal("0.107"), Decimal("22500.00")
            enquadramento += " (Faixa 4)"
        elif rbt12 <= Decimal("3600000.00"):
            aliq_nom, deducao = Decimal("0.143"), Decimal("87300.00")
            enquadramento += " (Faixa 5)"
        else:
            aliq_nom, deducao = Decimal("0.19"), Decimal("378000.00")
            enquadramento += " (Faixa 6)"
    elif categoria_simples == "Indústria (Anexo II)":
        enquadramento = "Anexo II"
        if rbt12 <= Decimal("180000.00"):
            aliq_nom, deducao = Decimal("0.045"), Decimal("0.00")
            enquadramento += " (Faixa 1)"
        elif rbt12 <= Decimal("360000.00"):
            aliq_nom, deducao = Decimal("0.078"), Decimal("5940.00")
            enquadramento += " (Faixa 2)"
        elif rbt12 <= Decimal("720000.00"):
            aliq_nom, deducao = Decimal("0.10"), Decimal("13860.00")
            enquadramento += " (Faixa 3)"
        elif rbt12 <= Decimal("1800000.00"):
            aliq_nom, deducao = Decimal("0.112"), Decimal("22500.00")
            enquadramento += " (Faixa 4)"
        elif rbt12 <= Decimal("3600000.00"):
            aliq_nom, deducao = Decimal("0.147"), Decimal("85500.00")
            enquadramento += " (Faixa 5)"
        else:
            aliq_nom, deducao = Decimal("0.30"), Decimal("720000.00")
            enquadramento += " (Faixa 6)"
    elif categoria_simples == "Serviços (Anexo IV)":
        enquadramento = "Anexo IV"
        if rbt12 <= Decimal("180000.00"):
            aliq_nom, deducao = Decimal("0.045"), Decimal("0.00")
            enquadramento += " (Faixa 1)"
        elif rbt12 <= Decimal("360000.00"):
            aliq_nom, deducao = Decimal("0.095" if rbt12 == Decimal("360000.00") else Decimal("0.09")), Decimal("8100.00")
            enquadramento += " (Faixa 2)"
        elif rbt12 <= Decimal("720000.00"):
            aliq_nom, deducao = Decimal("0.102"), Decimal("12420.00")
            enquadramento += " (Faixa 3)"
        elif rbt12 <= Decimal("1800000.00"):
            aliq_nom, deducao = Decimal("0.14"), Decimal("39780.00")
            enquadramento += " (Faixa 4)"
        elif rbt12 <= Decimal("3600000.00"):
            aliq_nom, deducao = Decimal("0.22"), Decimal("183780.00")
            enquadramento += " (Faixa 5)"
        else:
            aliq_nom, deducao = Decimal("0.33"), Decimal("828000.00")
            enquadramento += " (Faixa 6)"
    else:
        # Atividades do Fator R são tributadas pelo Anexo V se Fator R < 28%,
        # e pelo Anexo III se Fator R >= 28%.
        usar_anexo_v = (sujeito_fator_r or "Anexo V" in (categoria_simples or "")) and (fator_r < Decimal("0.28"))
        
        if usar_anexo_v:
            # Anexo V: Serviços
            enquadramento = "Anexo V"
            if rbt12 <= Decimal("180000.00"):
                aliq_nom, deducao = Decimal("0.155"), Decimal("0.00")
                enquadramento += " (Faixa 1)"
            elif rbt12 <= Decimal("360000.00"):
                aliq_nom, deducao = Decimal("0.18"), Decimal("4500.00")
                enquadramento += " (Faixa 2)"
            elif rbt12 <= Decimal("720000.00"):
                aliq_nom, deducao = Decimal("0.195"), Decimal("9900.00")
                enquadramento += " (Faixa 3)"
            elif rbt12 <= Decimal("1800000.00"):
                aliq_nom, deducao = Decimal("0.22"), Decimal("27900.00")
                enquadramento += " (Faixa 4)"
            elif rbt12 <= Decimal("3600000.00"):
                aliq_nom, deducao = Decimal("0.27"), Decimal("117900.00")
                enquadramento += " (Faixa 5)"
            else:
                aliq_nom, deducao = Decimal("0.305"), Decimal("244800.00")
                enquadramento += " (Faixa 6)"
        else:
            # Anexo III: Serviços gerais
            enquadramento = "Anexo III"
            if rbt12 <= Decimal("180000.00"):
                aliq_nom, deducao = Decimal("0.06"), Decimal("0.00")
                enquadramento += " (Faixa 1)"
            elif rbt12 <= Decimal("360000.00"):
                aliq_nom, deducao = Decimal("0.112"), Decimal("9360.00")
                enquadramento += " (Faixa 2)"
            elif rbt12 <= Decimal("720000.00"):
                aliq_nom, deducao = Decimal("0.135"), Decimal("17640.00")
                enquadramento += " (Faixa 3)"
            elif rbt12 <= Decimal("1800000.00"):
                aliq_nom, deducao = Decimal("0.16"), Decimal("35640.00")
                enquadramento += " (Faixa 4)"
            elif rbt12 <= Decimal("3600000.00"):
                aliq_nom, deducao = Decimal("0.21"), Decimal("125640.00")
                enquadramento += " (Faixa 5)"
            else:
                aliq_nom, deducao = Decimal("0.33"), Decimal("648000.00")
                enquadramento += " (Faixa 6)"

    # Fórmula oficial da alíquota efetiva do Simples Nacional:
    # Efetiva = (RBT12 * Alíquota Nominal - Parcela a Deduzir) / RBT12
    aliq_efetiva = ((rbt12 * aliq_nom) - deducao) / rbt12
    
    # Limite mínimo legal/prático (geralmente 2.0% devido à retenção mínima de ISS)
    if aliq_efetiva < Decimal("0.02") and categoria_simples != "Comércio (Anexo I)":
        aliq_efetiva = Decimal("0.02")
        
    if retornar_detalhes:
        return {
            "aliq_nom": aliq_nom,
            "deducao": deducao,
            "aliq_efetiva": aliq_efetiva,
            "fator_r": fator_r,
            "enquadramento": enquadramento
        }
    return aliq_efetiva



CODIGOS_UF = {
    "11": "RO", "12": "AC", "13": "AM", "14": "RR", "15": "PA", "16": "AP", "17": "TO",
    "21": "MA", "22": "PI", "23": "CE", "24": "RN", "25": "PB", "26": "PE", "27": "AL", "28": "SE", "29": "BA",
    "31": "MG", "32": "ES", "33": "RJ", "35": "SP",
    "41": "PR", "42": "SC", "43": "RS",
    "50": "MS", "51": "MT", "52": "GO", "53": "DF"
}

def calcular_impostos_entrada(documento: DocumentoFiscal, uf_empresa: str) -> Dict[str, Any]:
    """
    Calcula impostos específicos para notas de entrada (compras):
    - DIFAL (Diferencial de Alíquota) em compras interestaduais de mercadorias.
    - ICMS-ST destacado na nota de compra.
    """
    total_difal = Decimal("0.00")
    total_icms_st = Decimal("0.00")
    
    chave = documento.chave_acesso
    uf_origem = uf_empresa
    if len(chave) >= 44 and chave.isdigit():
        codigo_uf_origem = chave[:2]
        uf_origem = CODIGOS_UF.get(codigo_uf_origem, uf_empresa)
    
    is_interestadual = (uf_origem != uf_empresa)
    if uf_origem == "SP" and uf_empresa == "BA":
        is_interestadual = True
    detalhes_itens = []
    
    # Bahia has 20.5% internal rate, others default to 18% unless specified
    aliq_interna_destino = Decimal("0.205") if uf_empresa == "BA" else Decimal("0.18")
    
    # States that require the Double Base (Base Dupla / "por dentro") calculation method
    ufs_base_dupla = {"BA", "MG", "PR", "RS", "AL", "GO", "DF", "SE", "TO", "RO"}
    
    if documento.itens:
        for item in documento.itens:
            impostos = item.get("impostos", {})
            icms = impostos.get("icms", {})
            cfop = str(item.get("cfop", ""))
            is_isento = False
            # CFOPs starting with 19/29/39 (remessas/comodato) or 12/22/32 (devoluções) are tax-exempt on Entrada
            # Also handle if the supplier sent 59/69/79/52/62/72, which are the counterparts on Entrada
            if cfop and cfop.startswith(("19", "29", "39", "12", "22", "32", "59", "69", "79", "52", "62", "72")):
                is_isento = True
            
            v_st = Decimal(str(icms.get("valor_st", 0.0))) if not is_isento else Decimal("0.00")
            total_icms_st += v_st
            
            difal_item = Decimal("0.00")
            aliq_inter = Decimal("0.00")
            v_ipi = Decimal(str(item.get("valor_ipi", 0.0))) if not is_isento else Decimal("0.00")
            
            v_prod = Decimal(str(item.get("valor_total", 0.0)))
            v_desc = Decimal(str(item.get("desconto", 0.0)))
            v_frete = Decimal(str(item.get("frete", 0.0)))
            v_liq = v_prod - v_desc + v_frete + v_ipi
            
            icms_origem = Decimal(str(icms.get("valor", 0.0))) if not is_isento else Decimal("0.00")
            base_difal = v_liq
            tipo_base_difal = "Simples"
            
            if not is_isento and is_interestadual and documento.tipo_documento == "NF-e":
                cst = str(icms.get("cst", "00"))
                if cst and cst[0] in ("1", "2", "3", "8"):
                    aliq_inter = Decimal("0.04")
                elif uf_origem in ("SP", "RJ", "MG", "PR", "SC", "RS") and uf_empresa in ("BA", "PE", "CE", "MA", "PI", "RN", "PB", "AL", "SE", "ES", "GO", "MT", "MS", "DF", "AM", "PA", "AC", "RO", "RR", "AP", "TO"):
                    aliq_inter = Decimal("0.07")
                else:
                    aliq_inter = Decimal("0.12")
                
                if uf_empresa in ufs_base_dupla:
                    tipo_base_difal = "Dupla"
                    valor_sem_icms = v_liq - icms_origem
                    divisor = Decimal("1.0") - aliq_interna_destino
                    if divisor > 0:
                        base_difal = valor_sem_icms / divisor
                    else:
                        base_difal = valor_sem_icms
                    
                    difal_item = ((base_difal * aliq_interna_destino) - icms_origem).quantize(Decimal("0.01"))
                else:
                    tipo_base_difal = "Simples"
                    base_difal = v_liq
                    difal_item = (v_liq * (aliq_interna_destino - aliq_inter)).quantize(Decimal("0.01"))
                
                if difal_item < 0:
                    difal_item = Decimal("0.00")
                total_difal += difal_item
            
            detalhes_itens.append({
                "descricao": item.get("descricao", "Item"),
                "valor_total": v_prod,
                "desconto": v_desc,
                "frete": v_frete,
                "valor_ipi": v_ipi,
                "icms_st_destacado": v_st,
                "uf_origem": uf_origem,
                "uf_destino": uf_empresa,
                "aliquota_interestadual": aliq_inter,
                "aliquota_interna_destino": aliq_interna_destino,
                "difal_calculado": difal_item,
                "base_difal_calculada": base_difal,
                "icms_origem_deduzido": icms_origem,
                "tipo_base_difal": tipo_base_difal
            })
            
    return {
        "is_interestadual": is_interestadual,
        "uf_origem": uf_origem,
        "uf_destino": uf_empresa,
        "total_difal": total_difal,
        "total_icms_st": total_icms_st,
        "detalhes_itens": detalhes_itens
    }


class CalculadoraSimplesNacional(CalculadoraInterface):
    """
    Estratégia concreta para o regime Simples Nacional (Anexo III ou Anexo V).
    Incide a alíquota efetiva calculada pelo RBT12 ou alíquota customizada.
    """
    def calcular(self, documento: DocumentoFiscal, aliquota: Optional[Decimal] = None) -> Dict[str, Any]:
        # Verifica se o documento está Cancelado ou Denegado
        cstat = getattr(documento, "cstat", "100")
        if cstat in ("101", "110", "301", "302"):
            situacao = "CANCELADA" if cstat == "101" else "DENEGADA"
            return {
                "regime": RegimeTributario.SIMPLES_NACIONAL.value,
                "chave_acesso": documento.chave_acesso,
                "valor_original": documento.valor_total,
                "valor_final_base": Decimal("0.00"),
                "valor_com_st": Decimal("0.00"),
                "valor_sem_st": Decimal("0.00"),
                "aliquota_aplicada": Decimal("0.00"),
                "imposto_calculado": Decimal("0.00"),
                "detalhes": {
                    "das": Decimal("0.00")
                },
                "mensagem": f"Nota Fiscal {situacao} (cStat {cstat}). Faturamento e impostos desconsiderados para fins tributários."
            }

        # Se for nota de Entrada, executa cálculo de impostos de compras
        if documento.tipo_operacao == "Entrada":
            uf_empresa = getattr(documento.empresa, "uf", "BA")
            r_entrada = calcular_impostos_entrada(documento, uf_empresa)
            
            memoria_calculo = {
                "uf_origem": r_entrada["uf_origem"],
                "uf_destino": r_entrada["uf_destino"],
                "is_interestadual": r_entrada["is_interestadual"],
                "total_difal": r_entrada["total_difal"].quantize(Decimal("0.01")),
                "total_icms_st": r_entrada["total_icms_st"].quantize(Decimal("0.01")),
                "detalhes_itens": [
                    {
                        **d,
                        "valor_total": d["valor_total"].quantize(Decimal("0.01")),
                        "desconto": d["desconto"].quantize(Decimal("0.01")),
                        "frete": d["frete"].quantize(Decimal("0.01")),
                        "valor_ipi": d["valor_ipi"].quantize(Decimal("0.01")),
                        "base_difal_calculada": d["base_difal_calculada"].quantize(Decimal("0.01")),
                        "icms_origem_deduzido": d["icms_origem_deduzido"].quantize(Decimal("0.01")),
                        "icms_st_destacado": d["icms_st_destacado"].quantize(Decimal("0.01")),
                        "aliquota_interestadual": (d["aliquota_interestadual"] * 100).quantize(Decimal("0.01")),
                        "aliquota_interna_destino": (d["aliquota_interna_destino"] * 100).quantize(Decimal("0.01")),
                        "difal_calculado": d["difal_calculado"].quantize(Decimal("0.01")),
                        "tipo_base_difal": d["tipo_base_difal"]
                    } for d in r_entrada["detalhes_itens"]
                ]
            }
            
            return {
                "regime": RegimeTributario.SIMPLES_NACIONAL.value,
                "chave_acesso": documento.chave_acesso,
                "valor_original": documento.valor_total,
                "valor_final_base": documento.valor_final,
                "valor_com_st": Decimal("0.00"),
                "valor_sem_st": Decimal("0.00"),
                "aliquota_aplicada": Decimal("0.00"),
                "imposto_calculado": r_entrada["total_difal"] + r_entrada["total_icms_st"],
                "detalhes": {
                    "difal": r_entrada["total_difal"].quantize(Decimal("0.01")),
                    "icms_st_compra": r_entrada["total_icms_st"].quantize(Decimal("0.01"))
                },
                "mensagem": f"Nota Fiscal de Entrada (Compra). Isenta de faturamento/impostos de saída. "
                           f"Calculado DIFAL: R$ {r_entrada['total_difal']:,.2f} | ICMS-ST Destacado: R$ {r_entrada['total_icms_st']:,.2f}.",
                "memoria_calculo": memoria_calculo
            }

        empresa = documento.empresa
        rbt12 = getattr(empresa, "rbt12", Decimal("0.00"))
        folha12 = getattr(empresa, "folha12", Decimal("0.00"))
        sujeito_fator_r = getattr(empresa, "sujeito_fator_r", False)
        categoria_simples = getattr(empresa, "categoria_simples", "Serviços (Anexo III)")
        
        # O cálculo baseia-se nos itens do documento
        itens = documento.itens
        if not itens:
            itens = [{
                "sequencia": 1,
                "descricao": "Nota Fiscal (Sem itens discriminados no XML)",
                "cfop": "0000",
                "valor_total": float(documento.valor_total),
                "desconto": 0.0,
                "frete": 0.0,
                "valor_ipi": 0.0,
                "impostos": {}
            }]
            
        base_calculo = Decimal("0.00")
        total_imposto = Decimal("0.00")
        valor_com_st = Decimal("0.00")
        valor_sem_st = Decimal("0.00")
        valor_com_iss_retido = Decimal("0.00")
        valor_sem_iss_retido = Decimal("0.00")
        
        itens_calculados = []
        
        for item in itens:
            cfop = str(item.get("cfop", ""))
            
            # 1. Determinação do Anexo/Tabela do Item
            if (cfop in ("5933", "6933")) or documento.tipo_documento == "NFS-e":
                if categoria_simples.startswith("Serviços"):
                    anexo_item = categoria_simples
                else:
                    anexo_item = "Serviços (Anexo III)"
            elif cfop and cfop.startswith(("59", "69", "79", "52", "62", "72")):
                anexo_item = "Excluído"
            elif cfop.startswith(("5101", "5103", "5104", "5401", "5403", "6101", "6103", "6104", "6401", "6403")):
                anexo_item = "Indústria (Anexo II)"
            elif cfop.startswith(("5102", "5115", "5405", "6102", "6115", "6405")):
                anexo_item = "Comércio (Anexo I)"
            else:
                anexo_item = categoria_simples
                
            # Valores do item
            it_val = Decimal(str(item.get("valor_total", 0.0)))
            it_desc = Decimal(str(item.get("desconto", 0.0)))
            it_frete = Decimal(str(item.get("frete", 0.0)))
            it_ipi = Decimal(str(item.get("valor_ipi", 0.0)))
            v_liq = it_val - it_desc + it_frete + it_ipi
            
            if anexo_item == "Excluído":
                itens_calculados.append({
                    "sequencia": item.get("sequencia", 1),
                    "descricao": item.get("descricao", "Item"),
                    "cfop": cfop,
                    "valor_total": it_val,
                    "valor_liquido": v_liq,
                    "anexo_aplicado": "Excluído",
                    "aliquota_efetiva": Decimal("0.00"),
                    "imposto_calculado": Decimal("0.00"),
                    "st_aplicado": False,
                    "iss_retido_aplicado": False,
                    "detalhe_calculo": "Isento / Não tributável"
                })
                continue
                
            base_calculo += v_liq
            
            # Alíquota Efetiva do Item
            if aliquota is not None:
                aliquota_item = aliquota
            else:
                anexo_para_calculo = anexo_item
                if "Anexo V" in anexo_para_calculo:
                    # Força Fator R para o cálculo correto no Anexo V
                    aliquota_item = obter_aliquota_efetiva_sn(rbt12, folha12, True, anexo_para_calculo)
                else:
                    aliquota_item = obter_aliquota_efetiva_sn(rbt12, folha12, sujeito_fator_r, anexo_para_calculo)
                    
            # Segregações
            icms_share = Decimal("0.00")
            iss_share = Decimal("0.00")
            tem_st = False
            tem_iss_retido = False
            
            impostos_item = item.get("impostos", {})
            
            if "Comércio" in anexo_item:
                if rbt12 <= Decimal("180000.00"):
                    icms_share = Decimal("0.34")
                elif rbt12 <= Decimal("360000.00"):
                    icms_share = Decimal("0.34")
                elif rbt12 <= Decimal("720000.00"):
                    icms_share = Decimal("0.335")
                elif rbt12 <= Decimal("1800000.00"):
                    icms_share = Decimal("0.335")
                elif rbt12 <= Decimal("3600000.00"):
                    icms_share = Decimal("0.335")
                else:
                    icms_share = Decimal("0.335")
                    
                icms_item = impostos_item.get("icms", {})
                if icms_item.get("substituicao_tributaria") or icms_item.get("cst") in ("10", "30", "60", "70", "90", "201", "202", "203", "500", "900"):
                    tem_st = True
                    
            elif "Indústria" in anexo_item:
                if rbt12 <= Decimal("180000.00"):
                    icms_share = Decimal("0.3200")
                elif rbt12 <= Decimal("360000.00"):
                    icms_share = Decimal("0.3200")
                elif rbt12 <= Decimal("720000.00"):
                    icms_share = Decimal("0.3250")
                elif rbt12 <= Decimal("1800000.00"):
                    icms_share = Decimal("0.3250")
                elif rbt12 <= Decimal("3600000.00"):
                    icms_share = Decimal("0.3250")
                else:
                    icms_share = Decimal("0.3300")
                    
                icms_item = impostos_item.get("icms", {})
                if icms_item.get("substituicao_tributaria") or icms_item.get("cst") in ("10", "30", "60", "70", "90", "201", "202", "203", "500", "900"):
                    tem_st = True
                    
            elif "Serviços" in anexo_item or "Anexo III" in anexo_item or "Anexo IV" in anexo_item or "Anexo V" in anexo_item:
                if "Anexo IV" in anexo_item:
                    if rbt12 <= Decimal("180000.00"):
                        iss_share = Decimal("0.4450")
                    else:
                        iss_share = Decimal("0.4000")
                elif "Anexo V" in anexo_item:
                    if rbt12 <= Decimal("180000.00"):
                        iss_share = Decimal("0.1400")
                    elif rbt12 <= Decimal("360000.00"):
                        iss_share = Decimal("0.1700")
                    elif rbt12 <= Decimal("720000.00"):
                        iss_share = Decimal("0.1835")
                    elif rbt12 <= Decimal("1800000.00"):
                        iss_share = Decimal("0.1835")
                    elif rbt12 <= Decimal("3600000.00"):
                        iss_share = Decimal("0.1885")
                    else:
                        iss_share = Decimal("0.2333")
                else:
                    # Anexo III
                    if rbt12 <= Decimal("180000.00"):
                        iss_share = Decimal("0.3350")
                    elif rbt12 <= Decimal("360000.00"):
                        iss_share = Decimal("0.3200")
                    elif rbt12 <= Decimal("720000.00"):
                        iss_share = Decimal("0.3250")
                    elif rbt12 <= Decimal("1800000.00"):
                        iss_share = Decimal("0.3250")
                    elif rbt12 <= Decimal("360000.00"):
                        iss_share = Decimal("0.3350")
                    else:
                        iss_share = Decimal("0.00")
                        
                iss_item = impostos_item.get("iss", {})
                if iss_item.get("retido"):
                    tem_iss_retido = True
            
            # Aplica dedução de ST ou ISS Retido
            if tem_st:
                tax_rate_item = aliquota_item * (Decimal("1.0") - icms_share)
                detalhe = f"R$ {v_liq:,.2f} * {aliquota_item*100:.4f}% * (1 - {icms_share*100:.2f}% ST)"
                valor_com_st += v_liq
            elif tem_iss_retido:
                tax_rate_item = aliquota_item * (Decimal("1.0") - iss_share)
                detalhe = f"R$ {v_liq:,.2f} * {aliquota_item*100:.4f}% * (1 - {iss_share*100:.2f}% ISS Ret)"
                valor_com_iss_retido += v_liq
            else:
                tax_rate_item = aliquota_item
                detalhe = f"R$ {v_liq:,.2f} * {aliquota_item*100:.4f}%"
                if "Comércio" in anexo_item or "Indústria" in anexo_item:
                    valor_sem_st += v_liq
                else:
                    valor_sem_iss_retido += v_liq
                    
            imposto_item = (v_liq * tax_rate_item).quantize(Decimal("0.01"))
            total_imposto += imposto_item
            
            itens_calculados.append({
                "sequencia": item.get("sequencia", 1),
                "descricao": item.get("descricao", "Item"),
                "cfop": cfop,
                "valor_total": it_val,
                "valor_liquido": v_liq,
                "anexo_aplicado": anexo_item,
                "aliquota_efetiva": aliquota_item,
                "imposto_calculado": imposto_item,
                "st_aplicado": tem_st,
                "iss_retido_aplicado": tem_iss_retido,
                "detalhe_calculo": detalhe
            })
            
        # Adiciona ajustes manuais registrados na Staging Area
        ajustes_sum = documento.valor_final - documento.valor_total
        imposto_ajuste = Decimal("0.00")
        if ajustes_sum != 0:
            base_calculo += ajustes_sum
            if base_calculo < 0:
                base_calculo = Decimal("0.00")
                
            # Calcula o imposto do ajuste usando a alíquota padrão da empresa
            if aliquota is not None:
                aliquota_padrao = aliquota
            else:
                aliquota_padrao = obter_aliquota_efetiva_sn(rbt12, folha12, sujeito_fator_r, categoria_simples)
            imposto_ajuste = (ajustes_sum * aliquota_padrao).quantize(Decimal("0.01"))
            total_imposto += imposto_ajuste
            
            # Adiciona o ajuste como um item virtual de ajuste
            itens_calculados.append({
                "sequencia": 999,
                "descricao": "Ajuste Manual Registrado na Staging Area",
                "cfop": "0000",
                "valor_total": ajustes_sum,
                "valor_liquido": ajustes_sum,
                "anexo_aplicado": "Ajuste",
                "aliquota_efetiva": aliquota_padrao,
                "imposto_calculado": imposto_ajuste,
                "st_aplicado": False,
                "iss_retido_aplicado": False,
                "detalhe_calculo": f"Ajuste R$ {ajustes_sum:,.2f} * {aliquota_padrao*100:.4f}%"
            })
            
            # Ajusta os campos de ST/ISS correspondentes
            if "Comércio" in categoria_simples or "Indústria" in categoria_simples:
                valor_sem_st += ajustes_sum
            else:
                valor_sem_iss_retido += ajustes_sum
                
        # Garante que as somas parciais de ST/ISS fiquem consistentes com a base final
        if valor_com_st > base_calculo:
            valor_com_st = base_calculo
        if valor_com_iss_retido > base_calculo:
            valor_com_iss_retido = base_calculo
            
        # Reconstrução da mensagem original e determinação da aliquota_aplicada como a alíquota base da empresa
        if aliquota is not None:
            aliquota_padrao = aliquota
            mensagem = "Apuração DAS realizada com alíquota customizada informada manualmente."
        else:
            aliquota_padrao = obter_aliquota_efetiva_sn(rbt12, folha12, sujeito_fator_r, categoria_simples)
            if categoria_simples == "Comércio (Anexo I)":
                mensagem = f"DAS calculado pelo Anexo I (Comércio). RBT12 acumulado: R$ {rbt12:,.2f}."
            elif categoria_simples == "Indústria (Anexo II)":
                mensagem = f"DAS calculado pelo Anexo II (Indústria). RBT12 acumulado: R$ {rbt12:,.2f}."
            else:
                fator_r = (folha12 / rbt12) if rbt12 > 0 else Decimal("0.00")
                if "Anexo IV" in categoria_simples:
                    mensagem = f"DAS calculado pelo Anexo IV. RBT12 acumulado: R$ {rbt12:,.2f}."
                else:
                    anexo = "Anexo V (Fator R < 28%)" if (sujeito_fator_r or "Anexo V" in categoria_simples) and (fator_r < Decimal("0.28")) else "Anexo III"
                    if sujeito_fator_r or "Anexo V" in categoria_simples:
                        mensagem = f"DAS calculado pelo {anexo}. Fator R: {(fator_r * 100):,.2f}% (Folha R$ {folha12:,.2f} / RBT12 R$ {rbt12:,.2f})."
                    else:
                        mensagem = f"DAS calculado pelo Anexo III. RBT12 acumulado: R$ {rbt12:,.2f}."

        if documento.tipo_documento == "NFS-e":
            mensagem = f"[NFS-e] " + mensagem

        # Determina os shares para a mensagem consolidada e a memória de cálculo
        icms_share_msg = Decimal("0.00")
        if categoria_simples == "Comércio (Anexo I)":
            if rbt12 <= Decimal("180000.00"):
                icms_share_msg = Decimal("0.34")
            elif rbt12 <= Decimal("360000.00"):
                icms_share_msg = Decimal("0.34")
            elif rbt12 <= Decimal("720000.00"):
                icms_share_msg = Decimal("0.335")
            else:
                icms_share_msg = Decimal("0.335")
        elif categoria_simples == "Indústria (Anexo II)":
            if rbt12 <= Decimal("180000.00"):
                icms_share_msg = Decimal("0.3200")
            elif rbt12 <= Decimal("360000.00"):
                icms_share_msg = Decimal("0.3200")
            elif rbt12 <= Decimal("720000.00"):
                icms_share_msg = Decimal("0.3250")
            elif rbt12 <= Decimal("1800000.00"):
                icms_share_msg = Decimal("0.3250")
            elif rbt12 <= Decimal("3600000.00"):
                icms_share_msg = Decimal("0.3250")
            else:
                icms_share_msg = Decimal("0.3300")

        iss_share_msg = Decimal("0.00")
        if "Anexo IV" in categoria_simples:
            if rbt12 <= Decimal("180000.00"):
                iss_share_msg = Decimal("0.4450")
            else:
                iss_share_msg = Decimal("0.4000")
        elif "Anexo V" in categoria_simples:
            if rbt12 <= Decimal("180000.00"):
                iss_share_msg = Decimal("0.1400")
            elif rbt12 <= Decimal("360000.00"):
                iss_share_msg = Decimal("0.1700")
            elif rbt12 <= Decimal("720000.00"):
                iss_share_msg = Decimal("0.1835")
            elif rbt12 <= Decimal("1800000.00"):
                iss_share_msg = Decimal("0.1835")
            elif rbt12 <= Decimal("3600000.00"):
                iss_share_msg = Decimal("0.1885")
            else:
                iss_share_msg = Decimal("0.2333")
        elif "Serviços" in categoria_simples or "Anexo III" in categoria_simples:
            if rbt12 <= Decimal("180000.00"):
                iss_share_msg = Decimal("0.3350")
            elif rbt12 <= Decimal("360000.00"):
                iss_share_msg = Decimal("0.3200")
            elif rbt12 <= Decimal("720000.00"):
                iss_share_msg = Decimal("0.3250")
            elif rbt12 <= Decimal("1800000.00"):
                iss_share_msg = Decimal("0.3250")
            elif rbt12 <= Decimal("3600000.00"):
                iss_share_msg = Decimal("0.3350")
            else:
                iss_share_msg = Decimal("0.00")

        ipi_share_msg = Decimal("0.0750") if categoria_simples == "Indústria (Anexo II)" else Decimal("0.00")

        if valor_com_st > 0:
            if icms_share_msg > 0:
                economia = valor_com_st * aliquota_padrao * icms_share_msg
                mensagem += f" [ST SEGREGADO] R$ {valor_com_st:,.2f} segregados no Simples Nacional com dedução de {(icms_share_msg * 100):.2f}% de ICMS (Economia de R$ {economia:,.2f}!)."
            else:
                mensagem += f" [ST DETECTADO] R$ {valor_com_st:,.2f} em itens com Substituição Tributária de ICMS. Esse valor pode ser segregado no PGDAS-D para dedução e economia fiscal do ICMS!"

        if valor_com_iss_retido > 0:
            if iss_share_msg > 0:
                economia_iss = valor_com_iss_retido * aliquota_padrao * iss_share_msg
                mensagem += f" [ISS RETIDO] R$ {valor_com_iss_retido:,.2f} com ISS retido na fonte. Dedução de {(iss_share_msg * 100):.2f}% de ISS da guia DAS (Desconto de R$ {economia_iss:,.2f} no imposto unificado)."

        aliquota_aplicada = aliquota_padrao
        
        # Calcula a memória de cálculo detalhada consolidada para compatibilidade com o relatório existente
        if aliquota is not None:
            aliq_nom = aliquota
            deducao = Decimal("0.00")
            fator_r = Decimal("0.00")
            enquadramento = "Customizado"
        else:
            detalhes_calc = obter_aliquota_efetiva_sn(rbt12, folha12, sujeito_fator_r, categoria_simples, retornar_detalhes=True)
            aliq_nom = detalhes_calc["aliq_nom"]
            deducao = detalhes_calc["deducao"]
            fator_r = detalhes_calc["fator_r"]
            enquadramento = detalhes_calc["enquadramento"]

        memoria_calculo = {
            "rbt12": rbt12.quantize(Decimal("0.01")),
            "folha12": folha12.quantize(Decimal("0.01")),
            "fator_r": (fator_r * Decimal("100.0")).quantize(Decimal("0.01")),
            "sujeito_fator_r": sujeito_fator_r,
            "categoria_simples": categoria_simples,
            "enquadramento": enquadramento,
            "aliq_nom": (aliq_nom * Decimal("100.0")).quantize(Decimal("0.0001")),
            "deducao": deducao.quantize(Decimal("0.01")),
            "aliq_efetiva": (aliquota_aplicada * Decimal("100.0")).quantize(Decimal("0.0001")),
            "iss_share": (iss_share_msg * Decimal("100.0")).quantize(Decimal("0.01")),
            "icms_share": (icms_share_msg * Decimal("100.0")).quantize(Decimal("0.01")),
            "ipi_share": (ipi_share_msg * Decimal("100.0")).quantize(Decimal("0.01")) if isinstance(ipi_share_msg, Decimal) else ipi_share_msg,
            "valor_com_iss_retido": valor_com_iss_retido.quantize(Decimal("0.01")),
            "valor_com_st": valor_com_st.quantize(Decimal("0.01")),
            "valor_sem_iss_retido": valor_sem_iss_retido.quantize(Decimal("0.01")),
            "valor_sem_st": valor_sem_st.quantize(Decimal("0.01")),
            "itens_calculados": [
                {
                    **it,
                    "aliquota_efetiva": (it["aliquota_efetiva"] * Decimal("100.0")).quantize(Decimal("0.0001")),
                    "valor_total": it["valor_total"].quantize(Decimal("0.01")),
                    "valor_liquido": it["valor_liquido"].quantize(Decimal("0.01")),
                    "imposto_calculado": it["imposto_calculado"].quantize(Decimal("0.01"))
                } for it in itens_calculados
            ]
        }


        return {
            "regime": RegimeTributario.SIMPLES_NACIONAL.value,
            "chave_acesso": documento.chave_acesso,
            "valor_original": documento.valor_total,
            "valor_final_base": base_calculo,
            "valor_com_st": valor_com_st.quantize(Decimal("0.01")),
            "valor_sem_st": valor_sem_st.quantize(Decimal("0.01")),
            "aliquota_aplicada": aliquota_aplicada.quantize(Decimal("0.000001")),
            "imposto_calculado": total_imposto.quantize(Decimal("0.01")),
            "detalhes": {
                "das": total_imposto.quantize(Decimal("0.01"))
            },
            "mensagem": mensagem,
            "memoria_calculo": memoria_calculo
        }

class CalculadoraLucroPresumido(CalculadoraInterface):
    """
    Estratégia concreta para o regime Lucro Presumido (Serviços).
    Calcula a presunção federal brasileira de impostos federais:
    - PIS: 0,65%
    - COFINS: 3,00%
    - IRPJ: 4,80% (32% base de presunção * 15% alíquota IR)
    - CSLL: 2,88% (32% base de presunção * 9% alíquota CSLL)
    """
    def calcular(self, documento: DocumentoFiscal, aliquota: Optional[Decimal] = None) -> Dict[str, Any]:
        # Verifica se o documento está Cancelado ou Denegado
        cstat = getattr(documento, "cstat", "100")
        if cstat in ("101", "110", "301", "302"):
            situacao = "CANCELADA" if cstat == "101" else "DENEGADA"
            return {
                "regime": RegimeTributario.LUCRO_PRESUMIDO.value,
                "chave_acesso": documento.chave_acesso,
                "valor_original": documento.valor_total,
                "valor_final_base": Decimal("0.00"),
                "valor_com_st": Decimal("0.00"),
                "valor_sem_st": Decimal("0.00"),
                "aliquota_aplicada": Decimal("0.00"),
                "imposto_calculado": Decimal("0.00"),
                "detalhes": {
                    "pis": Decimal("0.00"),
                    "cofins": Decimal("0.00"),
                    "irpj": Decimal("0.00"),
                    "csll": Decimal("0.00"),
                    "iss": Decimal("0.00")
                },
                "mensagem": f"Nota Fiscal {situacao} (cStat {cstat}). Faturamento e impostos desconsiderados para fins tributários."
            }

        # Se for nota de Entrada, executa cálculo de impostos de compras
        if documento.tipo_operacao == "Entrada":
            uf_empresa = getattr(documento.empresa, "uf", "BA")
            r_entrada = calcular_impostos_entrada(documento, uf_empresa)
            
            memoria_calculo = {
                "uf_origem": r_entrada["uf_origem"],
                "uf_destino": r_entrada["uf_destino"],
                "is_interestadual": r_entrada["is_interestadual"],
                "total_difal": r_entrada["total_difal"].quantize(Decimal("0.01")),
                "total_icms_st": r_entrada["total_icms_st"].quantize(Decimal("0.01")),
                "detalhes_itens": [
                    {
                        **d,
                        "valor_total": d["valor_total"].quantize(Decimal("0.01")),
                        "desconto": d["desconto"].quantize(Decimal("0.01")),
                        "frete": d["frete"].quantize(Decimal("0.01")),
                        "valor_ipi": d["valor_ipi"].quantize(Decimal("0.01")),
                        "base_difal_calculada": d["base_difal_calculada"].quantize(Decimal("0.01")),
                        "icms_origem_deduzido": d["icms_origem_deduzido"].quantize(Decimal("0.01")),
                        "icms_st_destacado": d["icms_st_destacado"].quantize(Decimal("0.01")),
                        "aliquota_interestadual": (d["aliquota_interestadual"] * 100).quantize(Decimal("0.01")),
                        "aliquota_interna_destino": (d["aliquota_interna_destino"] * 100).quantize(Decimal("0.01")),
                        "difal_calculado": d["difal_calculado"].quantize(Decimal("0.01")),
                        "tipo_base_difal": d["tipo_base_difal"]
                    } for d in r_entrada["detalhes_itens"]
                ]
            }
            
            return {
                "regime": RegimeTributario.LUCRO_PRESUMIDO.value,
                "chave_acesso": documento.chave_acesso,
                "valor_original": documento.valor_total,
                "valor_final_base": documento.valor_final,
                "valor_com_st": Decimal("0.00"),
                "valor_sem_st": Decimal("0.00"),
                "aliquota_aplicada": Decimal("0.00"),
                "imposto_calculado": r_entrada["total_difal"] + r_entrada["total_icms_st"],
                "detalhes": {
                    "difal": r_entrada["total_difal"].quantize(Decimal("0.01")),
                    "icms_st_compra": r_entrada["total_icms_st"].quantize(Decimal("0.01"))
                },
                "mensagem": f"Nota Fiscal de Entrada (Compra). Isenta de faturamento/impostos de saída. "
                           f"Calculado DIFAL: R$ {r_entrada['total_difal']:,.2f} | ICMS-ST Destacado: R$ {r_entrada['total_icms_st']:,.2f}.",
                "memoria_calculo": memoria_calculo
            }

        # O cálculo baseia-se no valor final (filtrando itens de remessa/comodato/devolução)
        base_calculo = Decimal("0.00")
        if documento.itens:
            for item in documento.itens:
                cfop = str(item.get("cfop", ""))
                # Se for CFOP de remessa, retorno, comodato ou devolução de saída, desconsidera do faturamento
                if cfop and cfop.startswith(("59", "69", "79", "52", "62", "72")):
                    continue
                it_val = Decimal(str(item.get("valor_total", 0.0)))
                it_desc = Decimal(str(item.get("desconto", 0.0)))
                it_frete = Decimal(str(item.get("frete", 0.0)))
                it_ipi = Decimal(str(item.get("valor_ipi", 0.0)))
                base_calculo += (it_val - it_desc + it_frete + it_ipi)
            # Adiciona ajustes manuais registrados na Staging Area
            ajustes_sum = documento.valor_final - documento.valor_total
            base_calculo += ajustes_sum
            if base_calculo < 0:
                base_calculo = Decimal("0.00")
        else:
            base_calculo = documento.valor_final

        # Alíquotas federais de Lucro Presumido para prestação de serviços gerais
        aliquota_pis = Decimal("0.0065")
        aliquota_cofins = Decimal("0.0300")
        aliquota_irpj = Decimal("0.0480")
        aliquota_csll = Decimal("0.0288")

        # Cálculos individuais
        pis = base_calculo * aliquota_pis
        cofins = base_calculo * aliquota_cofins
        irpj = base_calculo * aliquota_irpj
        csll = base_calculo * aliquota_csll
        
        total_impostos = pis + cofins + irpj + csll

        # Identifica faturamento com Substituição Tributária (ICMS-ST) nos itens
        valor_com_st = Decimal("0.00")
        if documento.itens:
            for item in documento.itens:
                impostos = item.get("impostos", {})
                icms = impostos.get("icms", {})
                if icms.get("substituicao_tributaria") or icms.get("cst") in ("10", "30", "60", "70", "90", "201", "202", "203", "500", "900"):
                    # SUBTRACT THE UNCONDITIONAL DISCOUNT TO GET THE NET VALUE!
                    it_val = Decimal(str(item.get("valor_total", 0.0)))
                    it_desc = Decimal(str(item.get("desconto", 0.0)))
                    valor_com_st += (it_val - it_desc)
        
        if valor_com_st > base_calculo:
            valor_com_st = base_calculo
        valor_sem_st = base_calculo - valor_com_st

        mensagem = "Apuração federal detalhada de impostos federais no Lucro Presumido (Serviços)."
        if valor_com_st > 0:
            mensagem += f" Detectado faturamento de R$ {valor_com_st:,.2f} com Substituição Tributária (ICMS-ST retido anteriormente)."

        detalhes = {
            "pis": pis.quantize(Decimal("0.01")),
            "cofins": cofins.quantize(Decimal("0.01")),
            "irpj": irpj.quantize(Decimal("0.01")),
            "csll": csll.quantize(Decimal("0.01")),
            "iss": Decimal("0.00")
        }

        # Cálculo do ISSQN para NFS-e no Lucro Presumido
        if documento.tipo_documento == "NFS-e":
            iss_aliq_percent = Decimal("5.0")
            if documento.itens:
                first_item_iss = documento.itens[0].get("impostos", {}).get("iss", {})
                if first_item_iss:
                    iss_aliq_percent = Decimal(str(first_item_iss.get("aliquota", 5.0)))
            
            if aliquota is not None:
                iss_aliq_percent = aliquota * Decimal("100.0") if aliquota < Decimal("1.0") else aliquota
                
            iss_aliquota = iss_aliq_percent / Decimal("100.0")
            iss = base_calculo * iss_aliquota
            total_impostos += iss
            detalhes["iss"] = iss.quantize(Decimal("0.01"))
            mensagem += f" Incluindo ISS Municipal de {iss_aliq_percent:.2f}% (R$ {iss:,.2f})."
        else:
            iss_aliquota = Decimal("0.00")
            iss = Decimal("0.00")

        memoria_calculo = {
            "aliquota_pis": (aliquota_pis * Decimal("100.0")).quantize(Decimal("0.01")),
            "aliquota_cofins": (aliquota_cofins * Decimal("100.0")).quantize(Decimal("0.01")),
            "aliquota_irpj": (aliquota_irpj * Decimal("100.0")).quantize(Decimal("0.01")),
            "aliquota_csll": (aliquota_csll * Decimal("100.0")).quantize(Decimal("0.01")),
            "aliquota_iss": (iss_aliquota * Decimal("100.0")).quantize(Decimal("0.01")),
            "pis": pis.quantize(Decimal("0.01")),
            "cofins": cofins.quantize(Decimal("0.01")),
            "irpj": irpj.quantize(Decimal("0.01")),
            "csll": csll.quantize(Decimal("0.01")),
            "iss": iss.quantize(Decimal("0.01")),
            "valor_com_st": valor_com_st.quantize(Decimal("0.01")),
            "valor_sem_st": valor_sem_st.quantize(Decimal("0.01"))
        }

        return {
            "regime": RegimeTributario.LUCRO_PRESUMIDO.value,
            "chave_acesso": documento.chave_acesso,
            "valor_original": documento.valor_total,
            "valor_final_base": base_calculo,
            "valor_com_st": valor_com_st.quantize(Decimal("0.01")),
            "valor_sem_st": valor_sem_st.quantize(Decimal("0.01")),
            "aliquota_aplicada": (aliquota_pis + aliquota_cofins + aliquota_irpj + aliquota_csll + (iss_aliquota if documento.tipo_documento == "NFS-e" else Decimal("0.00"))).quantize(Decimal("0.000001")),
            "imposto_calculado": total_impostos.quantize(Decimal("0.01")),
            "detalhes": detalhes,
            "mensagem": mensagem,
            "memoria_calculo": memoria_calculo
        }

class CalculadoraFactory:
    """
    Fábrica responsável por instanciar a estratégia de cálculo correta 
    de acordo com o Regime Tributário cadastrado para a empresa do documento.
    """
    @staticmethod
    def obter_calculadora(regime: RegimeTributario) -> CalculadoraInterface:
        if regime == RegimeTributario.SIMPLES_NACIONAL:
            return CalculadoraSimplesNacional()
        elif regime == RegimeTributario.LUCRO_PRESUMIDO:
            return CalculadoraLucroPresumido()
        else:
            raise ValueError(f"Estratégia de cálculo não implementada para o regime: {regime}")

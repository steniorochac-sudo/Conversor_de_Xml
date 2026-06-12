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
    detalhes_itens = []
    
    # Bahia has 20.5% internal rate, others default to 18% unless specified
    aliq_interna_destino = Decimal("0.205") if uf_empresa == "BA" else Decimal("0.18")
    
    # States that require the Double Base (Base Dupla / "por dentro") calculation method
    ufs_base_dupla = {"BA", "MG", "PR", "RS", "AL", "GO", "DF", "SE", "TO", "RO"}
    
    if documento.itens:
        for item in documento.itens:
            impostos = item.get("impostos", {})
            icms = impostos.get("icms", {})
            v_st = Decimal(str(icms.get("valor_st", 0.0)))
            total_icms_st += v_st
            
            difal_item = Decimal("0.00")
            aliq_inter = Decimal("0.00")
            v_ipi = Decimal(str(item.get("valor_ipi", 0.0)))
            
            v_prod = Decimal(str(item.get("valor_total", 0.0)))
            v_desc = Decimal(str(item.get("desconto", 0.0)))
            v_frete = Decimal(str(item.get("frete", 0.0)))
            v_liq = v_prod - v_desc + v_frete + v_ipi
            
            icms_origem = Decimal(str(icms.get("valor", 0.0)))
            base_difal = v_liq
            tipo_base_difal = "Simples"
            
            if is_interestadual and documento.tipo_documento == "NF-e":
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
        
        # O cálculo baseia-se no valor final (XML original + ajustes de staging)
        base_calculo = documento.valor_final
        categoria_simples = getattr(empresa, "categoria_simples", "Serviços (Anexo III)")
        
        # Determina a alíquota a aplicar
        if aliquota is not None:
            aliquota_aplicada = aliquota
            mensagem = "Apuração DAS realizada com alíquota customizada informada manualmente."
        else:
            rbt12 = getattr(empresa, "rbt12", Decimal("0.00"))
            folha12 = getattr(empresa, "folha12", Decimal("0.00"))
            sujeito_fator_r = getattr(empresa, "sujeito_fator_r", False)
            if "Anexo V" in categoria_simples:
                sujeito_fator_r = True
            
            aliquota_aplicada = obter_aliquota_efetiva_sn(rbt12, folha12, sujeito_fator_r, categoria_simples)
            
            # Monta mensagem descritiva amigável
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

        # Fração de ICMS no Simples Nacional para segregação de ST
        icms_share = Decimal("0.00")
        if categoria_simples == "Comércio (Anexo I)":
            rbt12 = getattr(empresa, "rbt12", Decimal("0.00"))
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
        elif categoria_simples == "Indústria (Anexo II)":
            rbt12 = getattr(empresa, "rbt12", Decimal("0.00"))
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

        # Fração de IPI no Anexo II (Indústria)
        ipi_share = Decimal("0.0750") if categoria_simples == "Indústria (Anexo II)" else Decimal("0.00")

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
        
        # Fração de ISS no Simples Nacional para segregação de retenção de ISS
        iss_share = Decimal("0.00")
        rbt12 = getattr(empresa, "rbt12", Decimal("0.00"))
        if "Anexo IV" in categoria_simples:
            if rbt12 <= Decimal("180000.00"):
                iss_share = Decimal("0.4450")
            else:
                iss_share = Decimal("0.4000")
        elif "Anexo V" in categoria_simples:
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
        elif "Serviços" in categoria_simples:
            if rbt12 <= Decimal("180000.00"):
                iss_share = Decimal("0.3350")
            elif rbt12 <= Decimal("360000.00"):
                iss_share = Decimal("0.3200")
            elif rbt12 <= Decimal("720000.00"):
                iss_share = Decimal("0.3250")
            elif rbt12 <= Decimal("1800000.00"):
                iss_share = Decimal("0.3250")
            elif rbt12 <= Decimal("3600000.00"):
                iss_share = Decimal("0.3350")
            else:
                iss_share = Decimal("0.00")

        # Identifica faturamento com ISS Retido nos itens
        valor_com_iss_retido = Decimal("0.00")
        if documento.itens:
            for item in documento.itens:
                impostos = item.get("impostos", {})
                iss = impostos.get("iss", {})
                if iss.get("retido"):
                    it_val = Decimal(str(item.get("valor_total", 0.0)))
                    it_desc = Decimal(str(item.get("desconto", 0.0)))
                    valor_com_iss_retido += (it_val - it_desc)

        if valor_com_iss_retido > base_calculo:
            valor_com_iss_retido = base_calculo
        valor_sem_iss_retido = base_calculo - valor_com_iss_retido

        # Cálculo segregado se houver ICMS-ST
        imposto_sem_st = valor_sem_st * aliquota_aplicada
        imposto_com_st = valor_com_st * aliquota_aplicada * (Decimal("1.0") - icms_share)
        imposto_das_produtos = imposto_sem_st + imposto_com_st

        # Cálculo segregado se houver ISS Retido
        imposto_sem_iss = valor_sem_iss_retido * aliquota_aplicada
        imposto_com_iss = valor_com_iss_retido * aliquota_aplicada * (Decimal("1.0") - iss_share)
        imposto_das_servicos = imposto_sem_iss + imposto_com_iss

        # Escolhe o imposto apropriado conforme categoria
        if categoria_simples in ("Comércio (Anexo I)", "Indústria (Anexo II)"):
            imposto_das = imposto_das_produtos
        else:
            imposto_das = imposto_das_servicos

        if valor_com_st > 0:
            if icms_share > 0:
                economia = valor_com_st * aliquota_aplicada * icms_share
                mensagem += f" [ST SEGREGADO] R$ {valor_com_st:,.2f} segregados no Simples Nacional com dedução de {(icms_share * 100):.2f}% de ICMS (Economia de R$ {economia:,.2f}!)."
            else:
                mensagem += f" [ST DETECTADO] R$ {valor_com_st:,.2f} em itens com Substituição Tributária de ICMS. Esse valor pode ser segregado no PGDAS-D para dedução e economia fiscal do ICMS!"

        if valor_com_iss_retido > 0:
            if iss_share > 0:
                economia_iss = valor_com_iss_retido * aliquota_aplicada * iss_share
                mensagem += f" [ISS RETIDO] R$ {valor_com_iss_retido:,.2f} com ISS retido na fonte. Dedução de {(iss_share * 100):.2f}% de ISS da guia DAS (Desconto de R$ {economia_iss:,.2f} no imposto unificado)."

        # Calcula a memória de cálculo detalhada
        if aliquota is not None:
            aliq_nom = aliquota_aplicada
            deducao = Decimal("0.00")
            fator_r = Decimal("0.00")
            enquadramento = "Customizado"
            rbt12_val = Decimal("0.00")
            folha12_val = Decimal("0.00")
            sujeito_fator_r_val = False
        else:
            rbt12_val = getattr(empresa, "rbt12", Decimal("0.00"))
            folha12_val = getattr(empresa, "folha12", Decimal("0.00"))
            sujeito_fator_r_val = getattr(empresa, "sujeito_fator_r", False)
            if "Anexo V" in categoria_simples:
                sujeito_fator_r_val = True
            detalhes_calc = obter_aliquota_efetiva_sn(rbt12_val, folha12_val, sujeito_fator_r_val, categoria_simples, retornar_detalhes=True)
            aliq_nom = detalhes_calc["aliq_nom"]
            deducao = detalhes_calc["deducao"]
            fator_r = detalhes_calc["fator_r"]
            enquadramento = detalhes_calc["enquadramento"]

        memoria_calculo = {
            "rbt12": rbt12_val.quantize(Decimal("0.01")),
            "folha12": folha12_val.quantize(Decimal("0.01")),
            "fator_r": (fator_r * Decimal("100.0")).quantize(Decimal("0.01")),
            "sujeito_fator_r": sujeito_fator_r_val,
            "categoria_simples": categoria_simples,
            "enquadramento": enquadramento,
            "aliq_nom": (aliq_nom * Decimal("100.0")).quantize(Decimal("0.0001")),
            "deducao": deducao.quantize(Decimal("0.01")),
            "aliq_efetiva": (aliquota_aplicada * Decimal("100.0")).quantize(Decimal("0.0001")),
            "iss_share": (iss_share * Decimal("100.0")).quantize(Decimal("0.01")),
            "icms_share": (icms_share * Decimal("100.0")).quantize(Decimal("0.01")),
            "ipi_share": (ipi_share * Decimal("100.0")).quantize(Decimal("0.01")),
            "valor_com_iss_retido": valor_com_iss_retido.quantize(Decimal("0.01")),
            "valor_com_st": valor_com_st.quantize(Decimal("0.01")),
            "valor_sem_iss_retido": valor_sem_iss_retido.quantize(Decimal("0.01")),
            "valor_sem_st": valor_sem_st.quantize(Decimal("0.01"))
        }


        return {
            "regime": RegimeTributario.SIMPLES_NACIONAL.value,
            "chave_acesso": documento.chave_acesso,
            "valor_original": documento.valor_total,
            "valor_final_base": base_calculo,
            "valor_com_st": valor_com_st.quantize(Decimal("0.01")),
            "valor_sem_st": valor_sem_st.quantize(Decimal("0.01")),
            "aliquota_aplicada": aliquota_aplicada.quantize(Decimal("0.000001")),
            "imposto_calculado": imposto_das.quantize(Decimal("0.01")),
            "detalhes": {
                "das": imposto_das.quantize(Decimal("0.01"))
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

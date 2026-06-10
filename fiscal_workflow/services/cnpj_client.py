import json
import urllib.request
from typing import Optional

def buscar_cnae_oficial(cnpj: str) -> Optional[str]:
    """
    Consulta a API pública cnpj.ws (e como fallback minhareceita.org) para recuperar 
    o código limpo (apenas dígitos) do CNAE principal registrado para o CNPJ fornecido.
    """
    cnpj_limpo = "".join(filter(str.isdigit, cnpj))
    if len(cnpj_limpo) != 14:
        return None
        
    # Tentativa 1: cnpj.ws
    url_ws = f"https://publica.cnpj.ws/cnpj/{cnpj_limpo}"
    try:
        req = urllib.request.Request(
            url_ws, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req, timeout=5) as response:
            dados = json.loads(response.read().decode("utf-8"))
            cnae_principal = dados.get("estabelecimento", {}).get("atividade_principal", {}).get("code", "")
            if cnae_principal:
                return "".join(filter(str.isdigit, str(cnae_principal)))
    except Exception as e:
        print(f"Aviso: Falha na API cnpj.ws para CNPJ {cnpj_limpo}. Motivo: {e}. Tentando fallback...")

    # Tentativa 2: minhareceita.org (Fallback público sem rate-limit rígido)
    url_fallback = f"https://minhareceita.org/{cnpj_limpo}"
    try:
        req = urllib.request.Request(
            url_fallback, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req, timeout=5) as response:
            dados = json.loads(response.read().decode("utf-8"))
            cnae_principal = dados.get("cnae_fiscal", "")
            if cnae_principal:
                return "".join(filter(str.isdigit, str(cnae_principal)))
    except Exception as e:
        print(f"Aviso: Falha no fallback da API MinhaReceita para CNPJ {cnpj_limpo}. Motivo: {e}")
        
    return None

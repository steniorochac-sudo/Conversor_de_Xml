import os
import json
import urllib.request
from pathlib import Path

CNAE_JSON_PATH = Path(__file__).resolve().parent.parent / "core" / "cnae_data.json"

def sync_cnaes_from_ibge() -> int:
    """
    Sincroniza os CNAEs do banco de dados local com as subclasses oficiais do IBGE.
    Preserva as regras fiscais existentes e aplica regras padrão baseadas na divisão do CNAE:
    - CNAEs iniciados em 45, 46, 47: Enquadrados no Anexo I (Comércio, alíquota inicial 4.0%)
    - Outros CNAEs: Enquadrados no Anexo III (Serviços, alíquota inicial 6.0%)
    """
    # 1. Carrega os dados locais existentes (com regras fiscais curadas)
    local_rules = {}
    if CNAE_JSON_PATH.exists():
        try:
            with open(CNAE_JSON_PATH, "r", encoding="utf-8") as f:
                local_rules = json.load(f)
        except Exception:
            pass

    # 2. Requisita subclasses oficiais da API do IBGE
    url = "https://servicodados.ibge.gov.br/api/v2/cnae/subclasses"
    try:
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req, timeout=15) as response:
            subclasses = json.loads(response.read().decode("utf-8"))
    except Exception as e:
        raise RuntimeError(f"Falha ao conectar com o serviço do IBGE: {str(e)}")

    if not isinstance(subclasses, list):
        raise ValueError("Resposta da API do IBGE em formato inesperado.")

    # 3. Mescla os dados do IBGE com as regras tributárias locais
    new_cnae_data = {}
    for sub in subclasses:
        cnae_code = sub.get("id", "").replace("/", "").replace("-", "").strip()
        descricao = sub.get("descricao", "").strip()
        
        if not cnae_code or not descricao:
            continue

        # Se já existe nas regras locais curadas, preserva todos os dados
        if cnae_code in local_rules:
            new_cnae_data[cnae_code] = local_rules[cnae_code]
            # Atualiza apenas a descrição caso tenha mudado
            new_cnae_data[cnae_code]["descricao"] = descricao
        else:
            # Caso contrário, aplica heurísticas fiscais padrão inteligentes
            first_two = cnae_code[:2]
            try:
                div_num = int(first_two)
            except ValueError:
                div_num = 0

            if first_two in ("45", "46", "47"):
                # Comércio
                new_cnae_data[cnae_code] = {
                    "descricao": descricao,
                    "anexo": "I",
                    "fator_r": False,
                    "aliquota": 4.0
                }
            elif 10 <= div_num <= 33:
                # Indústria (Anexo II)
                new_cnae_data[cnae_code] = {
                    "descricao": descricao,
                    "anexo": "II",
                    "fator_r": False,
                    "aliquota": 4.5
                }
            else:
                # Serviços (Default Geral Anexo III)
                # Verifica se é TI, Medicina ou Psicologia que costuma cair em Fator R
                sujeito_fator_r = False
                aliquota_inicial = 6.0
                anexo = "III"
                if first_two in ("62", "86", "73", "74", "69"):
                    sujeito_fator_r = True
                    anexo = "V"
                
                new_cnae_data[cnae_code] = {
                    "descricao": descricao,
                    "anexo": anexo,
                    "fator_r": sujeito_fator_r,
                    "aliquota": aliquota_inicial
                }

    # 4. Salva a base unificada de volta no JSON
    with open(CNAE_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(new_cnae_data, f, ensure_ascii=False, indent=2)

    return len(new_cnae_data)

if __name__ == "__main__":
    print("Iniciando sincronização de CNAEs...")
    try:
        total = sync_cnaes_from_ibge()
        print(f"Sincronização concluída! {total} CNAEs mapeados e salvos em cnae_data.json.")
    except Exception as err:
        print("Erro durante a sincronização:", err)

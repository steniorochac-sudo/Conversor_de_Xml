import pdfplumber
import pandas as pd
import re
import os
import requests
import pyodbc # Garanta que o pyodbc também está importado

URL_PDF = "https://mbusca.sefaz.ba.gov.br/DITRI/normas_complementares/decretos/decreto_2012_13780_ricms_anexo_1_vigente_2026.pdf"
CAMINHO_PDF_LOCAL = "decreto_icms_ba.pdf"
# --- ADICIONE ESTA LINHA ---
BANCO_ACCESS = r'C:\Users\StenioRochaCardoso\Cabral & Sousa Ltda\Intranet Cabral & Sousa - 🔎PÚBLICA\🏢 PÚBLICA - FISCAL\01 - STÊNIO ROCHA\Apuração de Impostos - Entrada\Base_NF_ENTRADA.accdb'

def limpar_valor(texto):
    if not texto: return 0.0
    # Pega números com vírgula/ponto antes do símbolo %
    match = re.search(r'([\d,.]+)\s*%', str(texto))
    if match:
        return float(match.group(1).replace(',', '.'))
    return 0.0

def extrair_mvas_inteligente(texto_completo):
    """Extração robusta que ignora erros de digitação e parênteses faltantes no PDF."""
    # Regex flexível: Procura o número + % + (Alíq + Valor alvo + % opcional + parênteses opcional)
    # O [^)]* permite qualquer caractere que não seja um parênteses de fechamento no meio
    mva4 = re.search(r'([\d,.]+)%\s*\(Al[íi]q\.?\s*4%?[^)]*\)?', texto_completo, re.I)
    mva7 = re.search(r'([\d,.]+)%\s*\(Al[íi]q\.?\s*7%?[^)]*\)?', texto_completo, re.I)
    mva12 = re.search(r'([\d,.]+)%\s*\(Al[íi]q\.?\s*12%?[^)]*\)?', texto_completo, re.I)
    
    val4 = limpar_valor(mva4.group(1) + "%") if mva4 else 0.0
    val7 = limpar_valor(mva7.group(1) + "%") if mva7 else 0.0
    val12 = limpar_valor(mva12.group(1) + "%") if mva12 else 0.0

    # Busca a MVA Original (o valor que sobra e que geralmente é > 10%)
    texto_para_original = re.sub(r'[\d,.]+%?\s*\(Al[íi]q\.?\s*\d+[^)]*\)?', '', texto_completo, flags=re.I)
    sobras = re.findall(r'([\d,.]+)%', texto_para_original)
    
    mva_original = 0.0
    if sobras:
        candidatos = [float(s.replace(',', '.')) for s in sobras]
        for c in reversed(candidatos):
            if c > 15.0: # Filtro para não pegar alíquotas internas menores
                mva_original = c
                break
            
    return mva_original, val4, val7, val12

def extrair_mva_ajustada(texto, aliquota_alvo):
    """Busca MVA Ajustada pelo padrão 'Valor% (Aliq. X%)'."""
    if not texto: return 0.0
    # O padrão aceita 'Aliq.', 'Aliq' (sem ponto) e variações de acento
    padrao = rf"([\d,.]+)%\s*\(Al[íi]q\.?\s*{aliquota_alvo}%\)"
    match = re.search(padrao, str(texto), re.IGNORECASE)
    return limpar_valor(match.group(1) + "%") if match else 0.0

def expandir_ncm(ncm_texto):
    ncm = re.sub(r'[.\s]', '', str(ncm_texto))
    if not ncm or not ncm.isdigit() or len(ncm) < 2: return "", ""
    return ncm.ljust(8, '0'), ncm.ljust(8, '9')

def executar_extracao(pdf_path):
    dados_finais = []
    print("Processando PDF de forma inteligente...")
    
    with pdfplumber.open(pdf_path) as pdf:
        item_atual = None
        
        for page in pdf.pages:
            tabela = page.extract_table()
            if not tabela: continue
            
            for row in tabela:
                # Limpa as células da linha
                row_limpa = [str(c).strip() if c else "" for c in row]
                
                # Pula linhas vazias ou cabeçalhos
                if not any(row_limpa) or "ITEM" in row_limpa[0].upper():
                    continue
                
                # Identifica se a linha é um novo ITEM (ex: 1.1, 3.5.2, 10)
                # Melhoria na Regex para ser mais precisa com números de itens
                item_match = re.match(r'^\d+(\.\d+)*$', row_limpa[0])
                
                if item_match:
                    # SALVAMENTO DO ITEM ANTERIOR
                    if item_atual:
                        # IMPORTANTE: Use .extend() para "achatar" a lista de registros
                        dados_finais.extend(processar_item_buffer(item_atual))
                    
                    # INÍCIO DO NOVO ITEM
                    item_atual = {
                        "ITEM": row_limpa[0],
                        "MVA_ORIGINAL_TEXTO": row_limpa[-1], # Pega o valor da última coluna
                        "TEXTO_COMPLETO": " ".join(row_limpa[1:]) # Junta descrição, CEST e NCM
                    }
                else:
                    # É CONTINUAÇÃO do item anterior
                    if item_atual:
                        item_atual["TEXTO_COMPLETO"] += " " + " ".join(row_limpa)
                        # Se a última coluna desta linha de continuação tiver valor, pode ser a MVA Original
                        if row_limpa[-1] and "%" in row_limpa[-1]:
                             item_atual["MVA_ORIGINAL_TEXTO"] = row_limpa[-1]
        
        # Adiciona o último item do PDF
        if item_atual:
            dados_finais.extend(processar_item_buffer(item_atual))

    # Transforma a lista de dicionários em DataFrame
    df = pd.DataFrame(dados_finais)
    
    # Define a ordem das colunas para o Excel/WinThor
    colunas_ordenadas = [
        'ITEM', 'CEST', 'NCM', 'NCM_INICIAL', 'NCM_FINAL', 
        'MVA_ORIGINAL', 'MVA_AJUSTADA_4', 'MVA_AJUSTADA_7', 'MVA_AJUSTADA_12', 'DESCRIÇÃO'
    ]
    
    if not df.empty:
        # REMOVE DUPLICADOS: Garante que combinações idênticas de ITEM+CEST+NCM não se repitam
        df = df.drop_duplicates(subset=['ITEM', 'CEST', 'NCM', 'MVA_ORIGINAL'])
        # Garante que todas as colunas existam antes de reordenar
        for col in colunas_ordenadas:
            if col not in df.columns: df[col] = ""
        df = df[colunas_ordenadas]
    
    return df

def expandir_intervalo_cest(texto):
    """Detecta '01.001.00 a 01.001.05' e retorna uma lista com todos os códigos no intervalo."""
    padrao = r'(\d{2}\.\d{3}\.\d{2})\s+a\s+(\d{2}\.\d{3}\.\d{2})'
    match = re.search(padrao, texto)
    if not match:
        return re.findall(r'\d{2}\.\d{3}\.\d{2}', texto) # Retorna CESTs isolados se não houver intervalo
    
    inicio = match.group(1)
    fim = match.group(2)
    
    prefixo = inicio[:7] # Ex: '17.044.'
    num_inicio = int(inicio[7:])
    num_fim = int(fim[7:])
    
    lista_expandida = [f"{prefixo}{str(i).zfill(2)}" for i in range(num_inicio, num_fim + 1)]
    return lista_expandida

def extrair_todos_ncms(texto):
    """Extrai NCMs ignorando anos de convênios e números de decretos."""
    # Limpeza de caracteres que podem confundir o regex
    limpo = texto.replace(' e ', ' ').replace(' , ', ' ').replace(',', ' ')
    # Busca números com formato de NCM (4 a 8 dígitos)
    encontrados = re.findall(r'\b\d{4}(?:\.\d{1,4})*\b', limpo)
    
    # Lista de "Ruído": Anos comuns e números de decretos da Bahia
    anos = [str(a) for a in range(1990, 2031)]
    leis_e_decretos = ["13780", "13.780", "14629", "14.629", "7014", "7.014", "41/08", "142/18", "97/10"]
    blacklist = anos + leis_e_decretos
    
    ncms_validos = []
    for n in encontrados:
        # Se o número estiver na blacklist ou for parte de uma data, ignora
        if n in blacklist: continue
        
        # NCMs reais no seu PDF geralmente têm pontos (ex: 2201.1) ou 8 dígitos
        # Se for só '2018', o filtro acima já removeu.
        ncms_validos.append(n)
        
    return ncms_validos

def extrair_mvas_todas(texto_completo):
    """Extrai MVA Original e Ajustadas identificando alíquotas pelo contexto."""
    # 1. Identifica as ajustadas primeiro (mantém como está)
    mva4 = re.search(r'([\d,.]+)%\s*\(Al[íi]q\.?\s*4%\s*\)', texto_completo, re.I)
    mva7 = re.search(r'([\d,.]+)%\s*\(Al[íi]q\.?\s*7%\s*\)', texto_completo, re.I)
    mva12 = re.search(r'([\d,.]+)%\s*\(Al[íi]q\.?\s*12%\s*\)', texto_completo, re.I)
    
    val4 = limpar_valor(mva4.group(1) + "%") if mva4 else 0.0
    val7 = limpar_valor(mva7.group(1) + "%") if mva7 else 0.0
    val12 = limpar_valor(mva12.group(1) + "%") if mva12 else 0.0
    ajustadas = [val4, val7, val12]

    # 2. Busca a MVA Original ignorando apenas o que está explicitamente marcado como alíquota
    # Removemos do texto as partes já identificadas como Ajustadas para não confundir
    texto_limpo = re.sub(r'[\d,.]+%?\s*\(Al[íi]q\.?\s*\d+%\s*\)', '', texto_completo, flags=re.I)
    
    # Agora buscamos qualquer percentual que sobrou
    sobras = re.findall(r'([\d,.]+)%', texto_limpo)
    candidatos = []
    for s in sobras:
        v = float(s.replace(',', '.'))
        if v > 0 and v not in ajustadas:
            candidatos.append(v)
    
    # O primeiro valor que sobrar e não for 4, 7 ou 12 (puros) é a nossa MVA Original
    mva_original = 0.0
    for c in candidatos:
        if c not in [4.0, 7.0, 12.0]: # Bloqueia apenas os números "secos" das alíquotas
            mva_original = c
            break
            
    return mva_original, val4, val7, val12

def processar_item_buffer(item):
    texto = item["TEXTO_COMPLETO"]
    mva_orig, mva4, mva7, mva12 = extrair_mvas_inteligente(texto)
    
    lista_ncms = extrair_todos_ncms(texto)
    lista_cests = expandir_intervalo_cest(texto)
    
    # FILTRO CRÍTICO: Se não tem NCM nem CEST, é apenas um título de grupo. 
    # Retornamos lista vazia para o Power Query não receber "lixo".
    if not lista_ncms and not lista_cests:
        return []
    
    registros = []
    for ncm in (lista_ncms if lista_ncms else [""]):
        ncm_ini, ncm_fim = expandir_ncm(ncm)
        for cest in (lista_cests if lista_cests else [""]):
            # LIMPEZA DO CEST: Remove pontos para o Excel/Power Query
            cest_limpo = cest.replace('.', '')
            
            registros.append({
                "ITEM": item["ITEM"],
                "CEST": cest_limpo,
                "NCM": ncm,
                "NCM_INICIAL": ncm_ini,
                "NCM_FINAL": ncm_fim,
                "MVA_ORIGINAL": mva_orig,
                "MVA_AJUSTADA_4": mva4,
                "MVA_AJUSTADA_7": mva7,
                "MVA_AJUSTADA_12": mva12,
                "DESCRIÇÃO": re.sub(r'\s+', ' ', texto).strip()
            })
    return registros

# --- EXECUÇÃO ---

def automatizar_mva(banco_access):
    # O download agora acontece apenas quando a função for chamada pelo botão
    print("📥 Baixando o PDF atualizado da SEFAZ-BA...")
    resposta = requests.get(URL_PDF, verify=False) 
    with open(CAMINHO_PDF_LOCAL, 'wb') as f:
        f.write(resposta.content)
    print("✅ Download concluído! Iniciando extração dos dados...")

    print("🔍 Analisando o PDF...")
    df = executar_extracao(CAMINHO_PDF_LOCAL)
    
    if df.empty:
        print("❌ Nenhum dado encontrado no PDF.")
        return

    # Conecta usando o caminho que veio lá da interface gráfica
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={banco_access};'
    
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    
        print("🧹 Limpando tabela antiga de MVAs no banco...")
        cursor.execute("DELETE FROM tblMVA_Bahia")
    
        print("💾 Salvando as novas MVAs no Access...")
        sql_insert = """INSERT INTO tblMVA_Bahia (
            ITEM, CEST, NCM, NCM_INICIAL, NCM_FINAL, 
            MVA_ORIGINAL, MVA_AJUSTADA_4, MVA_AJUSTADA_7, MVA_AJUSTADA_12, DESCRIÇÃO
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
    
        for index, row in df.iterrows():
            valores = (
                str(row['ITEM']) if pd.notna(row['ITEM']) else "",
                str(row['CEST']) if pd.notna(row['CEST']) else "",
                str(row['NCM']) if pd.notna(row['NCM']) else "",
                str(row['NCM_INICIAL']) if pd.notna(row['NCM_INICIAL']) else "",
                str(row['NCM_FINAL']) if pd.notna(row['NCM_FINAL']) else "",
                float(row['MVA_ORIGINAL']) if pd.notna(row['MVA_ORIGINAL']) else 0.0,
                float(row['MVA_AJUSTADA_4']) if pd.notna(row['MVA_AJUSTADA_4']) else 0.0,
                float(row['MVA_AJUSTADA_7']) if pd.notna(row['MVA_AJUSTADA_7']) else 0.0,
                float(row['MVA_AJUSTADA_12']) if pd.notna(row['MVA_AJUSTADA_12']) else 0.0,
                str(row['DESCRIÇÃO']) if pd.notna(row['DESCRIÇÃO']) else ""
            )
            cursor.execute(sql_insert, valores)

        conn.commit()
        print(f"🚀 Sucesso! {len(df)} linhas atualizadas no banco de dados com base no decreto da SEFAZ.")
    
    except Exception as e:
        print(f"🔥 Erro na gravação do banco: {e}")
    finally:
        if 'cursor' in locals(): cursor.close()
        if 'conn' in locals(): conn.close()
        
if __name__ == "__main__":
    automatizar_mva()

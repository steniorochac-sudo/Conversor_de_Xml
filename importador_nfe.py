import os
import xml.etree.ElementTree as ET
import pyodbc
from datetime import datetime
import time

# --- CONFIGURAÇÕES (USE O 'r' ANTES DAS ASPAS) ---
PASTA_XML = r"C:\Users\StenioRochaCardoso\Cabral & Sousa Ltda\Intranet Cabral & Sousa - 🔎PÚBLICA\🏢 PÚBLICA - FISCAL\01 - STÊNIO ROCHA\Apuração de Impostos - Entrada\Xmls"
BANCO_ACCESS = r'C:\Users\StenioRochaCardoso\Cabral & Sousa Ltda\Intranet Cabral & Sousa - 🔎PÚBLICA\🏢 PÚBLICA - FISCAL\01 - STÊNIO ROCHA\Apuração de Impostos - Entrada\Base_NF_ENTRADA.accdb'

# =======================================================================================
# GUIA RÁPIDO PARA INICIANTES: COMO ADICIONAR NOVAS COLUNAS NO BANCO
# ---------------------------------------------------------------------------------------
# PASSO 0 (No Access): Antes de mexer no Python, abra o seu banco de dados no Access 
# em 'Modo Design' e crie a coluna nova lá (ex: 'vOutro' do tipo Moeda ou Texto).
# 
# Depois, no Python, você fará apenas 3 alterações no loop dos itens (veja abaixo).
# =======================================================================================


def p(tag, root, ns):
    """Busca texto da tag com segurança no namespace nfe"""
    try:
        elemento = root.find(f'.//nfe:{tag}', ns)
        return elemento.text if elemento is not None else None
    except:
        return None

def processar():
    tempo_inicio = time.time()
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={BANCO_ACCESS};'
    
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        print("🚀 Conectado ao Access! Iniciando leitura...")

        # --- CARGA INCREMENTAL ---
        # 1. Busca todas as chaves que já estão no banco
        cursor.execute("SELECT DISTINCT Chave_NFe FROM tblNotasDetalhado")
        notas_existentes = set(row[0] for row in cursor.fetchall())
        print(f"📌 {len(notas_existentes)} notas já existem no banco e serão ignoradas.")
        # -----------------------------------

        # O os.walk varre a pasta principal e todas as subpastas dentro dela
        for diretorio_atual, subpastas, arquivos in os.walk(PASTA_XML):
            for arquivo in arquivos:
                if not arquivo.lower().endswith('.xml'): continue
                
                try:
                    caminho = os.path.join(diretorio_atual, arquivo)
                    tree = ET.parse(caminho)
                    root = tree.getroot()
                    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
                    
                    infNFe = root.find('.//nfe:infNFe', ns)
                    if infNFe is None: continue
                    
                    chave = infNFe.attrib['Id'][3:]
                    
                    # ---  O FILTRO ---
                    # 2. Se a nota já está no banco, pula para o próximo arquivo!
                    if chave in notas_existentes:
                        continue 
                    # --------------------------
                    ide = infNFe.find('nfe:ide', ns)
                    emit = infNFe.find('nfe:emit', ns)
                    dest = infNFe.find('nfe:dest', ns)
                    
                    # Dados da Capa
                    n_nf = p('nNF', ide, ns)
                    d_emi = p('dhEmi', ide, ns) or p('dEmi', ide, ns)
                    d_emi_dt = datetime.fromisoformat(d_emi[:10]) if d_emi else None
                
                    # O período é o nome da pasta onde o XML está localizado, assumindo que a estrutura seja algo como "...\2026\02-2026\arquivo.xml"
                    # Exemplo: se o diretorio_atual for "...\2026\02-2026", o período será "02-2026"
                    periodo = os.path.basename(diretorio_atual)

                    # Loop nos itens (Produtos)
                    for det in infNFe.findall('nfe:det', ns):
                        prod = det.find('nfe:prod', ns)
                        imposto = det.find('nfe:imposto', ns)
                        icms = imposto.find('.//nfe:ICMS', ns) if imposto is not None else None
                        
                        # 1. Lemos os valores de ICMS Normal
                        icms_vbc = float(p('vBC', icms, ns) or 0)
                        icms_picms = float(p('pICMS', icms, ns) or 0)
                        icms_vicms = float(p('vICMS', icms, ns) or 0)
                        
                        # 2. Lemos os valores de Crédito do Simples Nacional
                        icms_pcredsn = float(p('pCredSN', icms, ns) or 0)
                        icms_vcredicmssn = float(p('vCredICMSSN', icms, ns) or 0)
                        
                        # ---> PASSO 1 (No Python): Extrair a nova informação do XML.
                        # Crie uma variável nova. Use 'p()' para buscar a tag.
                        # Se for texto, use: nova_variavel = p('NomeDaTag', det, ns)
                        # Se for número, use: nova_variavel = float(p('NomeDaTag', det, ns) or 0)
                        
                        # Adicionamos as novas colunas no SQL (agora com 34 pontos de interrogação)
                        sql = """INSERT INTO tblNotasDetalhado (
                            Chave_NFe, Periodo, Numero_NF, Data_Emissao, 
                            Emitente_CNPJ, Emitente_Nome, Emitente_UF, Emitente_IE, Emitente_CRT,
                            Destinatario_CNPJ, Destinatario_Nome, Destinatario_UF,
                            Produto_cProd, Produto_xProd, Produto_cEAN, CEST, Produto_NCM, Produto_CFOP, Unidade,
                            Produto_qCom, Produto_vUnCom, Produto_vProd, vIPI, Produto_vDesc, Produto_vFrete,
                            ICMS_CST, ICMS_Item_vBC, ICMS_Item_pICMS, ICMS_Item_vICMS,
                            ICMS_Item_pCredSN, ICMS_Item_vCredICMSSN,
                            vPIS, vCOFINS, cStat                           
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
                        # ---> PASSO 2: Digite o NOME EXATO da nova coluna do Access aqui, antes do 'cStat'
                        
                        # ---> PASSO 3: Para cada coluna nova que você adicionou no Passo 2,
                        # adicione um novo ponto de interrogação '?' no final do bloco VALUES acima.
                        
                        valores = (
                            chave, periodo, n_nf, d_emi_dt,
                            p('CNPJ', emit, ns), p('xNome', emit, ns), p('UF', emit, ns), p('IE', emit, ns), p('CRT', emit, ns),
                            p('CNPJ', dest, ns) or p('CPF', dest, ns), p('xNome', dest, ns), p('UF', dest, ns),
                            p('cProd', prod, ns), p('xProd', prod, ns), p('cEAN', prod, ns), p('CEST', prod, ns),
                            p('NCM', prod, ns), p('CFOP', prod, ns), p('uCom', prod, ns),
                            float(p('qCom', prod, ns) or 0), float(p('vUnCom', prod, ns) or 0), float(p('vProd', prod, ns) or 0),
                            float(p('vIPI', imposto, ns) or 0), float(p('vDesc', prod, ns) or 0), float(p('vFrete', prod, ns) or 0),
                            p('CST', icms, ns) or p('CSOSN', icms, ns),
                            icms_vbc, icms_picms, icms_vicms, 
                            icms_pcredsn, icms_vcredicmssn, # <- Colunas exclusivas do Simples inseridas aqui
                            float(p('vPIS', imposto, ns) or 0), float(p('vCOFINS', imposto, ns) or 0), "100"
                            
                            # ---> PASSO 4: Coloque sua 'nova_variavel' aqui, NA MESMA ORDEM 
                            # em que você digitou o nome da coluna lá no Passo 2!
                        )
                        cursor.execute(sql, valores)
                        
                        # --- NOVIDADE: LEITURA DOS TOTAIS DA NOTA ---
                    bloco_total = infNFe.find('.//nfe:total/nfe:ICMSTot', ns)
                    if bloco_total is not None:
                        tot_vbc = float(p('vBC', bloco_total, ns) or 0)
                        tot_vicms = float(p('vICMS', bloco_total, ns) or 0)
                        tot_vbcst = float(p('vBCST', bloco_total, ns) or 0)
                        tot_vst = float(p('vST', bloco_total, ns) or 0)
                        tot_vfcp = float(p('vFCP', bloco_total, ns) or 0)
                        tot_vpis = float(p('vPIS', bloco_total, ns) or 0)
                        tot_vcofins = float(p('vCOFINS', bloco_total, ns) or 0)
                        tot_vnf = float(p('vNF', bloco_total, ns) or 0)
                        
                        sql_totais = """INSERT INTO tblNotasTotais (
                            Chave_NFe, Periodo, Numero_NF, Data_Emissao, 
                            Emitente_CNPJ, Emitente_Nome, 
                            vBC, vICMS, vBCST, vST, vFCP, vPIS, vCOFINS, vNF
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
                        
                        valores_totais = (
                            chave, periodo, n_nf, d_emi_dt,
                            p('CNPJ', emit, ns), p('xNome', emit, ns),
                            tot_vbc, tot_vicms, tot_vbcst, tot_vst, tot_vfcp, tot_vpis, tot_vcofins, tot_vnf
                        )
                        cursor.execute(sql_totais, valores_totais)
                    # ---------------------------------------------
                    
                    print(f"✅ Nota {n_nf} importada.")
                    conn.commit()
                    
                except Exception as e_file:
                    print(f"❌ Erro no arquivo {arquivo}: {e_file}")
                    
        cursor.close()
        conn.close()
        
        tempo_fim = time.time() 
        duracao = tempo_fim - tempo_inicio
        minutos = int(duracao // 60)
        segundos = int(duracao % 60)
        
        print("\n========================================================")
        print("--- PROCESSO CONCLUÍDO COM SUCESSO! ---")
        print(f"⏱️ Tempo total de processamento: {minutos} minutos e {segundos} segundos.")
        print("========================================================\n")

    except Exception as e:
        print(f"🔥 Erro de conexão ou SQL: {e}")

if __name__ == "__main__":
    processar()
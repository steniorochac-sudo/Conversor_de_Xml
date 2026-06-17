import os
import xml.etree.ElementTree as ET
import pyodbc
from datetime import datetime
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import sys
import extrator_mva

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

import re

def obter_periodo(diretorio_atual, data_emissao_dt):
    """
    Determina o período fiscal (MM-YYYY) a partir da data de emissão ou da pasta.
    """
    # 1. Tenta extrair a partir da data de emissão (Mais preciso)
    if data_emissao_dt:
        return data_emissao_dt.strftime("%m-%Y")
        
    # 2. Tenta extrair um padrão de período (ex: 05-2026, 05/2026, 202605) do caminho completo
    # Procura por MM-YYYY ou MM_YYYY ou MM/YYYY
    match = re.search(r'\b(0[1-9]|1[0-2])[-/_](\d{4})\b', diretorio_atual)
    if match:
        return f"{match.group(1)}-{match.group(2)}"
        
    # Procura por YYYYMM ou MMYYYY
    match_digits = re.search(r'\b(\d{6})\b', diretorio_atual)
    if match_digits:
        digits = match_digits.group(1)
        # Se os primeiros 4 dígitos parecem um ano (ex: 202605)
        if 2000 <= int(digits[:4]) <= 2100:
            return f"{digits[4:6]}-{digits[:4]}"
        # Se os últimos 4 dígitos parecem um ano (ex: 052026)
        elif 2000 <= int(digits[2:]) <= 2100:
            return f"{digits[:2]}-{digits[2:]}"
            
    # 3. Fallback para o nome da pasta atual
    return os.path.basename(diretorio_atual)


# 1. Adicione os parâmetros na função:
def processar(pasta_xml, banco_access, barra_progresso=None):
    tempo_inicio = time.time()
    
    # Conta total de arquivos XML para a barra de progresso
    total_arquivos = sum(len([a for a in arq if a.lower().endswith('.xml')]) for _, _, arq in os.walk(pasta_xml))
    arquivos_processados = 0
    
    # 2. Atualize a variável do banco (minúscula agora):
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={banco_access};'    
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
        for diretorio_atual, subpastas, arquivos in os.walk(pasta_xml):
            for arquivo in arquivos:
                if not arquivo.lower().endswith('.xml'): continue
                
                arquivos_processados += 1
                if barra_progresso and total_arquivos > 0:
                    barra_progresso['value'] = (arquivos_processados / total_arquivos) * 100
                
                try:
                    caminho = os.path.join(diretorio_atual, arquivo)
                    with open(caminho, 'rb') as f:
                        xml_content = f.read()
                    
                    try:
                        from fiscal_workflow.services.parsers import parse_documento_fiscal
                        lista_notas = parse_documento_fiscal(xml_content)
                    except Exception as e_parse:
                        print(f"❌ Erro ao ler/parsear {arquivo}: {e_parse}")
                        continue

                    # Voltar ao XML estruturado para NFe para compatibilidade
                    tree = ET.parse(caminho)
                    root = tree.getroot()
                    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

                    for dados_nota in lista_notas:
                        chave = dados_nota["chave_acesso"]
                        
                        # ---  O FILTRO ---
                        # 2. Se a nota já está no banco, pula!
                        if chave in notas_existentes:
                            continue 
                        # --------------------------
                        
                        # Dados da Capa
                        n_nf = dados_nota["numero_nf"]
                        d_emi = dados_nota["data_emissao"]
                        d_emi_dt = datetime.fromisoformat(d_emi[:10]) if d_emi else None
                    
                        # O período é detectado dinamicamente com base na emissão ou pasta
                        periodo = obter_periodo(diretorio_atual, d_emi_dt)

                        # Loop nos itens
                        for item in dados_nota["itens"]:
                            icms = item["impostos"].get("icms", {})
                            pis = item["impostos"].get("pis", {})
                            cofins = item["impostos"].get("cofins", {})
                            
                            icms_vbc = float(icms.get("valor_base_calculo", 0.0) or 0)
                            icms_picms = float(icms.get("aliquota", 0.0) or 0)
                            icms_vicms = float(icms.get("valor", 0.0) or 0)
                            icms_pcredsn = float(icms.get("aliquota_credito_sn", 0.0) or 0)
                            icms_vcredicmssn = float(icms.get("valor_credito_sn", 0.0) or 0)
                            
                            # Para NFS-e/NFe
                            icms_vbcfcpstret = float(icms.get("valor_base_calculo_fcp_st_ret", 0.0) or 0)
                            icms_pfcpstret = float(icms.get("aliquota_fcp_st_ret", 0.0) or 0)
                            icms_pmvast = float(icms.get("aliquota_mva_st", 0.0) or 0)
                            icms_vbcst = float(icms.get("valor_base_calculo_st", 0.0) or 0)
                            icms_picmsst = float(icms.get("aliquota_st", 0.0) or 0)
                            icms_vicmsst = float(icms.get("valor_st", 0.0) or 0)
                            
                            pis_cst = pis.get("cst", "")
                            pis_vbc = float(pis.get("valor_base_calculo", 0.0) or 0)
                            pis_ppis = float(pis.get("aliquota", 0.0) or 0)
                            pis_vpis = float(pis.get("valor", 0.0) or 0)
                            
                            cofins_cst = cofins.get("cst", "")
                            cofins_vbc = float(cofins.get("valor_base_calculo", 0.0) or 0)
                            cofins_pcofins = float(cofins.get("aliquota", 0.0) or 0)
                            cofins_vcofins = float(cofins.get("valor", 0.0) or 0)

                            sql = """INSERT INTO tblNotasDetalhado (
                                Chave_NFe, Periodo, Numero_NF, Data_Emissao, 
                                Emitente_CNPJ, Emitente_Nome, Emitente_UF, Emitente_IE, Emitente_CRT,
                                Destinatario_CNPJ, Destinatario_Nome, Destinatario_UF,
                                Produto_cProd, Produto_xProd, Produto_cEAN, CEST, Produto_NCM, Produto_CFOP, Unidade,
                                Produto_qCom, Produto_vUnCom, Produto_vProd, vIPI, Produto_vDesc, Produto_vFrete,
                                vBCFCPSTRet, pFCPSTRet, ICMS_CST, ICMS_Item_vBC, ICMS_Item_pICMS, 
                                ICMS_Item_vCredICMSSN, ICMS_Item_vICMS, ICMS_Item_pCredSN, 
                                ICMS_pMVAST, vBC_ST, pICMSST, vICMSST,
                                PIS_CST, PIS_vBC, PIS_pPIS, vPIS, 
                                COFINS_CST, COFINS_vBC, COFINS_pCOFINS, vCOFINS, 
                                cStat
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
                            
                            valores = (
                                chave, periodo, n_nf, d_emi_dt,
                                dados_nota["emitente_cnpj"], dados_nota["emitente_razao_social"], "", "", dados_nota["emitente_crt"],
                                dados_nota.get("destinatario_cnpj", ""), dados_nota.get("destinatario_name", dados_nota.get("destinatario_nome", "")), dados_nota.get("destinatario_uf", ""),
                                item["codigo_produto"], item["descricao"], "", "", item["ncm"], item["cfop"], "UN",
                                float(item["quantidade"] or 0), float(item["valor_unitario"] or 0), float(item["valor_total"] or 0),
                                0.0, float(item["desconto"] or 0), float(item["frete"] or 0),
                                icms_vbcfcpstret, icms_pfcpstret, icms.get("cst", ""), icms_vbc, icms_picms, 
                                icms_vcredicmssn, icms_vicms, icms_pcredsn, 
                                icms_pmvast, icms_vbcst, icms_picmsst, icms_vicmsst,
                                pis_cst, pis_vbc, pis_ppis, pis_vpis, 
                                cofins_cst, cofins_vbc, cofins_pcofins, cofins_vcofins, 
                                dados_nota["cstat"]
                            )
                            cursor.execute(sql, valores)
                            
                        # --- LEITURA DOS TOTAIS DA NOTA ---
                        tot_vbc = 0.0
                        tot_vicms = 0.0
                        tot_vbcst = 0.0
                        tot_vst = 0.0
                        tot_vfcp = 0.0
                        tot_vpis = 0.0
                        tot_vcofins = 0.0
                        tot_vnf = float(dados_nota["valor_total"] or 0)
                        
                        if dados_nota["tipo_documento"] != "NFS-e":
                            bloco_total = root.find('.//nfe:total/nfe:ICMSTot', ns)
                            if bloco_total is not None:
                                tot_vbc = float(p('vBC', bloco_total, ns) or 0)
                                tot_vicms = float(p('vICMS', bloco_total, ns) or 0)
                                tot_vbcst = float(p('vBCST', bloco_total, ns) or 0)
                                tot_vst = float(p('vST', bloco_total, ns) or 0)
                                tot_vfcp = float(p('vFCP', bloco_total, ns) or 0)
                                tot_vpis = float(p('vPIS', bloco_total, ns) or 0)
                                tot_vcofins = float(p('vCOFINS', bloco_total, ns) or 0)
                        else:
                            first_item = dados_nota["itens"][0]
                            tot_vbc = float(first_item["impostos"].get("iss", {}).get("valor_base_calculo", 0.0) or 0)
                            
                        sql_totais = """INSERT INTO tblNotasTotais (
                            Chave_NFe, Periodo, Numero_NF, Data_Emissao, 
                            Emitente_CNPJ, Emitente_Nome, 
                            vBC, vICMS, vBCST, vST, vFCP, vPIS, vCOFINS, vNF
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
                        
                        valores_totais = (
                            chave, periodo, n_nf, d_emi_dt,
                            dados_nota["emitente_cnpj"], dados_nota["emitente_razao_social"],
                            tot_vbc, tot_vicms, tot_vbcst, tot_vst, tot_vfcp, tot_vpis, tot_vcofins, tot_vnf
                        )
                        cursor.execute(sql_totais, valores_totais)
                        
                        print(f"✅ Nota {n_nf} ({dados_nota['tipo_documento']}) importada.")
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


# --- CLASSE PARA REDIRECIONAR O PRINT PARA O TERMINAL DA TELA ---
class RedirecionadorConsole:
    def __init__(self, widget_texto):
        self.widget_texto = widget_texto

    def write(self, texto):
        self.widget_texto.insert(tk.END, texto)
        self.widget_texto.see(tk.END) # Rola automaticamente para baixo
        self.widget_texto.update_idletasks()

    def flush(self):
        pass

def iniciar_interface():
    # Cria a janela principal
    root = tk.Tk()
    root.title("Importador de XML")
    root.geometry("650x550") # Janela aumentada para caber o terminal
    root.resizable(False, False)

    # Variáveis para armazenar os caminhos escolhidos
    pasta_var = tk.StringVar()
    banco_var = tk.StringVar()

    # --- FUNÇÕES DOS BOTÕES ---
    def buscar_pasta():
        pasta = filedialog.askdirectory(title="Selecione a Pasta dos XMLs")
        if pasta: pasta_var.set(pasta)

    def buscar_banco():
        banco = filedialog.askopenfilename(title="Selecione o Banco", filetypes=[("Access", "*.accdb;*.mdb")])
        if banco: banco_var.set(banco)

    def executar():
        if not pasta_var.get() or not banco_var.get():
            messagebox.showwarning("Atenção", "Por favor, selecione a pasta e o banco antes de iniciar.")
            return
        
        btn_iniciar.config(text="Processando...", bg="#757575", state="disabled")
        btn_mva.config(state="disabled") # Bloqueia o outro botão por segurança
        barra_progresso['value'] = 0
        terminal.delete(1.0, tk.END)
        root.update() 
        
        stdout_original = sys.stdout
        sys.stdout = RedirecionadorConsole(terminal)
        
        # Cria uma função interna para rodar em segundo plano
        def tarefa_segundo_plano():
            try:
                processar(pasta_var.get(), banco_var.get(), barra_progresso)
                messagebox.showinfo("Sucesso", "Importação concluída com sucesso! Verifique o log.")
            except Exception as e:
                print(f"🔥 Erro crítico: {e}")
                messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
            finally:
                sys.stdout = stdout_original 
                btn_iniciar.config(text="Iniciar Importação", bg="#4CAF50", state="normal")
                btn_mva.config(state="normal")

        # Inicia a thread paralela (daemon=True garante que ela morre se você fechar o app)
        threading.Thread(target=tarefa_segundo_plano, daemon=True).start()
    
    def executar_mva():
        if not banco_var.get():
            messagebox.showwarning("Atenção", "Selecione o Banco Access antes de atualizar a MVA.")
            return
        
        btn_iniciar.config(state="disabled")
        btn_mva.config(text="Baixando e Atualizando...", bg="#757575", state="disabled")
        terminal.delete(1.0, tk.END)
        root.update() 
        
        stdout_original = sys.stdout
        sys.stdout = RedirecionadorConsole(terminal)
        
        def tarefa_mva_segundo_plano():
            try:
                extrator_mva.automatizar_mva(banco_var.get())
                messagebox.showinfo("Sucesso", "Tabela de MVA atualizada com sucesso no banco!")
            except Exception as e:
                print(f"🔥 Erro crítico: {e}")
                messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
            finally:
                sys.stdout = stdout_original 
                btn_iniciar.config(state="normal")
                btn_mva.config(text="Atualizar Tabela MVA", bg="#2196F3", state="normal")

        threading.Thread(target=tarefa_mva_segundo_plano, daemon=True).start()
        

    # --- LAYOUT DA TELA ---
    # Define uma cor de fundo moderna (um cinza bem claro, padrão de sistemas web)
    cor_fundo = "#F5F6F8" 
    root.configure(bg=cor_fundo)

    style = ttk.Style()
    # Tenta usar o tema nativo moderno do Windows ('vista' ou 'xpnative') para tirar a cara de Win98
    temas = style.theme_names()
    if 'vista' in temas:
        style.theme_use('vista')
    elif 'xpnative' in temas:
        style.theme_use('xpnative')
    elif 'clam' in temas:
        style.theme_use('clam')

    # Força todos os textos e molduras a usarem a mesma cor de fundo da janela
    style.configure("TLabel", background=cor_fundo)
    style.configure("TLabelframe", background=cor_fundo)
    style.configure("TLabelframe.Label", background=cor_fundo, font=("Segoe UI", 9, "bold"), foreground="#333333")

    # Cabeçalho
    ttk.Label(root, text="Importador de XML", font=("Segoe UI", 18, "bold"), foreground="#1A237E").pack(pady=(15, 5))
    ttk.Label(root, text="Módulo de Importação de Notas Fiscais", font=("Segoe UI", 10)).pack(pady=(0, 15))

    # Área de Seleção 1: Pasta
    frame_pasta = ttk.LabelFrame(root, text=" 1. Pasta de Arquivos XML ", padding=(10, 10))
    frame_pasta.pack(fill="x", padx=20, pady=5)
    ttk.Entry(frame_pasta, textvariable=pasta_var, state="readonly", width=68).pack(side="left", padx=(0, 10))
    ttk.Button(frame_pasta, text="Procurar...", command=buscar_pasta).pack(side="left")

    # Área de Seleção 2: Banco de Dados
    frame_banco = ttk.LabelFrame(root, text=" 2. Banco de Dados Access ", padding=(10, 10))
    frame_banco.pack(fill="x", padx=20, pady=5)
    ttk.Entry(frame_banco, textvariable=banco_var, state="readonly", width=68).pack(side="left", padx=(0, 10))
    ttk.Button(frame_banco, text="Procurar...", command=buscar_banco).pack(side="left")

    # --- BOTÕES DE AÇÃO ---
    # Aplica a mesma cor de fundo no frame que segura os botões
    frame_botoes = tk.Frame(root, bg=cor_fundo)
    frame_botoes.pack(fill="x", padx=20, pady=10)

    btn_iniciar = tk.Button(frame_botoes, text="Iniciar Importação", font=("Segoe UI", 11, "bold"), bg="#4CAF50", fg="white", relief="flat", command=executar, height=2, width=30)
    btn_iniciar.pack(side="left", padx=(0, 10), expand=True, fill="x")

    btn_mva = tk.Button(frame_botoes, text="Atualizar Tabela MVA", font=("Segoe UI", 11, "bold"), bg="#2196F3", fg="white", relief="flat", command=executar_mva, height=2, width=30)
    btn_mva.pack(side="right", expand=True, fill="x")
    # --- ÁREA DO TERMINAL E BARRA DE PROGRESSO ---
    frame_log = ttk.LabelFrame(root, text=" Progresso da Importação ", padding=(10, 10))
    frame_log.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    barra_progresso = ttk.Progressbar(frame_log, orient="horizontal", mode="determinate")
    barra_progresso.pack(fill="x", pady=(0, 5))

    scroll = ttk.Scrollbar(frame_log)
    scroll.pack(side="right", fill="y")
    
    terminal = tk.Text(frame_log, height=8, bg="black", fg="#00FF00", font=("Consolas", 9), yscrollcommand=scroll.set)
    terminal.pack(fill="both", expand=True)
    scroll.config(command=terminal.yview)

    # Inicia o programa
    root.mainloop()

if __name__ == "__main__":
    # --- FECHA A TELA DE CARREGAMENTO (Splash Screen do PyInstaller) ---
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass # Ignora o erro ao rodar pelo VS Code
        
    iniciar_interface()
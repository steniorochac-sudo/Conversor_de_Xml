from decimal import Decimal
from datetime import datetime
import json
from typing import List, Optional
from fastapi import FastAPI, Depends, HTTPException, UploadFile, File, Form, status, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse
from sqlalchemy.orm import Session
import asyncio
import logging
from logging.handlers import RotatingFileHandler
import os

# Configura o Logger para gravar no arquivo na raiz do projeto
log_file = "fiscal_workflow.log"


file_handler = RotatingFileHandler(log_file, maxBytes=1024 * 1024 * 5, backupCount=3, encoding="utf-8")
file_handler.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
file_handler.setFormatter(formatter)

# Configura logger padrão da aplicação
app_logger = logging.getLogger("fiscal_workflow")
app_logger.setLevel(logging.INFO)
app_logger.addHandler(file_handler)

# Adiciona o handler nos loggers do uvicorn para capturar acessos e inicialização
for logger_name in ("uvicorn", "uvicorn.error", "uvicorn.access"):
    l = logging.getLogger(logger_name)
    l.addHandler(file_handler)

# Importa a conexão com o banco de dados
from fiscal_workflow.db.database import engine, get_db, SessionLocal
from fiscal_workflow.models.models import Base, Empresa, DocumentoFiscal, AjusteDocumento, StatusApuracao, RegimeTributario
from fiscal_workflow.schemas.schemas import EmpresaCreate, EmpresaResponse, EmpresaUpdate, AjusteCreate, DocumentoResponse
from fiscal_workflow.services.parsers import parse_documento_fiscal
from fiscal_workflow.services.calculadoras import CalculadoraFactory
from fiscal_workflow.core.dashboard_template import HTML_CONTENT
from fiscal_workflow.services.cnae_syncer import sync_cnaes_from_ibge, CNAE_JSON_PATH

# Inicializa as tabelas do banco de dados na inicialização da API
# (Garante facilidade em novos ambientes e testes automatizados)
Base.metadata.create_all(bind=engine)

# Migração em runtime: garante que a coluna 'tipo_operacao' exista na tabela documentos_fiscais e rbt12/folha12/sujeito_fator_r em empresas
from sqlalchemy import text
with engine.connect() as conn:
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ALTER COLUMN chave_acesso TYPE VARCHAR(60)"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ADD COLUMN tipo_operacao VARCHAR(20) DEFAULT 'Saída'"))
    except Exception:
        pass # Ignora se a coluna já existir (comum em PostgreSQL ou após rodar o app uma vez)
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ADD COLUMN cstat VARCHAR(10) DEFAULT '100'"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE empresas ADD COLUMN rbt12 NUMERIC(15, 2) DEFAULT 0.00"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE empresas ADD COLUMN folha12 NUMERIC(15, 2) DEFAULT 0.00"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE empresas ADD COLUMN sujeito_fator_r BOOLEAN DEFAULT FALSE"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE empresas ADD COLUMN categoria_simples VARCHAR(50) DEFAULT 'Serviços (Anexo III)'"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE empresas ADD COLUMN cnae VARCHAR(7)"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ADD COLUMN data_emissao TIMESTAMP"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ADD COLUMN data_competencia TIMESTAMP"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE empresas ADD COLUMN uf VARCHAR(2) DEFAULT 'BA'"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ADD COLUMN numero_nf VARCHAR(20)"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ADD COLUMN emitente_nome VARCHAR(255)"))
    except Exception:
        pass
    try:
        with conn.begin():
            conn.execute(text("ALTER TABLE documentos_fiscais ADD COLUMN destinatario_nome VARCHAR(255)"))
    except Exception:
        pass

class HeartbeatManager:
    def __init__(self):
        self.active_connections = 0
        self.shutdown_task = None
        self.has_connected_once = False
        
    def connect(self):
        self.active_connections += 1
        self.has_connected_once = True
        if self.shutdown_task:
            self.shutdown_task.cancel()
            self.shutdown_task = None
            
    def disconnect(self):
        self.active_connections -= 1
        if self.has_connected_once and self.active_connections <= 0:
            self.shutdown_task = asyncio.create_task(self.delayed_shutdown(8))
            
    async def delayed_shutdown(self, delay: int):
        try:
            await asyncio.sleep(delay)
            app_logger.info("Nenhuma conexão ativa detectada. Desligando o servidor uvicorn...")
            import os
            import signal
            os.kill(os.getpid(), signal.SIGINT)
        except asyncio.CancelledError:
            pass

heartbeat_manager = HeartbeatManager()

def obter_caminho_copia_xml(
    empresa_nome: str,
    data_competencia: Optional[datetime],
    tipo_operacao: str,
    tipo_documento: str,
    chave_acesso: str
):
    """Calcula o caminho estruturado onde o XML da nota fiscal deve ser armazenado com base na competência."""
    import re
    from pathlib import Path
    empresa_saneada = re.sub(r'[\/\\\:\*\?\"\<\>\|]', '_', empresa_nome).strip()
    if not empresa_saneada:
        empresa_saneada = "Empresa_Nao_Identificada"

    if data_competencia:
        periodo = data_competencia.strftime("%Y%m")
    else:
        periodo = datetime.now().strftime("%Y%m")

    movimento = "Entradas" if tipo_operacao == "Entrada" else "Saídas"
    tipo_doc_limpo = "NFSe" if "nfs" in tipo_documento.lower() else "NFe"
    base_dir = Path("armazenamento_xml")
    return base_dir / empresa_saneada / periodo / movimento / tipo_doc_limpo / f"{chave_acesso}.xml"

def salvar_copia_xml(
    xml_content: bytes,
    empresa_nome: str,
    data_competencia: Optional[datetime],
    tipo_operacao: str,
    tipo_documento: str,
    chave_acesso: str
):
    """
    Salva uma cópia do arquivo XML de forma organizada na pasta raiz.
    """
    try:
        target_file = obter_caminho_copia_xml(empresa_nome, data_competencia, tipo_operacao, tipo_documento, chave_acesso)
        target_file.parent.mkdir(parents=True, exist_ok=True)
        with open(target_file, "wb") as f:
            f.write(xml_content)
        app_logger.info(f"Cópia do XML salva com sucesso em: {target_file}")
    except Exception as e:
        app_logger.error(f"Erro ao salvar cópia do XML para chave {chave_acesso}: {str(e)}")

def sincronizar_arquivo_xml(
    empresa_nome: str,
    data_competencia_antiga: Optional[datetime],
    data_competencia_nova: Optional[datetime],
    tipo_operacao: str,
    tipo_documento: str,
    chave_acesso: str
):
    """
    Move o arquivo XML físico da pasta antiga para a nova pasta se houver mudança de período de competência.
    Limpa diretórios vazios remanescentes.
    """
    try:
        old_file = obter_caminho_copia_xml(empresa_nome, data_competencia_antiga, tipo_operacao, tipo_documento, chave_acesso)
        new_file = obter_caminho_copia_xml(empresa_nome, data_competencia_nova, tipo_operacao, tipo_documento, chave_acesso)
        
        if old_file != new_file and old_file.exists():
            new_file.parent.mkdir(parents=True, exist_ok=True)
            import shutil
            shutil.move(str(old_file), str(new_file))
            app_logger.info(f"Arquivo XML físico movido de {old_file} para {new_file}")
            
            # Limpa pastas vazias órfãs de forma recursiva
            try:
                parent = old_file.parent
                for _ in range(4):
                    if parent.exists() and not any(parent.iterdir()):
                        parent.rmdir()
                        parent = parent.parent
                    else:
                        break
            except Exception:
                pass
    except Exception as e:
        app_logger.error(f"Erro ao mover arquivo XML físico da nota {chave_acesso}: {str(e)}")

def obter_periodo_do_caminho(caminho: str) -> Optional[datetime]:
    """
    Tenta extrair um período (ano e mês) a partir do caminho do arquivo (ex: webkitRelativePath).
    Retorna um objeto datetime no dia 1 do mês, ou None.
    """
    if not caminho:
        return None
    import re
    # Procura por MM-YYYY, MM_YYYY, MM/YYYY
    match = re.search(r'\b(0[1-9]|1[0-2])[-/_](\d{4})\b', caminho)
    if match:
        mes = int(match.group(1))
        ano = int(match.group(2))
        return datetime(ano, mes, 1)
    # Procura por YYYYMM ou MMYYYY (6 dígitos)
    match_digits = re.search(r'\b(\d{6})\b', caminho)
    if match_digits:
        digits = match_digits.group(1)
        if 2000 <= int(digits[:4]) <= 2100:
            ano = int(digits[:4])
            mes = int(digits[4:6])
            if 1 <= mes <= 12:
                return datetime(ano, mes, 1)
        elif 2000 <= int(digits[2:]) <= 2100:
            mes = int(digits[:2])
            ano = int(digits[2:])
            if 1 <= mes <= 12:
                return datetime(ano, mes, 1)
    # Procura por YYYY-MM ou YYYY_MM (7 caracteres com separador)
    match_yyyy_mm = re.search(r'\b(\d{4})[-/_](0[1-9]|1[0-2])\b', caminho)
    if match_yyyy_mm:
        ano = int(match_yyyy_mm.group(1))
        mes = int(match_yyyy_mm.group(2))
        return datetime(ano, mes, 1)
    return None

app = FastAPI(
    title="Workflow Modular Fiscal",
    description="API de ingestão de XMLs, Staging Area e motor de apuração fiscal",
    version="1.0.0"
)

@app.on_event("startup")
def startup_backfill():
    """Realiza o backfill automático dos campos novos das notas já cadastradas no banco."""
    db = SessionLocal()
    try:
        documentos = db.query(DocumentoFiscal).all()
        for doc in documentos:
            atualizado = False
            # 1. Backfill do número da nota (extraído da chave de acesso)
            if not doc.numero_nf and doc.chave_acesso:
                try:
                    if len(doc.chave_acesso) == 44:
                        num_str = doc.chave_acesso[25:34].lstrip("0")
                        doc.numero_nf = num_str if num_str else "0"
                        atualizado = True
                    elif doc.chave_acesso.startswith("NFSE"):
                        doc.numero_nf = doc.chave_acesso[4:]
                        atualizado = True
                except Exception:
                    pass
                
            # 2. Backfill dos parceiros
            if not doc.emitente_nome:
                if doc.tipo_operacao == "Saída":
                    doc.emitente_nome = doc.empresa.razao_social
                    atualizado = True
                else:
                    doc.emitente_nome = "Fornecedor Não Informado"
                    atualizado = True
            
            if not doc.destinatario_nome:
                if doc.tipo_operacao == "Entrada":
                    doc.destinatario_nome = doc.empresa.razao_social
                    atualizado = True
                else:
                    doc.destinatario_nome = "Cliente Não Informado"
                    atualizado = True
                    
            if not doc.data_competencia:
                if doc.data_emissao:
                    doc.data_competencia = datetime(doc.data_emissao.year, doc.data_emissao.month, 1)
                    atualizado = True
                    
            if atualizado:
                db.add(doc)
        db.commit()
        app_logger.info("Backfill de campos novos de notas fiscais concluído com sucesso.")
    except Exception as e:
        app_logger.error(f"Erro ao executar backfill de notas fiscais: {str(e)}")
    finally:
        db.close()

@app.websocket("/ws/heartbeat")
async def websocket_heartbeat(websocket: WebSocket):
    await websocket.accept()
    heartbeat_manager.connect()
    try:
        while True:
            # Mantém a conexão aberta esperando mensagens (ou desconexão)
            await websocket.receive_text()
    except WebSocketDisconnect:
        pass
    finally:
        heartbeat_manager.disconnect()

@app.get("/", response_class=HTMLResponse)
def read_root():
    """Retorna o painel visual (Dashboard) de staging e apuração fiscal."""
    return HTMLResponse(content=HTML_CONTENT)

@app.get("/api/logs")
def get_system_logs():
    """Retorna as últimas 200 linhas do arquivo de log do sistema."""
    log_path = "fiscal_workflow.log"
    if not os.path.exists(log_path):
        return {"logs": ["Nenhum log gerado ainda no servidor."]}
    
    try:
        with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
            last_lines = lines[-200:]
            return {"logs": [line.strip() for line in last_lines]}
    except Exception as e:
        return {"logs": [f"Erro ao ler arquivo de logs: {str(e)}"]}

@app.post("/api/logs/clear")
def clear_system_logs():
    """Limpa o conteúdo do arquivo de log."""
    log_path = "fiscal_workflow.log"
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("")
        app_logger.info("Histórico de logs limpo pelo usuário.")
        return {"status": "success", "message": "Logs limpos com sucesso."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao limpar logs: {str(e)}")


# ==========================================
# ENDPOINTS DE CADASTRO DE EMPRESA
# ==========================================

@app.post("/empresas", response_model=EmpresaResponse, status_code=status.HTTP_201_CREATED)
def criar_empresa(empresa: EmpresaCreate, db: Session = Depends(get_db)):
    """Cadastra uma nova Empresa no sistema, validando o CNPJ único."""
    empresa_existente = db.query(Empresa).filter(Empresa.cnpj == empresa.cnpj).first()
    if empresa_existente:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, 
            detail=f"Já existe uma empresa cadastrada com o CNPJ {empresa.cnpj}."
        )
    
    nova_empresa = Empresa(
        cnpj=empresa.cnpj,
        razao_social=empresa.razao_social,
        regime_tributario=empresa.regime_tributario,
        rbt12=empresa.rbt12 if empresa.rbt12 is not None else Decimal("0.00"),
        folha12=empresa.folha12 if empresa.folha12 is not None else Decimal("0.00"),
        sujeito_fator_r=empresa.sujeito_fator_r if empresa.sujeito_fator_r is not None else False,
        categoria_simples=empresa.categoria_simples if empresa.categoria_simples is not None else "Serviços (Anexo III)",
        cnae=empresa.cnae,
        uf=empresa.uf if empresa.uf is not None else "BA"
    )
    db.add(nova_empresa)
    db.commit()
    db.refresh(nova_empresa)
    return nova_empresa

@app.get("/empresas", response_model=List[EmpresaResponse])
def listar_empresas(db: Session = Depends(get_db)):
    """Retorna a lista de todas as empresas cadastradas."""
    return db.query(Empresa).all()

@app.put("/empresas/{empresa_id}", response_model=EmpresaResponse)
def editar_empresa(empresa_id: int, empresa_update: EmpresaUpdate, db: Session = Depends(get_db)):
    """Edita a razão social ou o regime tributário de uma Empresa cadastrada."""
    empresa = db.query(Empresa).filter(Empresa.id == empresa_id).first()
    if not empresa:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, 
            detail=f"Empresa com ID {empresa_id} não encontrada."
        )
    
    empresa.razao_social = empresa_update.razao_social
    empresa.regime_tributario = empresa_update.regime_tributario
    empresa.rbt12 = empresa_update.rbt12
    empresa.folha12 = empresa_update.folha12
    empresa.sujeito_fator_r = empresa_update.sujeito_fator_r
    empresa.categoria_simples = empresa_update.categoria_simples
    empresa.cnae = empresa_update.cnae
    empresa.uf = empresa_update.uf
    
    db.commit()
    db.refresh(empresa)
    return empresa

# ==========================================
# ENDPOINTS DO FLUXO DE DOCUMENTOS
# ==========================================

@app.post("/documentos/upload", response_model=List[DocumentoResponse], status_code=status.HTTP_201_CREATED)
async def upload_xml(
    empresa_id: Optional[int] = Form(None, description="ID da empresa para forçar o vínculo e a importação."),
    tipo_operacao_forcada: Optional[str] = Form(None, description="Forçar tipo de operação ('Entrada' ou 'Saída')."),
    data_competencia: Optional[str] = Form(None, description="Data da competência/entrada (formato YYYY-MM-DD ou YYYY-MM)."),
    files: List[UploadFile] = File(..., description="Arquivos XML de NF-e / NFC-e / NFS-e (lote)"),
    db: Session = Depends(get_db)
):
    """
    Recebe múltiplos arquivos XML em lote, realiza o parse de cada um,
    resolve a empresa (via parâmetro ou autodetecção pelo CNPJ) e salva na Staging Area.
    """
    documentos_salvos = []
    
    for file in files:
        if not file.filename.lower().endswith(".xml"):
            continue
            
        try:
            xml_content = await file.read()
            lista_notas = parse_documento_fiscal(xml_content)
        except Exception:
            continue

        for dados_nota in lista_notas:
            # 1. Parse da data de emissão real do XML
            real_data_emissao = None
            if dados_nota.get("data_emissao"):
                try:
                    real_data_emissao = datetime.fromisoformat(dados_nota["data_emissao"])
                except Exception:
                    pass

            # 2. Valida se a chave de acesso já foi importada
            doc_existente = db.query(DocumentoFiscal).filter(
                DocumentoFiscal.chave_acesso == dados_nota["chave_acesso"]
            ).first()
            
            if doc_existente:
                dados_atualizados = False
                tipo_op_resolvido = doc_existente.tipo_operacao
                
                # Resolve dt_competencia usando a hierarquia para o tipo resolvido
                dt_competencia = None
                if data_competencia:
                    try:
                        if len(data_competencia) == 10:
                            dt_competencia = datetime.fromisoformat(data_competencia)
                        elif len(data_competencia) == 7:
                            dt_competencia = datetime.strptime(data_competencia, "%Y-%m")
                        if dt_competencia:
                            dt_competencia = datetime(dt_competencia.year, dt_competencia.month, 1)
                    except Exception:
                        pass
                
                if tipo_op_resolvido == "Entrada":
                    if not dt_competencia:
                        dt_competencia = obter_periodo_do_caminho(file.filename)
                    if not dt_competencia and real_data_emissao:
                        dt_competencia = datetime(real_data_emissao.year, real_data_emissao.month, 1)
                else:
                    if not dt_competencia and real_data_emissao:
                        dt_competencia = datetime(real_data_emissao.year, real_data_emissao.month, 1)
                
                if not dt_competencia:
                    hoje = datetime.now()
                    dt_competencia = datetime(hoje.year, hoje.month, 1)
                
                data_competencia_antiga = doc_existente.data_competencia
                
                # Atualiza a competência se ela mudou ou se foi forçada
                if dt_competencia and (doc_existente.data_competencia is None or data_competencia):
                    if doc_existente.data_competencia != dt_competencia:
                        doc_existente.data_competencia = dt_competencia
                        dados_atualizados = True
                
                # Atualiza a emissão real se ela for diferente ou nula no banco
                if real_data_emissao and doc_existente.data_emissao != real_data_emissao:
                    doc_existente.data_emissao = real_data_emissao
                    dados_atualizados = True
                
                nova_cstat = dados_nota.get("cstat", "100")
                if doc_existente.cstat != nova_cstat:
                    doc_existente.cstat = nova_cstat
                    dados_atualizados = True
                
                if dados_atualizados:
                    db.commit()
                    db.refresh(doc_existente)
                    # Sincroniza o arquivo físico de pasta caso a competência mude
                    if data_competencia_antiga != dt_competencia:
                        sincronizar_arquivo_xml(
                            empresa_nome=doc_existente.empresa.razao_social,
                            data_competencia_antiga=data_competencia_antiga,
                            data_competencia_nova=dt_competencia,
                            tipo_operacao=doc_existente.tipo_operacao,
                            tipo_documento=doc_existente.tipo_documento,
                            chave_acesso=doc_existente.chave_acesso
                        )
                
                # Salva uma cópia estruturada do arquivo XML importado
                salvar_copia_xml(
                    xml_content=dados_nota.get("xml_content", xml_content),
                    empresa_nome=doc_existente.empresa.razao_social,
                    data_competencia=doc_existente.data_competencia,
                    tipo_operacao=doc_existente.tipo_operacao,
                    tipo_documento=doc_existente.tipo_documento,
                    chave_acesso=doc_existente.chave_acesso
                )
                
                documentos_salvos.append(doc_existente)
                continue

            # 3. Resolução da Empresa Dona da Nota (Parâmetro Explicito ou autodetecção por CNPJ cadastrado)
            tipo_operacao = tipo_operacao_forcada or dados_nota.get("tipo_operacao", "Saída")
            
            empresa = None
            if empresa_id:
                emp_candidate = db.query(Empresa).filter(Empresa.id == empresa_id).first()
                if not emp_candidate:
                    raise HTTPException(
                        status_code=status.HTTP_404_NOT_FOUND,
                        detail=f"Empresa com ID {empresa_id} especificada no formulário não encontrada."
                    )
                if tipo_operacao_forcada:
                    empresa = emp_candidate
                else:
                    if emp_candidate.cnpj == dados_nota.get("destinatario_cnpj"):
                        empresa = emp_candidate
                        tipo_operacao = "Entrada"
                    elif emp_candidate.cnpj == dados_nota.get("emitente_cnpj"):
                        empresa = emp_candidate
                        tipo_operacao = "Saída"
                    else:
                        empresa = None

            if not empresa:
                # 1. Se o destinatário for uma empresa cadastrada, é Entrada
                dest_cnpj = dados_nota.get("destinatario_cnpj")
                if dest_cnpj:
                    empresa = db.query(Empresa).filter(Empresa.cnpj == dest_cnpj).first()
                    if empresa:
                        tipo_operacao = "Entrada"
                
                # 2. Se não encontrou e o emitente for uma empresa cadastrada, é Saída
                if not empresa:
                    emit_cnpj = dados_nota.get("emitente_cnpj")
                    if emit_cnpj:
                        empresa = db.query(Empresa).filter(Empresa.cnpj == emit_cnpj).first()
                        if empresa:
                            tipo_operacao = "Saída"
                
                # 3. Se não for nenhuma empresa cadastrada, autocadastra
                if not empresa:
                    is_entrada = tipo_operacao == "Entrada"
                    cnpj_empresa = dados_nota.get("destinatario_cnpj") if is_entrada else dados_nota.get("emitente_cnpj")
                    if not cnpj_empresa:
                        cnpj_empresa = dados_nota.get("emitente_cnpj") or dados_nota.get("destinatario_cnpj")
                        
                    if not cnpj_empresa:
                        continue # Pula se o XML não tiver dados identificáveis da empresa
                        
                    empresa = db.query(Empresa).filter(Empresa.cnpj == cnpj_empresa).first()
                    
                    if not empresa:
                        crt = dados_nota.get("emitente_crt") if not is_entrada else "1"
                        if crt in ("1", "2"):
                            regime = RegimeTributario.SIMPLES_NACIONAL
                        elif crt == "3":
                            regime = RegimeTributario.LUCRO_PRESUMIDO
                        else:
                            regime = RegimeTributario.SIMPLES_NACIONAL

                        from fiscal_workflow.services.cnpj_client import buscar_cnae_oficial
                        cnae_resolvido = buscar_cnae_oficial(cnpj_empresa)
                        
                        sujeito_fator_r = False
                        categoria_simples = "Serviços (Anexo III)"
                        
                        if cnae_resolvido:
                            try:
                                if CNAE_JSON_PATH.exists():
                                    with open(CNAE_JSON_PATH, "r", encoding="utf-8") as f:
                                        regras_locais = json.load(f)
                                    if cnae_resolvido in regras_locais:
                                        regra = regras_locais[cnae_resolvido]
                                        sujeito_fator_r = regra.get("fator_r", False)
                                        if regra.get("anexo") == "I":
                                            categoria_simples = "Comércio (Anexo I)"
                                        else:
                                            categoria_simples = "Serviços (Anexo III)"
                                    else:
                                        if cnae_resolvido[:2] in ("45", "46", "47"):
                                            categoria_simples = "Comércio (Anexo I)"
                                        elif cnae_resolvido[:2] in ("62", "86", "73", "74", "69"):
                                            sujeito_fator_r = True
                            except Exception:
                                pass

                        razao_social = (dados_nota.get("destinatario_nome") if is_entrada else dados_nota.get("emitente_razao_social")) or f"Empresa CNPJ {cnpj_empresa}"
                        uf_empresa = (dados_nota.get("destinatario_uf") if is_entrada else "BA") or "BA"

                        empresa = Empresa(
                             cnpj=cnpj_empresa,
                             razao_social=razao_social,
                             regime_tributario=regime,
                             rbt12=Decimal("0.00"),
                             folha12=Decimal("0.00"),
                             sujeito_fator_r=sujeito_fator_r,
                             categoria_simples=categoria_simples,
                             cnae=cnae_resolvido,
                             uf=uf_empresa
                        )
                        db.add(empresa)
                        db.commit()
                        db.refresh(empresa)

            # 4. Salva a nota fiscal na Staging Area
            dt_competencia = None
            if data_competencia:
                try:
                    if len(data_competencia) == 10:
                        dt_competencia = datetime.fromisoformat(data_competencia)
                    elif len(data_competencia) == 7:
                        dt_competencia = datetime.strptime(data_competencia, "%Y-%m")
                    if dt_competencia:
                        dt_competencia = datetime(dt_competencia.year, dt_competencia.month, 1)
                except Exception:
                    pass

            if tipo_operacao == "Entrada":
                if not dt_competencia:
                    dt_competencia = obter_periodo_do_caminho(file.filename)
                if not dt_competencia and real_data_emissao:
                    dt_competencia = datetime(real_data_emissao.year, real_data_emissao.month, 1)
            else:
                if not dt_competencia and real_data_emissao:
                    dt_competencia = datetime(real_data_emissao.year, real_data_emissao.month, 1)

            if not dt_competencia:
                hoje = datetime.now()
                dt_competencia = datetime(hoje.year, hoje.month, 1)

            novo_doc = DocumentoFiscal(
                empresa_id=empresa.id,
                chave_acesso=dados_nota["chave_acesso"],
                tipo_documento=dados_nota["tipo_documento"],
                tipo_operacao=tipo_operacao,
                valor_total=Decimal(str(dados_nota["valor_total"])),
                status_apuracao=StatusApuracao.PENDENTE,
                cstat=dados_nota.get("cstat", "100"),
                itens=dados_nota["itens"],
                data_emissao=real_data_emissao,
                data_competencia=dt_competencia,
                numero_nf=dados_nota.get("numero_nf"),
                emitente_nome=dados_nota.get("emitente_razao_social"),
                destinatario_nome=dados_nota.get("destinatario_nome")
            )

            db.add(novo_doc)
            db.commit()
            db.refresh(novo_doc)
            
            # Salva uma cópia estruturada do arquivo XML importado
            salvar_copia_xml(
                xml_content=dados_nota.get("xml_content", xml_content),
                empresa_nome=empresa.razao_social,
                data_competencia=dt_competencia,
                tipo_operacao=tipo_operacao,
                tipo_documento=dados_nota["tipo_documento"],
                chave_acesso=dados_nota["chave_acesso"]
            )
            
            documentos_salvos.append(novo_doc)

    return documentos_salvos


@app.get("/documentos", response_model=List[DocumentoResponse])
def listar_documentos(
    empresa_id: Optional[int] = None, 
    status: Optional[StatusApuracao] = None,
    mes: Optional[int] = None,
    ano: Optional[int] = None,
    tipo_operacao: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """Lista todos os documentos importados (Staging Area) com filtros opcionais."""
    from sqlalchemy import extract
    query = db.query(DocumentoFiscal)
    if empresa_id:
        query = query.filter(DocumentoFiscal.empresa_id == empresa_id)
    if status:
        query = query.filter(DocumentoFiscal.status_apuracao == status)
    if mes:
        query = query.filter(extract('month', DocumentoFiscal.data_competencia) == mes)
    if ano:
        query = query.filter(extract('year', DocumentoFiscal.data_competencia) == ano)
    if tipo_operacao:
        query = query.filter(DocumentoFiscal.tipo_operacao == tipo_operacao)
    
    return query.all()

# ==========================================
# ENDPOINTS DE AJUSTES MANUAIS (AUDITORIA)
# ==========================================

@app.post("/documentos/{documento_id}/ajustes", response_model=DocumentoResponse)
def registrar_ajuste(
    documento_id: int, 
    ajuste: AjusteCreate, 
    db: Session = Depends(get_db)
):
    """
    Registra um ajuste manual auditável em um documento fiscal na Staging Area.
    Modifica automaticamente o status do documento para 'Em Revisão'.
    Bloqueia alterações caso o período esteja 'Encerrado'.
    """
    # 1. Busca o documento
    doc = db.query(DocumentoFiscal).filter(DocumentoFiscal.id == documento_id).first()
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, 
            detail=f"Documento fiscal com ID {documento_id} não encontrado."
        )

    # 2. Regra de Bloqueio (Snapshot consolidado)
    if doc.status_apuracao == StatusApuracao.ENCERRADO:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, 
            detail="Não é possível registrar ajustes manuais. Este período fiscal já está Encerrado/Consolidado."
        )

    # 3. Registra o log de ajuste
    novo_ajuste = AjusteDocumento(
        documento_id=doc.id,
        valor_total_ajuste=ajuste.valor_total_ajuste,
        justificativa=ajuste.justificativa,
        usuario=ajuste.usuario
    )
    db.add(novo_ajuste)

    # 4. Modifica o status para sinalizar revisão em progresso
    doc.status_apuracao = StatusApuracao.EM_REVISAO
    
    db.commit()
    db.refresh(doc)
    return doc


@app.post("/documentos/{documento_id}/encerrar", response_model=DocumentoResponse)
def encerrar_apuracao(documento_id: int, db: Session = Depends(get_db)):
    """
    Consolida e encerra o período fiscal deste documento, gerando um snapshot.
    Impede novas alterações manuais.
    """
    doc = db.query(DocumentoFiscal).filter(DocumentoFiscal.id == documento_id).first()
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, 
            detail=f"Documento fiscal com ID {documento_id} não encontrado."
        )

    doc.status_apuracao = StatusApuracao.ENCERRADO
    db.commit()
    db.refresh(doc)
    return doc

# ==========================================
# MOTOR DE CÁLCULO DINÂMICO (STRATEGY PATTERN)
# ==========================================

@app.get("/documentos/{documento_id}/apurar")
def apurar_documento(
    documento_id: int, 
    aliquota: Optional[float] = None, 
    db: Session = Depends(get_db)
):
    """
    Executa a esteira de cálculo de impostos dinamicamente (Strategy Pattern).
    Obtém a empresa do documento, identifica o seu regime tributário 
    e injeta a calculadora (estratégia) correta.
    """
    # 1. Busca o documento
    doc = db.query(DocumentoFiscal).filter(DocumentoFiscal.id == documento_id).first()
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, 
            detail=f"Documento fiscal com ID {documento_id} não encontrado."
        )

    # 2. Obtém a empresa e o regime
    empresa = doc.empresa

    # 3. Injeta a calculadora correta a partir do Regime Tributário
    try:
        calculadora = CalculadoraFactory.obter_calculadora(empresa.regime_tributario)
        
        # Converte a alíquota opcional float para Decimal
        aliquota_decimal = Decimal(str(aliquota)) if aliquota is not None else None
        
        # Roda o cálculo dinâmico da estratégia
        apuracao = calculadora.calcular(doc, aliquota_decimal)
        return apuracao
    except ValueError as val_err:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, 
            detail=str(val_err)
        )


@app.get("/empresas/{empresa_id}/consolidado")
def apurar_consolidado_empresa(
    empresa_id: int, 
    mes: Optional[int] = None,
    ano: Optional[int] = None,
    db: Session = Depends(get_db)
):
    """
    Realiza a apuração consolidada de impostos de todas as notas fiscais
    cadastradas na Staging Area para a empresa selecionada.
    """
    # 1. Busca a empresa
    empresa = db.query(Empresa).filter(Empresa.id == empresa_id).first()
    if not empresa:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, 
            detail=f"Empresa com ID {empresa_id} não encontrada."
        )

    # 2. Busca todas as notas da empresa
    from sqlalchemy import extract
    query = db.query(DocumentoFiscal).filter(DocumentoFiscal.empresa_id == empresa_id)
    if mes:
        query = query.filter(extract('month', DocumentoFiscal.data_competencia) == mes)
    if ano:
        query = query.filter(extract('year', DocumentoFiscal.data_competencia) == ano)
    documentos = query.all()

    # Separa os documentos por tipo de operação para apuração correta
    documentos_saida = [d for d in documentos if d.tipo_operacao == "Saída"]
    documentos_entrada = [d for d in documentos if d.tipo_operacao == "Entrada"]

    # 3. Inicializa calculadora tributária
    try:
        calculadora = CalculadoraFactory.obter_calculadora(empresa.regime_tributario)
    except ValueError as val_err:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, 
            detail=str(val_err)
        )

    total_faturamento = Decimal("0.00")
    total_imposto = Decimal("0.00")
    
    # Detalhes consolidados dependendo do regime
    detalhes = {}
    if empresa.regime_tributario == RegimeTributario.SIMPLES_NACIONAL:
        detalhes = {"das": Decimal("0.00")}
    elif empresa.regime_tributario == RegimeTributario.LUCRO_PRESUMIDO:
        detalhes = {
            "pis": Decimal("0.00"),
            "cofins": Decimal("0.00"),
            "irpj": Decimal("0.00"),
            "csll": Decimal("0.00")
        }

    active_docs_count = 0
    canceled_docs_count = 0

    # 4. Apura cada nota de saída (venda/faturamento) e consolida
    for doc in documentos_saida:
        apuracao_nota = calculadora.calcular(doc)
        if doc.cstat in ("101", "110", "301", "302"):
            canceled_docs_count += 1
            continue
            
        active_docs_count += 1
        total_faturamento += Decimal(str(apuracao_nota["valor_final_base"]))
        total_imposto += Decimal(str(apuracao_nota["imposto_calculado"]))
        
        for k, v in apuracao_nota["detalhes"].items():
            if k in detalhes:
                detalhes[k] += Decimal(str(v))

    # 4.1. Apura cada nota de entrada (compra)
    total_compras = Decimal("0.00")
    total_difal = Decimal("0.00")
    total_icms_st_compra = Decimal("0.00")
    active_entradas_count = 0
    
    for doc in documentos_entrada:
        if doc.cstat in ("101", "110", "301", "302"):
            continue
        active_entradas_count += 1
        apuracao_compra = calculadora.calcular(doc)
        total_compras += Decimal(str(apuracao_compra["valor_final_base"]))
        total_difal += Decimal(str(apuracao_compra["detalhes"].get("difal", 0.0)))
        total_icms_st_compra += Decimal(str(apuracao_compra["detalhes"].get("icms_st_compra", 0.0)))

    # Alíquota efetiva consolidada de saída
    aliquota_efetiva = (total_imposto / total_faturamento) if total_faturamento > 0 else Decimal("0.00")

    # Consolida a memória de cálculo das notas de saída ativas
    memoria_calculo = {}
    if active_docs_count > 0:
        ref_mem = None
        for doc in documentos_saida:
            if doc.cstat not in ("101", "110", "301", "302"):
                apuracao_nota = calculadora.calcular(doc)
                if "memoria_calculo" in apuracao_nota:
                    ref_mem = apuracao_nota["memoria_calculo"]
                    break
        
        if ref_mem:
            memoria_calculo = {**ref_mem}
            if empresa.regime_tributario == RegimeTributario.SIMPLES_NACIONAL:
                memoria_calculo["valor_com_iss_retido"] = Decimal("0.00")
                memoria_calculo["valor_com_st"] = Decimal("0.00")
                memoria_calculo["valor_sem_iss_retido"] = Decimal("0.00")
                memoria_calculo["valor_sem_st"] = Decimal("0.00")
                
                for doc in documentos_saida:
                    if doc.cstat in ("101", "110", "301", "302"):
                        continue
                    apur = calculadora.calcular(doc)
                    m = apur.get("memoria_calculo", {})
                    memoria_calculo["valor_com_iss_retido"] += Decimal(str(m.get("valor_com_iss_retido", 0)))
                    memoria_calculo["valor_com_st"] += Decimal(str(m.get("valor_com_st", 0)))
                    memoria_calculo["valor_sem_iss_retido"] += Decimal(str(m.get("valor_sem_iss_retido", 0)))
                    memoria_calculo["valor_sem_st"] += Decimal(str(m.get("valor_sem_st", 0)))
            elif empresa.regime_tributario == RegimeTributario.LUCRO_PRESUMIDO:
                memoria_calculo["pis"] = Decimal("0.00")
                memoria_calculo["cofins"] = Decimal("0.00")
                memoria_calculo["irpj"] = Decimal("0.00")
                memoria_calculo["csll"] = Decimal("0.00")
                memoria_calculo["iss"] = Decimal("0.00")
                memoria_calculo["valor_com_st"] = Decimal("0.00")
                memoria_calculo["valor_sem_st"] = Decimal("0.00")
                
                for doc in documentos_saida:
                    if doc.cstat in ("101", "110", "301", "302"):
                        continue
                    apur = calculadora.calcular(doc)
                    m = apur.get("memoria_calculo", {})
                    memoria_calculo["pis"] += Decimal(str(m.get("pis", 0)))
                    memoria_calculo["cofins"] += Decimal(str(m.get("cofins", 0)))
                    memoria_calculo["irpj"] += Decimal(str(m.get("irpj", 0)))
                    memoria_calculo["csll"] += Decimal(str(m.get("csll", 0)))
                    memoria_calculo["iss"] += Decimal(str(m.get("iss", 0)))
                    memoria_calculo["valor_com_st"] += Decimal(str(m.get("valor_com_st", 0)))
                    memoria_calculo["valor_sem_st"] += Decimal(str(m.get("valor_sem_st", 0)))

    return {
        "empresa_cnpj": empresa.cnpj,
        "empresa_razao_social": empresa.razao_social,
        "regime": empresa.regime_tributario.value,
        "quantidade_documentos": len(documentos),
        "quantidade_ativos": active_docs_count,
        "quantidade_cancelados": canceled_docs_count,
        "total_faturamento": total_faturamento,
        "total_imposto": total_imposto,
        "aliquota_efetiva_consolidada": aliquota_efetiva,
        "detalhes": detalhes,
        "memoria_calculo": memoria_calculo,
        "mes": mes,
        "ano": ano,
        "compras": {
            "total_compras": total_compras,
            "total_difal": total_difal,
            "total_icms_st": total_icms_st_compra,
            "quantidade_entradas": active_entradas_count
        }
    }


@app.delete("/documentos/{documento_id}", status_code=status.HTTP_200_OK)
def deletar_documento(documento_id: int, db: Session = Depends(get_db)):
    """Exclui um documento fiscal da Staging Area e seus ajustes associados em cascata."""
    doc = db.query(DocumentoFiscal).filter(DocumentoFiscal.id == documento_id).first()
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, 
            detail=f"Documento fiscal com ID {documento_id} não encontrado."
        )
    db.delete(doc)
    db.commit()
    return {"detail": "Documento fiscal apagado com sucesso!"}


@app.delete("/documentos", status_code=status.HTTP_200_OK)
def deletar_documentos_filtrados(
    empresa_id: int,
    mes: Optional[int] = None,
    ano: Optional[int] = None,
    db: Session = Depends(get_db)
):
    """Exclui documentos fiscais de uma empresa, opcionalmente filtrando por período (mês/ano)."""
    from sqlalchemy import extract
    query = db.query(DocumentoFiscal).filter(DocumentoFiscal.empresa_id == empresa_id)
    if mes:
        query = query.filter(extract('month', DocumentoFiscal.data_competencia) == mes)
    if ano:
        query = query.filter(extract('year', DocumentoFiscal.data_competencia) == ano)
    
    docs_to_delete = query.all()
    count = len(docs_to_delete)
    
    for doc in docs_to_delete:
        db.delete(doc)
        
    db.commit()
    return {"detail": f"{count} documentos fiscais excluídos com sucesso!"}


@app.post("/documentos/excluir-em-lote", status_code=status.HTTP_200_OK)
def excluir_documentos_lote(
    payload: dict,
    db: Session = Depends(get_db)
):
    """Exclui múltiplos documentos fiscais informados por uma lista de IDs."""
    ids = payload.get("ids", [])
    if not ids:
        return {"detail": "Nenhum ID fornecido para exclusão."}
        
    docs_to_delete = db.query(DocumentoFiscal).filter(DocumentoFiscal.id.in_(ids)).all()
    count = len(docs_to_delete)
    for doc in docs_to_delete:
        db.delete(doc)
    db.commit()
    return {"detail": f"{count} documentos fiscais excluídos com sucesso!"}


@app.post("/documentos/encerrar-em-lote", status_code=status.HTTP_200_OK)
def encerrar_documentos_lote(
    payload: dict,
    db: Session = Depends(get_db)
):
    """Encerra múltiplos documentos fiscais informados por uma lista de IDs."""
    ids = payload.get("ids", [])
    if not ids:
        return {"detail": "Nenhum ID fornecido para encerramento."}
        
    docs_to_close = db.query(DocumentoFiscal).filter(DocumentoFiscal.id.in_(ids)).all()
    count = 0
    for doc in docs_to_close:
        if doc.status_apuracao != StatusApuracao.ENCERRADO:
            doc.status_apuracao = StatusApuracao.ENCERRADO
            count += 1
    db.commit()
    return {"detail": f"{count} documentos fiscais encerrados com sucesso!"}


@app.post("/documentos/competencia-em-lote", status_code=status.HTTP_200_OK)
def editar_competencia_documentos_lote(
    payload: dict,
    db: Session = Depends(get_db)
):
    """Altera a data de competência de múltiplos documentos fiscais em lote."""
    ids = payload.get("ids", [])
    nova_competencia = payload.get("data_competencia")
    
    if not ids:
        raise HTTPException(status_code=400, detail="Nenhum ID fornecido.")
    if not nova_competencia:
        raise HTTPException(status_code=400, detail="Campo 'data_competencia' é obrigatório.")
        
    try:
        dt = None
        if len(nova_competencia) == 10:
            dt = datetime.fromisoformat(nova_competencia)
        elif len(nova_competencia) == 7:
            dt = datetime.strptime(nova_competencia, "%Y-%m")
        else:
            raise ValueError()
        if dt:
            dt = datetime(dt.year, dt.month, 1)
    except Exception:
        raise HTTPException(status_code=400, detail="Formato de data inválido. Use YYYY-MM ou YYYY-MM-DD.")
        
    docs_to_update = db.query(DocumentoFiscal).filter(DocumentoFiscal.id.in_(ids)).all()
    count = 0
    for doc in docs_to_update:
        if doc.status_apuracao != StatusApuracao.ENCERRADO:
            data_competencia_antiga = doc.data_competencia
            doc.data_competencia = dt
            count += 1
            # Move o arquivo físico de pasta caso a competência mude
            sincronizar_arquivo_xml(
                empresa_nome=doc.empresa.razao_social,
                data_competencia_antiga=data_competencia_antiga,
                data_competencia_nova=dt,
                tipo_operacao=doc.tipo_operacao,
                tipo_documento=doc.tipo_documento,
                chave_acesso=doc.chave_acesso
            )
            
    db.commit()
    return {"detail": f"Competência de {count} documentos fiscais atualizada com sucesso!"}


@app.put("/documentos/{documento_id}/competencia", response_model=DocumentoResponse)
def editar_competencia_documento(
    documento_id: int,
    payload: dict,
    db: Session = Depends(get_db)
):
    """Edita a data de competência (data_emissao) de um documento fiscal na Staging Area."""
    nova_competencia = payload.get("data_competencia") # Esperado YYYY-MM ou YYYY-MM-DD
    if not nova_competencia:
        raise HTTPException(status_code=400, detail="Campo 'data_competencia' é obrigatório.")
        
    doc = db.query(DocumentoFiscal).filter(DocumentoFiscal.id == documento_id).first()
    if not doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND, 
            detail=f"Documento fiscal com ID {documento_id} não encontrado."
        )
        
    if doc.status_apuracao == StatusApuracao.ENCERRADO:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, 
            detail="Não é possível editar a competência de um documento em período Encerrado."
        )
        
    try:
        dt = None
        if len(nova_competencia) == 10:
            dt = datetime.fromisoformat(nova_competencia)
        elif len(nova_competencia) == 7:
            dt = datetime.strptime(nova_competencia, "%Y-%m")
        else:
            raise ValueError()
        
        if dt:
            dt = datetime(dt.year, dt.month, 1)
        
        data_competencia_antiga = doc.data_competencia
        doc.data_competencia = dt
        db.commit()
        db.refresh(doc)
        
        # Move o arquivo físico de pasta caso a competência mude
        sincronizar_arquivo_xml(
            empresa_nome=doc.empresa.razao_social,
            data_competencia_antiga=data_competencia_antiga,
            data_competencia_nova=dt,
            tipo_operacao=doc.tipo_operacao,
            tipo_documento=doc.tipo_documento,
            chave_acesso=doc.chave_acesso
        )
        
        return doc
    except Exception:
        raise HTTPException(status_code=400, detail="Formato de data inválido. Use YYYY-MM ou YYYY-MM-DD.")


@app.post("/system/reset", status_code=status.HTTP_200_OK)
def resetar_banco(db: Session = Depends(get_db)):
    """Limpa completamente todas as tabelas do banco de dados (Ajustes, Documentos e Empresas)."""
    db.query(AjusteDocumento).delete()
    db.query(DocumentoFiscal).delete()
    db.query(Empresa).delete()
    db.commit()
    return {"detail": "Banco de dados totalmente resetado com sucesso!"}


# ==========================================
# ENDPOINTS DE CNAES
# ==========================================

@app.get("/api/cnaes")
def buscar_cnaes(q: Optional[str] = None):
    """
    Busca CNAEs por código ou descrição. Retorna uma lista formatada de correspondências.
    """
    try:
        if not CNAE_JSON_PATH.exists():
            return []
        with open(CNAE_JSON_PATH, "r", encoding="utf-8") as f:
            cnaes = json.load(f)
    except Exception:
        return []

    resultados = []
    if q:
        q_lower = q.lower().strip()
        for code, info in cnaes.items():
            if q_lower in code or q_lower in info.get("descricao", "").lower():
                resultados.append({
                    "codigo": code,
                    "descricao": info.get("descricao"),
                    "anexo": info.get("anexo"),
                    "fator_r": info.get("fator_r"),
                    "aliquota": info.get("aliquota")
                })
    else:
        for code, info in cnaes.items():
            resultados.append({
                "codigo": code,
                "descricao": info.get("descricao"),
                "anexo": info.get("anexo"),
                "fator_r": info.get("fator_r"),
                "aliquota": info.get("aliquota")
            })
    return resultados[:50] # Limita a 50 resultados para manter a performance


@app.post("/api/cnaes/sync")
def sincronizar_cnaes():
    """
    Sincroniza os CNAEs a partir do IBGE e atualiza o arquivo local.
    """
    try:
        total = sync_cnaes_from_ibge()
        return {"detail": "Sincronização concluída com sucesso!", "total": total}
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Erro ao sincronizar CNAEs: {str(e)}"
        )

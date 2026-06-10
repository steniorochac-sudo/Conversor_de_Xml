from decimal import Decimal
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
from fiscal_workflow.db.database import engine, get_db
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

app = FastAPI(
    title="Workflow Modular Fiscal",
    description="API de ingestão de XMLs, Staging Area e motor de apuração fiscal",
    version="1.0.0"
)

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
        cnae=empresa.cnae
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
    
    db.commit()
    db.refresh(empresa)
    return empresa

# ==========================================
# ENDPOINTS DO FLUXO DE DOCUMENTOS
# ==========================================

@app.post("/documentos/upload", response_model=List[DocumentoResponse], status_code=status.HTTP_201_CREATED)
async def upload_xml(
    empresa_id: Optional[int] = Form(None, description="[DEPRECATED] Ignorado pelo backend. O emitente sempre será autodetectado e cadastrado a partir do XML."),
    files: List[UploadFile] = File(..., description="Arquivos XML de NF-e / NFC-e / NFS-e (lote)"),
    db: Session = Depends(get_db)
):
    """
    Recebe múltiplos arquivos XML em lote, realiza o parse de cada um,
    autodetecta/autocadastra emitentes e salva na Staging Area. Pula duplicados de forma resiliente.
    """
    documentos_salvos = []
    
    for file in files:
        # Pula arquivos vazios ou não XML
        if not file.filename.lower().endswith(".xml"):
            continue
            
        try:
            # 1. Lê e decodifica o arquivo XML enviado
            xml_content = await file.read()
            lista_notas = parse_documento_fiscal(xml_content)
        except Exception:
            # Se um arquivo do lote falhar, apenas o pulamos para não quebrar a importação inteira
            continue

        for dados_nota in lista_notas:
            # 2. Valida se a chave de acesso já foi importada
            doc_existente = db.query(DocumentoFiscal).filter(
                DocumentoFiscal.chave_acesso == dados_nota["chave_acesso"]
            ).first()
            
            if doc_existente:
                # Se a nota já existe, verifica se o novo XML traz uma situação de cancelamento/atualização
                nova_cstat = dados_nota.get("cstat", "100")
                if doc_existente.cstat != nova_cstat:
                    doc_existente.cstat = nova_cstat
                    db.commit()
                    db.refresh(doc_existente)
                documentos_salvos.append(doc_existente)
                continue

            # 3. Resolução do Emitente (Empresa) - Sempre via Autodetecção pelo CNPJ do XML
            cnpj_emitente = dados_nota["emitente_cnpj"]
            if not cnpj_emitente:
                continue # Pula se o XML não tiver dados do emitente
                
            empresa = db.query(Empresa).filter(Empresa.cnpj == cnpj_emitente).first()
            
            # Se a empresa não existir, autocadastra
            if not empresa:
                crt = dados_nota.get("emitente_crt")
                if crt in ("1", "2"):
                    regime = RegimeTributario.SIMPLES_NACIONAL
                elif crt == "3":
                    regime = RegimeTributario.LUCRO_PRESUMIDO
                else:
                    regime = RegimeTributario.SIMPLES_NACIONAL

                # Busca o CNAE oficial via API externa de CNPJ
                from fiscal_workflow.services.cnpj_client import buscar_cnae_oficial
                cnae_resolvido = buscar_cnae_oficial(cnpj_emitente)
                
                sujeito_fator_r = False
                categoria_simples = "Serviços (Anexo III)"
                
                if cnae_resolvido:
                    # Tenta carregar as regras fiscais do CNAE local
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
                                # Aplica heurísticas de fallback caso o CNAE não esteja catalogado
                                if cnae_resolvido[:2] in ("45", "46", "47"):
                                    categoria_simples = "Comércio (Anexo I)"
                                elif cnae_resolvido[:2] in ("62", "86", "73", "74", "69"):
                                    sujeito_fator_r = True
                    except Exception:
                        pass

                empresa = Empresa(
                    cnpj=cnpj_emitente,
                    razao_social=dados_nota["emitente_razao_social"] or f"Empresa CNPJ {cnpj_emitente}",
                    regime_tributario=regime,
                    rbt12=Decimal("0.00"),
                    folha12=Decimal("0.00"),
                    sujeito_fator_r=sujeito_fator_r,
                    categoria_simples=categoria_simples,
                    cnae=cnae_resolvido
                )
                db.add(empresa)
                db.commit()
                db.refresh(empresa)

            # 4. Salva a nota fiscal na Staging Area
            novo_doc = DocumentoFiscal(
                empresa_id=empresa.id,
                chave_acesso=dados_nota["chave_acesso"],
                tipo_documento=dados_nota["tipo_documento"],
                tipo_operacao=dados_nota.get("tipo_operacao", "Saída"),
                valor_total=Decimal(str(dados_nota["valor_total"])),
                status_apuracao=StatusApuracao.PENDENTE,
                cstat=dados_nota.get("cstat", "100"),
                itens=dados_nota["itens"]
            )

            db.add(novo_doc)
            db.commit()
            db.refresh(novo_doc)
            documentos_salvos.append(novo_doc)

    return documentos_salvos


@app.get("/documentos", response_model=List[DocumentoResponse])
def listar_documentos(
    empresa_id: Optional[int] = None, 
    status: Optional[StatusApuracao] = None,
    db: Session = Depends(get_db)
):
    """Lista todos os documentos importados (Staging Area) com filtros opcionais."""
    query = db.query(DocumentoFiscal)
    if empresa_id:
        query = query.filter(DocumentoFiscal.empresa_id == empresa_id)
    if status:
        query = query.filter(DocumentoFiscal.status_apuracao == status)
    
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
def apurar_consolidado_empresa(empresa_id: int, db: Session = Depends(get_db)):
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
    documentos = db.query(DocumentoFiscal).filter(DocumentoFiscal.empresa_id == empresa_id).all()

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

    # 4. Apura cada nota e consolida
    for doc in documentos:
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

    # Alíquota efetiva consolidada
    aliquota_efetiva = (total_imposto / total_faturamento) if total_faturamento > 0 else Decimal("0.00")

    # Consolida a memória de cálculo das notas ativas
    memoria_calculo = {}
    if active_docs_count > 0:
        ref_mem = None
        for doc in documentos:
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
                
                for doc in documentos:
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
                
                for doc in documentos:
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
        "memoria_calculo": memoria_calculo
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

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from fiscal_workflow.core.config import settings

# Ajustes especiais de conexão para SQLite (evita erros de multi-thread no FastAPI)
connect_args = {}
if settings.DATABASE_URL.startswith("sqlite"):
    connect_args["check_same_thread"] = False

# Criação do motor do banco de dados (engine)
engine = create_engine(
    settings.DATABASE_URL, 
    connect_args=connect_args
)

# Pool de sessões de banco de dados
SessionLocal = sessionmaker(
    autocommit=False, 
    autoflush=False, 
    bind=engine
)

def get_db():
    """
    Dependency do FastAPI para prover sessões de banco de dados
    com ciclo de vida gerenciado automaticamente (abre, serve a requisição, fecha).
    """
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

import os
from pathlib import Path
from dotenv import load_dotenv

# Encontra o diretório base do projeto para localizar o arquivo .env
BASE_DIR = Path(__file__).resolve().parent.parent.parent
ENV_PATH = BASE_DIR / ".env"

# Carrega o .env se existir
if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)

class Settings:
    PROJECT_NAME: str = "Workflow Modular Fiscal"
    
    # URL de Conexão:
    # 1. Se estiver configurada no .env (ex: Neon.tech), utiliza ela.
    # 2. Caso contrário, cai de volta (fallback) no SQLite local na raiz do projeto.
    DATABASE_URL: str = os.getenv(
        "DATABASE_URL",
        f"sqlite:///{BASE_DIR}/fiscal_workflow.db"
    )

settings = Settings()

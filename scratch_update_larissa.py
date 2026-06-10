from sqlalchemy import create_engine, text
from fiscal_workflow.core.config import settings

engine = create_engine(settings.DATABASE_URL)
with engine.connect() as conn:
    # Atualiza a empresa
    res = conn.execute(
        text("UPDATE empresas SET categoria_simples = :cat WHERE cnpj = :cnpj"),
        {"cat": "Serviços (Anexo IV)", "cnpj": "47754125000170"}
    )
    conn.commit()
    print("Atualização concluída com sucesso.")

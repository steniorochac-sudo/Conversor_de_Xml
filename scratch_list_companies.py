import os
from sqlalchemy import create_engine, text
from fiscal_workflow.core.config import settings

engine = create_engine(settings.DATABASE_URL)
with engine.connect() as conn:
    res = conn.execute(text("SELECT id, cnpj, razao_social, regime_tributario, categoria_simples, rbt12 FROM empresas")).fetchall()
    print("Empresas cadastradas:")
    for row in res:
        print(f"ID: {row[0]} | CNPJ: {row[1]} | Razão: {row[2]} | Regime: {row[3]} | Categoria: {row[4]} | RBT12: {row[5]}")

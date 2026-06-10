import os
from sqlalchemy import create_engine, text
from fiscal_workflow.core.config import settings
from fiscal_workflow.services.parsers import parse_documento_fiscal

file_path = "Notas de exemplo/NotasFiscais mohara.xml"

try:
    with open(file_path, "rb") as f:
        content = f.read()
    
    # Parse das notas
    notas = parse_documento_fiscal(content)
    print(f"Total de notas parsed: {len(notas)}")
    for i, nota in enumerate(notas):
        print(f"Nota {i+1}:")
        print(f"  Chave: {nota['chave_acesso']}")
        print(f"  Numero: {nota['numero_nf']}")
        print(f"  Emit CNPJ: {nota['emitente_cnpj']}")
        print(f"  Emit Razao: {nota['emitente_razao_social']}")
        print(f"  Emit CRT: {nota['emitente_crt']}")
        print(f"  Dest CNPJ: {nota.get('destinatario_cnpj')}")
        print(f"  Dest Nome: {nota.get('destinatario_nome')}")
        
    # Conecta ao banco para ver se a empresa existe
    engine = create_engine(settings.DATABASE_URL)
    with engine.connect() as conn:
        cnpj_list = [n['emitente_cnpj'] for n in notas]
        for cnpj in set(cnpj_list):
            res = conn.execute(text("SELECT id, cnpj, razao_social FROM empresas WHERE cnpj = :cnpj"), {"cnpj": cnpj}).fetchone()
            print(f"Empresa no banco para CNPJ {cnpj}: {res}")
            
            # Conta se existem documentos associados a esta empresa
            if res:
                doc_count = conn.execute(text("SELECT COUNT(*) FROM documentos_fiscais WHERE empresa_id = :id"), {"id": res[0]}).scalar()
                print(f"  Documentos da empresa no banco: {doc_count}")

except Exception as e:
    print(f"Erro: {e}")

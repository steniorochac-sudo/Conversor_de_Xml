import unittest
from decimal import Decimal
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.exc import IntegrityError

# Importa os modelos do nosso pacote recém-criado
from fiscal_workflow.models import Base, Empresa, DocumentoFiscal, RegimeTributario, StatusApuracao

class TestFiscalModels(unittest.TestCase):
    def setUp(self):
        """Configura um banco de dados SQLite em memória para cada caso de teste."""
        self.engine = create_engine("sqlite:///:memory:")
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.session = self.Session()

    def tearDown(self):
        """Fecha a sessão e limpa as tabelas ao fim do teste."""
        self.session.close()
        Base.metadata.drop_all(self.engine)

    def test_criar_empresa_e_documento_com_sucesso(self):
        """Valida a inserção e integridade de relacionamento de Empresa e Documento."""
        # 1. Cria a Empresa com Enum estrito
        empresa = Empresa(
            cnpj="12345678000199",
            razao_social="Stenio Software Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL
        )
        self.session.add(empresa)
        self.session.commit()

        self.assertIsNotNone(empresa.id)
        self.assertEqual(empresa.regime_tributario, RegimeTributario.SIMPLES_NACIONAL)

        # 2. Cria o Documento Fiscal associado
        doc = DocumentoFiscal(
            empresa_id=empresa.id,
            chave_acesso="35230512345678000199550010000001231234567890",
            tipo_documento="NF-e",
            valor_total=Decimal("1500.50"),
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc)
        self.session.commit()

        self.assertIsNotNone(doc.id)
        self.assertEqual(doc.status_apuracao, StatusApuracao.PENDENTE)
        self.assertEqual(doc.valor_total, Decimal("1500.50"))

        # 3. Testa os relacionamentos bidirecionais
        # Empresa -> Documento
        self.assertEqual(len(empresa.documentos), 1)
        self.assertEqual(empresa.documentos[0].chave_acesso, doc.chave_acesso)

        # Documento -> Empresa
        self.assertEqual(doc.empresa.razao_social, "Stenio Software Ltda")

    def test_cnpj_unico(self):
        """Valida que o banco rejeita CNPJs duplicados."""
        emp1 = Empresa(
            cnpj="11111111000111",
            razao_social="Empresa Um",
            regime_tributario=RegimeTributario.LUCRO_PRESUMIDO
        )
        emp2 = Empresa(
            cnpj="11111111000111",  # CNPJ Duplicado
            razao_social="Empresa Dois",
            regime_tributario=RegimeTributario.LUCRO_REAL
        )
        
        self.session.add(emp1)
        self.session.commit()
        
        self.session.add(emp2)
        with self.assertRaises(IntegrityError):
            self.session.commit()

    def test_chave_acesso_unica(self):
        """Valida que o banco de dados impede chaves de acesso duplicadas."""
        emp = Empresa(
            cnpj="22222222000122",
            razao_social="Empresa Teste Unicidade",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL
        )
        self.session.add(emp)
        self.session.commit()

        doc1 = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="44444444444444444444444444444444444444444444",
            tipo_documento="NF-e",
            valor_total=Decimal("100.00")
        )
        doc2 = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="44444444444444444444444444444444444444444444",  # Chave duplicada
            tipo_documento="NF-e",
            valor_total=Decimal("200.00")
        )

        self.session.add(doc1)
        self.session.commit()

        self.session.add(doc2)
        with self.assertRaises(IntegrityError):
            self.session.commit()

    def test_cascata_delecao_empresa(self):
        """Valida que ao excluir uma empresa, todos os seus documentos fiscais são removidos automaticamente."""
        emp = Empresa(
            cnpj="33333333000133",
            razao_social="Empresa a Deletar",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL
        )
        self.session.add(emp)
        self.session.commit()

        doc = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="55555555555555555555555555555555555555555555",
            tipo_documento="NFC-e",
            valor_total=Decimal("50.00")
        )
        self.session.add(doc)
        self.session.commit()

        # Garante que está no banco
        self.assertEqual(self.session.query(DocumentoFiscal).count(), 1)

        # Deleta a empresa
        self.session.delete(emp)
        self.session.commit()

        # Verifica se o documento também sumiu (cascading delete)
        self.assertEqual(self.session.query(DocumentoFiscal).count(), 0)

if __name__ == "__main__":
    unittest.main()

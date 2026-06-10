import io
import unittest
from decimal import Decimal
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.pool import StaticPool
from sqlalchemy.orm import sessionmaker

# Importa o app FastAPI e as dependências
from fiscal_workflow.main import app
from fiscal_workflow.db.database import get_db
from fiscal_workflow.models.models import Base, StatusApuracao
from fiscal_workflow.tests.test_parsers import MOCK_NFE_XML

# 1. Configura um banco de dados SQLite isolado em memória com Pool Estático para os testes de API
engine_test = create_engine(
    "sqlite:///:memory:", 
    connect_args={"check_same_thread": False},
    poolclass=StaticPool
)
SessionTest = sessionmaker(autocommit=False, autoflush=False, bind=engine_test)

def override_get_db():
    """Sobrescreve a dependência get_db do FastAPI para usar o banco de testes."""
    db = SessionTest()
    try:
        yield db
    finally:
        db.close()

# Injeta a dependência no FastAPI
app.dependency_overrides[get_db] = override_get_db

class TestFiscalAPI(unittest.TestCase):
    def setUp(self):
        """Inicializa as tabelas do banco de testes antes de cada caso de teste."""
        Base.metadata.create_all(bind=engine_test)
        self.client = TestClient(app)

    def tearDown(self):
        """Limpa as tabelas ao fim do caso de teste."""
        Base.metadata.drop_all(bind=engine_test)

    def test_endpoint_raiz(self):
        """Verifica se o endpoint inicial responde servindo o Dashboard HTML com sucesso."""
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn("<!DOCTYPE html>", response.text)
        self.assertIn("Antigravity Fiscal", response.text)

    def test_fluxo_completo_api(self):
        """Testa o fluxo completo da API: Cadastro de Empresa -> Upload de Nota -> Ajustes -> Fechamento."""
        
        # 1. Cadastra uma nova Empresa
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio Software Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        empresa_data = response_empresa.json()
        empresa_id = empresa_data["id"]
        self.assertEqual(empresa_data["cnpj"], "12345678000199")

        # 2. Upload de XML da Nota Fiscal (Simula múltiplos arquivos enviados via Multipart Form)
        xml_file = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        response_upload = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe_teste.xml", xml_file, "text/xml"))]
        )
        self.assertEqual(response_upload.status_code, 201)
        doc_list = response_upload.json()
        self.assertEqual(len(doc_list), 1)
        doc_data = doc_list[0]
        doc_id = doc_data["id"]
        
        # Garante que os campos brutos e enums estão corretos
        self.assertEqual(doc_data["chave_acesso"], "35230512345678000199550010000001231234567890")
        self.assertEqual(doc_data["tipo_documento"], "NF-e")
        self.assertEqual(doc_data["status_apuracao"], "Pendente")
        self.assertEqual(float(doc_data["valor_total"]), 4350.00)
        self.assertEqual(float(doc_data["valor_final"]), 4350.00) # Inicialmente sem ajustes
        self.assertEqual(len(doc_data["itens"]), 1)

        # 3. Consulta a Staging Area (GET /documentos)
        response_list = self.client.get("/documentos")
        self.assertEqual(response_list.status_code, 200)
        self.assertEqual(len(response_list.json()), 1)
        self.assertEqual(response_list.json()[0]["id"], doc_id)

        # 4. Registra um Ajuste Manual de Auditoria (Acréscimo de +100.00)
        response_ajuste = self.client.post(
            f"/documentos/{doc_id}/ajustes",
            json={
                "valor_total_ajuste": 100.00,
                "justificativa": "Correção de glosa de frete acordada com o cliente",
                "usuario": "stenio_auditor"
            }
        )
        self.assertEqual(response_ajuste.status_code, 200)
        doc_ajustado_data = response_ajuste.json()

        # O status deve transicionar para "Em Revisão"
        self.assertEqual(doc_ajustado_data["status_apuracao"], "Em Revisão")
        # O valor final deve somar dinamicamente o valor original + ajuste (4350.00 + 100.00 = 4450.00)
        self.assertEqual(float(doc_ajustado_data["valor_final"]), 4450.00)
        
        # Verifica se o log de auditoria está anexado no retorno
        self.assertEqual(len(doc_ajustado_data["ajustes"]), 1)
        self.assertEqual(doc_ajustado_data["ajustes"][0]["usuario"], "stenio_auditor")
        self.assertEqual(float(doc_ajustado_data["ajustes"][0]["valor_total_ajuste"]), 100.00)

        # 5. Consolida e Encerra o Período Fiscal (POST /documentos/{id}/encerrar)
        response_encerrar = self.client.post(f"/documentos/{doc_id}/encerrar")
        self.assertEqual(response_encerrar.status_code, 200)
        self.assertEqual(response_encerrar.json()["status_apuracao"], "Encerrado")

        # 6. Valida o Snapshot e Bloqueio: Novos ajustes devem ser sumariamente REJEITADOS
        response_bloqueado = self.client.post(
            f"/documentos/{doc_id}/ajustes",
            json={
                "valor_total_ajuste": -50.00,
                "justificativa": "Tentativa de ajuste após fechamento",
                "usuario": "stenio_auditor"
            }
        )
        # Deve retornar erro 400 Bad Request
        self.assertEqual(response_bloqueado.status_code, 400)
        self.assertIn("já está Encerrado", response_bloqueado.json()["detail"])

    def test_upload_multiple_files_batch(self):
        """Testa o upload de múltiplos arquivos XML em lote com autodetecção e autocadastro."""
        # Cria um segundo mock XML com CNPJ e chave de acesso diferentes
        mock_nfe_xml_2 = MOCK_NFE_XML.replace(
            "35230512345678000199550010000001231234567890", "35230599999999000199550010000001241234567890"
        ).replace(
            "12345678000199", "99999999000199"
        ).replace(
            "Stenio Software Ltda", "Stenio Software Filial"
        ).replace(
            "<CRT>1</CRT>", "<CRT>3</CRT>"
        ).replace(
            "<nNF>123</nNF>", "<nNF>124</nNF>"
        )

        xml_file_1 = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        xml_file_2 = io.BytesIO(mock_nfe_xml_2.encode('utf-8'))

        # Envia ambos os arquivos no mesmo request, sem passar empresa_id (autodetectar ambos!)
        response = self.client.post(
            "/documentos/upload",
            files=[
                ("files", ("nfe_1.xml", xml_file_1, "text/xml")),
                ("files", ("nfe_2.xml", xml_file_2, "text/xml"))
            ]
        )
        self.assertEqual(response.status_code, 201)
        doc_list = response.json()
        
        # Devem ter sido retornados exatamente 2 documentos fiscais
        self.assertEqual(len(doc_list), 2)
        
        # Verifica se as duas empresas foram autocadastradas com seus respectivos regimes tributários
        response_empresas = self.client.get("/empresas")
        self.assertEqual(response_empresas.status_code, 200)
        empresas = response_empresas.json()
        
        # Deve haver 2 empresas cadastradas no total
        self.assertEqual(len(empresas), 2)
        
        empresa_1 = next(e for e in empresas if e["cnpj"] == "12345678000199")
        empresa_2 = next(e for e in empresas if e["cnpj"] == "99999999000199")
        
        # Empresa 1 deve ser Simples Nacional (CRT 1)
        self.assertEqual(empresa_1["regime_tributario"], "Simples Nacional")
        # Empresa 2 deve ser Lucro Presumido (CRT 3)
        self.assertEqual(empresa_2["regime_tributario"], "Lucro Presumido")

    def test_consolidated_taxes_calculation(self):
        """Testa o cálculo de impostos consolidado de uma empresa no Simples Nacional."""
        # 1. Cadastra uma nova Empresa
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "11111111000199",
                "razao_social": "Empresa Teste Consolidado Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        empresa_id = response_empresa.json()["id"]

        # 2. Faz o upload de 2 notas para esta empresa
        # XML 1
        xml_file_1 = io.BytesIO(MOCK_NFE_XML.replace(
            "12345678000199", "11111111000199"
        ).encode('utf-8'))
        
        # XML 2 (Valor 4350.00, NF 124)
        mock_nfe_xml_2 = MOCK_NFE_XML.replace(
            "35230512345678000199550010000001231234567890", "35230511111111000199550010000001241234567890"
        ).replace(
            "12345678000199", "11111111000199"
        ).replace(
            "<nNF>123</nNF>", "<nNF>124</nNF>"
        )
        xml_file_2 = io.BytesIO(mock_nfe_xml_2.encode('utf-8'))

        response_upload = self.client.post(
            "/documentos/upload",
            files=[
                ("files", ("nfe_1.xml", xml_file_1, "text/xml")),
                ("files", ("nfe_2.xml", xml_file_2, "text/xml"))
            ]
        )
        self.assertEqual(response_upload.status_code, 201)

        # 3. Chama o endpoint consolidado
        response_consolidado = self.client.get(f"/empresas/{empresa_id}/consolidado")
        self.assertEqual(response_consolidado.status_code, 200)
        data = response_consolidado.json()

        # Faturamento total: 4350.00 * 2 = 8700.00
        self.assertEqual(float(data["total_faturamento"]), 8700.00)
        # Imposto calculado (Simples Nacional 6%): 8700.00 * 0.06 = 522.00
        self.assertEqual(float(data["total_imposto"]), 522.00)
        self.assertEqual(float(data["aliquota_efetiva_consolidada"]), 0.06)
        self.assertEqual(float(data["detalhes"]["das"]), 522.00)
        
        # Valida a presença e chaves da memoria_calculo
        self.assertIn("memoria_calculo", data)
        mc = data["memoria_calculo"]
        self.assertEqual(float(mc["aliq_efetiva"]), 6.00)
        self.assertEqual(mc["categoria_simples"], "Serviços (Anexo III)")

    def test_editar_empresa_endpoint(self):
        """Testa o endpoint PUT /empresas/{id} de edição de empresa."""
        # 1. Cadastra uma nova Empresa no regime Lucro Presumido
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "22222222000199",
                "razao_social": "JJ Weiss S/A",
                "regime_tributario": "Lucro Presumido"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        empresa_id = response_empresa.json()["id"]

        # 2. Edita o regime tributário para Simples Nacional e altera a Razão Social
        response_update = self.client.put(
            f"/empresas/{empresa_id}",
            json={
                "razao_social": "JJ Weiss Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        self.assertEqual(response_update.status_code, 200)
        data = response_update.json()
        
        self.assertEqual(data["id"], empresa_id)
        self.assertEqual(data["cnpj"], "22222222000199")
        self.assertEqual(data["razao_social"], "JJ Weiss Ltda")
        self.assertEqual(data["regime_tributario"], "Simples Nacional")

    def test_cancelamento_nfe_fluxo_completo(self):
        """Valida o fluxo completo de cancelamento de NF-e na API (Parser, Upsert, Apuração e Consolidado)."""
        # 1. Cadastra uma nova Empresa
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "55555555000199",
                "razao_social": "Empresa Cancelamento Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        empresa_id = response_empresa.json()["id"]

        # 2. Faz o upload da nota fiscal normal (Autorizada, cStat 100)
        xml_file_1 = io.BytesIO(MOCK_NFE_XML.replace(
            "12345678000199", "55555555000199"
        ).encode('utf-8'))
        
        response_upload_1 = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe_normal.xml", xml_file_1, "text/xml"))]
        )
        self.assertEqual(response_upload_1.status_code, 201)
        doc_data_1 = response_upload_1.json()[0]
        self.assertEqual(doc_data_1["cstat"], "100")
        doc_id = doc_data_1["id"]

        # 3. Faz o upload do mesmo XML, porém com o protocolo de cancelamento (cStat 101) para testar o Upsert Reativo
        xml_cancelada = MOCK_NFE_XML.replace(
            "12345678000199", "55555555000199"
        ).replace(
            "</infNFe>\n    </NFe>", 
            "</infNFe>\n    </NFe>\n    <protNFe>\n        <infProt>\n            <cStat>101</cStat>\n        </infProt>\n    </protNFe>"
        )
        xml_file_2 = io.BytesIO(xml_cancelada.encode('utf-8'))
        
        response_upload_2 = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe_cancelada.xml", xml_file_2, "text/xml"))]
        )
        self.assertEqual(response_upload_2.status_code, 201)
        doc_data_2 = response_upload_2.json()[0]
        
        # Garante que o status da nota existente no banco foi atualizado reativamente para "101"
        self.assertEqual(doc_data_2["id"], doc_id)
        self.assertEqual(doc_data_2["cstat"], "101")

        # 4. Testa a apuração individual da nota cancelada (Deve retornar impostos e bases zerados)
        response_apuracao = self.client.get(f"/documentos/{doc_id}/apurar")
        self.assertEqual(response_apuracao.status_code, 200)
        apuracao = response_apuracao.json()
        self.assertEqual(float(apuracao["valor_final_base"]), 0.00)
        self.assertEqual(float(apuracao["imposto_calculado"]), 0.00)
        self.assertIn("CANCELADA", apuracao["mensagem"])

        # 5. Testa o faturamento consolidado da empresa (Deve ignorar totalmente a nota cancelada)
        response_consolidado = self.client.get(f"/empresas/{empresa_id}/consolidado")
        self.assertEqual(response_consolidado.status_code, 200)
        consolidado = response_consolidado.json()
        self.assertEqual(float(consolidado["total_faturamento"]), 0.00)
        self.assertEqual(float(consolidado["total_imposto"]), 0.00)
        self.assertEqual(consolidado["quantidade_ativos"], 0)
        self.assertEqual(consolidado["quantidade_cancelados"], 1)

    def test_deletar_documento_individual(self):
        """Valida a exclusão individual de uma nota fiscal na staging area e seus efeitos."""
        # 1. Cadastra uma nova Empresa e uma nota
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "66666666000199",
                "razao_social": "Empresa Delete Individual Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        empresa_id = response_empresa.json()["id"]

        xml_file = io.BytesIO(MOCK_NFE_XML.replace("12345678000199", "66666666000199").encode('utf-8'))
        response_upload = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe.xml", xml_file, "text/xml"))]
        )
        doc_id = response_upload.json()[0]["id"]

        # Adiciona um ajuste
        self.client.post(
            f"/documentos/{doc_id}/ajustes",
            json={
                "valor_total_ajuste": 50.00,
                "justificativa": "Ajuste de teste",
                "usuario": "teste"
            }
        )

        # 2. Deleta a nota fiscal
        response_delete = self.client.delete(f"/documentos/{doc_id}")
        self.assertEqual(response_delete.status_code, 200)

        # 3. Verifica se a nota sumiu da busca
        response_list = self.client.get("/documentos")
        self.assertEqual(len(response_list.json()), 0)

    def test_reset_sistema_global(self):
        """Valida que o endpoint POST /system/reset limpa completamente todas as tabelas."""
        # 1. Cadastra dados
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "77777777000199",
                "razao_social": "Empresa Reset Global Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        empresa_id = response_empresa.json()["id"]

        xml_file = io.BytesIO(MOCK_NFE_XML.replace("12345678000199", "77777777000199").encode('utf-8'))
        self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe.xml", xml_file, "text/xml"))]
        )

        # 2. Roda o reset global
        response_reset = self.client.post("/system/reset")
        self.assertEqual(response_reset.status_code, 200)

        # 3. Valida que tudo está limpo
        response_empresas = self.client.get("/empresas")
        self.assertEqual(len(response_empresas.json()), 0)

        response_documentos = self.client.get("/documentos")
        self.assertEqual(len(response_documentos.json()), 0)

    def test_cnaes_api_search(self):
        """Testa a listagem e busca filtrada de CNAEs."""
        # 1. Testa a busca geral (sem filtro)
        response_all = self.client.get("/api/cnaes")
        self.assertEqual(response_all.status_code, 200)
        cnaes = response_all.json()
        self.assertTrue(len(cnaes) > 0)
        
        # 2. Testa a busca filtrada por código existente (ex: '6201501')
        response_filtered = self.client.get("/api/cnaes?q=6201501")
        self.assertEqual(response_filtered.status_code, 200)
        filtered_cnaes = response_filtered.json()
        self.assertEqual(len(filtered_cnaes), 1)
        self.assertEqual(filtered_cnaes[0]["codigo"], "6201501")
        self.assertEqual(filtered_cnaes[0]["anexo"], "III")
        self.assertTrue(filtered_cnaes[0]["fator_r"])
        
        # 3. Testa busca por termo de descrição
        response_desc = self.client.get("/api/cnaes?q=programa")
        self.assertEqual(response_desc.status_code, 200)
        desc_cnaes = response_desc.json()
        self.assertTrue(len(desc_cnaes) > 0)
        
        # 4. Testa busca com termo inexistente
        response_none = self.client.get("/api/cnaes?q=cnaeinexistente123")
        self.assertEqual(response_none.status_code, 200)
        self.assertEqual(len(response_none.json()), 0)

    def test_autocadastro_com_cnae_cnpj_api(self):
        """Verifica se o autocadastro do emitente consulta a API de CNPJ e preenche as regras fiscais do CNAE."""
        from unittest.mock import patch
        
        # 1. Configura mock para retornar CNAE de TI
        with patch("fiscal_workflow.services.cnpj_client.buscar_cnae_oficial") as mock_buscar_cnae:
            mock_buscar_cnae.return_value = "6201501"
            
            # 2. Upload de nota fiscal com emitente não cadastrado
            xml_file = io.BytesIO(MOCK_NFE_XML.replace("12345678000199", "99999999000199").encode('utf-8'))
            response = self.client.post(
                "/documentos/upload",
                files=[("files", ("nfe_teste_cnpj.xml", xml_file, "text/xml"))]
            )
            self.assertEqual(response.status_code, 201)
            
            # 3. Consulta a empresa criada para verificar se o CNAE e os enquadramentos foram importados
            response_empresas = self.client.get("/empresas")
            empresas = response_empresas.json()
            
            empresa_ti = next((emp for emp in empresas if emp["cnpj"] == "99999999000199"), None)
            self.assertIsNotNone(empresa_ti)
            self.assertEqual(empresa_ti["cnae"], "6201501")
            self.assertEqual(empresa_ti["categoria_simples"], "Serviços (Anexo III)")
            self.assertTrue(empresa_ti["sujeito_fator_r"])

if __name__ == "__main__":
    unittest.main()

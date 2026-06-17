import io
import unittest
from decimal import Decimal
from datetime import datetime
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
        self.assertEqual(doc_data["numero_nf"], "123")
        self.assertEqual(doc_data["emitente_nome"], "Stenio Software Ltda")
        self.assertEqual(doc_data["destinatario_nome"], "Cliente Exemplo SA")
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
        # Imposto calculado (Simples Nacional 4%): 8700.00 * 0.04 = 348.00
        self.assertEqual(float(data["total_imposto"]), 348.00)
        self.assertEqual(float(data["aliquota_efetiva_consolidada"]), 0.04)
        self.assertEqual(float(data["detalhes"]["das"]), 348.00)
        
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

    def test_period_filtering(self):
        """Testa o filtro de período (mês e ano) nas rotas de listagem e consolidação."""
        # 1. Cadastra uma Empresa
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio Software Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        emp_id = response_empresa.json()["id"]

        # 2. Insere documentos fiscais diretamente no banco com datas diferentes para testar o filtro
        from datetime import datetime
        db = next(override_get_db())
        from fiscal_workflow.models.models import DocumentoFiscal
        
        doc1 = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="35230512345678000199550010000001231234567891",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            data_emissao=datetime(2026, 5, 15),
            data_competencia=datetime(2026, 5, 1)
        )
        doc2 = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="35230512345678000199550010000001231234567892",
            tipo_documento="NF-e",
            valor_total=Decimal("2000.00"),
            data_emissao=datetime(2026, 6, 20),
            data_competencia=datetime(2026, 6, 1)
        )
        db.add(doc1)
        db.add(doc2)
        db.commit()

        # 3. Testa listagem filtrando por Maio/2026
        resp_list = self.client.get(f"/documentos?empresa_id={emp_id}&mes=5&ano=2026")
        self.assertEqual(resp_list.status_code, 200)
        docs = resp_list.json()
        self.assertEqual(len(docs), 1)
        self.assertEqual(float(docs[0]["valor_total"]), 1000.00)

        # 4. Testa listagem filtrando por Junho/2026
        resp_list2 = self.client.get(f"/documentos?empresa_id={emp_id}&mes=6&ano=2026")
        self.assertEqual(resp_list2.status_code, 200)
        docs2 = resp_list2.json()
        self.assertEqual(len(docs2), 1)
        self.assertEqual(float(docs2[0]["valor_total"]), 2000.00)

        # 5. Testa consolidação de impostos filtrando por Maio/2026
        resp_cons = self.client.get(f"/empresas/{emp_id}/consolidado?mes=5&ano=2026")
        self.assertEqual(resp_cons.status_code, 200)
        cons = resp_cons.json()
        self.assertEqual(float(cons["total_faturamento"]), 1000.00)

        # 6. Testa consolidação de impostos filtrando por Junho/2026
        resp_cons2 = self.client.get(f"/empresas/{emp_id}/consolidado?mes=6&ano=2026")
        self.assertEqual(resp_cons2.status_code, 200)
        cons2 = resp_cons2.json()
        self.assertEqual(float(cons2["total_faturamento"]), 2000.00)

    def test_excluir_periodo_api(self):
        """Testa o endpoint de exclusão de notas fiscais por competência (DELETE /documentos)."""
        # 1. Cadastra uma Empresa
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio Software Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        emp_id = response_empresa.json()["id"]

        # 2. Insere documentos
        from datetime import datetime
        db = next(override_get_db())
        from fiscal_workflow.models.models import DocumentoFiscal
        
        doc1 = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="35230512345678000199550010000001231234567891",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            data_emissao=datetime(2026, 5, 15),
            data_competencia=datetime(2026, 5, 1)
        )
        doc2 = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="35230512345678000199550010000001231234567892",
            tipo_documento="NF-e",
            valor_total=Decimal("2000.00"),
            data_emissao=datetime(2026, 6, 20),
            data_competencia=datetime(2026, 6, 1)
        )
        db.add(doc1)
        db.add(doc2)
        db.commit()

        # 3. Exclui por período (Maio/2026)
        resp_del = self.client.delete(f"/documentos?empresa_id={emp_id}&mes=5&ano=2026")
        self.assertEqual(resp_del.status_code, 200)
        self.assertIn("1 documentos fiscais", resp_del.json()["detail"])

        # 4. Verifica se sobrou apenas a nota de Junho/2026
        resp_list = self.client.get(f"/documentos?empresa_id={emp_id}")
        self.assertEqual(len(resp_list.json()), 1)
        self.assertEqual(float(resp_list.json()[0]["valor_total"]), 2000.00)

    def test_excluir_lote_ids_api(self):
        """Testa o endpoint de exclusão de notas fiscais por lista de IDs (POST /documentos/excluir-em-lote)."""
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio Software Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        emp_id = response_empresa.json()["id"]

        # Insere documentos
        from datetime import datetime
        db = next(override_get_db())
        from fiscal_workflow.models.models import DocumentoFiscal
        
        doc1 = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="35230512345678000199550010000001231234567891",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            data_emissao=datetime(2026, 5, 15)
        )
        doc2 = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="35230512345678000199550010000001231234567892",
            tipo_documento="NF-e",
            valor_total=Decimal("2000.00"),
            data_emissao=datetime(2026, 6, 20)
        )
        db.add(doc1)
        db.add(doc2)
        db.commit()

        # Exclui em lote
        resp_del = self.client.post("/documentos/excluir-em-lote", json={"ids": [doc1.id, doc2.id]})
        self.assertEqual(resp_del.status_code, 200)
        self.assertIn("2 documentos fiscais", resp_del.json()["detail"])

        # Verifica se banco ficou vazio para essa empresa
        resp_list = self.client.get(f"/documentos?empresa_id={emp_id}")
        self.assertEqual(len(resp_list.json()), 0)

    def test_consolidado_compras_e_vendas_exclusao_api(self):
        """Testa que notas de entrada são excluídas da apuração consolidada de faturamento e incluídas no relatório de compras."""
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio Software Ltda",
                "regime_tributario": "Simples Nacional",
                "uf": "BA"
            }
        )
        emp_id = response_empresa.json()["id"]

        # Insere um documento de Saída (Faturamento) e um de Entrada (Compra Interestadual de SP "35" para BA "BA")
        db = next(override_get_db())
        from fiscal_workflow.models.models import DocumentoFiscal
        from datetime import datetime
        
        doc_saida = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="29230512345678000199550010000001231234567891", # BA -> BA
            tipo_documento="NF-e",
            tipo_operacao="Saída",
            valor_total=Decimal("5000.00"),
            data_emissao=datetime(2026, 5, 10),
            data_competencia=datetime(2026, 5, 1)
        )
        doc_entrada = DocumentoFiscal(
            empresa_id=emp_id,
            chave_acesso="35230599999999000199550010000001231234567892", # SP -> BA (Interestadual)
            tipo_documento="NF-e",
            tipo_operacao="Entrada",
            valor_total=Decimal("1000.00"),
            data_emissao=datetime(2026, 5, 15),
            data_competencia=datetime(2026, 5, 1),
            itens=[{
                "sequencia": 1,
                "codigo_produto": "PROD001",
                "descricao": "Item Compra SP",
                "valor_total": 1000.00,
                "desconto": 0.0,
                "frete": 0.0,
                "impostos": {
                    "icms": {
                        "cst": "00",
                        "valor_st": 80.00
                    }
                }
            }]
        )
        db.add(doc_saida)
        db.add(doc_entrada)
        db.commit()

        # Consulta consolidado do período (Maio/2026)
        resp_cons = self.client.get(f"/empresas/{emp_id}/consolidado?mes=5&ano=2026")
        self.assertEqual(resp_cons.status_code, 200)
        res_data = resp_cons.json()

        # Faturamento de saída deve ser exatamente 5000.00 (Entrada foi excluída)
        self.assertEqual(float(res_data["total_faturamento"]), 5000.00)
        # Imposto calculado de saída deve ser 6% de 5000 = 300.00 (DAS)
        self.assertEqual(float(res_data["total_imposto"]), 300.00)

        # Dados de compras devem estar preenchidos
        compras = res_data["compras"]
        self.assertEqual(float(compras["total_compras"]), 1000.00)
        # DIFAL (BA usa base dupla): 257.86
        self.assertEqual(float(compras["total_difal"]), 257.86)
        self.assertEqual(float(compras["total_icms_st"]), 80.00)
        self.assertEqual(compras["quantidade_entradas"], 1)

    def test_upload_nota_entrada_resolucao_empresa(self):
        """Testa que notas de entrada cadastram a nota e resolvem a empresa pelo Destinatário."""
        # 1. Cria um XML mock com tpNF=0 (Entrada), Destinatário específico e UF no enderDest
        mock_xml_entrada = MOCK_NFE_XML.replace(
            "<tpNF>1</tpNF>", "<tpNF>0</tpNF>"
        ).replace(
            "<dest>\n                <CNPJ>98765432000188</CNPJ>\n                <xNome>Cliente Exemplo SA</xNome>\n                <IE>444333222111</IE>\n            </dest>",
            "<dest>\n                <CNPJ>99988877000166</CNPJ>\n                <xNome>Angi Compras Ltda</xNome>\n                <enderDest>\n                    <UF>SC</UF>\n                </enderDest>\n            </dest>"
        ).replace(
            "35230512345678000199550010000001231234567890", "35230512345678000199550010000001231234567899"
        )

        xml_file = io.BytesIO(mock_xml_entrada.encode('utf-8'))

        # 2. Executa o upload da nota de entrada
        response = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe_entrada.xml", xml_file, "text/xml"))]
        )
        self.assertEqual(response.status_code, 201)
        doc_list = response.json()
        self.assertEqual(len(doc_list), 1)
        doc_data = doc_list[0]
        self.assertEqual(doc_data["tipo_operacao"], "Entrada")

        # 3. Verifica se a empresa destinatária foi cadastrada automaticamente (e não a emitente)
        response_empresas = self.client.get("/empresas")
        self.assertEqual(response_empresas.status_code, 200)
        empresas = response_empresas.json()

        # O CNPJ emitente "12345678000199" NÃO deve estar cadastrado como empresa (ele é fornecedor da compra)
        emitente_cadastrado = any(e["cnpj"] == "12345678000199" for e in empresas)
        self.assertFalse(emitente_cadastrado, "O emitente de uma nota de entrada não deve ser cadastrado como empresa.")

        # O CNPJ destinatário "99988877000166" DEVE estar cadastrado como a empresa dona da nota
        compradora = next((e for e in empresas if e["cnpj"] == "99988877000166"), None)
        self.assertIsNotNone(compradora)
        self.assertEqual(compradora["razao_social"], "Angi Compras Ltda")
        self.assertEqual(compradora["uf"], "SC")
        self.assertEqual(compradora["regime_tributario"], "Simples Nacional")

        # O documento fiscal deve pertencer à empresa compradora
        self.assertEqual(doc_data["empresa_id"], compradora["id"])

    def test_upload_entrada_com_parametros_forcados(self):
        """Testa que o upload com parâmetros de formulário força a associação e a operação (Entrada)."""
        # 1. Cadastra uma Empresa ativa
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "99988877000166",
                "razao_social": "Angi Compras Ltda",
                "regime_tributario": "Simples Nacional",
                "uf": "SC"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        emp_id = response_empresa.json()["id"]

        # 2. Prepara um XML normal de Saída (tpNF=1)
        # O XML original diz que é uma venda (Saída) do emitente 12345678000199
        xml_file = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))

        # 3. Executa o upload passando parâmetros para forçar Entrada da empresa selecionada e competência
        response = self.client.post(
            "/documentos/upload",
            data={
                "empresa_id": emp_id,
                "tipo_operacao_forcada": "Entrada",
                "data_competencia": "2026-05"
            },
            files=[("files", ("nfe_terceiros.xml", xml_file, "text/xml"))]
        )
        self.assertEqual(response.status_code, 201)
        doc_list = response.json()
        self.assertEqual(len(doc_list), 1)
        doc_data = doc_list[0]

        # 4. Valida se a operação foi forçada para Entrada e vinculada à empresa certa
        self.assertEqual(doc_data["tipo_operacao"], "Entrada")
        self.assertEqual(doc_data["empresa_id"], emp_id)
        
        # A data de competência deve ter sido forçada para a competência indicada (2026-05-01)
        self.assertTrue(doc_data["data_competencia"].startswith("2026-05-01"))
        # A data de emissão deve conter a data real do XML (2023-05-15)
        self.assertTrue(doc_data["data_emissao"].startswith("2023-05-15"))

    def test_upload_entrada_autodetect_empresa_cadastrada(self):
        """Testa que se a empresa destinatária está cadastrada, a nota é associada como Entrada sem duplicar ou cadastrar fornecedor."""
        # 1. Cadastra a empresa que será a destinatária (com o CNPJ destinatário do MOCK_NFE_XML: 98765432000188)
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "98765432000188",
                "razao_social": "Cliente Exemplo SA (Nossa Empresa)",
                "regime_tributario": "Simples Nacional",
                "uf": "BA"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        emp_id = response_empresa.json()["id"]

        # 2. Upload da nota de compra (emitente = 12345678000199, destinatario = 98765432000188)
        xml_file = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        response_upload = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe_compra.xml", xml_file, "text/xml"))]
        )
        self.assertEqual(response_upload.status_code, 201)
        doc_list = response_upload.json()
        self.assertEqual(len(doc_list), 1)
        doc_data = doc_list[0]

        # O documento deve ter sido associado a Cliente Exemplo SA (ID emp_id) e tipo_operacao = Entrada
        self.assertEqual(doc_data["empresa_id"], emp_id)
        self.assertEqual(doc_data["tipo_operacao"], "Entrada")

        # Garante que o fornecedor/emitente (12345678000199) não foi cadastrado como empresa ativa no banco
        response_empresas = self.client.get("/empresas")
        self.assertEqual(response_empresas.status_code, 200)
        empresas = response_empresas.json()
        
        # Só deve existir uma empresa cadastrada (a que criamos manualmente)
        self.assertEqual(len(empresas), 1)
        self.assertEqual(empresas[0]["cnpj"], "98765432000188")

    def test_xml_storage_hierarchy(self):
        """Testa se o XML importado é salvo corretamente na hierarquia de pastas proposta."""
        import shutil
        from pathlib import Path
        
        # 1. Limpa qualquer armazenamento de teste prévio apenas para a empresa de teste
        test_storage = Path("armazenamento_xml")
        test_company_storage = test_storage / "Stenio_Software_Tech_Ltda"
        shutil.rmtree(test_company_storage, ignore_errors=True)
        
        # 2. Cadastra uma empresa emitente com caracteres inválidos no nome
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio/Software\\Tech:Ltda", # Caracteres inválidos / \ :
                "regime_tributario": "Simples Nacional",
                "uf": "BA"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        
        # 3. Upload do XML mockado
        xml_file = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        response_upload = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe_venda.xml", xml_file, "text/xml"))]
        )
        self.assertEqual(response_upload.status_code, 201)
        
        # 4. Verifica se o caminho esperado foi gerado corretamente
        # Razão social saneada: "Stenio_Software_Tech_Ltda"
        expected_path = test_company_storage / "202305" / "Saídas" / "NFe" / "35230512345678000199550010000001231234567890.xml"
        
        self.assertTrue(expected_path.exists(), f"Caminho esperado não foi criado: {expected_path}")
        
        # 5. Valida o conteúdo salvo
        with open(expected_path, "r", encoding="utf-8") as f:
            saved_content = f.read()
        self.assertEqual(saved_content, MOCK_NFE_XML)
        
        # Limpa apenas o diretório de teste
        shutil.rmtree(test_company_storage, ignore_errors=True)

    def test_obter_periodo_do_caminho(self):
        """Testa o extrator obter_periodo_do_caminho com diferentes formatos de pasta."""
        from fiscal_workflow.main import obter_periodo_do_caminho
        self.assertEqual(obter_periodo_do_caminho("05-2026/nota.xml"), datetime(2026, 5, 1))
        self.assertEqual(obter_periodo_do_caminho("pasta/202606/nota.xml"), datetime(2026, 6, 1))
        self.assertEqual(obter_periodo_do_caminho("pasta/07_2026/nota.xml"), datetime(2026, 7, 1))
        self.assertEqual(obter_periodo_do_caminho("2026-08/nota.xml"), datetime(2026, 8, 1))
        self.assertEqual(obter_periodo_do_caminho("sem_data/nota.xml"), None)

    def test_upload_entrada_period_hierarchy(self):
        """Testa a hierarquia de resolução de período de Entrada na API de upload (pasta vs emissão)."""
        # 1. Cadastra a empresa destinatária para ser tratada como Entrada
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "98765432000188",
                "razao_social": "Cliente Exemplo SA",
                "regime_tributario": "Simples Nacional"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        
        # 2. Upload de nota fiscal com caminho de subpasta contendo o período '2026-10'
        # A emissão original do MOCK_NFE_XML é 2023-05-15
        xml_file = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        response_upload = self.client.post(
            "/documentos/upload",
            files=[("files", ("subpasta/2026-10/nfe.xml", xml_file, "text/xml"))]
        )
        self.assertEqual(response_upload.status_code, 201)
        doc_data = response_upload.json()[0]
        
        # Como é Entrada, deve priorizar a pasta para definir a competência, mas manter a emissão real
        self.assertEqual(doc_data["tipo_operacao"], "Entrada")
        self.assertTrue(doc_data["data_competencia"].startswith("2026-10-01"))
        self.assertTrue(doc_data["data_emissao"].startswith("2023-05-15"))

    def test_editar_competencia_documento_endpoint(self):
        """Testa o endpoint PUT /documentos/{id}/competencia para alterar competência."""
        # 1. Cadastra a empresa e faz upload de nota
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio Software Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        emp_id = response_empresa.json()["id"]
        
        xml_file = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        response_upload = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe.xml", xml_file, "text/xml"))]
        )
        doc_id = response_upload.json()[0]["id"]
        
        # 2. Altera a competência da nota
        response_put = self.client.put(
            f"/documentos/{doc_id}/competencia",
            json={"data_competencia": "2026-07"}
        )
        self.assertEqual(response_put.status_code, 200)
        doc_data = response_put.json()
        self.assertTrue(doc_data["data_competencia"].startswith("2026-07-01"))
        
        # 3. Verifica se a alteração persistiu no banco
        response_get = self.client.get("/documentos")
        doc_get = response_get.json()[0]
        self.assertTrue(doc_get["data_competencia"].startswith("2026-07-01"))

    def test_editar_competencia_lote_endpoint(self):
        """Testa o endpoint POST /documentos/competencia-em-lote para alterar competência de múltiplas notas."""
        # 1. Cadastra a empresa e faz upload de 2 notas
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Stenio Software Ltda",
                "regime_tributario": "Simples Nacional"
            }
        )
        emp_id = response_empresa.json()["id"]
        
        xml_file_1 = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        mock_xml_2 = MOCK_NFE_XML.replace(
            "35230512345678000199550010000001231234567890", "35230512345678000199550010000001241234567890"
        )
        xml_file_2 = io.BytesIO(mock_xml_2.encode('utf-8'))
        
        response_upload = self.client.post(
            "/documentos/upload",
            files=[
                ("files", ("nfe1.xml", xml_file_1, "text/xml")),
                ("files", ("nfe2.xml", xml_file_2, "text/xml"))
            ]
        )
        docs = response_upload.json()
        ids = [doc["id"] for doc in docs]
        
        # 2. Altera a competência em lote
        response_post = self.client.post(
            "/documentos/competencia-em-lote",
            json={"ids": ids, "data_competencia": "2026-11"}
        )
        self.assertEqual(response_post.status_code, 200)
        self.assertIn("Competência de 2 documentos fiscais", response_post.json()["detail"])
        
        # 3. Verifica se as alterações persistiram no banco
        response_get = self.client.get("/documentos")
        for doc in response_get.json():
            if doc["id"] in ids:
                self.assertTrue(doc["data_competencia"].startswith("2026-11-01"))

    def test_xml_storage_sincronizacao_competencia(self):
        """Testa se o arquivo XML físico é movido de pasta quando a competência é alterada."""
        import shutil
        from pathlib import Path
        
        # 1. Limpa diretório de testes
        test_storage = Path("armazenamento_xml")
        test_company_storage = test_storage / "Sincronia Ltda"
        shutil.rmtree(test_company_storage, ignore_errors=True)
        
        # 2. Cadastra empresa e faz upload de XML
        response_empresa = self.client.post(
            "/empresas",
            json={
                "cnpj": "12345678000199",
                "razao_social": "Sincronia Ltda",
                "regime_tributario": "Simples Nacional",
                "uf": "BA"
            }
        )
        self.assertEqual(response_empresa.status_code, 201)
        
        xml_file = io.BytesIO(MOCK_NFE_XML.encode('utf-8'))
        response_upload = self.client.post(
            "/documentos/upload",
            files=[("files", ("nfe_venda.xml", xml_file, "text/xml"))]
        )
        doc_id = response_upload.json()[0]["id"]
        
        # Caminho original da nota: 202305
        original_path = test_company_storage / "202305" / "Saídas" / "NFe" / "35230512345678000199550010000001231234567890.xml"
        self.assertTrue(original_path.exists())
        
        # 3. Altera competência para 2026-07
        response_put = self.client.put(
            f"/documentos/{doc_id}/competencia",
            json={"data_competencia": "2026-07"}
        )
        self.assertEqual(response_put.status_code, 200)
        
        # Novo caminho esperado da nota: 202607
        new_path = test_company_storage / "202607" / "Saídas" / "NFe" / "35230512345678000199550010000001231234567890.xml"
        
        # O arquivo no caminho antigo deve ter sido movido (não existir mais)
        self.assertFalse(original_path.exists())
        # O arquivo deve existir no novo caminho
        self.assertTrue(new_path.exists())
        
        # Limpa o diretório de testes
        shutil.rmtree(test_company_storage, ignore_errors=True)

if __name__ == "__main__":
    unittest.main()

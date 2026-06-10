import unittest
from decimal import Decimal
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

# Importa o parser
from fiscal_workflow.services.parsers import parse_nfe
# Importa a modelagem
from fiscal_workflow.models import Base, Empresa, DocumentoFiscal, RegimeTributario, StatusApuracao

# XML Mock de NF-e estruturado e completo baseado nos padrões da SEFAZ
MOCK_NFE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
    <NFe>
        <infNFe Id="NFe35230512345678000199550010000001231234567890" versao="4.00">
            <ide>
                <cUF>35</cUF>
                <cNF>12345678</cNF>
                <natOp>Venda de mercadoria</natOp>
                <mod>55</mod>
                <serie>1</serie>
                <nNF>123</nNF>
                <dhEmi>2023-05-15T10:30:00-03:00</dhEmi>
                <tpNF>1</tpNF>
                <idDest>1</idDest>
                <cMunFG>3550308</cMunFG>
                <tpImp>1</tpImp>
                <tpEmis>1</tpEmis>
                <cDV>0</cDV>
                <tpAmb>2</tpAmb>
                <finNFe>1</finNFe>
                <indPres>1</indPres>
                <procEmi>0</procEmi>
                <verProc>4.0.0</verProc>
            </ide>
            <emit>
                <CNPJ>12345678000199</CNPJ>
                <xNome>Stenio Software Ltda</xNome>
                <xFant>Stenio Tech</xFant>
                <IE>111222333444</IE>
                <CRT>1</CRT>
            </emit>
            <dest>
                <CNPJ>98765432000188</CNPJ>
                <xNome>Cliente Exemplo SA</xNome>
                <IE>444333222111</IE>
            </dest>
            <det nSeq="1">
                <prod>
                    <cProd>PROD001</cProd>
                    <cEAN>7891234567890</cEAN>
                    <xProd>Notebook de Ultima Geracao</xProd>
                    <NCM>84713012</NCM>
                    <CFOP>5102</CFOP>
                    <uCom>UN</uCom>
                    <qCom>2.0000</qCom>
                    <vUnCom>2250.0000</vUnCom>
                    <vProd>4500.00</vProd>
                    <vDesc>200.00</vDesc>
                    <vFrete>50.00</vFrete>
                </prod>
                <imposto>
                    <ICMS>
                        <ICMSSN101>
                            <orig>0</orig>
                            <CSOSN>101</CSOSN>
                            <pCredSN>2.50</pCredSN>
                            <vCredICMSSN>112.50</vCredICMSSN>
                        </ICMSSN101>
                    </ICMS>
                    <PIS>
                        <PISAliq>
                            <CST>01</CST>
                            <vBC>4300.00</vBC>
                            <pPIS>0.65</pPIS>
                            <vPIS>27.95</vPIS>
                        </PISAliq>
                    </PIS>
                    <COFINS>
                        <COFINSAliq>
                            <CST>01</CST>
                            <vBC>4300.00</vBC>
                            <pCOFINS>3.00</pCOFINS>
                            <vCOFINS>129.00</vCOFINS>
                        </COFINSAliq>
                    </COFINS>
                </imposto>
            </det>
            <total>
                <ICMSTot>
                    <vBC>0.00</vBC>
                    <vICMS>0.00</vICMS>
                    <vProd>4500.00</vProd>
                    <vDesc>200.00</vDesc>
                    <vFrete>50.00</vFrete>
                    <vNF>4350.00</vNF>
                </ICMSTot>
            </total>
        </infNFe>
    </NFe>
</nfeProc>
"""

class TestNfeParser(unittest.TestCase):
    def setUp(self):
        """Configura banco em memória para testar a persistência pós-parse."""
        self.engine = create_engine("sqlite:///:memory:")
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.session = self.Session()

    def tearDown(self):
        self.session.close()
        Base.metadata.drop_all(self.engine)

    def test_parse_nfe_dados_cabecalho_e_itens(self):
        """Testa se a extração lxml de cabeçalho e itens brutos funciona corretamente."""
        xml_bytes = MOCK_NFE_XML.encode('utf-8')
        dados = parse_nfe(xml_bytes)

        # 1. Validação do Cabeçalho da Nota
        self.assertEqual(dados["chave_acesso"], "35230512345678000199550010000001231234567890")
        self.assertEqual(dados["numero_nf"], "123")
        self.assertEqual(dados["data_emissao"], "2023-05-15T10:30:00-03:00")
        self.assertEqual(dados["tipo_documento"], "NF-e")
        self.assertEqual(dados["emitente_cnpj"], "12345678000199")
        self.assertEqual(dados["emitente_razao_social"], "Stenio Software Ltda")
        self.assertEqual(dados["valor_total"], 4350.00)

        # 2. Validação dos Itens Extraitos (CST, CFOP, vBC, vICMS, etc.)
        self.assertEqual(len(dados["itens"]), 1)
        item = dados["itens"][0]
        self.assertEqual(item["sequencia"], 1)
        self.assertEqual(item["codigo_produto"], "PROD001")
        self.assertEqual(item["descricao"], "Notebook de Ultima Geracao")
        self.assertEqual(item["cfop"], "5102")
        self.assertEqual(item["ncm"], "84713012")
        self.assertEqual(item["quantidade"], 2.0)
        self.assertEqual(item["valor_unitario"], 2250.0)
        self.assertEqual(item["valor_total"], 4500.0)
        self.assertEqual(item["desconto"], 200.0)
        self.assertEqual(item["frete"], 50.0)

        # 3. Validação dos Tributos do Item
        impostos = item["impostos"]
        
        # ICMS Simples Nacional (CSOSN 101)
        self.assertEqual(impostos["icms"]["cst"], "101")
        self.assertEqual(impostos["icms"]["aliquota_credito_sn"], 2.50)
        self.assertEqual(impostos["icms"]["valor_credito_sn"], 112.50)

        # PIS
        self.assertEqual(impostos["pis"]["cst"], "01")
        self.assertEqual(impostos["pis"]["valor_base_calculo"], 4300.00)
        self.assertEqual(impostos["pis"]["aliquota"], 0.65)
        self.assertEqual(impostos["pis"]["valor"], 27.95)

        # COFINS
        self.assertEqual(impostos["cofins"]["cst"], "01")
        self.assertEqual(impostos["cofins"]["valor_base_calculo"], 4300.00)
        self.assertEqual(impostos["cofins"]["aliquota"], 3.00)
        self.assertEqual(impostos["cofins"]["valor"], 129.00)

    def test_persistencia_dados_extraidos_banco(self):
        """Testa se os dados extraídos pelo parser podem ser salvos no banco com o JSON de itens."""
        xml_bytes = MOCK_NFE_XML.encode('utf-8')
        dados = parse_nfe(xml_bytes)

        # 1. Cria a empresa emitente no banco
        empresa = Empresa(
            cnpj=dados["emitente_cnpj"],
            razao_social=dados["emitente_razao_social"],
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL
        )
        self.session.add(empresa)
        self.session.commit()

        # 2. Cria o Documento Fiscal associando os dados e a lista de itens mapeada em JSON
        doc = DocumentoFiscal(
            empresa_id=empresa.id,
            chave_acesso=dados["chave_acesso"],
            tipo_documento=dados["tipo_documento"],
            valor_total=Decimal(str(dados["valor_total"])),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=dados["itens"] # Persiste como JSON
        )
        self.session.add(doc)
        self.session.commit()

        # Limpa o cache da sessão para forçar a leitura do banco de dados física
        self.session.expire_all()

        # 3. Consulta o documento salvo e valida
        doc_salvo = self.session.query(DocumentoFiscal).filter_by(chave_acesso=dados["chave_acesso"]).first()
        self.assertIsNotNone(doc_salvo)
        self.assertEqual(doc_salvo.valor_total, Decimal("4350.00"))
        self.assertEqual(doc_salvo.status_apuracao, StatusApuracao.PENDENTE)
        
        # Valida que o JSON de itens foi perfeitamente recuperado
        self.assertEqual(len(doc_salvo.itens), 1)
        item = doc_salvo.itens[0]
        self.assertEqual(item["codigo_produto"], "PROD001")
        self.assertEqual(item["cfop"], "5102")
        self.assertEqual(item["impostos"]["icms"]["cst"], "101")
        self.assertEqual(item["impostos"]["pis"]["valor"], 27.95)

    def test_parse_nfe_cancelada(self):
        """Testa se o parser identifica o status cancelado (cStat 101) a partir do protocolo."""
        xml_cancelada = MOCK_NFE_XML.replace("</infNFe>\n    </NFe>", "</infNFe>\n    </NFe>\n    <protNFe>\n        <infProt>\n            <cStat>101</cStat>\n        </infProt>\n    </protNFe>")
        dados = parse_nfe(xml_cancelada.encode('utf-8'))
        self.assertEqual(dados["cstat"], "101")

if __name__ == "__main__":
    unittest.main()

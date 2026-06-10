import unittest
from decimal import Decimal
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from fiscal_workflow.services.parsers import parse_documento_fiscal, parse_nfse
from fiscal_workflow.services.calculadoras import CalculadoraFactory
from fiscal_workflow.models import Base, Empresa, DocumentoFiscal, RegimeTributario, StatusApuracao

MOCK_NFSE_XML = """<?xml version="1.0" encoding="utf-8"?>
<NFSe xmlns="http://www.sped.fazenda.gov.br/nfse" versao="1.00">
  <infNFSe Id="NFS29116591233887520000117000000000176726060000167275">
    <xLocEmi>Guajeru</xLocEmi>
    <xLocPrestacao>Cordeiros</xLocPrestacao>
    <nNFSe>1767</nNFSe>
    <cLocIncid>2911659</cLocIncid>
    <xLocIncid>Guajeru</xLocIncid>
    <xTribNac>Assessoria e consultoria em informatica</xTribNac>
    <xTribMun>999</xTribMun>
    <xNBS>Servicos de processamento de dados</xNBS>
    <verAplic>1.00</verAplic>
    <ambGer>1</ambGer>
    <tpEmis>2</tpEmis>
    <cStat>100</cStat>
    <dhProc>2026-06-02T10:33:16-03:00</dhProc>
    <nDFSe>16727</nDFSe>
    <emit>
      <CNPJ>33887520000117</CNPJ>
      <IM>220023782</IM>
      <xNome>SIMPLE TECNOLOGIA EM GESTAO PUBLICA LTDA</xNome>
      <xFant>SIMPLE Tecnologia Em Gestao Publica LTDA</xFant>
      <enderNac>
        <xLgr>RUA   OSVALDO JOSE DE DEUS</xLgr>
        <nro>317</nro>
        <xBairro>CENTRO</xBairro>
        <cMun>2911659</cMun>
        <UF>BA</UF>
        <CEP>46205000</CEP>
      </enderNac>
      <fone>77981395301</fone>
      <email>contato@simples.srv.br</email>
    </emit>
    <valores>
      <vBC>1800.00</vBC>
      <pAliqAplic>2.17</pAliqAplic>
      <vISSQN>39.06</vISSQN>
      <vLiq>1800.00</vLiq>
    </valores>
    <DPS versao="1.00">
      <infDPS Id="DPS291165923388752000011700001000000000001767">
        <tpAmb>1</tpAmb>
        <dhEmi>2026-06-02T10:33:16-03:00</dhEmi>
        <verAplic>1.00</verAplic>
        <serie>1</serie>
        <nDPS>1767</nDPS>
        <dCompet>2026-05-01</dCompet>
        <tpEmit>1</tpEmit>
        <cLocEmi>2911659</cLocEmi>
        <prest>
          <CNPJ>33887520000117</CNPJ>
          <IM>22.002378-2</IM>
          <fone>77981395301</fone>
          <email>contato@simples.srv.br</email>
          <regTrib>
            <opSimpNac>3</opSimpNac>
            <regApTribSN>1</regApTribSN>
            <regEspTrib>0</regEspTrib>
          </regTrib>
        </prest>
        <toma>
          <CNPJ>13694468000175</CNPJ>
          <xNome>MUNICIPIO DE CORDEIROS</xNome>
          <end>
            <endNac>
              <cMun>2909000</cMun>
              <CEP>46280000</CEP>
            </endNac>
            <xLgr>PC JOSE MOREIRA CORDEIRO</xLgr>
            <nro>104</nro>
            <xBairro>Centro</xBairro>
          </end>
        </toma>
        <serv>
          <locPrest>
            <cLocPrestacao>2909000</cLocPrestacao>
          </locPrest>
          <cServ>
            <cTribNac>010601</cTribNac>
            <cTribMun>999</cTribMun>
            <xDescServ>PRESTACAO DE SERVICOS ESPECIALIZADOS EM ASSESSORIA COM SUPORTE TECNICO MANUTENCAO E OUTROS SERVICOS EM TECNOLOGIA DA INFORMACAO ORGANIZACAO DE ROTINAS DE COMPRAS E SERVICOS</xDescServ>
            <cNBS>115090000</cNBS>
          </cServ>
          <infoCompl>
            <xInfComp>Lei 1274112  Carga tributaria aprox: Uniao 1333% Estado 000% Municipio 217% Substitui IBPT DADOS BANCARIOS  SIMPLE TECEM GESTPUBLICA LTDA  BANCO: 104 CAIXA ECONOMICA FEDERAL AG: 0947 CONTA CORRENTE: 5783304057 OPER: 1292 CHAVE PIX CNPJ  33887520000117</xInfComp>
          </infoCompl>
        </serv>
        <valores>
          <vServPrest>
            <vServ>1800.00</vServ>
          </vServPrest>
          <trib>
            <tribMun>
              <tribISSQN>1</tribISSQN>
              <tpRetISSQN>1</tpRetISSQN>
              <pAliq>2.17</pAliq>
            </tribMun>
            <totTrib>
              <indTotTrib>0</indTotTrib>
            </totTrib>
          </trib>
        </valores>
      </infDPS>
    </DPS>
  </infNFSe>
</NFSe>
"""

class TestNfseParserAndCalculators(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:")
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.session = self.Session()

    def tearDown(self):
        self.session.close()
        Base.metadata.drop_all(self.engine)

    def test_parse_nfse_success(self):
        xml_bytes = MOCK_NFSE_XML.encode('utf-8')
        dados_list = parse_documento_fiscal(xml_bytes)
        self.assertEqual(len(dados_list), 1)
        dados = dados_list[0]

        self.assertEqual(dados["chave_acesso"], "29116591233887520000117000000000176726060000167275")
        self.assertEqual(dados["numero_nf"], "1767")
        self.assertEqual(dados["tipo_documento"], "NFS-e")
        self.assertEqual(dados["emitente_cnpj"], "33887520000117")
        self.assertEqual(dados["emitente_razao_social"], "SIMPLE TECNOLOGIA EM GESTAO PUBLICA LTDA")
        self.assertEqual(dados["valor_total"], 1800.00)

        # Item virtual
        self.assertEqual(len(dados["itens"]), 1)
        item = dados["itens"][0]
        self.assertEqual(item["codigo_produto"], "010601")
        self.assertEqual(item["descricao"], "PRESTACAO DE SERVICOS ESPECIALIZADOS EM ASSESSORIA COM SUPORTE TECNICO MANUTENCAO E OUTROS SERVICOS EM TECNOLOGIA DA INFORMACAO ORGANIZACAO DE ROTINAS DE COMPRAS E SERVICOS")
        self.assertEqual(item["valor_total"], 1800.00)
        self.assertEqual(item["impostos"]["iss"]["valor_base_calculo"], 1800.00)
        self.assertEqual(item["impostos"]["iss"]["aliquota"], 2.17)
        self.assertEqual(item["impostos"]["iss"]["valor"], 39.06)

    def test_lucro_presumido_calculation_with_iss(self):
        xml_bytes = MOCK_NFSE_XML.encode('utf-8')
        dados = parse_documento_fiscal(xml_bytes)[0]

        empresa = Empresa(
            cnpj=dados["emitente_cnpj"],
            razao_social=dados["emitente_razao_social"],
            regime_tributario=RegimeTributario.LUCRO_PRESUMIDO
        )
        self.session.add(empresa)
        self.session.commit()

        doc = DocumentoFiscal(
            empresa_id=empresa.id,
            chave_acesso=dados["chave_acesso"],
            tipo_documento=dados["tipo_documento"],
            valor_total=Decimal(str(dados["valor_total"])),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=dados["itens"]
        )
        self.session.add(doc)
        self.session.commit()

        calculadora = CalculadoraFactory.obter_calculadora(RegimeTributario.LUCRO_PRESUMIDO)
        res = calculadora.calcular(doc)

        self.assertEqual(res["regime"], "Lucro Presumido")
        # Federal taxes: 1800 * (0.0065 + 0.0300 + 0.0480 + 0.0288) = 1800 * 0.1133 = 203.94
        # Municipal ISS: 1800 * 2.17% = 39.06
        # Total: 203.94 + 39.06 = 243.00
        self.assertEqual(res["imposto_calculado"], Decimal("243.00"))
        self.assertEqual(res["detalhes"]["iss"], Decimal("39.06"))
        self.assertEqual(res["detalhes"]["pis"], Decimal("11.70"))

    def test_simples_nacional_calculation(self):
        xml_bytes = MOCK_NFSE_XML.encode('utf-8')
        dados = parse_documento_fiscal(xml_bytes)[0]

        empresa = Empresa(
            cnpj=dados["emitente_cnpj"],
            razao_social=dados["emitente_razao_social"],
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("200000.00"), # Anexo III, Faixa 2: nominal 11.2%, dedução 9360.00
            categoria_simples="Serviços (Anexo III)"
        )
        self.session.add(empresa)
        self.session.commit()

        doc = DocumentoFiscal(
            empresa_id=empresa.id,
            chave_acesso=dados["chave_acesso"],
            tipo_documento=dados["tipo_documento"],
            valor_total=Decimal(str(dados["valor_total"])),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=dados["itens"]
        )
        self.session.add(doc)
        self.session.commit()

        calculadora = CalculadoraFactory.obter_calculadora(RegimeTributario.SIMPLES_NACIONAL)
        res = calculadora.calcular(doc)

        self.assertEqual(res["regime"], "Simples Nacional")
        # Effective rate calculation:
        # (200000 * 0.112 - 9360) / 200000 = (22400 - 9360) / 200000 = 13040 / 200000 = 0.0652 (6.52%)
        # Tax: 1800 * 6.52% = 117.36
        self.assertEqual(res["imposto_calculado"], Decimal("117.36"))
        self.assertEqual(res["detalhes"]["das"], Decimal("117.36"))

    def test_parse_abrasf_multiple_nfse(self):
        # XML mock com duas notas no padrão ABRASF
        mock_abrasf = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <ListaNfse>
            <CompNfse>
                <Nfse versao="2.04">
                    <InfNfse Id="2c9180829df2fe48019df34974be3a08">
                        <Numero>202600000000014</Numero>
                        <DataEmissao>2026-05-04</DataEmissao>
                        <ValoresNfse>
                            <BaseCalculo>4400.00</BaseCalculo>
                            <Aliquota>2.01</Aliquota>
                            <ValorIss>88.44</ValorIss>
                            <ValorLiquidoNfse>4400.00</ValorLiquidoNfse>
                        </ValoresNfse>
                        <PrestadorServico>
                            <RazaoSocial>CBF MEDICINA INTEGRADA LTDA</RazaoSocial>
                        </PrestadorServico>
                        <DeclaracaoPrestacaoServico>
                            <InfDeclaracaoPrestacaoServico>
                                <Prestador>
                                    <CpfCnpj>
                                        <Cnpj>61877417000121</Cnpj>
                                    </CpfCnpj>
                                </Prestador>
                                <TomadorServico>
                                    <IdentificacaoTomador>
                                        <CpfCnpj>
                                            <Cnpj>11402446000169</Cnpj>
                                        </CpfCnpj>
                                    </IdentificacaoTomador>
                                    <RazaoSocial>FUNDO MUNIC SAUDE PLANALTO</RazaoSocial>
                                </TomadorServico>
                                <Servico>
                                    <Valores>
                                        <ValorServicos>4400.00</ValorServicos>
                                    </Valores>
                                    <CodigoServicoNacional>040101</CodigoServicoNacional>
                                    <Discriminacao>04 PLANTOES MEDICOS</Discriminacao>
                                </Servico>
                                <OptanteSimplesNacional>1</OptanteSimplesNacional>
                            </InfDeclaracaoPrestacaoServico>
                        </DeclaracaoPrestacaoServico>
                    </InfNfse>
                </Nfse>
            </CompNfse>
            <CompNfse>
                <Nfse versao="2.04">
                    <InfNfse Id="2c9180829df2fe48019df351008b3f7b">
                        <Numero>202600000000015</Numero>
                        <DataEmissao>2026-05-04</DataEmissao>
                        <ValoresNfse>
                            <BaseCalculo>4400.00</BaseCalculo>
                            <Aliquota>2.01</Aliquota>
                            <ValorIss>88.44</ValorIss>
                            <ValorLiquidoNfse>4400.00</ValorLiquidoNfse>
                        </ValoresNfse>
                        <PrestadorServico>
                            <RazaoSocial>CBF MEDICINA INTEGRADA LTDA</RazaoSocial>
                        </PrestadorServico>
                        <DeclaracaoPrestacaoServico>
                            <InfDeclaracaoPrestacaoServico>
                                <Prestador>
                                    <CpfCnpj>
                                        <Cnpj>61877417000121</Cnpj>
                                    </CpfCnpj>
                                </Prestador>
                                <TomadorServico>
                                    <IdentificacaoTomador>
                                        <CpfCnpj>
                                            <Cnpj>11402446000169</Cnpj>
                                        </CpfCnpj>
                                    </IdentificacaoTomador>
                                    <RazaoSocial>FUNDO MUNIC SAUDE PLANALTO</RazaoSocial>
                                </TomadorServico>
                                <Servico>
                                    <Valores>
                                        <ValorServicos>4400.00</ValorServicos>
                                    </Valores>
                                    <CodigoServicoNacional>040101</CodigoServicoNacional>
                                    <Discriminacao>04 PLANTOES MEDICOS OUTRO</Discriminacao>
                                </Servico>
                                <OptanteSimplesNacional>1</OptanteSimplesNacional>
                            </InfDeclaracaoPrestacaoServico>
                        </DeclaracaoPrestacaoServico>
                    </InfNfse>
                </Nfse>
            </CompNfse>
        </ListaNfse>
        """
        dados_list = parse_documento_fiscal(mock_abrasf.encode('utf-8'))
        self.assertEqual(len(dados_list), 2)
        
        self.assertEqual(dados_list[0]["chave_acesso"], "2c9180829df2fe48019df34974be3a08")
        self.assertEqual(dados_list[0]["numero_nf"], "202600000000014")
        self.assertEqual(dados_list[0]["emitente_cnpj"], "61877417000121")
        self.assertEqual(dados_list[0]["destinatario_cnpj"], "11402446000169")
        
        self.assertEqual(dados_list[1]["chave_acesso"], "2c9180829df2fe48019df351008b3f7b")
        self.assertEqual(dados_list[1]["numero_nf"], "202600000000015")

if __name__ == "__main__":
    unittest.main()

import unittest
from decimal import Decimal
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from fiscal_workflow.models.models import Base, Empresa, DocumentoFiscal, AjusteDocumento, RegimeTributario, StatusApuracao
from fiscal_workflow.services.calculadoras import (
    CalculadoraFactory,
    CalculadoraSimplesNacional,
    CalculadoraLucroPresumido
)

class TestTaxCalculators(unittest.TestCase):
    def setUp(self):
        """Configura banco em memória e Session para criar documentos fictícios."""
        self.engine = create_engine("sqlite:///:memory:")
        Base.metadata.create_all(self.engine)
        self.Session = sessionmaker(bind=self.engine)
        self.session = self.Session()

    def tearDown(self):
        self.session.close()
        Base.metadata.drop_all(self.engine)

    def test_factory_calculadora_injecao(self):
        """Verifica se a factory retorna a estratégia correta de cálculo para cada regime."""
        calc_simples = CalculadoraFactory.obter_calculadora(RegimeTributario.SIMPLES_NACIONAL)
        self.assertIsInstance(calc_simples, CalculadoraSimplesNacional)

        calc_presumido = CalculadoraFactory.obter_calculadora(RegimeTributario.LUCRO_PRESUMIDO)
        self.assertIsInstance(calc_presumido, CalculadoraLucroPresumido)

        with self.assertRaises(ValueError):
            # Lucro Real ainda não foi implementado no motor
            CalculadoraFactory.obter_calculadora(RegimeTributario.LUCRO_REAL)

    def test_calculo_simples_nacional_com_e_sem_ajuste(self):
        """Valida a apuração de imposto no Simples Nacional e a reatividade aos ajustes manuais."""
        # 1. Cadastro de Empresa e Nota Fiscal (Original: 1000.00)
        emp = Empresa(
            cnpj="12345678000199",
            razao_social="Simples Comércio Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL
        )
        self.session.add(emp)
        self.session.commit()

        doc = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="11112222333344445555666677778888999900001111",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc)
        self.session.commit()

        # 2. Cálculo inicial (Sem ajustes, alíquota padrão 6%)
        calc = CalculadoraFactory.obter_calculadora(emp.regime_tributario)
        res_inicial = calc.calcular(doc)

        self.assertEqual(res_inicial["valor_final_base"], Decimal("1000.00"))
        self.assertEqual(res_inicial["imposto_calculado"], Decimal("60.00")) # 6% de 1000.00

        # 3. Adiciona um ajuste manual de +150.00
        ajuste = AjusteDocumento(
            documento_id=doc.id,
            valor_total_ajuste=Decimal("150.00"),
            justificativa="Acréscimo de serviço complementar"
        )
        self.session.add(ajuste)
        self.session.commit()

        # 4. Novo cálculo (Deve incidir sobre o valor_final = 1150.00)
        res_ajustado = calc.calcular(doc)
        self.assertEqual(res_ajustado["valor_final_base"], Decimal("1150.00"))
        self.assertEqual(res_ajustado["imposto_calculado"], Decimal("69.00")) # 6% de 1150.00

    def test_calculo_lucro_presumido_com_e_sem_ajuste(self):
        """Valida a apuração federal (PIS/COFINS/IRPJ/CSLL) no Lucro Presumido."""
        # 1. Cadastro de Empresa e Nota Fiscal (Original: 10000.00)
        emp = Empresa(
            cnpj="98765432000188",
            razao_social="Presumido Serviços Ltda",
            regime_tributario=RegimeTributario.LUCRO_PRESUMIDO
        )
        self.session.add(emp)
        self.session.commit()

        doc = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="22223333444455556666777788889999000011112222",
            tipo_documento="NF-e",
            valor_total=Decimal("10000.00"),
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc)
        self.session.commit()

        # 2. Cálculo inicial (Sem ajustes)
        calc = CalculadoraFactory.obter_calculadora(emp.regime_tributario)
        res_inicial = calc.calcular(doc)

        self.assertEqual(res_inicial["valor_final_base"], Decimal("10000.00"))
        
        # PIS: 0.65% de 10000 = 65.00
        self.assertEqual(res_inicial["detalhes"]["pis"], Decimal("65.00"))
        # COFINS: 3% de 10000 = 300.00
        self.assertEqual(res_inicial["detalhes"]["cofins"], Decimal("300.00"))
        # IRPJ: 4.8% de 10000 = 480.00
        self.assertEqual(res_inicial["detalhes"]["irpj"], Decimal("480.00"))
        # CSLL: 2.88% de 10000 = 288.00
        self.assertEqual(res_inicial["detalhes"]["csll"], Decimal("288.00"))
        
        # Total Impostos: 65 + 300 + 480 + 288 = 1133.00
        self.assertEqual(res_inicial["imposto_calculado"], Decimal("1133.00"))

        # 3. Adiciona um ajuste negativo de -2000.00
        ajuste = AjusteDocumento(
            documento_id=doc.id,
            valor_total_ajuste=Decimal("-2000.00"),
            justificativa="Desconto comercial incondicional não faturado"
        )
        self.session.add(ajuste)
        self.session.commit()

        # 4. Novo cálculo (Deve incidir sobre o valor_final = 8000.00)
        res_ajustado = calc.calcular(doc)
        self.assertEqual(res_ajustado["valor_final_base"], Decimal("8000.00"))
        
        # PIS: 0.65% de 8000 = 52.00
        self.assertEqual(res_ajustado["detalhes"]["pis"], Decimal("52.00"))
        # COFINS: 3% de 8000 = 240.00
        self.assertEqual(res_ajustado["detalhes"]["cofins"], Decimal("240.00"))
        # IRPJ: 4.8% de 8000 = 384.00
        self.assertEqual(res_ajustado["detalhes"]["irpj"], Decimal("384.00"))
        # CSLL: 2.88% de 8000 = 230.40 (arredondado para duas casas)
        self.assertEqual(res_ajustado["detalhes"]["csll"], Decimal("230.40"))
        
        # Total: 52 + 240 + 384 + 230.40 = 906.40
        self.assertEqual(res_ajustado["imposto_calculado"], Decimal("906.40"))

    def test_simples_nacional_fator_r_e_faixas(self):
        """Valida o cálculo de alíquotas efetivas por faixas e regras de Fator R no Simples Nacional."""
        # 1. Caso 1: Empresa no Simples Nacional, Anexo III, Faixa 2 (RBT12 = 300.000,00)
        # Alíquota nominal: 11.20% | Parcela a Deduzir: R$ 9.360,00
        # Efetiva: (300000 * 0.112 - 9360) / 300000 = 8.08%
        emp_iii = Empresa(
            cnpj="44444444000199",
            razao_social="Serviços Médicos III Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("300000.00"),
            folha12=Decimal("0.00"), # Sem folha / sem Fator R
            sujeito_fator_r=False
        )
        self.session.add(emp_iii)
        self.session.commit()

        doc_iii = DocumentoFiscal(
            empresa_id=emp_iii.id,
            chave_acesso="33334444555566667777888899990000111122223333",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc_iii)
        self.session.commit()

        calc = CalculadoraFactory.obter_calculadora(emp_iii.regime_tributario)
        res_iii = calc.calcular(doc_iii)
        self.assertEqual(res_iii["aliquota_aplicada"].quantize(Decimal("0.0001")), Decimal("0.0808"))
        self.assertEqual(res_iii["imposto_calculado"], Decimal("80.80")) # 8.08% de 1000.00
        self.assertIn("Anexo III", res_iii["mensagem"])

        # 2. Caso 2: Empresa no Simples Nacional sujeito a Fator R com Fator R < 28% (RBT12 = 300.000, Folha = 30.000 -> 10% Fator R)
        # Deve aplicar Anexo V, Faixa 2
        # Alíquota nominal: 18.00% | Parcela a Deduzir: R$ 4.500,00
        # Efetiva: (300000 * 0.18 - 4500) / 300000 = 16.50%
        emp_v = Empresa(
            cnpj="55555555000199",
            razao_social="Tecnologia V Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("300000.00"),
            folha12=Decimal("300000.00") * Decimal("0.10"), # Fator R = 10%
            sujeito_fator_r=True
        )
        self.session.add(emp_v)
        self.session.commit()

        doc_v = DocumentoFiscal(
            empresa_id=emp_v.id,
            chave_acesso="44445555666677778888999900001111222233334444",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc_v)
        self.session.commit()

        res_v = calc.calcular(doc_v)
        self.assertEqual(res_v["aliquota_aplicada"].quantize(Decimal("0.0001")), Decimal("0.1650"))
        self.assertEqual(res_v["imposto_calculado"], Decimal("165.00")) # 16.50% de 1000.00
        self.assertIn("Anexo V", res_v["mensagem"])

        # 3. Caso 3: Empresa no Simples Nacional sujeito a Fator R com Fator R >= 28% (RBT12 = 300.000, Folha = 90.000 -> 30% Fator R)
        # Deve voltar para o Anexo III, Faixa 2
        # Efetiva: 8.08%
        emp_fator_ok = Empresa(
            cnpj="66666666000199",
            razao_social="Design OK Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("300000.00"),
            folha12=Decimal("300000.00") * Decimal("0.30"), # Fator R = 30%
            sujeito_fator_r=True
        )
        self.session.add(emp_fator_ok)
        self.session.commit()

        doc_fator_ok = DocumentoFiscal(
            empresa_id=emp_fator_ok.id,
            chave_acesso="55556666777788889999000011112222333344445555",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc_fator_ok)
        self.session.commit()

        res_fator_ok = calc.calcular(doc_fator_ok)
        self.assertEqual(res_fator_ok["aliquota_aplicada"].quantize(Decimal("0.0001")), Decimal("0.0808"))
        self.assertEqual(res_fator_ok["imposto_calculado"], Decimal("80.80"))
        self.assertIn("Anexo III", res_fator_ok["mensagem"])

    def test_calculadora_nota_cancelada(self):
        """Verifica se as calculadoras desconsideram e zeram impostos para notas canceladas."""
        emp = Empresa(
            cnpj="77777777000199",
            razao_social="Cancelada Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL
        )
        self.session.add(emp)
        self.session.commit()

        doc = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="77776666777788889999000011112222333344445555",
            tipo_documento="NF-e",
            valor_total=Decimal("1000.00"),
            cstat="101",
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc)
        self.session.commit()

        calc = CalculadoraFactory.obter_calculadora(emp.regime_tributario)
        res = calc.calcular(doc)

        self.assertEqual(res["valor_final_base"], Decimal("0.00"))
        self.assertEqual(res["imposto_calculado"], Decimal("0.00"))
        self.assertIn("CANCELADA", res["mensagem"])

    def test_simples_nacional_iss_retido(self):
        """Verifica se a CalculadoraSimplesNacional realiza o desconto da parcela de ISS para notas com retenção."""
        # Empresa com RBT12 correspondente a Alíquota de 7.3110% (Anexo III, Faixa 2)
        # Fração de ISS na Faixa 2: 32.00%
        emp = Empresa(
            cnpj="88889999000122",
            razao_social="Serviços com Retencao Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("240679.17"),
            folha12=Decimal("0.00"),
            sujeito_fator_r=False,
            categoria_simples="Serviços (Anexo III)"
        )
        self.session.add(emp)
        self.session.commit()

        # Documento com item que possui ISS retido
        doc = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="88886666777788889999000011112222333344445555",
            tipo_documento="NFS-e",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=[
                {
                    "sequencia": 1,
                    "codigo_produto": "140101",
                    "descricao": "RASTREAMENTO",
                    "valor_total": 1000.0,
                    "impostos": {
                        "iss": {
                            "valor_base_calculo": 1000.0,
                            "aliquota": 2.0,
                            "valor": 20.0,
                            "retido": True
                        }
                    }
                }
            ]
        )
        self.session.add(doc)
        self.session.commit()

        calc = CalculadoraFactory.obter_calculadora(emp.regime_tributario)
        res = calc.calcular(doc)

        # Alíquota Efetiva: 7.311005%
        # Com ISS retido, deve ser reduzida em 32.00%:
        # Alíquota final: 7.311005% * 0.68 = 4.971483%
        # Imposto calculado: 1000.00 * 4.971483% = 49.71
        self.assertEqual(res["aliquota_aplicada"].quantize(Decimal("0.000001")), Decimal("0.073110"))
        self.assertEqual(res["imposto_calculado"], Decimal("49.71"))
        self.assertIn("ISS RETIDO", res["mensagem"])
        
        # Testes adicionais da Memória de Cálculo
        self.assertIn("memoria_calculo", res)
        mc = res["memoria_calculo"]
        self.assertEqual(mc["rbt12"], Decimal("240679.17"))
        self.assertEqual(mc["iss_share"], Decimal("32.00"))
        self.assertEqual(mc["aliq_efetiva"], Decimal("7.3110"))
        self.assertEqual(mc["valor_com_iss_retido"], Decimal("1000.00"))

    def test_simples_nacional_anexo_iv(self):
        """Verifica o cálculo de alíquota efetiva e partilha de ISS no Anexo IV."""
        # Empresa do Anexo IV, Faixa 2 (RBT12 = 200k)
        emp = Empresa(
            cnpj="77777777000166",
            razao_social="Limpeza e Conservacao Anexo IV Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("200000.00"),
            folha12=Decimal("0.00"),
            sujeito_fator_r=False,
            categoria_simples="Serviços (Anexo IV)"
        )
        self.session.add(emp)
        self.session.commit()

        # Documento com ISS retido
        doc = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="77776666777788889999000011112222333344445555",
            tipo_documento="NFS-e",
            valor_total=Decimal("10000.00"),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=[
                {
                    "sequencia": 1,
                    "codigo_produto": "0702",
                    "descricao": "LIMPEZA PREDIAL",
                    "valor_total": 10000.0,
                    "impostos": {
                        "iss": {
                            "valor_base_calculo": 10000.0,
                            "aliquota": 2.0,
                            "valor": 200.0,
                            "retido": True
                        }
                    }
                }
            ]
        )
        self.session.add(doc)
        self.session.commit()

        calc = CalculadoraFactory.obter_calculadora(emp.regime_tributario)
        res = calc.calcular(doc)

        # Alíquota efetiva: ((200000 * 0.09) - 8100) / 200000 = 4.95%
        # Com ISS retido, deduz 40.0%: 4.95% * (1 - 0.40) = 2.97%
        # Imposto calculado: 10000 * 2.97% = 297.00
        self.assertEqual(res["aliquota_aplicada"].quantize(Decimal("0.0001")), Decimal("0.0495"))
        self.assertEqual(res["imposto_calculado"], Decimal("297.00"))
        
        mc = res["memoria_calculo"]
        self.assertEqual(mc["iss_share"], Decimal("40.00"))
        self.assertEqual(mc["aliq_efetiva"], Decimal("4.9500"))

    def test_simples_nacional_anexo_ii(self):
        """Verifica o cálculo de alíquota efetiva, partilha de ICMS-ST e IPI no Anexo II (Indústria)."""
        # Empresa da Indústria, Faixa 2 (RBT12 = 200k)
        emp = Empresa(
            cnpj="77777777000188",
            razao_social="Metalurgica Anexo II Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("200000.00"),
            folha12=Decimal("0.00"),
            sujeito_fator_r=False,
            categoria_simples="Indústria (Anexo II)"
        )
        self.session.add(emp)
        self.session.commit()

        # Nota fiscal com ICMS-ST
        doc = DocumentoFiscal(
            empresa_id=emp.id,
            chave_acesso="77776666777788889999000011112222333344445558",
            tipo_documento="NF-e",
            valor_total=Decimal("10000.00"),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=[
                {
                    "sequencia": 1,
                    "codigo_produto": "1001",
                    "descricao": "PRODUTO METALURGICO",
                    "valor_total": 10000.0,
                    "impostos": {
                        "icms": {
                            "cst": "10",
                            "substituicao_tributaria": True
                        }
                    }
                }
            ]
        )
        self.session.add(doc)
        self.session.commit()

        calc = CalculadoraFactory.obter_calculadora(emp.regime_tributario)
        res = calc.calcular(doc)

        # Alíquota efetiva: ((200000 * 0.078) - 5940) / 200000 = 4.83%
        # Com ST, deduz 32.0% de ICMS: 4.83% * (1 - 0.32) = 3.2844%
        # Imposto calculado: 10000 * 3.2844% = 328.44
        self.assertEqual(res["aliquota_aplicada"].quantize(Decimal("0.0001")), Decimal("0.0483"))
        self.assertEqual(res["imposto_calculado"], Decimal("328.44"))
        
        mc = res["memoria_calculo"]
        self.assertEqual(mc["icms_share"], Decimal("32.00"))
        self.assertEqual(mc["ipi_share"], Decimal("7.50"))
        self.assertEqual(mc["aliq_efetiva"], Decimal("4.8300"))

    def test_simples_nacional_anexo_v_fator_r(self):
        """Verifica a aplicação do Anexo V e a transição para Anexo III via Fator R."""
        # Caso A: Fator R < 28% -> Tributado pelo Anexo V (RBT12 = 200k, Folha = 40k -> Fator R = 20%)
        emp_v = Empresa(
            cnpj="99999999000101",
            razao_social="TI Fator R Baixo Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("200000.00"),
            folha12=Decimal("40000.00"),
            sujeito_fator_r=True,
            categoria_simples="Serviços (Anexo V - Fator R)"
        )
        self.session.add(emp_v)
        
        # Caso B: Fator R >= 28% -> Volta para o Anexo III (RBT12 = 200k, Folha = 60k -> Fator R = 30%)
        emp_iii = Empresa(
            cnpj="99999999000102",
            razao_social="TI Fator R Alto Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            rbt12=Decimal("200000.00"),
            folha12=Decimal("60000.00"),
            sujeito_fator_r=True,
            categoria_simples="Serviços (Anexo V - Fator R)"
        )
        self.session.add(emp_iii)
        self.session.commit()

        # Nota fiscal com ISS Retido (Caso A)
        doc_v = DocumentoFiscal(
            empresa_id=emp_v.id,
            chave_acesso="77776666777788889999000011112222333344445559",
            tipo_documento="NFS-e",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=[
                {
                    "sequencia": 1,
                    "codigo_produto": "0101",
                    "descricao": "DESENVOLVIMENTO",
                    "valor_total": 1000.0,
                    "impostos": {
                        "iss": {
                            "retido": True
                        }
                    }
                }
            ]
        )
        self.session.add(doc_v)

        # Nota fiscal sem retenção (Caso B)
        doc_iii = DocumentoFiscal(
            empresa_id=emp_iii.id,
            chave_acesso="77776666777788889999000011112222333344445560",
            tipo_documento="NFS-e",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE
        )
        self.session.add(doc_iii)
        self.session.commit()

        calc = CalculadoraFactory.obter_calculadora(RegimeTributario.SIMPLES_NACIONAL)
        
        # Teste Caso A:
        # AE Anexo V, Faixa 2: ((200000 * 0.18) - 4500) / 200000 = 15.75%
        # Com ISS retido, deduz 17.0% de ISS: 15.75% * (1 - 0.17) = 13.0725%
        # Imposto: 1000.00 * 13.0725% = 130.725 -> rounds to 130.72 (ROUND_HALF_EVEN)
        res_v = calc.calcular(doc_v)
        self.assertEqual(res_v["aliquota_aplicada"].quantize(Decimal("0.0001")), Decimal("0.1575"))
        self.assertEqual(res_v["imposto_calculado"], Decimal("130.72"))

        # Teste Caso B:
        # AE Anexo III, Faixa 2: ((200000 * 0.112) - 9360) / 200000 = 6.52%
        # Sem retenção: 1000.00 * 6.52% = 65.20
        res_iii = calc.calcular(doc_iii)
        self.assertEqual(res_iii["aliquota_aplicada"].quantize(Decimal("0.0001")), Decimal("0.0652"))
        self.assertEqual(res_iii["imposto_calculado"], Decimal("65.20"))

    def test_calculo_difal_e_impostos_entrada(self):
        """Valida que notas de Entrada calculam apenas DIFAL (com Base Simples/Dupla e IPI) e ICMS-ST de compra."""
        # 1. Caso Base Dupla com IPI (Bahia)
        emp_ba = Empresa(
            cnpj="12345678000199",
            razao_social="Simples Comércio BA Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            uf="BA"
        )
        self.session.add(emp_ba)
        self.session.commit()

        # Nota de entrada interestadual (origem São Paulo "35" -> destino Bahia "BA")
        # Valor total: 1000.00, IPI: 50.00, ICMS destacado (Origem): 70.00
        # Alíquota interestadual: 7% (SP -> BA)
        # Alíquota interna BA: 20.5% (Base Dupla)
        # Base Líquida = 1000.00 - 0.00 + 0.00 + 50.00 = 1050.00
        # Valor Sem ICMS = 1050.00 - 70.00 = 980.00
        # divisor = 1 - 0.205 = 0.795
        # Base DIFAL = 980.00 / 0.795 = 1232.7044...
        # DIFAL = (1232.7044... * 0.205) - 70.00 = 252.7044... - 70.00 = 182.70
        # ICMS-ST destacado de compra: 50.00
        doc_ba = DocumentoFiscal(
            empresa_id=emp_ba.id,
            chave_acesso="35230512345678000199550010000001231234567890", # SP code = 35
            tipo_documento="NF-e",
            tipo_operacao="Entrada",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=[{
                "sequencia": 1,
                "codigo_produto": "PROD001",
                "descricao": "Item Interestadual BA",
                "valor_total": 1000.00,
                "desconto": 0.0,
                "frete": 0.0,
                "valor_ipi": 50.00,
                "impostos": {
                    "icms": {
                        "cst": "00",
                        "valor_st": 50.00,
                        "valor": 70.00
                    }
                }
            }]
        )
        self.session.add(doc_ba)
        self.session.commit()

        calc_ba = CalculadoraFactory.obter_calculadora(emp_ba.regime_tributario)
        res_ba = calc_ba.calcular(doc_ba)

        self.assertEqual(res_ba["detalhes"]["difal"], Decimal("182.70"))
        self.assertEqual(res_ba["detalhes"]["icms_st_compra"], Decimal("50.00"))
        self.assertEqual(res_ba["imposto_calculado"], Decimal("232.70")) # 182.70 + 50.00
        self.assertIn("Nota Fiscal de Entrada (Compra)", res_ba["mensagem"])

        # 2. Caso Base Simples (Pernambuco - PE, não está em ufs_base_dupla)
        emp_pe = Empresa(
            cnpj="12345678000188",
            razao_social="Simples Comércio PE Ltda",
            regime_tributario=RegimeTributario.SIMPLES_NACIONAL,
            uf="PE"
        )
        self.session.add(emp_pe)
        self.session.commit()

        # Nota de entrada interestadual (origem São Paulo "35" -> destino Pernambuco "PE")
        # Base Líquida = 1000.00 - 0.00 + 0.00 + 50.00 = 1050.00
        # Alíquota interestadual: 7% (SP -> PE)
        # Alíquota interna PE: 18% (Base Simples)
        # DIFAL = 1050.00 * (18% - 7%) = 1050.00 * 11% = 115.50
        doc_pe = DocumentoFiscal(
            empresa_id=emp_pe.id,
            chave_acesso="35230512345678000188550010000001231234567890", # SP code = 35
            tipo_documento="NF-e",
            tipo_operacao="Entrada",
            valor_total=Decimal("1000.00"),
            status_apuracao=StatusApuracao.PENDENTE,
            itens=[{
                "sequencia": 1,
                "codigo_produto": "PROD002",
                "descricao": "Item Interestadual PE",
                "valor_total": 1000.00,
                "desconto": 0.0,
                "frete": 0.0,
                "valor_ipi": 50.00,
                "impostos": {
                    "icms": {
                        "cst": "00",
                        "valor_st": 0.00,
                        "valor": 70.00
                    }
                }
            }]
        )
        self.session.add(doc_pe)
        self.session.commit()

        calc_pe = CalculadoraFactory.obter_calculadora(emp_pe.regime_tributario)
        res_pe = calc_pe.calcular(doc_pe)

        self.assertEqual(res_pe["detalhes"]["difal"], Decimal("115.50"))

if __name__ == "__main__":
    unittest.main()

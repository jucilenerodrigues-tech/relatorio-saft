"""
Testes do apuramento. Correr com: python -m unittest

Cada teste monta um SAF-T mínimo em memória e confirma uma regra:
sinal das linhas, exclusão de anulados, numeração e conferência dos totais.
"""

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from analisar_saft import apurar, ler_saft, reunir_avisos

MODELO = """<?xml version="1.0" encoding="UTF-8"?>
<AuditFile xmlns="urn:OECD:StandardAuditFile-Tax:PT_1.04_01">
  <Header>
    <AuditFileVersion>1.04_01</AuditFileVersion>
    <TaxRegistrationNumber>500000000</TaxRegistrationNumber>
    <CompanyName>Empresa de Teste, Lda.</CompanyName>
    <FiscalYear>2026</FiscalYear>
    <StartDate>2026-01-01</StartDate>
    <EndDate>2026-12-31</EndDate>
    <CurrencyCode>EUR</CurrencyCode>
  </Header>
  <MasterFiles>
    <Customer>
      <CustomerID>C001</CustomerID>
      <CustomerTaxID>501111111</CustomerTaxID>
      <CompanyName>Cliente Um, Lda.</CompanyName>
    </Customer>
    <Customer>
      <CustomerID>C999</CustomerID>
      <CustomerTaxID>999999990</CustomerTaxID>
      <CompanyName>Consumidor final</CompanyName>
    </Customer>
  </MasterFiles>
  <SourceDocuments>
    <SalesInvoices>
      <NumberOfEntries>{numero}</NumberOfEntries>
      <TotalDebit>{debito}</TotalDebit>
      <TotalCredit>{credito}</TotalCredit>
      {documentos}
    </SalesInvoices>
  </SourceDocuments>
</AuditFile>
"""

LINHA = """
        <Line>
          <LineNumber>{n}</LineNumber>
          <ProductCode>A1</ProductCode>
          <ProductDescription>Artigo</ProductDescription>
          <Quantity>1.00</Quantity>
          <UnitOfMeasure>UN</UnitOfMeasure>
          <UnitPrice>{base}</UnitPrice>
          <Description>Artigo</Description>
          <{lado}Amount>{base}</{lado}Amount>
          <Tax>
            <TaxType>IVA</TaxType>
            <TaxCountryRegion>PT</TaxCountryRegion>
            <TaxCode>NOR</TaxCode>
            <TaxPercentage>23.00</TaxPercentage>
          </Tax>
        </Line>"""

DOCUMENTO = """
      <Invoice>
        <InvoiceNo>{numero}</InvoiceNo>
        <DocumentStatus>
          <InvoiceStatus>{estado}</InvoiceStatus>
        </DocumentStatus>
        <Period>1</Period>
        <InvoiceDate>2026-01-10</InvoiceDate>
        <InvoiceType>{tipo}</InvoiceType>
        <CustomerID>{cliente}</CustomerID>{linhas}
        <DocumentTotals>
          <TaxPayable>{imposto}</TaxPayable>
          <NetTotal>{base}</NetTotal>
          <GrossTotal>{total}</GrossTotal>
        </DocumentTotals>
      </Invoice>"""


def documento(numero, tipo, linhas, total=None, estado="N", cliente="C001"):
    """linhas: lista de (valor, "Credit" ou "Debit")."""
    corpo = "".join(
        LINHA.format(n=indice, base=f"{base:.2f}", lado=lado)
        for indice, (base, lado) in enumerate(linhas, start=1)
    )
    liquido = sum(b if l == "Credit" else -b for b, l in linhas)
    imposto = (abs(liquido) * Decimal("0.23")).quantize(Decimal("0.01"))
    bruto = abs(liquido) + imposto
    return DOCUMENTO.format(
        numero=numero,
        tipo=tipo,
        estado=estado,
        cliente=cliente,
        linhas=corpo,
        base=f"{abs(liquido):.2f}",
        imposto=f"{imposto:.2f}",
        total=f"{bruto if total is None else total:.2f}",
    )


def apurar_ficheiro(documentos, credito, debito, numero=None):
    conteudo = MODELO.format(
        numero=len(documentos) if numero is None else numero,
        credito=f"{credito:.2f}",
        debito=f"{debito:.2f}",
        documentos="".join(documentos),
    )
    with tempfile.TemporaryDirectory() as pasta:
        caminho = Path(pasta) / "teste.xml"
        caminho.write_text(conteudo, encoding="utf-8")
        dados = ler_saft(caminho)
    apuramento = apurar(dados)
    return apuramento, reunir_avisos(apuramento)


def base_total(apuramento):
    return sum(m["base"] for m in apuramento["por_mes"].values())


class TesteSinalDasLinhas(unittest.TestCase):
    def teste_nota_de_credito_reduz_a_faturacao(self):
        documentos = [
            documento("FT 2026/1", "FT", [(Decimal("100"), "Credit")]),
            documento("NC 2026/1", "NC", [(Decimal("40"), "Debit")]),
        ]
        apuramento, avisos = apurar_ficheiro(documentos, credito=100, debito=40)
        self.assertEqual(base_total(apuramento), Decimal("60"))
        self.assertEqual(avisos, [])

    def teste_linha_a_debito_numa_fatura_tambem_subtrai(self):
        """Um desconto lançado como linha a débito dentro de uma fatura."""
        documentos = [
            documento(
                "FT 2026/1",
                "FT",
                [(Decimal("100"), "Credit"), (Decimal("25"), "Debit")],
            )
        ]
        apuramento, avisos = apurar_ficheiro(documentos, credito=100, debito=25)
        self.assertEqual(base_total(apuramento), Decimal("75"))
        self.assertEqual(avisos, [])


class TesteAnulados(unittest.TestCase):
    def teste_anulado_fica_fora_das_somas(self):
        documentos = [
            documento("FT 2026/1", "FT", [(Decimal("100"), "Credit")]),
            documento("FT 2026/2", "FT", [(Decimal("50"), "Credit")], estado="A"),
        ]
        apuramento, _ = apurar_ficheiro(documentos, credito=100, debito=0, numero=2)
        self.assertEqual(base_total(apuramento), Decimal("100"))
        self.assertEqual(len(apuramento["anulados"]), 1)

    def teste_anulado_nao_abre_falha_na_numeracao(self):
        documentos = [
            documento("FT 2026/1", "FT", [(Decimal("10"), "Credit")]),
            documento("FT 2026/2", "FT", [(Decimal("10"), "Credit")], estado="A"),
            documento("FT 2026/3", "FT", [(Decimal("10"), "Credit")]),
        ]
        apuramento, _ = apurar_ficheiro(documentos, credito=20, debito=0, numero=3)
        self.assertEqual(apuramento["falhas_numeracao"], [])


class TesteConferencias(unittest.TestCase):
    def teste_numero_em_falta_na_serie(self):
        documentos = [
            documento("FT 2026/1", "FT", [(Decimal("10"), "Credit")]),
            documento("FT 2026/3", "FT", [(Decimal("10"), "Credit")]),
        ]
        apuramento, _ = apurar_ficheiro(documentos, credito=20, debito=0)
        self.assertEqual(apuramento["falhas_numeracao"][0]["em_falta"], [2])

    def teste_documento_que_nao_fecha(self):
        documentos = [
            documento("FT 2026/1", "FT", [(Decimal("100"), "Credit")], total=200)
        ]
        apuramento, _ = apurar_ficheiro(documentos, credito=100, debito=0)
        self.assertEqual(len(apuramento["totais_errados"]), 1)

    def teste_total_de_controlo_errado(self):
        documentos = [documento("FT 2026/1", "FT", [(Decimal("100"), "Credit")])]
        _, avisos = apurar_ficheiro(documentos, credito=999, debito=0)
        self.assertTrue(any(a["tipo"] == "Total a crédito" for a in avisos))

    def teste_fatura_sem_nif_acima_do_limite(self):
        documentos = [
            documento(
                "FT 2026/1", "FT", [(Decimal("2000"), "Credit")], cliente="C999"
            )
        ]
        _, avisos = apurar_ficheiro(documentos, credito=2000, debito=0)
        self.assertTrue(any(a["tipo"] == "Fatura sem NIF do cliente" for a in avisos))

    def teste_ficheiro_sem_problemas_nao_gera_avisos(self):
        documentos = [
            documento("FT 2026/1", "FT", [(Decimal("100"), "Credit")]),
            documento("FT 2026/2", "FT", [(Decimal("200"), "Credit")]),
        ]
        _, avisos = apurar_ficheiro(documentos, credito=300, debito=0)
        self.assertEqual(avisos, [])


if __name__ == "__main__":
    unittest.main()

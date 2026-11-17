"""
Gera um ficheiro SAF-T (PT) fictício, na estrutura da versão 1.04_01.

Serve para que qualquer pessoa possa experimentar o analisador sem precisar
de um ficheiro real. Todos os dados são inventados: empresa, NIF, clientes,
artigos e valores.

A semente aleatória é fixa, por isso o ficheiro gerado é sempre o mesmo.

Aviso: as assinaturas (Hash) e os códigos ATCUD são texto de preenchimento.
Num ficheiro real são produzidos por software certificado pela AT.
"""

from __future__ import annotations

import random
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree as ET

SEMENTE = 20260101
ANO = 2026
NAMESPACE = "urn:OECD:StandardAuditFile-Tax:PT_1.04_01"
FICHEIRO_SAIDA = Path("dados") / f"saft_exemplo_{ANO}.xml"

EMPRESA = {
    "nif": "508431276",
    "nome": "Casa das Ferragens Monteiro, Lda.",
    "morada": "Rua da Estação, 142",
    "localidade": "Braga",
    "codigo_postal": "4700-223",
}

# Taxas de IVA do continente em vigor.
TAXAS = {
    "NOR": Decimal("23.00"),
    "INT": Decimal("13.00"),
    "RED": Decimal("6.00"),
    "ISE": Decimal("0.00"),
}

DESCRICAO_TAXAS = {
    "NOR": "Taxa normal",
    "INT": "Taxa intermédia",
    "RED": "Taxa reduzida",
    "ISE": "Isenta",
}

# ProductType: P = produto, S = serviço.
ARTIGOS = [
    ("A001", "P", "Parafuso inox M6 (caixa 100 un.)", "UN", Decimal("12.40"), "NOR"),
    ("A002", "P", "Tinta acrílica branca 5 L", "UN", Decimal("28.90"), "NOR"),
    ("A003", "P", "Berbequim percutor 750 W", "UN", Decimal("89.00"), "NOR"),
    ("A004", "P", "Fita isoladora 19 mm", "UN", Decimal("1.85"), "NOR"),
    ("A005", "P", "Escada alumínio 6 degraus", "UN", Decimal("64.50"), "NOR"),
    ("A006", "P", "Luvas de proteção (par)", "PAR", Decimal("4.20"), "NOR"),
    ("A007", "P", "Cimento cola 25 kg", "SC", Decimal("9.75"), "RED"),
    ("A008", "P", "Argamassa de reboco 30 kg", "SC", Decimal("7.30"), "RED"),
    ("S001", "S", "Corte de chave", "UN", Decimal("3.50"), "NOR"),
    ("S002", "S", "Afiação de ferramenta", "UN", Decimal("6.00"), "NOR"),
    ("S003", "S", "Aluguer de martelo demolidor (dia)", "DIA", Decimal("35.00"), "NOR"),
    ("S004", "S", "Transporte e entrega ao domicílio", "UN", Decimal("15.00"), "INT"),
    ("S005", "S", "Formação em segurança no trabalho", "HORA", Decimal("45.00"), "ISE"),
]

CLIENTES_EMPRESA = [
    ("C001", "501234562", "Construções Vale Verde, Lda.", "Braga"),
    ("C002", "502987654", "Remodelações Silva & Pinto, Lda.", "Guimarães"),
    ("C003", "503445127", "Eletricidade Norte, Unipessoal Lda.", "Barcelos"),
    ("C004", "504778391", "Canalizações Minho, Lda.", "Braga"),
    ("C005", "505112268", "Condomínio Edifício Sol Nascente", "Braga"),
    ("C006", "506334419", "Carpintaria Almeida, Lda.", "Famalicão"),
    ("C007", "507221845", "Pinturas Rápidas do Cávado, Lda.", "Braga"),
    ("C008", "508669032", "Serralharia Castro, Lda.", "Póvoa de Lanhoso"),
    ("C009", "509887713", "Imobiliária Casa Boa, Lda.", "Braga"),
    ("C010", "510443926", "Jardins e Espaços Verdes Lima, Lda.", "Amares"),
]

# NIF genérico de consumidor final, usado nas faturas simplificadas.
CLIENTE_FINAL = ("C999", "999999990", "Consumidor final", "Desconhecido")

ISENCAO = ("M07", "Isento artigo 9.º do CIVA (ou similar)")

CENTIMO = Decimal("0.01")


def arredondar(valor: Decimal) -> Decimal:
    """Arredonda a duas casas, como se arredonda dinheiro."""
    return valor.quantize(CENTIMO, rounding=ROUND_HALF_UP)


def texto(pai: ET.Element, etiqueta: str, valor) -> ET.Element:
    elemento = ET.SubElement(pai, etiqueta)
    elemento.text = str(valor)
    return elemento


def hash_ficticio(aleatorio: random.Random) -> str:
    """Devolve texto com o aspeto de uma assinatura, sem o ser."""
    alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    return "".join(aleatorio.choice(alfabeto) for _ in range(172)) + "=="


def dias_uteis_do_ano(ano: int) -> list[date]:
    dia = date(ano, 1, 1)
    fim = date(ano, 12, 31)
    dias = []
    while dia <= fim:
        if dia.weekday() < 6:  # a loja abre de segunda a sábado
            dias.append(dia)
        dia += timedelta(days=1)
    return dias


def construir_linhas(aleatorio: random.Random, numero_linhas: int) -> list[dict]:
    """Cria as linhas de um documento e calcula base e imposto de cada uma."""
    linhas = []
    for numero in range(1, numero_linhas + 1):
        codigo, _, descricao, unidade, preco, codigo_taxa = aleatorio.choice(ARTIGOS)
        quantidade = Decimal(aleatorio.choice([1, 1, 1, 2, 2, 3, 4, 5, 10]))
        base = arredondar(preco * quantidade)
        percentagem = TAXAS[codigo_taxa]
        imposto = arredondar(base * percentagem / Decimal("100"))
        linhas.append(
            {
                "numero": numero,
                "codigo": codigo,
                "descricao": descricao,
                "unidade": unidade,
                "preco": preco,
                "quantidade": quantidade,
                "base": base,
                "codigo_taxa": codigo_taxa,
                "percentagem": percentagem,
                "imposto": imposto,
            }
        )
    return linhas


def gerar_documentos(aleatorio: random.Random) -> list[dict]:
    """
    Produz os documentos de venda do ano.

    Inclui, de propósito, quatro situações que o analisador deve apanhar:
      1. documentos anulados;
      2. uma falha na numeração da série FT;
      3. uma fatura (FT) emitida a consumidor final acima de 1 000 EUR;
      4. um documento cujas linhas não somam o total declarado.
    """
    dias = dias_uteis_do_ano(ANO)
    documentos = []
    contadores = {"FT": 0, "FS": 0, "FR": 0, "NC": 0}

    # Números da série FT reservados para as situações especiais.
    ft_com_falha = 118          # este número nunca é emitido
    ft_consumidor_final = 205   # fatura acima de 1 000 EUR sem NIF do cliente
    ft_total_errado = 311       # linhas não somam o total declarado

    for dia in dias:
        # Mais movimento no fim do ano e menos em agosto.
        documentos_do_dia = aleatorio.randint(2, 6)
        if dia.month == 8:
            documentos_do_dia = max(1, documentos_do_dia - 2)
        if dia.month in (11, 12):
            documentos_do_dia += 1

        for _ in range(documentos_do_dia):
            # A maioria das vendas ao balcão é fatura simplificada.
            tipo = aleatorio.choices(
                ["FS", "FT", "FR", "NC"], weights=[55, 30, 12, 3], k=1
            )[0]
            contadores[tipo] += 1
            if tipo == "FT" and contadores[tipo] == ft_com_falha:
                contadores[tipo] += 1  # salta um número da série

            numero_serie = contadores[tipo]
            numero_documento = f"{tipo} {ANO}/{numero_serie}"

            if tipo == "FS":
                cliente = CLIENTE_FINAL
                linhas = construir_linhas(aleatorio, aleatorio.randint(1, 3))
            else:
                cliente = aleatorio.choice(CLIENTES_EMPRESA)
                linhas = construir_linhas(aleatorio, aleatorio.randint(1, 6))

            # Situação 3: fatura a consumidor final acima de 1 000 EUR.
            if tipo == "FT" and numero_serie == ft_consumidor_final:
                cliente = CLIENTE_FINAL
                linhas = construir_linhas(aleatorio, 6)
                while sum(l["base"] + l["imposto"] for l in linhas) < Decimal("1000"):
                    linhas += construir_linhas(aleatorio, 2)
                for indice, linha in enumerate(linhas, start=1):
                    linha["numero"] = indice

            base_total = arredondar(sum(l["base"] for l in linhas))
            imposto_total = arredondar(sum(l["imposto"] for l in linhas))
            total_declarado = arredondar(base_total + imposto_total)

            # Situação 4: o total declarado não bate com a soma das linhas.
            if tipo == "FT" and numero_serie == ft_total_errado:
                total_declarado = arredondar(total_declarado + Decimal("0.02"))

            # Situação 1: alguns documentos anulados ao longo do ano.
            anulado = aleatorio.random() < 0.004
            hora = f"{aleatorio.randint(9, 19):02d}:{aleatorio.randint(0, 59):02d}:{aleatorio.randint(0, 59):02d}"

            documentos.append(
                {
                    "numero": numero_documento,
                    "tipo": tipo,
                    "serie": f"{tipo} {ANO}",
                    "sequencia": numero_serie,
                    "data": dia,
                    "periodo": dia.month,
                    "cliente": cliente,
                    "linhas": linhas,
                    "base": base_total,
                    "imposto": imposto_total,
                    "total": total_declarado,
                    "anulado": anulado,
                    "hora": hora,
                }
            )

    documentos.sort(key=lambda d: (d["data"], d["tipo"], d["sequencia"]))
    return documentos


def escrever_xml(documentos: list[dict], aleatorio: random.Random) -> ET.ElementTree:
    ET.register_namespace("", NAMESPACE)
    raiz = ET.Element(f"{{{NAMESPACE}}}AuditFile")

    # ------------------------------------------------------------------ Header
    cabecalho = ET.SubElement(raiz, "Header")
    texto(cabecalho, "AuditFileVersion", "1.04_01")
    texto(cabecalho, "CompanyID", EMPRESA["nif"])
    texto(cabecalho, "TaxRegistrationNumber", EMPRESA["nif"])
    texto(cabecalho, "TaxAccountingBasis", "F")  # F = faturação
    texto(cabecalho, "CompanyName", EMPRESA["nome"])
    morada = ET.SubElement(cabecalho, "CompanyAddress")
    texto(morada, "AddressDetail", EMPRESA["morada"])
    texto(morada, "City", EMPRESA["localidade"])
    texto(morada, "PostalCode", EMPRESA["codigo_postal"])
    texto(morada, "Country", "PT")
    texto(cabecalho, "FiscalYear", ANO)
    texto(cabecalho, "StartDate", f"{ANO}-01-01")
    texto(cabecalho, "EndDate", f"{ANO}-12-31")
    texto(cabecalho, "CurrencyCode", "EUR")
    texto(cabecalho, "DateCreated", date(ANO + 1, 1, 15).isoformat())
    texto(cabecalho, "TaxEntity", "Global")
    texto(cabecalho, "ProductCompanyTaxID", "999999990")
    texto(cabecalho, "SoftwareCertificateNumber", "0")
    texto(cabecalho, "ProductID", "Exemplo Ficticio/Exemplo Ficticio")
    texto(cabecalho, "ProductVersion", "1.0")

    # ------------------------------------------------------------- MasterFiles
    ficheiros_mestre = ET.SubElement(raiz, "MasterFiles")

    for codigo, nif, nome, localidade in CLIENTES_EMPRESA + [CLIENTE_FINAL]:
        cliente = ET.SubElement(ficheiros_mestre, "Customer")
        texto(cliente, "CustomerID", codigo)
        texto(cliente, "AccountID", "Desconhecido")
        texto(cliente, "CustomerTaxID", nif)
        texto(cliente, "CompanyName", nome)
        endereco = ET.SubElement(cliente, "BillingAddress")
        texto(endereco, "AddressDetail", "Desconhecido")
        texto(endereco, "City", localidade)
        texto(endereco, "PostalCode", "Desconhecido")
        texto(endereco, "Country", "PT")
        texto(cliente, "SelfBillingIndicator", "0")

    for codigo, tipo_artigo, descricao, _, _, _ in ARTIGOS:
        artigo = ET.SubElement(ficheiros_mestre, "Product")
        texto(artigo, "ProductType", tipo_artigo)
        texto(artigo, "ProductCode", codigo)
        texto(artigo, "ProductDescription", descricao)
        texto(artigo, "ProductNumberCode", codigo)

    tabela_impostos = ET.SubElement(ficheiros_mestre, "TaxTable")
    for codigo_taxa, percentagem in TAXAS.items():
        entrada = ET.SubElement(tabela_impostos, "TaxTableEntry")
        texto(entrada, "TaxType", "IVA")
        texto(entrada, "TaxCountryRegion", "PT")
        texto(entrada, "TaxCode", codigo_taxa)
        texto(entrada, "Description", DESCRICAO_TAXAS[codigo_taxa])
        texto(entrada, "TaxPercentage", f"{percentagem:.2f}")

    # ---------------------------------------------------------- SourceDocuments
    documentos_origem = ET.SubElement(raiz, "SourceDocuments")
    vendas = ET.SubElement(documentos_origem, "SalesInvoices")

    # Nas notas de crédito os valores entram a débito; nos restantes, a crédito.
    # Os totais de controlo excluem os documentos anulados, mas o número de
    # registos inclui-os (Portaria n.º 302/2016, campos 4.1.1 a 4.1.3).
    validos = [d for d in documentos if not d["anulado"]]
    total_credito = arredondar(sum(d["base"] for d in validos if d["tipo"] != "NC"))
    total_debito = arredondar(sum(d["base"] for d in validos if d["tipo"] == "NC"))

    texto(vendas, "NumberOfEntries", len(documentos))
    texto(vendas, "TotalDebit", f"{total_debito:.2f}")
    texto(vendas, "TotalCredit", f"{total_credito:.2f}")

    for documento in documentos:
        codigo_cliente, _, _, _ = documento["cliente"]
        fatura = ET.SubElement(vendas, "Invoice")
        texto(fatura, "InvoiceNo", documento["numero"])
        texto(fatura, "ATCUD", f"JFT7PQ2X-{documento['sequencia']}")

        estado = ET.SubElement(fatura, "DocumentStatus")
        texto(estado, "InvoiceStatus", "A" if documento["anulado"] else "N")
        texto(
            estado,
            "InvoiceStatusDate",
            f"{documento['data'].isoformat()}T{documento['hora']}",
        )
        texto(estado, "SourceID", "Balcao")
        texto(estado, "SourceBilling", "P")  # P = produzido na aplicação

        texto(fatura, "Hash", hash_ficticio(aleatorio))
        texto(fatura, "HashControl", "1")
        texto(fatura, "Period", documento["periodo"])
        texto(fatura, "InvoiceDate", documento["data"].isoformat())
        texto(fatura, "InvoiceType", documento["tipo"])

        regimes = ET.SubElement(fatura, "SpecialRegimes")
        texto(regimes, "SelfBillingIndicator", "0")
        texto(regimes, "CashVATSchemeIndicator", "0")
        texto(regimes, "ThirdPartiesBillingIndicator", "0")

        texto(fatura, "SourceID", "Balcao")
        texto(
            fatura,
            "SystemEntryDate",
            f"{documento['data'].isoformat()}T{documento['hora']}",
        )
        texto(fatura, "CustomerID", codigo_cliente)

        for linha in documento["linhas"]:
            elemento = ET.SubElement(fatura, "Line")
            texto(elemento, "LineNumber", linha["numero"])
            texto(elemento, "ProductCode", linha["codigo"])
            texto(elemento, "ProductDescription", linha["descricao"])
            texto(elemento, "Quantity", f"{linha['quantidade']:.2f}")
            texto(elemento, "UnitOfMeasure", linha["unidade"])
            texto(elemento, "UnitPrice", f"{linha['preco']:.2f}")
            texto(elemento, "TaxPointDate", documento["data"].isoformat())
            texto(elemento, "Description", linha["descricao"])
            # Nota de crédito devolve valor ao cliente: entra a débito.
            campo_valor = "DebitAmount" if documento["tipo"] == "NC" else "CreditAmount"
            texto(elemento, campo_valor, f"{linha['base']:.2f}")

            imposto = ET.SubElement(elemento, "Tax")
            texto(imposto, "TaxType", "IVA")
            texto(imposto, "TaxCountryRegion", "PT")
            texto(imposto, "TaxCode", linha["codigo_taxa"])
            texto(imposto, "TaxPercentage", f"{linha['percentagem']:.2f}")

            if linha["codigo_taxa"] == "ISE":
                texto(elemento, "TaxExemptionReason", ISENCAO[1])
                texto(elemento, "TaxExemptionCode", ISENCAO[0])

        totais = ET.SubElement(fatura, "DocumentTotals")
        texto(totais, "TaxPayable", f"{documento['imposto']:.2f}")
        texto(totais, "NetTotal", f"{documento['base']:.2f}")
        texto(totais, "GrossTotal", f"{documento['total']:.2f}")

    return ET.ElementTree(raiz)


def principal() -> None:
    aleatorio = random.Random(SEMENTE)
    documentos = gerar_documentos(aleatorio)
    arvore = escrever_xml(documentos, aleatorio)

    FICHEIRO_SAIDA.parent.mkdir(parents=True, exist_ok=True)
    ET.indent(arvore, space="  ")
    arvore.write(FICHEIRO_SAIDA, encoding="UTF-8", xml_declaration=True)

    tamanho = FICHEIRO_SAIDA.stat().st_size / (1024 * 1024)
    anulados = sum(1 for d in documentos if d["anulado"])
    print(f"Ficheiro gerado: {FICHEIRO_SAIDA}")
    print(f"  Documentos: {len(documentos)} ({anulados} anulados)")
    print(f"  Tamanho: {tamanho:.1f} MB")


if __name__ == "__main__":
    principal()

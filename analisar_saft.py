"""
Lê um ficheiro SAF-T (PT) de faturação e produz um relatório em Excel.

O relatório responde ao que o gabinete precisa de saber sobre o mês ou o ano
de um cliente: quanto faturou, com que IVA, a quem vendeu, e o que no ficheiro
merece uma segunda vista antes de servir de base à declaração.

Utilização:
    python analisar_saft.py
    python analisar_saft.py --ficheiro o_meu_saft.xml --saida resultado/cliente.xlsx

Estrutura suportada: SAF-T (PT) versão 1.04_01 (Portaria n.º 302/2016).
"""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FICHEIRO_PREDEFINIDO = Path("dados") / "saft_exemplo_2026.xml"
SAIDA_PREDEFINIDA = Path("resultado") / "relatorio_saft.xlsx"

NOMES_TIPOS = {
    "FT": "Fatura",
    "FS": "Fatura simplificada",
    "FR": "Fatura-recibo",
    "ND": "Nota de débito",
    "NC": "Nota de crédito",
}

NIF_CONSUMIDOR_FINAL = "999999990"
LIMITE_FATURA_SEM_NIF = Decimal("1000.00")
TOLERANCIA = Decimal("0.01")
CENTIMO = Decimal("0.01")

# Estados que excluem o documento das somas: A = anulado, F = já faturado.
# O número de registos do ficheiro conta-os à mesma.
ESTADOS_EXCLUIDOS = {"A", "F"}

MESES = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

ZERO = Decimal("0.00")


# --------------------------------------------------------------------- leitura

def sem_namespace(etiqueta: str) -> str:
    """Devolve o nome da etiqueta sem o namespace do SAF-T."""
    return etiqueta.rsplit("}", 1)[-1]


def valor(elemento: ET.Element | None, etiqueta: str, predefinido: str = "") -> str:
    if elemento is None:
        return predefinido
    for filho in elemento:
        if sem_namespace(filho.tag) == etiqueta:
            return (filho.text or "").strip()
    return predefinido


def para_decimal(texto: str) -> Decimal:
    if not texto:
        return ZERO
    try:
        return Decimal(texto.replace(",", "."))
    except InvalidOperation:
        return ZERO


def para_data(texto: str) -> date | None:
    try:
        return datetime.strptime(texto[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def ler_saft(caminho: Path) -> dict:
    """
    Percorre o ficheiro uma única vez e liberta cada documento da memória
    depois de o ler. Um SAF-T anual pode ter centenas de megabytes.
    """
    cabecalho: dict[str, str] = {}
    clientes: dict[str, dict] = {}
    taxas: dict[str, Decimal] = {}
    documentos: list[dict] = []
    declarados = {"numero": 0, "credito": ZERO, "debito": ZERO}

    # O "start" diz quando entramos na secção da faturação: um SAF-T completo
    # tem contagens e totais noutras secções, que não são estes.
    contexto = ET.iterparse(caminho, events=("start", "end"))
    seccao_vendas: ET.Element | None = None

    for evento, elemento in contexto:
        etiqueta = sem_namespace(elemento.tag)

        if evento == "start":
            if etiqueta == "SalesInvoices":
                seccao_vendas = elemento
            continue

        if etiqueta == "Header":
            morada = None
            for filho in elemento:
                if sem_namespace(filho.tag) == "CompanyAddress":
                    morada = filho
            cabecalho = {
                "versao": valor(elemento, "AuditFileVersion"),
                "nif": valor(elemento, "TaxRegistrationNumber"),
                "empresa": valor(elemento, "CompanyName"),
                "localidade": valor(morada, "City"),
                "ano": valor(elemento, "FiscalYear"),
                "inicio": valor(elemento, "StartDate"),
                "fim": valor(elemento, "EndDate"),
                "moeda": valor(elemento, "CurrencyCode"),
                "programa": valor(elemento, "ProductID"),
            }
            elemento.clear()

        elif etiqueta == "Customer":
            codigo = valor(elemento, "CustomerID")
            clientes[codigo] = {
                "nif": valor(elemento, "CustomerTaxID"),
                "nome": valor(elemento, "CompanyName"),
            }
            elemento.clear()

        elif etiqueta == "TaxTableEntry":
            taxas[valor(elemento, "TaxCode")] = para_decimal(
                valor(elemento, "TaxPercentage")
            )
            elemento.clear()

        elif etiqueta == "NumberOfEntries" and seccao_vendas is not None:
            declarados["numero"] = int(elemento.text or 0)

        elif etiqueta == "TotalCredit" and seccao_vendas is not None:
            declarados["credito"] = para_decimal(elemento.text or "")

        elif etiqueta == "TotalDebit" and seccao_vendas is not None:
            declarados["debito"] = para_decimal(elemento.text or "")

        elif etiqueta == "Invoice":
            documentos.append(ler_documento(elemento))
            elemento.clear()
            # Esvaziar o elemento não basta: enquanto o pai o referir, continua
            # em memória. Limpar o pai a cada documento mantém a leitura leve
            # mesmo num ficheiro de centenas de megabytes.
            if seccao_vendas is not None:
                seccao_vendas.clear()

        elif etiqueta == "SalesInvoices":
            seccao_vendas = None

    return {
        "cabecalho": cabecalho,
        "clientes": clientes,
        "taxas": taxas,
        "documentos": documentos,
        "declarados": declarados,
    }


def ler_documento(elemento: ET.Element) -> dict:
    numero = valor(elemento, "InvoiceNo")
    tipo = valor(elemento, "InvoiceType")

    estado_elemento = None
    totais_elemento = None
    linhas: list[dict] = []

    for filho in elemento:
        nome = sem_namespace(filho.tag)
        if nome == "DocumentStatus":
            estado_elemento = filho
        elif nome == "DocumentTotals":
            totais_elemento = filho
        elif nome == "Line":
            linhas.append(ler_linha(filho))

    # A numeração vem no formato "FT 2026/123": a série é tudo antes da barra.
    if "/" in numero:
        serie, _, sufixo = numero.rpartition("/")
        try:
            sequencia = int(sufixo)
        except ValueError:
            sequencia = None
    else:
        serie, sequencia = numero, None

    return {
        "numero": numero,
        "tipo": tipo,
        "serie": serie,
        "sequencia": sequencia,
        "estado": valor(estado_elemento, "InvoiceStatus", "N"),
        "data": para_data(valor(elemento, "InvoiceDate")),
        "cliente": valor(elemento, "CustomerID"),
        "linhas": linhas,
        "base_declarada": para_decimal(valor(totais_elemento, "NetTotal")),
        "imposto_declarado": para_decimal(valor(totais_elemento, "TaxPayable")),
        "total_declarado": para_decimal(valor(totais_elemento, "GrossTotal")),
    }


def ler_linha(elemento: ET.Element) -> dict:
    imposto_elemento = None
    for filho in elemento:
        if sem_namespace(filho.tag) == "Tax":
            imposto_elemento = filho

    # Cada linha vai a crédito ou a débito. Uma linha a débito subtrai, esteja
    # numa nota de crédito ou numa fatura com um desconto lançado à parte.
    credito = valor(elemento, "CreditAmount")
    debito = valor(elemento, "DebitAmount")
    lado = "D" if debito and not credito else "C"
    base = para_decimal(debito if lado == "D" else credito)
    percentagem = para_decimal(valor(imposto_elemento, "TaxPercentage"))

    return {
        "artigo": valor(elemento, "ProductCode"),
        "descricao": valor(elemento, "ProductDescription"),
        "quantidade": para_decimal(valor(elemento, "Quantity")),
        "base": base,
        "lado": lado,
        "codigo_taxa": valor(imposto_elemento, "TaxCode"),
        "percentagem": percentagem,
        # O imposto não consta da linha do SAF-T: calcula-se a partir da taxa,
        # arredondando como arredonda um programa de faturação.
        "imposto": (base * percentagem / Decimal("100")).quantize(
            CENTIMO, rounding=ROUND_HALF_UP
        ),
        "motivo_isencao": valor(elemento, "TaxExemptionCode"),
    }


# ------------------------------------------------------------------- apuramento

def apurar(dados: dict) -> dict:
    """Soma os documentos válidos e reúne os avisos de conferência."""
    documentos = dados["documentos"]
    clientes = dados["clientes"]
    cabecalho = dados["cabecalho"]

    inicio = para_data(cabecalho.get("inicio", ""))
    fim = para_data(cabecalho.get("fim", ""))

    por_mes: dict[int, dict] = defaultdict(
        lambda: {"documentos": 0, "base": ZERO, "imposto": ZERO}
    )
    por_tipo: dict[str, dict] = defaultdict(
        lambda: {"documentos": 0, "base": ZERO, "imposto": ZERO}
    )
    por_taxa: dict[tuple[int, str], dict] = defaultdict(
        lambda: {"percentagem": ZERO, "base": ZERO, "imposto": ZERO}
    )
    por_cliente: dict[str, dict] = defaultdict(
        lambda: {"documentos": 0, "base": ZERO, "imposto": ZERO}
    )
    por_artigo: dict[str, dict] = defaultdict(
        lambda: {"descricao": "", "quantidade": ZERO, "base": ZERO}
    )

    anulados: list[dict] = []
    totais_errados: list[dict] = []
    sem_nif_acima_limite: list[dict] = []
    fora_do_periodo: list[dict] = []
    numeros_repetidos: list[str] = []

    vistos: set[str] = set()
    sequencias: dict[str, set[int]] = defaultdict(set)

    base_credito = ZERO
    base_debito = ZERO

    for documento in documentos:
        numero = documento["numero"]
        if numero in vistos:
            numeros_repetidos.append(numero)
        vistos.add(numero)

        # Separa o que a linha acrescenta do que a linha devolve.
        credito_documento = sum(
            (l["base"] for l in documento["linhas"] if l["lado"] == "C"), ZERO
        )
        debito_documento = sum(
            (l["base"] for l in documento["linhas"] if l["lado"] == "D"), ZERO
        )
        base_linhas = credito_documento - debito_documento
        imposto_linhas = sum(
            (
                l["imposto"] if l["lado"] == "C" else -l["imposto"]
                for l in documento["linhas"]
            ),
            ZERO,
        )

        # Conferência interna: as linhas somam o total declarado no documento?
        # Os totais de uma nota de crédito vêm declarados em valor positivo.
        diferenca = abs(base_linhas + imposto_linhas) - documento["total_declarado"]
        if abs(diferenca) > TOLERANCIA:
            totais_errados.append(
                {
                    "numero": numero,
                    "data": documento["data"],
                    "soma_linhas": abs(base_linhas + imposto_linhas),
                    "declarado": documento["total_declarado"],
                    "diferenca": diferenca,
                }
            )

        # Um documento anulado continua a ocupar o seu número na série, por
        # isso entra aqui antes de ser excluído das somas.
        if documento["sequencia"] is not None:
            sequencias[documento["serie"]].add(documento["sequencia"])

        if documento["estado"] in ESTADOS_EXCLUIDOS:
            if documento["estado"] == "A":
                anulados.append(
                    {
                        "numero": numero,
                        "data": documento["data"],
                        "cliente": clientes.get(documento["cliente"], {}).get("nome", ""),
                        "total": documento["total_declarado"],
                    }
                )
            continue  # anulado ou já faturado não entra nas somas

        data_documento = documento["data"]
        if data_documento and inicio and fim:
            if not (inicio <= data_documento <= fim):
                fora_do_periodo.append(
                    {"numero": numero, "data": data_documento}
                )

        cliente = clientes.get(documento["cliente"], {})
        if (
            documento["tipo"] in {"FT", "FR"}
            and cliente.get("nif") == NIF_CONSUMIDOR_FINAL
            and documento["total_declarado"] >= LIMITE_FATURA_SEM_NIF
        ):
            sem_nif_acima_limite.append(
                {
                    "numero": numero,
                    "data": data_documento,
                    "total": documento["total_declarado"],
                }
            )

        # Totais de controlo do ficheiro: uma soma por lado, como na Portaria.
        base_credito += credito_documento
        base_debito += debito_documento

        mes = data_documento.month if data_documento else 0
        por_mes[mes]["documentos"] += 1
        por_mes[mes]["base"] += base_linhas
        por_mes[mes]["imposto"] += imposto_linhas

        # As notas de crédito ficam com sinal negativo também aqui, para que
        # a coluna possa ser somada e dar o mesmo que a folha dos meses.
        por_tipo[documento["tipo"]]["documentos"] += 1
        por_tipo[documento["tipo"]]["base"] += base_linhas
        por_tipo[documento["tipo"]]["imposto"] += imposto_linhas

        chave_cliente = documento["cliente"]
        por_cliente[chave_cliente]["documentos"] += 1
        por_cliente[chave_cliente]["base"] += base_linhas
        por_cliente[chave_cliente]["imposto"] += imposto_linhas

        for linha in documento["linhas"]:
            sinal = Decimal("1") if linha["lado"] == "C" else Decimal("-1")
            chave_taxa = (mes, linha["codigo_taxa"])
            por_taxa[chave_taxa]["percentagem"] = linha["percentagem"]
            por_taxa[chave_taxa]["base"] += sinal * linha["base"]
            por_taxa[chave_taxa]["imposto"] += sinal * linha["imposto"]

            artigo = por_artigo[linha["artigo"]]
            artigo["descricao"] = linha["descricao"]
            artigo["quantidade"] += sinal * linha["quantidade"]
            artigo["base"] += sinal * linha["base"]

    # Falhas na numeração sequencial de cada série.
    falhas: list[dict] = []
    for serie, numeros in sorted(sequencias.items()):
        if not numeros:
            continue
        em_falta = sorted(set(range(min(numeros), max(numeros) + 1)) - numeros)
        if em_falta:
            falhas.append(
                {
                    "serie": serie,
                    "primeiro": min(numeros),
                    "ultimo": max(numeros),
                    "em_falta": em_falta,
                }
            )

    declarados = dados["declarados"]
    conferencia_ficheiro = {
        "documentos_contados": len(documentos),
        "documentos_declarados": declarados["numero"],
        "credito_calculado": base_credito,
        "credito_declarado": declarados["credito"],
        "debito_calculado": base_debito,
        "debito_declarado": declarados["debito"],
    }

    return {
        "por_mes": dict(por_mes),
        "por_tipo": dict(por_tipo),
        "por_taxa": dict(por_taxa),
        "por_cliente": dict(por_cliente),
        "por_artigo": dict(por_artigo),
        "anulados": anulados,
        "totais_errados": totais_errados,
        "sem_nif_acima_limite": sem_nif_acima_limite,
        "fora_do_periodo": fora_do_periodo,
        "numeros_repetidos": numeros_repetidos,
        "falhas_numeracao": falhas,
        "conferencia": conferencia_ficheiro,
    }


def reunir_avisos(apuramento: dict) -> list[dict]:
    """Transforma cada conferência falhada numa linha de aviso legível."""
    avisos: list[dict] = []
    conferencia = apuramento["conferencia"]

    if conferencia["documentos_contados"] != conferencia["documentos_declarados"]:
        avisos.append(
            {
                "tipo": "Contagem de documentos",
                "gravidade": "Alta",
                "detalhe": (
                    f"O ficheiro declara {conferencia['documentos_declarados']} "
                    f"documentos e contém {conferencia['documentos_contados']}."
                ),
            }
        )

    for nome, calculado, declarado in (
        ("crédito", conferencia["credito_calculado"], conferencia["credito_declarado"]),
        ("débito", conferencia["debito_calculado"], conferencia["debito_declarado"]),
    ):
        diferenca = calculado - declarado
        if abs(diferenca) > TOLERANCIA:
            avisos.append(
                {
                    "tipo": f"Total a {nome}",
                    "gravidade": "Alta",
                    "detalhe": (
                        f"Soma das linhas {calculado:.2f} EUR contra "
                        f"{declarado:.2f} EUR declarados no ficheiro "
                        f"(diferença de {diferenca:.2f} EUR)."
                    ),
                }
            )

    for documento in apuramento["totais_errados"]:
        avisos.append(
            {
                "tipo": "Documento não fecha",
                "gravidade": "Alta",
                "detalhe": (
                    f"{documento['numero']}: as linhas somam "
                    f"{documento['soma_linhas']:.2f} EUR e o documento declara "
                    f"{documento['declarado']:.2f} EUR."
                ),
            }
        )

    for falha in apuramento["falhas_numeracao"]:
        em_falta = falha["em_falta"]
        lista = ", ".join(str(n) for n in em_falta[:10])
        if len(em_falta) > 10:
            lista += f" (e mais {len(em_falta) - 10})"
        avisos.append(
            {
                "tipo": "Falha na numeração",
                "gravidade": "Alta",
                "detalhe": (
                    f"Série {falha['serie']} vai de {falha['primeiro']} a "
                    f"{falha['ultimo']} e não tem o número {lista}."
                ),
            }
        )

    for numero in apuramento["numeros_repetidos"]:
        avisos.append(
            {
                "tipo": "Número repetido",
                "gravidade": "Alta",
                "detalhe": f"O documento {numero} aparece mais do que uma vez.",
            }
        )

    for documento in apuramento["sem_nif_acima_limite"]:
        avisos.append(
            {
                "tipo": "Fatura sem NIF do cliente",
                "gravidade": "Média",
                "detalhe": (
                    f"{documento['numero']} de {documento['total']:.2f} EUR "
                    f"emitida a consumidor final."
                ),
            }
        )

    for documento in apuramento["fora_do_periodo"]:
        avisos.append(
            {
                "tipo": "Data fora do período",
                "gravidade": "Média",
                "detalhe": (
                    f"{documento['numero']} tem data de {documento['data']}, "
                    f"fora do período do ficheiro."
                ),
            }
        )

    if apuramento["anulados"]:
        total = sum(d["total"] for d in apuramento["anulados"])
        avisos.append(
            {
                "tipo": "Documentos anulados",
                "gravidade": "Informação",
                "detalhe": (
                    f"{len(apuramento['anulados'])} documentos anulados, "
                    f"{total:.2f} EUR, excluídos das somas. Ver a folha Anulados."
                ),
            }
        )

    return avisos


# ----------------------------------------------------------------------- Excel

AZUL = "1F3864"
CINZA = "F2F2F2"
MOEDA = '#,##0.00'
INTEIRO = '#,##0'

ESTILO_TITULO = Font(bold=True, color="FFFFFF", size=11)
PREENCHIMENTO_TITULO = PatternFill("solid", fgColor=AZUL)
BORDA_BAIXO = Border(bottom=Side(style="thin", color="BFBFBF"))


def escrever_folha(
    livro: Workbook,
    nome: str,
    colunas: list[tuple[str, str, int]],
    linhas: list[list],
    congelar: str = "A2",
) -> None:
    """
    Escreve uma folha com cabeçalho formatado.

    colunas: lista de (título, formato numérico ou "", largura).
    """
    folha = livro.create_sheet(nome)

    for indice, (titulo, _, largura) in enumerate(colunas, start=1):
        celula = folha.cell(row=1, column=indice, value=titulo)
        celula.font = ESTILO_TITULO
        celula.fill = PREENCHIMENTO_TITULO
        celula.alignment = Alignment(horizontal="center", vertical="center")
        folha.column_dimensions[get_column_letter(indice)].width = largura

    for numero_linha, valores in enumerate(linhas, start=2):
        for indice, conteudo in enumerate(valores, start=1):
            celula = folha.cell(row=numero_linha, column=indice)
            celula.value = float(conteudo) if isinstance(conteudo, Decimal) else conteudo
            formato = colunas[indice - 1][1]
            if formato:
                celula.number_format = formato
            if colunas[indice - 1][2] >= 60:  # colunas de texto longo
                celula.alignment = Alignment(wrap_text=True, vertical="top")
            celula.border = BORDA_BAIXO

    folha.freeze_panes = congelar
    if linhas:
        folha.auto_filter.ref = (
            f"A1:{get_column_letter(len(colunas))}{len(linhas) + 1}"
        )


def escrever_resumo(livro: Workbook, dados: dict, apuramento: dict, avisos: list) -> None:
    folha = livro.create_sheet("Resumo")
    folha.column_dimensions["A"].width = 34
    folha.column_dimensions["B"].width = 46

    cabecalho = dados["cabecalho"]
    conferencia = apuramento["conferencia"]

    base = sum(m["base"] for m in apuramento["por_mes"].values())
    imposto = sum(m["imposto"] for m in apuramento["por_mes"].values())

    notas_credito = apuramento["por_tipo"].get("NC", {"base": ZERO, "documentos": 0})

    linha = 1

    def seccao(titulo: str) -> None:
        nonlocal linha
        celula = folha.cell(row=linha, column=1, value=titulo)
        celula.font = ESTILO_TITULO
        celula.fill = PREENCHIMENTO_TITULO
        folha.cell(row=linha, column=2).fill = PREENCHIMENTO_TITULO
        linha += 1

    def campo(etiqueta: str, conteudo, formato: str = "") -> None:
        nonlocal linha
        folha.cell(row=linha, column=1, value=etiqueta).font = Font(bold=True)
        celula = folha.cell(row=linha, column=2)
        celula.value = float(conteudo) if isinstance(conteudo, Decimal) else conteudo
        if formato:
            celula.number_format = formato
            celula.alignment = Alignment(horizontal="left")
        linha += 1

    seccao("Empresa")
    campo("Designação", cabecalho.get("empresa", ""))
    campo("NIF", cabecalho.get("nif", ""))
    campo("Localidade", cabecalho.get("localidade", ""))
    campo("Programa de faturação", cabecalho.get("programa", ""))
    linha += 1

    seccao("Ficheiro")
    campo("Versão do SAF-T", cabecalho.get("versao", ""))
    campo("Exercício", cabecalho.get("ano", ""))
    campo("Período", f"{cabecalho.get('inicio', '')} a {cabecalho.get('fim', '')}")
    campo("Moeda", cabecalho.get("moeda", ""))
    campo("Documentos no ficheiro", conferencia["documentos_contados"], INTEIRO)
    campo("Documentos anulados", len(apuramento["anulados"]), INTEIRO)
    linha += 1

    seccao("Faturação do período (líquida de notas de crédito)")
    campo("Base tributável", base, MOEDA)
    campo("IVA liquidado", imposto, MOEDA)
    campo("Total", base + imposto, MOEDA)
    campo("Notas de crédito emitidas", notas_credito["documentos"], INTEIRO)
    campo("Valor das notas de crédito", notas_credito["base"], MOEDA)
    linha += 1

    seccao("Conferência")
    altos = sum(1 for a in avisos if a["gravidade"] == "Alta")
    medios = sum(1 for a in avisos if a["gravidade"] == "Média")
    campo("Avisos de gravidade alta", altos, INTEIRO)
    campo("Avisos de gravidade média", medios, INTEIRO)
    campo(
        "Total a crédito declarado",
        conferencia["credito_declarado"],
        MOEDA,
    )
    campo(
        "Total a crédito calculado",
        conferencia["credito_calculado"],
        MOEDA,
    )
    campo(
        "Total a débito declarado",
        conferencia["debito_declarado"],
        MOEDA,
    )
    campo(
        "Total a débito calculado",
        conferencia["debito_calculado"],
        MOEDA,
    )
    linha += 1

    folha.cell(
        row=linha,
        column=1,
        value="Relatório gerado por analisar_saft.py. Confira sempre contra o original.",
    ).font = Font(italic=True, color="808080")


def gerar_excel(dados: dict, apuramento: dict, avisos: list, caminho: Path) -> None:
    livro = Workbook()
    livro.remove(livro.active)

    escrever_resumo(livro, dados, apuramento, avisos)

    escrever_folha(
        livro,
        "Avisos",
        [("Gravidade", "", 14), ("Tipo", "", 30), ("Detalhe", "", 95)],
        [[a["gravidade"], a["tipo"], a["detalhe"]] for a in avisos]
        or [["Informação", "Sem avisos", "Nenhuma conferência falhou."]],
    )

    escrever_folha(
        livro,
        "Vendas por mês",
        [
            ("Mês", "", 14),
            ("Documentos", INTEIRO, 14),
            ("Base tributável", MOEDA, 18),
            ("IVA", MOEDA, 16),
            ("Total", MOEDA, 18),
        ],
        [
            [
                MESES[mes - 1].capitalize() if 1 <= mes <= 12 else "Sem data",
                valores["documentos"],
                valores["base"],
                valores["imposto"],
                valores["base"] + valores["imposto"],
            ]
            for mes, valores in sorted(apuramento["por_mes"].items())
        ],
    )

    escrever_folha(
        livro,
        "IVA por taxa",
        [
            ("Mês", "", 14),
            ("Código", "", 12),
            ("Taxa %", '0.00', 12),
            ("Base tributável", MOEDA, 18),
            ("IVA", MOEDA, 16),
        ],
        [
            [
                MESES[mes - 1].capitalize() if 1 <= mes <= 12 else "Sem data",
                codigo,
                valores["percentagem"],
                valores["base"],
                valores["imposto"],
            ]
            for (mes, codigo), valores in sorted(apuramento["por_taxa"].items())
        ],
    )

    escrever_folha(
        livro,
        "Tipos de documento",
        [
            ("Código", "", 12),
            ("Tipo", "", 26),
            ("Documentos", INTEIRO, 14),
            ("Base tributável", MOEDA, 18),
            ("IVA", MOEDA, 16),
        ],
        [
            [
                codigo,
                NOMES_TIPOS.get(codigo, codigo),
                valores["documentos"],
                valores["base"],
                valores["imposto"],
            ]
            for codigo, valores in sorted(
                apuramento["por_tipo"].items(),
                key=lambda item: item[1]["base"],
                reverse=True,
            )
        ],
    )

    clientes = dados["clientes"]
    escrever_folha(
        livro,
        "Clientes",
        [
            ("Cliente", "", 40),
            ("NIF", "", 14),
            ("Documentos", INTEIRO, 14),
            ("Base tributável", MOEDA, 18),
            ("IVA", MOEDA, 16),
        ],
        [
            [
                clientes.get(codigo, {}).get("nome", codigo),
                clientes.get(codigo, {}).get("nif", ""),
                valores["documentos"],
                valores["base"],
                valores["imposto"],
            ]
            for codigo, valores in sorted(
                apuramento["por_cliente"].items(),
                key=lambda item: item[1]["base"],
                reverse=True,
            )
        ],
    )

    escrever_folha(
        livro,
        "Artigos",
        [
            ("Código", "", 12),
            ("Descrição", "", 44),
            ("Quantidade", '#,##0.00', 14),
            ("Base tributável", MOEDA, 18),
        ],
        [
            [codigo, valores["descricao"], valores["quantidade"], valores["base"]]
            for codigo, valores in sorted(
                apuramento["por_artigo"].items(),
                key=lambda item: item[1]["base"],
                reverse=True,
            )
        ],
    )

    escrever_folha(
        livro,
        "Anulados",
        [
            ("Documento", "", 20),
            ("Data", "", 14),
            ("Cliente", "", 40),
            ("Total", MOEDA, 16),
        ],
        [
            [d["numero"], str(d["data"]), d["cliente"], d["total"]]
            for d in apuramento["anulados"]
        ],
    )

    caminho.parent.mkdir(parents=True, exist_ok=True)
    livro.save(caminho)


# ------------------------------------------------------------------- no ecrã

def mostrar_resumo(dados: dict, apuramento: dict, avisos: list, saida: Path) -> None:
    cabecalho = dados["cabecalho"]
    base = sum(m["base"] for m in apuramento["por_mes"].values())
    imposto = sum(m["imposto"] for m in apuramento["por_mes"].values())

    print()
    print(f"{cabecalho.get('empresa', '(sem nome)')}  ·  NIF {cabecalho.get('nif', '')}")
    print(f"Período {cabecalho.get('inicio', '')} a {cabecalho.get('fim', '')}"
          f"  ·  SAF-T {cabecalho.get('versao', '')}")
    print("-" * 62)
    print(f"{'Documentos':<28}{apuramento['conferencia']['documentos_contados']:>14,}")
    print(f"{'Anulados (excluídos)':<28}{len(apuramento['anulados']):>14,}")
    print(f"{'Base tributável':<28}{base:>14,.2f} EUR")
    print(f"{'IVA liquidado':<28}{imposto:>14,.2f} EUR")
    print(f"{'Total':<28}{base + imposto:>14,.2f} EUR")
    print("-" * 62)

    if avisos:
        print("Avisos:")
        for aviso in avisos:
            print(f"  [{aviso['gravidade']:<10}] {aviso['tipo']}: {aviso['detalhe']}")
    else:
        print("Nenhuma conferência falhou.")

    print()
    print(f"Relatório gravado em {saida}")


# ---------------------------------------------------------------------- início

def principal(argumentos: list[str] | None = None) -> int:
    analisador = argparse.ArgumentParser(
        description="Lê um ficheiro SAF-T (PT) e grava um relatório em Excel."
    )
    analisador.add_argument(
        "--ficheiro",
        type=Path,
        default=FICHEIRO_PREDEFINIDO,
        help="caminho do ficheiro SAF-T em XML",
    )
    analisador.add_argument(
        "--saida",
        type=Path,
        default=SAIDA_PREDEFINIDA,
        help="caminho do relatório Excel a gravar",
    )
    opcoes = analisador.parse_args(argumentos)

    if not opcoes.ficheiro.exists():
        print(f"Não encontrei o ficheiro {opcoes.ficheiro}.", file=sys.stderr)
        print("Execute primeiro: python gerar_saft.py", file=sys.stderr)
        return 1

    dados = ler_saft(opcoes.ficheiro)
    if not dados["documentos"]:
        print("O ficheiro não tem documentos de venda (SalesInvoices).", file=sys.stderr)
        return 1

    apuramento = apurar(dados)
    avisos = reunir_avisos(apuramento)
    gerar_excel(dados, apuramento, avisos, opcoes.saida)
    mostrar_resumo(dados, apuramento, avisos, opcoes.saida)
    return 0


if __name__ == "__main__":
    raise SystemExit(principal())

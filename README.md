# Relatório automático a partir do SAF-T (PT)

Script em Python que lê o ficheiro SAF-T (PT) de faturação e devolve, em Excel, o que o gabinete precisa de saber sobre o período de um cliente: quanto faturou, com que IVA, a quem vendeu, e o que no ficheiro merece uma segunda vista antes de servir de base à declaração.

É o trabalho que, feito à mão, obriga a abrir o ficheiro no Excel, filtrar linhas e montar tabelas dinâmicas a cada fecho.

**Dados fictícios.** A empresa, os clientes, os NIF e os valores são inventados e gerados por `gerar_saft.py`. Não é utilizado nenhum ficheiro real de qualquer empresa.

## O problema

O SAF-T é um XML com uma linha por artigo de cada documento. Um ano de uma pequena empresa dá facilmente centenas de milhares de linhas. A informação está lá toda, mas não está somada: para saber a base tributável de maio por taxa de IVA, ou quanto pesaram as notas de crédito no ano, é preciso construir o apuramento de cada vez.

E há o que não se vê ao olhar para os totais: um documento cujas linhas não somam o total declarado, um número em falta numa série, uma fatura acima de 1 000 EUR emitida a consumidor final.

## O que o script faz

1. Percorre o ficheiro uma única vez, libertando cada documento da memória depois de o ler.
2. Exclui das somas os documentos anulados, mas mantém-nos na contagem e numa folha à parte.
3. Trata as notas de crédito como redução da faturação, não como venda.
4. Apura vendas por mês, por tipo de documento, por taxa de IVA, por cliente e por artigo.
5. Corre cinco conferências sobre o ficheiro e lista o que falhou.
6. Grava um Excel com oito folhas e mostra um resumo no terminal.

### Conferências

| Conferência | O que procura |
| --- | --- |
| Totais de controlo | A soma das linhas bate com o `TotalCredit` e o `TotalDebit` declarados no ficheiro |
| Documento não fecha | As linhas de um documento somam o `GrossTotal` que ele próprio declara |
| Falha na numeração | Números em falta dentro de cada série (os anulados contam como presentes) |
| Fatura sem NIF | Faturas e faturas-recibo acima de 1 000 EUR emitidas a consumidor final |
| Data fora do período | Documentos com data fora do intervalo declarado no cabeçalho |

Os totais de controlo seguem a regra da Portaria n.º 302/2016: o número de registos **inclui** os documentos anulados, os totais a crédito e a débito **excluem**-nos.

## Resultado com os dados de exemplo

Um ano de uma loja de ferragens: 1263 documentos, processados em menos de um segundo.

| | Valor |
| --- | --- |
| Documentos | 1263 (5 anulados) |
| Base tributável | 267 859,60 EUR |
| IVA liquidado | 50 718,52 EUR |
| Total | 318 578,12 EUR |
| Notas de crédito | 30 documentos, −8 332,60 EUR |

O gerador planta de propósito quatro situações, e o analisador apanha as quatro:

```
[Alta      ] Documento não fecha: FT 2026/311: as linhas somam 969.82 EUR
             e o documento declara 969.84 EUR.
[Alta      ] Falha na numeração: Série FT 2026 vai de 1 a 371 e não tem o número 118.
[Média     ] Fatura sem NIF do cliente: FT 2026/205 de 1030.32 EUR emitida
             a consumidor final.
[Informação] Documentos anulados: 5 documentos anulados, 885.54 EUR,
             excluídos das somas.
```

Os dados são gerados com uma semente fixa, por isso quem executar o projeto obtém exatamente estes números.

## Folhas do relatório

| Folha | Conteúdo |
| --- | --- |
| Resumo | Empresa, período, totais e estado das conferências |
| Avisos | Uma linha por conferência falhada, com gravidade |
| Vendas por mês | Documentos, base, IVA e total de cada mês |
| IVA por taxa | Base e imposto por taxa e por mês, para a declaração periódica |
| Tipos de documento | Peso de FT, FS, FR, ND e NC |
| Clientes | Faturação por cliente, do maior para o menor |
| Artigos | Quantidade e valor por artigo ou serviço |
| Anulados | Os documentos excluídos das somas |

As folhas dos meses, das taxas, dos tipos e dos clientes somam todas o mesmo valor.

## Como utilizar

```
pip install -r requirements.txt

python gerar_saft.py      # cria dados/saft_exemplo_2026.xml
python analisar_saft.py   # grava resultado/relatorio_saft.xlsx
```

Com um ficheiro próprio:

```
python analisar_saft.py --ficheiro o_meu_saft.xml --saida resultado/cliente.xlsx
```

## Testes

```
python -m unittest
```

Nove testes, cada um sobre uma regra do apuramento: as notas de crédito reduzem a faturação, uma linha a débito dentro de uma fatura também subtrai, um documento anulado sai das somas mas não abre falha na numeração, e cada uma das conferências dispara quando deve.

## Estrutura

```
analisar_saft.py    leitura do SAF-T, apuramento e relatório
gerar_saft.py       gerador do ficheiro fictício
teste_analisar.py   testes do apuramento
requirements.txt    dependências
```

As pastas `dados/` e `resultado/` são criadas ao executar os dois scripts.

## Limites

- Lê a secção de faturação (`SalesInvoices`). Não lê a contabilidade (`GeneralLedgerEntries`), os documentos de transporte nem os recibos.
- Estrutura suportada: SAF-T (PT) versão 1.04_01.
- O ficheiro de exemplo tem assinaturas (`Hash`) e códigos ATCUD de preenchimento. Num ficheiro real são produzidos por software certificado pela AT, e este script não os valida.
- O relatório é um apoio à conferência, não substitui a revisão de quem assina a declaração.

## Adaptação a cada gabinete

Este repositório é uma demonstração. Num projeto real, o relatório é ajustado ao que cada gabinete precisa de entregar: indicadores próprios, comparação com o período homólogo, um ficheiro por cliente gerado em lote a partir de uma pasta, envio automático por e-mail, ou exportação para o formato que o software de contabilidade importa.

## Tecnologias

Python 3.10+, openpyxl. A leitura do XML usa `xml.etree.ElementTree` com `iterparse`, sem dependências externas.

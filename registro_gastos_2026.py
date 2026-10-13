
import pandas as pd

# Leitura do arquivo de viagens do Portal da Transparência
viagem = pd.read_csv(
    '2026_Viagem.csv',
    encoding='latin1',
    sep=';'
)

# TRATAMENTO E SOMATÓRIO DOS GASTOS POR ÓRGÃO
# -----------------------------------------------------------------

# Identifica a coluna do órgão solicitante.
possiveis_colunas_orgao = [
    'Nome órgão solicitante',
    'Nome do órgão solicitante'
]

coluna_orgao = next(
    (
        coluna for coluna in possiveis_colunas_orgao
        if coluna in viagem.columns
    ),
    None
)

# Verifica se a coluna do órgão existe.
if coluna_orgao is None:
    raise KeyError(
        'Coluna do órgão solicitante não encontrada. '
        f'Colunas disponíveis: {viagem.columns.tolist()}'
    )

# Define as colunas financeiras.
colunas_gastos = [
    'Valor diárias',
    'Valor passagens',
    'Valor devolução',
    'Valor outros gastos'
]

# Verifica se todas as colunas financeiras existem.
colunas_ausentes = [
    coluna for coluna in colunas_gastos
    if coluna not in viagem.columns
]

if colunas_ausentes:
    raise KeyError(
        f'Colunas financeiras não encontradas: {colunas_ausentes}'
    )

# -----------------------------------------------------------------
# PREPARAÇÃO DOS DADOS

# Copia somente as colunas necessárias para o relatório.
relatorio = viagem[[coluna_orgao] + colunas_gastos].copy()

# Converte os valores monetários para números.
# Exemplo: '1250,50' passa a ser 1250.50.


for coluna in colunas_gastos:
    relatorio[coluna] = pd.to_numeric(
        relatorio[coluna]
        .astype('string')
        .str.strip()
        .str.replace(',', '.', regex=False),
        errors='coerce'
    )

# Substitui valores ausentes por zero para os cálculos.
relatorio[colunas_gastos] = relatorio[colunas_gastos].fillna(0)

# Padroniza órgãos ausentes ou com nome vazio.
relatorio[coluna_orgao] = (
    relatorio[coluna_orgao]
    .fillna('(Órgão não informado)')
    .astype('string')
    .str.strip()
    .replace('', '(Órgão não informado)')
)

# Calcula o total líquido de cada registro.
# Diárias + passagens + outros gastos - devoluções.

relatorio['Total líquido'] = (
    relatorio['Valor diárias']
    + relatorio['Valor passagens']
    + relatorio['Valor outros gastos']
    - relatorio['Valor devolução']
)

# -----------------------------------------------------------------
# AGRUPAMENTO POR ÓRGÃO SOLICITANTE

# Soma os gastos e conta os registros de cada órgão 
# roupby(nativo do pandas) agrupa as linhas que possuem o mesmo valor em uma coluna
#agg ele faz os calculos apos o agrupamento
gastos_por_orgao = (
    relatorio
    .groupby(coluna_orgao, dropna=False)
    .agg(
        **{
            'Quantidade de registros': ('Total líquido', 'size'),
            'Total de diárias': ('Valor diárias', 'sum'),
            'Total de passagens': ('Valor passagens', 'sum'),
            'Total de devoluções': ('Valor devolução', 'sum'),
            'Total de outros gastos': ('Valor outros gastos', 'sum'),
            'Total líquido': ('Total líquido', 'sum')
        }
    )
    .reset_index()
)

# Padroniza o nome da coluna do órgão.
gastos_por_orgao = gastos_por_orgao.rename(
    columns={coluna_orgao: 'Órgão solicitante'}
)

# Ordena do maior para o menor gasto líquido.
gastos_por_orgao = gastos_por_orgao.sort_values(
    by='Total líquido',
    ascending=False
).reset_index(drop=True)

# -----------------------------------------------------------------
# IMPRESSÃO DO RELATÓRIO

# Cria uma cópia para formatar os valores somente na exibição.
relatorio_exibicao = gastos_por_orgao.copy()

colunas_monetarias = [
    'Total de diárias',
    'Total de passagens',
    'Total de devoluções',
    'Total de outros gastos',
    'Total líquido'
]

# Formata os valores no padrão monetário brasileiro.
for coluna in colunas_monetarias:
    relatorio_exibicao[coluna] = (
        relatorio_exibicao[coluna]
        .map(lambda valor: f'R$ {valor:,.2f}')
        .str.replace(',', 'X', regex=False)
        .str.replace('.', ',', regex=False)
        .str.replace('X', '.', regex=False)
    )

print('\n' + '=' * 100)
print('RELATÓRIO DE GASTOS POR ÓRGÃO SOLICITANTE')
print('=' * 100)
print(relatorio_exibicao.to_string(index=False))

# Calcula e imprime o total líquido geral.
total_geral = gastos_por_orgao['Total líquido'].sum()

total_geral_formatado = (
    f'R$ {total_geral:,.2f}'
    .replace(',', 'X')
    .replace('.', ',')
    .replace('X', '.')
)

print('\n' + '=' * 60)
print('TOTAL LÍQUIDO GERAL')
print('=' * 60)
print(total_geral_formatado)

# -----------------------------------------------------------------
# EXPORTAÇÃO PARA EXCEL

# Cria um Excel com duas abas:
# 1. Resumo por órgão.
# 2. Detalhamento dos registros.

arquivo_excel = 'Relatorio_Gastos_2026.xlsx'

with pd.ExcelWriter(
    arquivo_excel,
    engine='openpyxl'
) as escritor:

    gastos_por_orgao.to_excel(
        escritor,
        sheet_name='Resumo por órgão',
        index=False
    )

    relatorio.to_excel(
        escritor,
        sheet_name='Detalhamento',
        index=False
    )

# -----------------------------------------------------------------
# EXPORTAÇÃO PARA CSV

# Exporta o resumo com separador de colunas ';',
# separador decimal ',' e codificação compatível com Excel.

arquivo_csv = 'Relatorio_Gastos_2026.csv'

gastos_por_orgao.to_csv(
    arquivo_csv,
    index=False,
    sep=';',
    decimal=',',
    encoding='utf-8-sig'
)

# Informa os arquivos gerados.
print('\nARQUIVOS EXPORTADOS:')
print(f'Excel: {arquivo_excel}')
print(f'CSV: {arquivo_csv}')

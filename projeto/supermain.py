import pandas as pd

df1 = pd.read_excel(
    r'C:\Users\micro\Documents\new project\carteira_gerais.xlsx'
)

df2 = pd.read_excel(
    r'C:\Users\micro\Documents\new project\dados_empresas.xlsx'
)

colunas_remover = [
    'nome completo',
    'situação',
    'valor contratado',
    'saldo devedor',
    'parcelas',
    'prox. venc.',
    'valor parcela',
    'atraso.1',
    'dividido em',
    'faltam qnt'
]

df = df1.drop(columns=colunas_remover)

df_final = pd.merge(
    df2,
    df,
    on='contrato',
    how='left'
)

colunas_remover1 = [
    'Unnamed: 23',
    'Unnamed: 24',
    'Unnamed: 25',
    'Unnamed: 26',
    'Unnamed: 27',
    'a',
    'Coluna1',
    'repe',
]

df_final1 = df_final.drop(columns=colunas_remover1)

df_final1.to_excel(
    r'C:\Users\micro\Documents\new project\arquivo_final.xlsx',
    index=False
)

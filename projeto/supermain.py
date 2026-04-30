import pandas as pd

df1 = pd.read_excel(
    r'C:\Users\micro\Documents\new_project\carteira_gerais.xlsx'
)

df2 = pd.read_excel(
    r'C:\Users\micro\Documents\new_project\google_sheets_geral.xlsx'
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

    'a',
    'Coluna1',
    'repe',
]

df_final1 = df_final.drop(columns=colunas_remover1)

df_final1.to_excel(
    r'C:\Users\micro\Documents\new_project\arquivo_final.xlsx',
    index=False
)

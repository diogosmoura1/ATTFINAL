import pandas as pd
from dicionary import empresas_dic

lista_dfs = []

for item in empresas_dic:
    caminho = item["carteira"]
    empresa = item["empresa"]

    try:
        df = pd.read_excel(caminho, sheet_name="CARTEIRA GERAL")
        lista_dfs.append(df)

    except Exception as e:
        print(f"Erro ao ler {empresa}: {e}")

df_final = pd.concat(lista_dfs, ignore_index=True)

df_final.to_excel("carteira_gerais.xlsx", index= False)
print("documento criando com sucesso! \a")
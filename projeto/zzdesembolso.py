import pandas as pd
from dicionary import empresas_dic


lista_dfs = []

for item in empresas_dic:
    caminho = item["carteira"]
    empresa = item["empresa"]


    try:
        df = pd.read_excel(caminho, sheet_name="desembolso")
        lista_dfs.append(df)
        df["empresa"] = empresa


    except Exception as e:
        print(f"Erro ao ler {empresa}: {e}")


df_final = pd.concat(lista_dfs, ignore_index=True)


df_final.to_excel("desembolso geral.xlsx", index= False)
import pandas as pd
from dicionary import empresas_dic


lista_dfs = []

for item in empresas_dic:
    caminho = item["new"]
    empresa = item["empresa"]


    try:
        df = pd.read_excel(caminho, sheet_name="resultado")
        lista_dfs.append(df)


    except Exception as e:
        print(f"Erro ao ler {empresa}: {e}")


df_final = pd.concat(lista_dfs, ignore_index=True)


df_final.to_excel("new coleta cpf geral.xlsx", index= False)
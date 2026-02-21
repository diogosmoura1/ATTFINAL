import pandas as pd
from dicionary import empresas

lista_clientes_novos = []
lista_cpf_nao_encontrado = []

def tratar_cpf(df):
    if "cpf" in df.columns:
        df["cpf"] = (
            df["cpf"]
            .astype(str)
            .str.replace(r'\D', '', regex=True)
            .str.zfill(11)
        )
    return df

def tratar_valores(df):
    colunas_para_tratar = ['valor contratado', 'valor parcela']
    
    for coluna in colunas_para_tratar:
        if coluna in df.columns:
            df[coluna] = (
                df[coluna]
                .astype(str)
                .str.replace(r'[R$\s.]', '', regex=True)
                .str.replace(',', '.', regex=False)
            )
            df[coluna] = pd.to_numeric(df[coluna], errors='coerce')
    
    return df


for nome_empresa, dados in empresas.items():
    
    print(f"Processando {nome_empresa}...")
    
    try:
        # ========= LEITURAS =========
        
        desembolso = pd.read_excel(dados["carteira"], sheet_name="desembolso")
        desembolso = desembolso.rename(columns={
            'nome completo': 'nome completo forms',
            'Carimbo de data/hora': 'horario preenchimento forms'
        })
        
        resultado = pd.read_excel(dados["new"], sheet_name="resultado")
        remover_colunas_resultado = ['e-mail','cep','cpf aval','endereço','bairro']
        resultado = resultado.drop(columns=remover_colunas_resultado, errors="ignore")
        resultado = resultado.rename(columns={
            'nome completo': 'nome completo 360',
            'número do contrato': 'contrato'
        })
        
        link_export = dados["link"].replace("/edit?usp=sharing", "/export?format=xlsx")
        google_sheets = pd.read_excel(link_export)
        remover_colunas_google = ['nome completo','situação','saldo devedor','prox. venc.','atraso']
        google_sheets = google_sheets.drop(columns=remover_colunas_google, errors="ignore")
        
        # ========= TRATAMENTOS =========
        
        resultado = tratar_cpf(resultado)
        desembolso = tratar_cpf(desembolso)
        
        resultado = pd.merge(google_sheets, resultado, on='contrato', how='right')
        resultado = tratar_valores(resultado)
        
        # ========= CLIENTES NOVOS =========
        
        df1 = pd.merge(resultado, desembolso, on='cpf', how='left')
        df1["empresa"] = nome_empresa
        
        lista_clientes_novos.append(df1)
        
        # ========= CPF NÃO ENCONTRADO =========
        
        df2 = pd.merge(resultado, desembolso, on='cpf', how='right')
        df2["contrato"] = df2["contrato"].replace(r"^\s*$", pd.NA, regex=True)
        df2 = df2[df2["contrato"].isna()]
        df2["empresa"] = nome_empresa
        
        lista_cpf_nao_encontrado.append(df2)
    
    except Exception as e:
        print(f"Erro em {nome_empresa}: {e}")


# ========= CONSOLIDANDO TUDO =========

clientes_novos_final = pd.concat(lista_clientes_novos, ignore_index=True)
cpf_nao_encontrado_final = pd.concat(lista_cpf_nao_encontrado, ignore_index=True)

nova_ordem = ["data de contratação","horario preenchimento forms","empresa","agente","contrato","TIPO DE CLIENTE","valor contratado","valor parcela","DESEMBOLSO",
              "PRAZO","parcelas","nome completo 360","nome completo forms","cpf","celular","cidade","nome aval","celular aval","ano","mês"]

clientes_novos_final = clientes_novos_final[nova_ordem]

clientes_novos_final.to_excel(
    r'C:\Users\micro\Documents\new project\comp_des_result.py\PLANILHAS SALVAS\CONSOLIDADO_CLIENTES_NOVOS.xlsx',
    index=False
)

cpf_nao_encontrado_final = cpf_nao_encontrado_final[nova_ordem]

cpf_nao_encontrado_final.to_excel(
    r'C:\Users\micro\Documents\new project\comp_des_result.py\PLANILHAS SALVAS\CONSOLIDADO_CPF_NAO_ENCONTRADO.xlsx',
    index=False
)

print("Processo finalizado com sucesso 🚀")
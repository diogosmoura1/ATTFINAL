import pandas as pd
from dicionary import empresas_dic


def ler_google_sheets(link):
    """
    Recebe o link do Google Sheets e retorna um DataFrame
    """
    sheet_id = link.split("/d/")[1].split("/")[0]
    url_csv = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    return pd.read_csv(url_csv)


def tratar_dataframe(df, empresa):
    df['EMPRESA'] = empresa

    colunas_para_tratar = ['valor contratado', 'saldo devedor', 'valor parcela']
    
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


def main():
    dataframes = []

    for item in empresas_dic:
        link = item['link']
        empresa = item['empresa']

        print(f"Puxando dados da {empresa}...")

        df = ler_google_sheets(link)
        df = tratar_dataframe(df, empresa)

        dataframes.append(df)

    # junta tudo em uma única aba
    df_final = pd.concat(dataframes, ignore_index=True)

    # salva o Excel final
    df_final.to_excel("google_sheets_geral.xlsx", index=False)

    print("Arquivo 'google_sheets_geral.xlsx' criado com sucesso!")


if __name__ == "__main__":
    main()
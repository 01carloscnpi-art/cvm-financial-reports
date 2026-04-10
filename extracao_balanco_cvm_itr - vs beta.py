import pandas as pd
import zipfile
import requests
import io

# 1. Baixar o arquivo ZIP da CVM
url = "https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/ITR/DADOS/itr_cia_aberta_2023.zip"
r = requests.get(url)
z = zipfile.ZipFile(io.BytesIO(r.content))

# 2. Extrair o CSV
with z.open("itr_cia_aberta_BPP_con_2023.csv") as f:
    df = pd.read_csv(f, sep=";", encoding="latin1")

# 3. Filtrar pelo CNPJ
df = df[df["CNPJ_CIA"] == "33.611.500/0001-19"]

# 4. Transformar valores
df["VL_CONTA"] = pd.to_numeric(df["VL_CONTA"], errors="coerce") 

# 5. Remover colunas desnecessárias
cols_remove = ["CNPJ_CIA","DT_REFER","VERSAO","DENOM_CIA","CD_CVM",
               "GRUPO_DFP","MOEDA","ESCALA_MOEDA","ORDEM_EXERC"]
df = df.drop(columns=cols_remove)

# 6. Pivotar incluindo CD_CONTA e DS_CONTA
df_pivot = df.pivot_table(
    index=["CD_CONTA", "DS_CONTA"],
    columns="DT_FIM_EXERC",
    values="VL_CONTA",
    aggfunc="first"
)

#7.  Visualizar as 10 primeiras linhas com formatação brasileira
df_pivot.head(10).style.format(
    lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
)

# 8. Resetar índice para que CD_CONTA e DS_CONTA fiquem como colunas
df_pivot = df_pivot.reset_index()

# 9. Exportar para Excel
output_path = r"ENDEREÇO\NOME DO ARQUIVO.xlsx"
df_pivot.to_excel(output_path, engine="openpyxl", index=False)

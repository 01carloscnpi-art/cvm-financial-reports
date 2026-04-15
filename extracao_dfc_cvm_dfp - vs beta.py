import pandas as pd
import requests
import zipfile
import io
import os

# Lista de anos desejados
anos = [2023, 2024, 2025]

# Base da URL
base_url = "https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/DFP/DADOS/dfp_cia_aberta_{}.zip"

# Dicionário para armazenar os DataFrames DFC pivotados
dfs_dfc_pivot = {}

# CNPJ alvo
cnpj_alvo = "33.611.500/0001-19"

# Colunas a remover
cols_remove = ["CNPJ_CIA","DT_REFER","VERSAO","DENOM_CIA","CD_CVM",
               "GRUPO_DFP","MOEDA","ESCALA_MOEDA","ORDEM_EXERC","DT_INI_EXERC"]

for ano in anos:
    url = base_url.format(ano)
    print(f"Baixando arquivo: {url}")
    
    r = requests.get(url)
    if r.status_code == 200:
        pasta_destino = f"./dados_cvm/{ano}"
        os.makedirs(pasta_destino, exist_ok=True)
        
        z = zipfile.ZipFile(io.BytesIO(r.content))
        z.extractall(pasta_destino)
        print(f"Arquivos de {ano} extraídos com sucesso em {pasta_destino}!")
        
        # Procurar apenas o arquivo DFC dentro do ZIP
        for nome_arquivo in z.namelist():
            if nome_arquivo.startswith("dfp_cia_aberta_DFC_MI_con_") and nome_arquivo.endswith(".csv"):
                print(f"Lendo {nome_arquivo}...")
                with z.open(nome_arquivo) as f:
                    df = pd.read_csv(f, sep=";", encoding="latin1")
                    
                    # 1. Filtrar pelo CNPJ
                    df = df[df["CNPJ_CIA"] == cnpj_alvo]
                    
                    # 2. Converter valores para numérico
                    df["VL_CONTA"] = pd.to_numeric(df["VL_CONTA"], errors="coerce")
                    
                    # 3. Remover colunas desnecessárias
                    df = df.drop(columns=cols_remove)
                    
                    # 4. Pivotar
                    df_pivot = df.pivot_table(
                        index=["CD_CONTA", "DS_CONTA"],
                        columns="DT_FIM_EXERC",
                        values="VL_CONTA",
                        aggfunc="first"
                    )
                    
                    # 5. Resetar índice para que CD_CONTA e DS_CONTA fiquem como colunas
                    df_pivot = df_pivot.reset_index()
                    
                    # 6. Guardar versão pivotada completa (todas as linhas)
                    dfs_dfc_pivot[ano] = df_pivot
    else:
        print(f"Falha ao baixar arquivo de {ano}. Status: {r.status_code}")

# 7. Exportar todos os anos para Excel (cada ano em uma aba separada)
output_path = r"C:\Users\econo\Downloads\DFC_DFP_2023_2025.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for ano, df in dfs_dfc_pivot.items():
        df.to_excel(writer, sheet_name=str(ano), index=False)

print(f"Arquivo Excel gerado em: {output_path}")

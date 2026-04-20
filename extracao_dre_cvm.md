import pandas as pd
import requests
import zipfile
import io
import os
import re

# =============================================================================
# 1. CONFIGURAÇÕES E PARÂMETROS (EXTRACT)
# =============================================================================
cnpj_alvo = "61.585.865/0001-51"  # Raia Drogasil
anos = [2023, 2024, 2025, 2026]

fontes = [
    ("https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/DFP/DADOS/dfp_cia_aberta_{}.zip", "dfp"),
    ("https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/ITR/DADOS/itr_cia_aberta_{}.zip", "itr")
]

# Dicionário de "De-Para" para abreviações e padronizações (TRANSFORM)
map_abreviacoes = {
    r'\bDEPREC\.\b': 'DEPRECIAÇÃO',
    r'\bAMORT\.\b': 'AMORTIZAÇÃO',
    r'\bINCID\.\b': 'INCIDENTES',
    r'\bOP\.\b': 'OPERACIONAIS',
    r'\bLIQ\.\b': 'LÍQUIDA',
    r'\bRESULT\.\b': 'RESULTADO',
    r'\bVEND\.\b': 'VENDAS',
}

# Dicionário de "De-Para" para nomes específicos de contas
de_para_nomes = {
    "IMPOSTOS INCIDENTES SOBRE VENDAS": "IMPOSTOS SOBRE VENDAS",
    "IMPOSTOS INCIDENTES SOBRE AS VENDAS": "IMPOSTOS SOBRE VENDAS",
    "RECEITA LÍQUIDA DE VENDAS E/OU SERVIÇOS": "RECEITA LÍQUIDA",
}

def limpar_descricao(texto):
    if not isinstance(texto, str): return texto
    texto = texto.upper().strip()
    # Aplica as substituições de abreviações
    for padrao, substituto in map_abreviacoes.items():
        texto = re.sub(padrao, substituto, texto)
    # Remove pontos e espaços duplos
    texto = texto.replace('.', '').replace('  ', ' ')
    return texto

# =============================================================================
# 2. PROCESSO DE EXTRAÇÃO (EXTRACT)
# =============================================================================
dados_brutos = []

for ano in anos:
    for url_base, prefixo in fontes:
        url = url_base.format(ano)
        print(f"Buscando {prefixo.upper()} {ano}...")
        try:
            r = requests.get(url, timeout=30)
            if r.status_code == 200:
                with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                    nome_csv = f"{prefixo}_cia_aberta_DRE_con_{ano}.csv"
                    if nome_csv in z.namelist():
                        with z.open(nome_csv) as f:
                            df = pd.read_csv(f, sep=";", encoding="latin1")
                            df = df[df["CNPJ_CIA"].str.strip() == cnpj_alvo].copy()
                            if not df.empty:
                                dados_brutos.append(df)
        except Exception as e:
            print(f"Erro ao processar {url}: {e}")

# =============================================================================
# 3. PROCESSO DE TRANSFORMAÇÃO (TRANSFORM)
# =============================================================================
if dados_brutos:
    df_full = pd.concat(dados_brutos, ignore_index=True)
    
    # 1. Normalização de Valores
    df_full["VL_CONTA"] = pd.to_numeric(df_full["VL_CONTA"], errors="coerce")
    
    # 2. Normalização de Descrições (Abreviações e Regex)
    df_full["DS_CONTA"] = df_full["DS_CONTA"].apply(limpar_descricao)
    
    # 3. Aplicação do De-Para de nomes específicos
    df_full["DS_CONTA"] = df_full["DS_CONTA"].replace(de_para_nomes)
    
    # 4. Agrupamento (ETL): Resolve duplicidade de contas com nomes ligeiramente diferentes
    # Agrupamos por CD_CONTA (âncora fiel) e pegamos a descrição normalizada
    df_agrupado = df_full.groupby(["CD_CONTA", "DS_CONTA", "DT_FIM_EXERC"])["VL_CONTA"].sum().reset_index()
    
    # 5. Pivotagem para Série Histórica
    df_pivot = df_agrupado.pivot(
        index=["CD_CONTA", "DS_CONTA"],
        columns="DT_FIM_EXERC",
        values="VL_CONTA"
    ).reset_index()
    
    # 6. Ordenação Cronológica das Colunas
    colunas_id = ["CD_CONTA", "DS_CONTA"]
    colunas_datas = sorted([c for c in df_pivot.columns if c not in colunas_id])
    df_final = df_pivot[colunas_id + colunas_datas]

    # =============================================================================
    # 4. CARGA (LOAD)
    # =============================================================================
    output_path = r"C:\Users\caminho\Serie_Historica_DRE_ETL.xlsx"
    
    # Fallback caso o caminho acima não exista
    if not os.path.exists(os.path.dirname(output_path)):
        output_path = "Serie_Historica_DRE_ETL.xlsx"
        
    df_final.to_excel(output_path, index=False)
    
    print(f"\n✓ Sucesso! Série histórica consolidada com {len(colunas_datas)} períodos.")
    print(f"✓ Arquivo gerado: {output_path}")
    
    # Mostra o resultado das primeiras linhas para validação no Jupyter
    display(df_final.head(10))
else:
    print("Nenhum dado foi encontrado.")

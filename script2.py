import pandas as pd
print("Iniciando")
# Nome do primeiro arquivo Excel e as abas a serem lidas
arquivo_excel = 'Pedidos.xlsx'
# nomes_abas = ['abril21', 'maio21', 'junho21', 'julho21', 'agosto21', 'setembro21', 'outubro21', 'novembro21', 'dezembro21', 'janeiro22', 'fevereiro22', 'março22', 'abril22', 'maio22', 'junho22', 'julho22', 'agosto22', 'setembro22', 'novembro22', 'dezembro22']
nomes_abas = ['Pedido1', 'Pedido2', 'Pedido3']

# Ler todas as abas do primeiro arquivo e armazenar em uma lista de DataFrames
df_abas = [pd.read_excel(arquivo_excel, sheet_name=aba) for aba in nomes_abas]

# Nome do segundo arquivo Excel e a aba específica a ser lida
arquivo_excel2 = 'Pedido4.xlsx'
nome_aba_arquivo2 = 'Pedido4'

# Adicionar a aba específica do segundo arquivo à lista de DataFrames
df_aba_arquivo2 = pd.read_excel(arquivo_excel2, sheet_name=nome_aba_arquivo2) 
df_abas.append(df_aba_arquivo2)

# Identificar as colunas comuns entre todos os DataFrames
colunas_comuns = set(df_abas[0].columns)
for df in df_abas[1:]:
    colunas_comuns.intersection_update(df.columns)

# Converter colunas comuns para uma lista para uso no indexador do DataFrame
colunas_comuns = list(colunas_comuns)

# Função para encontrar linhas únicas que não estão presentes nos outros DataFrames
def encontrar_diferencas(df_base, outros_dfs):
    diferencas = df_base[~df_base[colunas_comuns].apply(tuple, 1).isin(
        pd.concat(outros_dfs)[colunas_comuns].apply(tuple, 1))]
    return diferencas

# Listar para armazenar todas as linhas únicas
df_diferencas_unicas = []

# Comparar cada DataFrame com os outros para encontrar as linhas únicas
for i, df_aba in enumerate(df_abas):
    outras_abas = df_abas[:i] + df_abas[i+1:]
    df_diferencas = encontrar_diferencas(df_aba, outras_abas)
    df_diferencas_unicas.append(df_diferencas)

# Concatenar todas as linhas únicas
df_final = pd.concat(df_diferencas_unicas).drop_duplicates(subset=colunas_comuns)

# Salvar o DataFrame final em um novo arquivo Excel
df_final.to_excel('Unial_final.xlsx', index=False)
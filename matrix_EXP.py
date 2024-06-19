import pandas as pd

# Caminho do arquivo Excel
file_path = r'C:\Users\moesios\Desktop\FASE2\EXP.xlsx'

# Ler as planilhas especificadas do arquivo Excel
df_estudos = pd.read_excel(file_path, sheet_name='ESTUDOS')
df_outros = pd.read_excel(file_path, sheet_name='OUTROS')
df_trabalho = pd.read_excel(file_path, sheet_name='TRABALHO')
df_ndomcililar = pd.read_excel(file_path, sheet_name='NDOMCILIAR')

# Função para transformar o dataframe de matriz para tabela
def transform_df(df):
    df = df.rename(columns={df.columns[0]: 'ZONA_ORIGEM'})
    df_melted = df.melt(id_vars=['ZONA_ORIGEM'], var_name='ZONA_DESTINO', value_name='VALOR')
    return df_melted

# Transformar cada dataframe
df_estudos_transformed = transform_df(df_estudos)
df_outros_transformed = transform_df(df_outros)
df_trabalho_transformed = transform_df(df_trabalho)
df_ndomcililar_transformed = transform_df(df_ndomcililar)

# Salvar os dataframes transformados em um novo arquivo Excel
output_path = r'C:\Users\moesios\Desktop\FASE2\arranjada_transformed.xlsx'
with pd.ExcelWriter(output_path) as writer:
    df_estudos_transformed.to_excel(writer, sheet_name='ESTUDOS', index=False)
    df_outros_transformed.to_excel(writer, sheet_name='OUTROS', index=False)
    df_trabalho_transformed.to_excel(writer, sheet_name='TRABALHO', index=False)
    df_ndomcililar_transformed.to_excel(writer, sheet_name='NDOMCILIAR', index=False)

print(f"Arquivo salvo em: {output_path}")

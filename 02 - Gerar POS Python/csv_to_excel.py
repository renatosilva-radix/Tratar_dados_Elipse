import pandas as pd

# Lista com os nomes dos arquivos CSV
csv_files = [
    r'E:\dados\Radix\projetos\docs\importador\Funcionalidades Botafogo v2.csv',
    r'E:\dados\Radix\projetos\docs\importador\Plantas Botafogo v2.csv'
]

# Nomes dos arquivos Excel de saída
output_plantas_excel = 'PRE_Plantas_e_Sistemas.xlsx'
output_funcionalidades_excel = 'PRE_Equipamentos_e_Funcionalidades.xlsx'

# Ler os arquivos CSV
df_planta = pd.read_csv(csv_files[1],sep=';')
df_funcionalidades = pd.read_csv(csv_files[0],sep=';')


# Converter colunas booleanas para strings
df_planta = df_planta.applymap(lambda x: str(x).replace(' ', '') if isinstance(x, (str, bool)) else x)
df_funcionalidades = df_funcionalidades.applymap(lambda x: str(x).replace(' ', '') if isinstance(x, (str, bool)) else x)


# Criar o primeiro arquivo Excel com a aba correspondente
with pd.ExcelWriter(output_plantas_excel, engine='openpyxl') as writer:
    df_planta.to_excel(writer, sheet_name='Lista de Plantas e Sistemas', index=False)

# Criar o segundo arquivo Excel com a aba correspondente
with pd.ExcelWriter(output_funcionalidades_excel, engine='openpyxl') as writer:
    df_funcionalidades.to_excel(writer, sheet_name='Equipamentos e Funcionalidades', index=False)

print(f'Arquivo Excel {output_plantas_excel} criado com sucesso.')
print(f'Arquivo Excel {output_funcionalidades_excel} criado com sucesso.')

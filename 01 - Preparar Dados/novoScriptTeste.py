import os
import pandas as pd

# Lista de IDs e Nomes
pastas = [
    (22, 'Colinas 2'),
    (23, 'Dunas')
]

# Caminho para os arquivos Excel
input_file_plantas = 'plantasExportar.xlsx'
input_file_novo = 'exportar.xlsx'

# Ler as planilhas Excel
df_plantas = pd.read_excel(input_file_plantas)
df_novo = pd.read_excel(input_file_novo)

# Obter a pasta onde os arquivos Excel estão localizados
directory = os.path.dirname(os.path.abspath(input_file_plantas))

# Obter os IDs únicos das colunas relevantes
ids_plantas = df_plantas['ID'].unique()
ids_novo = df_novo['ID na Plantas e Sistemas e/ou Áreas'].unique()

# Criar pastas e arquivos Excel
for id_value, nome in pastas:
    # Criar a pasta com base na lista
    folder_name = f"{id_value} - {nome}"
    os.makedirs(folder_name, exist_ok=True)
    print(f'Pasta criada: {folder_name}')
    
    # Verificar e salvar o arquivo Excel (.xlsm) do primeiro Excel
    if id_value in ids_plantas:
        df_filtered_plantas = df_plantas[df_plantas['ID'] == id_value]
        output_file_plantas = os.path.join(folder_name, f'Plantas e Sistemas - {folder_name}.xlsm')
        with pd.ExcelWriter(output_file_plantas, engine='openpyxl', mode='w') as writer:
            df_filtered_plantas.to_excel(writer, index=False, sheet_name='Lista de Plantas e Sistemas')
        print(f'Salvo: {output_file_plantas}')
    else:
        print(f'ID {id_value} não encontrado no arquivo {input_file_plantas}.')
    
    # Verificar e salvar o arquivo Excel (.xlsx) do segundo Excel
    if id_value in ids_novo:
        df_filtered_novo = df_novo[df_novo['ID na Plantas e Sistemas e/ou Áreas'] == id_value]
        output_file_novo = os.path.join(folder_name, f'Levantamento Plantas e Funcionalidades - {folder_name}.xlsx')
        with pd.ExcelWriter(output_file_novo, engine='openpyxl', mode='w') as writer:
            df_filtered_novo.to_excel(writer, index=False, sheet_name='Equipamentos e Funcionalidades')
        print(f'Salvo: {output_file_novo}')
    else:
        print(f'ID {id_value} não encontrado no arquivo {input_file_novo}.')

print('Processo concluído.')

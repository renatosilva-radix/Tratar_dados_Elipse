import os
import pandas as pd

def converter_excel_para_csv(arquivo_excel_1, arquivo_excel_2, diretorio_saida, planta_atual):
    # Carrega as abas dos arquivos Excel
    df_plantas_sistemas = pd.read_excel(arquivo_excel_1, sheet_name='Lista de Plantas e Sistemas')
    df_funcionalidades = pd.read_excel(arquivo_excel_2, sheet_name='Equipamentos e Funcionalidades')

    # Define os caminhos dos arquivos CSV de saída
    csv_plantas_sistemas = os.path.join(diretorio_saida, f"{planta_atual}Plantas_e_Sistemas.csv")
    csv_funcionalidades = os.path.join(diretorio_saida, f"{planta_atual}Levantamento_Plantas_e_Funcionalidades.csv")

    # Converte as abas em arquivos CSV separados por ';'
    df_plantas_sistemas.to_csv(csv_plantas_sistemas, sep=';', index=False, encoding='utf-8-sig')
    df_funcionalidades.to_csv(csv_funcionalidades, sep=';', index=False, encoding='utf-8-sig')

    # Informar ao usuário que o processo foi concluído
    print(f"Arquivos CSV criados com sucesso!\n- {csv_plantas_sistemas}\n- {csv_funcionalidades}")

def encontrar_arquivos(diretorio_base):
    arquivos = {}
    for root, dirs, files in os.walk(diretorio_base):
        for file in files:
            if file.startswith("Levantamento Plantas e Funcionalidades") and file.endswith((".xlsx", ".xlsm")):
                arquivos["funcionalidades"] = os.path.join(root, file)
            if file.startswith("Plantas e Sistemas") and file.endswith((".xlsx", ".xlsm")):
                arquivos["plantas_sistemas"] = os.path.join(root, file)
    return arquivos

def criar_pasta_destino(diretorio_origem, diretorio_destino):
    # Calcula o caminho relativo da origem para a base de PRE_Python
    pasta_relativa = os.path.relpath(diretorio_origem, start=diretorio_pre_python)
    # Cria o caminho de destino completo em POS_Python
    caminho_destino = os.path.join(diretorio_destino, pasta_relativa)
    os.makedirs(caminho_destino, exist_ok=True)
    return caminho_destino

def main():
    # Obtém o diretório onde o script está localizado
    diretorio_principal = os.path.dirname(os.path.abspath(__file__))
    if not diretorio_principal:
        print("Não foi possível determinar o diretório principal.")
        return

    # Diretórios base
    global diretorio_pre_python
    diretorio_pre_python = os.path.join(diretorio_principal, 'PRE_Python')
    diretorio_pos_python = os.path.join(diretorio_principal, 'POS_Python')

    # Cria o diretório POS_Python se não existir
    os.makedirs(diretorio_pos_python, exist_ok=True)

    # Encontra arquivos e processa cada pasta
    for root, dirs, files in os.walk(diretorio_pre_python):
        arquivos = encontrar_arquivos(root)
        if 'funcionalidades' in arquivos and 'plantas_sistemas' in arquivos:
            pasta_destino = criar_pasta_destino(root, diretorio_pos_python)
            converter_excel_para_csv(arquivos['plantas_sistemas'],
                                     arquivos['funcionalidades'],
                                     pasta_destino,
                                     planta_atual='')  # Defina o valor de `planta_atual` conforme necessário

if __name__ == "__main__":
    main()

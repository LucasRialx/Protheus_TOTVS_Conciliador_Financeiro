import os
import pandas as pd

# Caminhos das pastas
input_folder = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosBase\\FINANCEIRO'
output_folder = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosGerados\\FINANCEIRO'

# Nome do arquivo específico que você deseja processar
specific_file = 'TITULOS_A_PAGAR.xlsx'  # Substitua pelo nome da sua planilha

# Cria a pasta de saída, se não existir
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Caminho completo do arquivo específico
file_path = os.path.join(input_folder, specific_file)

# Verifica se o arquivo existe
if os.path.exists(file_path):
    # Lê a aba "1-Titulos a pagar" da planilha
    df = pd.read_excel(file_path, sheet_name='1-Titulos a pagar')
    
    # Filtra os valores da coluna A para manter apenas "DP", "FT" e "NF"
    df_filtered = df[df['Tp'].isin(['DP', 'FT', 'NF'])].copy()  # Use .copy() para evitar SettingWithCopyWarning
    
    # Concatena as colunas 'Codigo' e 'Loja' mantendo os zeros
    df_filtered['BC_concatenado'] = df_filtered['Codigo'].astype(str).str.zfill(6) + df_filtered['Loja'].astype(str).str.zfill(2)
    
    # Soma as colunas 'Tit Vencidos Valor corrigido' e 'Titulos a vencer Valor nominal', preenchendo valores ausentes com 0
    df_filtered['DF_soma'] = df_filtered['Tit Vencidos Valor corrigido'].fillna(0) + df_filtered['Titulos a vencer Valor nominal'].fillna(0)
    
    # Agrupa pelo campo concatenado (B + C) e soma os valores de D + F
    df_grouped = df_filtered.groupby('BC_concatenado')['DF_soma'].sum().reset_index()
    
    # Define o nome do novo arquivo
    new_filename = f'FORN_{specific_file}'
    new_file_path = os.path.join(output_folder, new_filename)
    
    # Salva o novo DataFrame contendo apenas as colunas concatenadas e a soma
    df_grouped.to_excel(new_file_path, index=False)

    print("Processamento concluído!")
else:
    print(f"O arquivo {specific_file} não foi encontrado na pasta especificada.")

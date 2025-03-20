import os
import pandas as pd

# Função para comparar as duas planilhas e identificar as diferenças nos valores dos itens
def compare_item_values(df_A, df_B, item_column_A, value_column_A, item_column_B, value_column_B):
    # Assegura que as colunas de itens sejam strings (mantendo os zeros à esquerda)
    df_A[item_column_A] = df_A[item_column_A].astype(str)
    df_B[item_column_B] = df_B[item_column_B].astype(str)
    
    # Mescla as duas planilhas com base nos itens (mantém todos os itens usando outer join)
    df_merged = pd.merge(df_A, df_B, left_on=item_column_A, right_on=item_column_B, how='outer', suffixes=('_C', '_F'))
    
    # Preenche os valores ausentes com 0 para evitar problemas de comparação
    df_merged[value_column_A] = df_merged[value_column_A].fillna(0)
    df_merged[value_column_B] = df_merged[value_column_B].fillna(0)
    
    # Converte as colunas de valores para float para garantir que a subtração funcione corretamente
    df_merged[value_column_A] = df_merged[value_column_A].astype(float)
    df_merged[value_column_B] = df_merged[value_column_B].astype(float)
    
    # Cria uma nova coluna com a diferença dos valores
    df_merged['DIFERENCA'] = df_merged[value_column_A] - df_merged[value_column_B]
    
    # Filtra as linhas onde a diferença de valores não é zero
    df_differences = df_merged[df_merged['DIFERENCA'].notna() & (df_merged['DIFERENCA'] != 0)]
    
    # Renomeia as colunas conforme solicitado
    df_differences = df_differences.rename(columns={
        item_column_A: 'CODIGO CLIENTE C.',
        value_column_A: 'VALOR CONTABIL',
        item_column_B: 'CODIGO CLIENTE F.',
        value_column_B: 'VALOR FINANCEIRO'
    })
    
    return df_differences

# Caminhos das pastas
input_folder_A = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosGerados\\CONTABIL'
input_folder_B = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosGerados\\FINANCEIRO'
output_folder = 'C:\\VSCode\\AutomacaoFinanceiro\\Resultado'

planilha_A = 'DATABASE_BALANCETE_CLIENTE.xlsx'
planilha_B = 'CLIENTES_TITULOS_A_RECEBER.xlsx'

# Carrega as planilhas forçando a leitura das colunas de código como strings para evitar perda de zeros à esquerda
df_A = pd.read_excel(os.path.join(input_folder_A, planilha_A), sheet_name='Sheet1', dtype={'Item Conta': str})
df_B = pd.read_excel(os.path.join(input_folder_B, planilha_B), sheet_name='Sheet1', dtype={'Concatenado_AB': str})

# Define as colunas para comparação
item_column_A = 'Item Conta'  # Coluna de itens na Planilha A
value_column_A = 'Saldo Atual'  # Coluna de valores na Planilha A
item_column_B = 'Concatenado_AB'  # Coluna de itens na Planilha B
value_column_B = 'Soma'  # Coluna de valores na Planilha B

# Compara os valores dos itens entre as duas planilhas
df_differences = compare_item_values(df_A, df_B, item_column_A, value_column_A, item_column_B, value_column_B)

# Define o caminho do arquivo de resultado
result_filename = 'RESULTADO_CLIENTES.xlsx'
result_file_path = os.path.join(output_folder, result_filename)

# Salva o DataFrame de diferenças no arquivo de saída
df_differences.to_excel(result_file_path, index=False)

print("Comparação concluída! As diferenças foram salvas.")

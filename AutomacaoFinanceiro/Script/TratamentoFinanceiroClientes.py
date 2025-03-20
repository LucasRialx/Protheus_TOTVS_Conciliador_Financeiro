import os
import pandas as pd

# Caminhos das pastas
input_folder = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosBase\\FINANCEIRO'
output_folder = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosGerados\\FINANCEIRO'

# Nome do arquivo específico que você deseja processar
specific_file = 'TITULOS_A_RECEBER.xlsx'

# Cria a pasta de saída, se não existir
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Caminho completo do arquivo específico
file_path = os.path.join(input_folder, specific_file)

# Verifica se o arquivo existe
if os.path.exists(file_path):
    # Lê a aba "2-Titulos a receber" da planilha
    df = pd.read_excel(file_path, sheet_name='2-Titulos a receber')

    # --------- FILTRA OS VALORES DA COLUNA 'TP' ---------
    filtros = ['BL', 'BOL', 'CC', 'CD', 'CF-', 'CH', 'CS-', 'DB', 'DP', 'IN-', 'IR-', 'NF', 'PD', 'PI-', 'TR','FI', 'US', 'R$']
    df_filtered = df[df['TP'].isin(filtros)].copy()

    # --------- PRIMEIRO TRATAMENTO: FILTRAR, CONCATENAR, SOMAR ---------

    # Filtra as linhas que **não** contêm "REDECARD" no campo 'Nome'
    df_normal = df_filtered[~df_filtered['Nome'].str.contains('REDECARD', case=False, na=False)].copy()

    # Concatenar Codigo e Loja, garantindo que ambos sejam strings
    df_normal['Concatenado_AB'] = df_normal['Codigo'].astype(str).str.zfill(6) + df_normal['Loja'].astype(str).str.zfill(2)

    # Soma os valores de "Tit Vencidos Valor Corrigido" e "Titulos a Vencer Valor Atual"
    df_normal['Soma'] = df_normal.apply(
        lambda row: (row['Tit Vencidos Valor Corrigido'] if pd.notna(row['Tit Vencidos Valor Corrigido']) else 0) + 
                    (row['Titulos a Vencer Valor Atual'] if pd.notna(row['Titulos a Vencer Valor Atual']) else 0),
        axis=1
    )

    # Agrupa os dados pelo "Concatenado_AB" e soma os valores
    df_normal_grouped = df_normal.groupby('Concatenado_AB')['Soma'].sum().reset_index()

    # --------- SEGUNDO TRATAMENTO: FILTRAR APENAS REDECARD ---------

    # Filtra as linhas que **contêm** "REDECARD" no campo 'Nome'
    df_redec = df_filtered[df_filtered['Nome'].str.contains('REDECARD', case=False, na=False)].copy()

    # Limpa e formata as colunas 'Cli. Origem' e 'Loj. Origem'
    df_redec['Cli. Origem'] = df_redec['Cli. Origem'].astype(str).str.replace('.0', '', regex=False).str.zfill(6)
    df_redec['Loj. Origem'] = df_redec['Loj. Origem'].astype(str).str.replace('.0', '', regex=False).str.zfill(2)

    # Concatenar Cli. Origem e Loj. Origem, garantindo que ambos sejam strings
    df_redec['Concatenado_AB'] = df_redec['Cli. Origem'] + df_redec['Loj. Origem'].astype(str).str.zfill(2)

    # Usar o valor da coluna "Valor Real" como a Soma
    df_redec['Soma'] = df_redec['Valor Real']

    # Agrupa os dados pelo "Concatenado_AB" e soma os valores
    df_redec_grouped = df_redec.groupby('Concatenado_AB')['Soma'].sum().reset_index()

    # --------- TERCEIRO TRATAMENTO: JUNTAR E SOMAR ITENS IGUAIS ---------

    # Concatena os dois DataFrames
    df_combined = pd.concat([df_normal_grouped, df_redec_grouped], ignore_index=True)

    # Agrupa novamente os dados para somar os itens iguais
    df_final = df_combined.groupby('Concatenado_AB')['Soma'].sum().reset_index()

    # Define o nome do arquivo final
    final_filename = 'CLIENTES_' + specific_file
    final_file_path = os.path.join(output_folder, final_filename)

    # Salva o DataFrame final
    df_final.to_excel(final_file_path, index=False)

    print("Processamento concluído! Dados combinados salvos.")
else:
    print(f"O arquivo {specific_file} não foi encontrado na pasta especificada.")

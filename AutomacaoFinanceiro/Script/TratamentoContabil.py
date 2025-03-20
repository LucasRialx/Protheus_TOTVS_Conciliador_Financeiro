# -*- coding: utf-8 -*-
# Projeto desenvolvido por Lucas Rial
# Copyright (c) 2025 Lucas Rial. Todos os direitos reservados.

import os
import pandas as pd
import re

# Função para obter os últimos 8 caracteres de uma string
def get_last_8_characters(text):
    if isinstance(text, str):
        # Retorna os últimos 8 caracteres
        return text[-8:] if len(text) >= 8 else text
    return text  # Retorna o texto original se não for string

# Função para extrair o número de "Saldo Atual", mantendo a vírgula como separador decimal
def extract_decimal_number_with_comma(text):
    if isinstance(text, str):
        # Remove caracteres indesejados, exceto números, pontos e vírgulas
        clean_text = re.sub(r'[^\d.,]', '', text)

        # Se houver ponto e vírgula, assumimos que o ponto é separador de milhar e a vírgula é o decimal
        if clean_text.count(',') == 1 and clean_text.count('.') > 0:
            # Remove o ponto de milhares e troca a vírgula pelo ponto decimal para conversão correta
            clean_text = clean_text.replace('.', '').replace(',', '.')

        # Se houver apenas vírgula, ela será o separador decimal
        elif clean_text.count(',') == 1:
            clean_text = clean_text.replace(',', '.')

        # Converte a string para float
        try:
            return float(clean_text)
        except ValueError:
            return None
    return None

# Caminhos das pastas
input_folder = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosBase\\CONTABIL'
output_folder = 'C:\\VSCode\\AutomacaoFinanceiro\\ArquivosGerados\\CONTABIL'

# Cria a pasta de saída, se não existir
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Itera sobre cada arquivo na pasta de entrada
for filename in os.listdir(input_folder):
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        # Caminho completo do arquivo
        file_path = os.path.join(input_folder, filename)
        
        # Lê a aba "2-Conta" da planilha
        df = pd.read_excel(file_path, sheet_name='2-Conta')
        
        # Aplica a função para obter os últimos 8 caracteres da coluna "Item Conta"
        df['Item Conta'] = df['Item Conta'].apply(get_last_8_characters)
        
        # Aplica a função para extrair números corretamente da coluna "Saldo Atual"
        df['Saldo Atual'] = df['Saldo Atual'].apply(extract_decimal_number_with_comma)
        
        # Preenche os valores vazios na coluna 'Saldo Atual' com zero
        df['Saldo Atual'] = df['Saldo Atual'].fillna(0)
        
        # Seleciona apenas as colunas 'Item Conta' e 'Saldo Atual'
        df_filtered = df[['Item Conta', 'Saldo Atual']]
        
        # Define o nome do novo arquivo
        new_filename = f'DATABASE_{filename}'
        new_file_path = os.path.join(output_folder, new_filename)
        
        # Salva o novo DataFrame contendo apenas as colunas 'Item Conta' e 'Saldo Atual'
        df_filtered.to_excel(new_file_path, index=False)

print("Processamento concluído!")

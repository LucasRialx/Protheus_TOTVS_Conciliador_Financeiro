# -*- coding: utf-8 -*-
# Projeto desenvolvido por Lucas Rial
# Copyright (c) 2025 Lucas Rial. Todos os direitos reservados.

import subprocess
import os
import time  # Importa o módulo time para usar a função sleep

# Caminho da pasta onde estão os scripts
scripts_folder = os.path.join('C:\\VSCode\\AutomacaoFinanceiro\\Script')

# Lista de scripts a serem executados
scripts = [
    'TratamentoContabil.py',
    'TratamentoFinanceiroAdtCliente.py',
    'TratamentoFinanceiroAdtFornecedor.py',
    'TratamentoFinanceiroAdtViagem.py',
    'TratamentoFinanceiroClientes.py',
    'TratamentoFinanceiroFornecedor.py',
    'ConflitoAdtCliente.py',
    'ConflitoAdtFornecedor.py',
    'ConflitoAdtViagem.py',
    'ConflitoCliente.py',
    'ConflitoFornecedor.py'
]

def run_script(script_name):
    try:
        # Executa o script e aguarda a conclusão
        result = subprocess.run(['python', os.path.join(scripts_folder, script_name)], capture_output=True, text=True)
        # Verifica se o script foi executado com sucesso
        if result.returncode == 0:
            print(f"Script '{script_name}' executado com sucesso!")
            print(result.stdout)  # Exibe a saída padrão do script
        else:
            print(f"Erro ao executar '{script_name}':")
            print(result.stderr)  # Exibe a saída de erro do script
    except Exception as e:
        print(f"Ocorreu um erro ao tentar executar '{script_name}': {e}")

if __name__ == "__main__":
    for script in scripts:
        run_script(script)
        time.sleep(3)  # Pausa de 3 segundos antes de executar o próximo script

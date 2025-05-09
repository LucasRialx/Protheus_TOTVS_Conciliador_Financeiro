# Protheus TOTVS - Conciliação Financeira

Este projeto tem como objetivo automatizar a conciliação contábil e financeira no sistema Protheus TOTVS, utilizando os códigos de clientes e fornecedores. A ferramenta cruza esses dados com os itens contábeis para facilitar a conferência e ajustes necessários. Empresas que utilizam essa metodologia podem adaptar o código às suas necessidades.

## Recursos

O sistema realiza a conciliação das seguintes entidades financeiras:
- **Clientes**
- **Fornecedores**
- **Adiantamentos**

## Relatórios Necessários

Para o correto funcionamento da automação, é necessário fornecer os seguintes relatórios:

- **Contas a Pagar Financeiro**
- **Contas a Receber Financeiro**
- **Balancetes Contábeis por Item Contábil**
  - O item contábil deve ser a concatenação do código, loja e fornecedor ou cliente.

## Estrutura do Projeto

A estrutura do projeto é organizada da seguinte forma:

```
Protheus_Conciliador/

AutomacaoFinanceiro
│
├── ArquivosGerados
│   ├── CONTABIL
│   │   ├── DATABASE_BALANCETE_CLIENTE.xlsx
│   │   ├── DATABASE_BALANCETE_FORNECEDOR.xlsx
│   │
│   ├── FINANCEIRO
│   │   ├── ADT_CLIENTES_TITULOS_A_RECEBER.xlsx
│   │   ├── ADT_FORN_TITULOS_A_PAGAR.xlsx
│   │   ├── ADT_VIAGEM_TITULOS_A_PAGAR.xlsx
│   │   ├── CLIENTES_TITULOS_A_RECEBER.xlsx
│   │   ├── FORN_TITULOS_A_PAGAR.xlsx
│
├── Resultado
│   ├── RESULTADO_FORNECEDOR.xlsx
│   ├── RESULTADO_ADT_CLIENTES.xlsx
│   ├── RESULTADO_ADT_FORNECEDOR.xlsx
│   ├── RESULTADO_ADT_VIAGEM.xlsx
│   ├── RESULTADO_CLIENTES.xlsx
│   ├── RESULTADO_FORNECEDOR.xlsx
│
├── Script
│   ├── ConflitoAdtCliente.py
│   ├── ConflitoAdtFornecedor.py
│   ├── ConflitoAdtViagem.py
│   ├── ConflitoCliente.py
│   ├── ConflitoFornecedor.py
│   ├── TratamentoContabil.py
│   ├── TratamentoFinanceiroAdtCliente.py
│   ├── TratamentoFinanceiroAdtFornecedor.py
│   ├── TratamentoFinanceiroAdtViagem.py
│   ├── TratamentoFinanceiroClientes.py
│   ├── TratamentoFinanceiroFornecedor.py
│
├── Processador.py

```

### Scripts Disponíveis

Os seguintes scripts são executados pelo `Processador.py`:

- **Tratamento Contábil**
  - `TratamentoContabil.py`
- **Tratamento Financeiro**
  - `TratamentoFinanceiroAdtCliente.py`
  - `TratamentoFinanceiroAdtFornecedor.py`
  - `TratamentoFinanceiroAdtViagem.py`
  - `TratamentoFinanceiroClientes.py`
  - `TratamentoFinanceiroFornecedor.py`
- **Conflitos Financeiros**
  - `ConflitoAdtCliente.py`
  - `ConflitoAdtFornecedor.py`
  - `ConflitoAdtViagem.py`
  - `ConflitoCliente.py`
  - `ConflitoFornecedor.py`

## Como Executar

1. Certifique-se de que todos os relatórios necessários estão nas pastas corretas e com a nomenclatura correta.
2. Verifique se todos os scripts estão na pasta `Script/`.

![image](https://github.com/user-attachments/assets/902289b8-31e9-4f7d-934b-dc5ace6d406b)



3. Execute o script principal `Processador.py`:

    ```sh
    python Processador.py
    ```

## Funcionamento

O `Processador.py` executa os scripts de forma sequencial, seguindo estes passos:

1. Define o caminho da pasta contendo os scripts.
2. Lista os scripts a serem executados.
3. Executa cada script da lista e aguarda sua conclusão.
4. Exibe a saída padrão ou a saída de erro do script.
5. Pausa por 3 segundos antes de iniciar a próxima execução.

## Resultado

O resultado é o apontamento de quais Fornecedores apresentam inconsistência entre Contábil e Financeiro em um arquivo XLSX (Excel);

![image](https://github.com/user-attachments/assets/fd6aec9a-e61b-4b27-80aa-c15586801201)

## Requisitos

- Python 3.x
- Bibliotecas utilizadas:
  - `subprocess`
  - `os`
  - `time`

## Aviso de Direitos Autorais
Este projeto foi desenvolvido por **Lucas Rial**. O código é de propriedade exclusiva de Lucas Rial e está licenciado sob a **Licença MIT**. A reutilização do código é permitida, desde que a autoria seja devidamente reconhecida. Se você utilizar ou modificar este código, deve incluir uma referência a **Lucas Rial** em qualquer distribuição ou derivação.

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para mais detalhes.


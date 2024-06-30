# **Sheet Comparator and Processor**

Este script Python compara dois arquivos CSV, remove linhas com valores específicos e converte o resultado para um arquivo Excel (XLSX). Além disso, permite ordenar o arquivo XLSX resultante por uma coluna específica.

Funcionalidades
Comparação de Arquivos CSV: Compara dois arquivos CSV com base em colunas específicas e gera um novo arquivo CSV contendo apenas as linhas correspondentes.
Remoção de Valores: Remove linhas de um arquivo CSV com base em palavras-chave específicas presentes em uma coluna.
Conversão de CSV para XLSX: Converte um arquivo CSV para o formato Excel (XLSX).
Ordenação de Arquivo XLSX: Ordena um arquivo XLSX com base em uma coluna especificada.
Estrutura do Código
ler_csv(file_path, delimiter=';', encoding='utf-8-sig'): Lê um arquivo CSV e retorna um DataFrame pandas.
escrever_csv(df, file_path, delimiter=';', encoding='utf-8-sig'): Escreve um DataFrame pandas em um arquivo CSV.
comparar_arquivos_csv(arquivo1, coluna1, arquivo2, coluna2, arquivo_saida): Compara dois arquivos CSV e escreve as linhas correspondentes em um novo arquivo CSV.
remover_valores(file_path, keywords, column_name): Remove linhas de um arquivo CSV que contenham palavras-chave específicas em uma coluna.
csv_para_xlsx(arquivo_csv): Converte um arquivo CSV para o formato Excel (XLSX).
ordenar_xlsx_por_coluna(arquivo_xlsx, nome_coluna): Ordena um arquivo XLSX com base em uma coluna especificada.
Como Usar
1. Comparar Arquivos CSV
Compara dois arquivos CSV e escreve as linhas correspondentes em um novo arquivo CSV.

python
Copy code
arquivo1 = 'relacao_servidores.csv'
arquivo2 = 'processos_completo.csv'
arquivo_saida = 'processos_verificados.csv'

comparar_arquivos_csv(arquivo1, 'NOME', arquivo2, 'NOME_PARTE_ATIVA', arquivo_saida)
2. Remover Valores
Remove linhas de um arquivo CSV que contenham palavras-chave específicas em uma coluna.

python
Copy code
file_path = 'processos_verificados.csv'
keywords = ["Mato Grosso", "município", "polícia"]
column_name = 'NOME_PARTE_ATIVA'

remover_valores(file_path, keywords, column_name)
3. Converter CSV para XLSX
Converte um arquivo CSV para o formato Excel (XLSX).

python
Copy code
arquivo_csv = 'processos_verificados.csv'
csv_para_xlsx(arquivo_csv)
4. Ordenar Arquivo XLSX
Ordena um arquivo XLSX com base em uma coluna especificada.

python
Copy code
arquivo_xlsx = 'processos_verificados.xlsx'
nome_coluna = 'NOME'

ordenar_xlsx_por_coluna(arquivo_xlsx, nome_coluna)
Pré-requisitos
Python 3.x
Pandas
Openpyxl
Você pode instalar os pacotes necessários usando o pip:

sh
Copy code
pip install pandas openpyxl
Exemplo Completo
Aqui está um exemplo completo de como usar todas as funcionalidades do script:

python
Copy code
import logging

arquivo1 = 'relacao_servidores.csv'
arquivo2 = 'processos_completo.csv'
arquivo_saida = 'processos_verificados.csv'
arquivo_xlsx_saida = 'processos_verificados.xlsx'

# Comparar arquivos CSV
comparar_arquivos_csv(arquivo1, 'NOME', arquivo2, 'NOME_PARTE_ATIVA', arquivo_saida)

# Remover valores indesejados
keywords = ["Estado de Mato Grosso", "Casa Civil", "Secretaria"]
remover_valores(arquivo_saida, keywords, 'NOME_PARTE_ATIVA')

# Converter CSV para XLSX
csv_para_xlsx(arquivo_saida)

# Ordenar XLSX por coluna
ordenar_xlsx_por_coluna(arquivo_xlsx_saida, 'NOME')
Contribuição
Sinta-se à vontade para abrir issues ou enviar pull requests para melhorias.

Licença
Este projeto está licenciado sob a Licença MIT - consulte o arquivo LICENSE para obter detalhes.
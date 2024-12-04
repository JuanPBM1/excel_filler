import pandas as pd
import openpyxl as xl
from datetime import datetime
from pathlib import Path

print(f"Programa para inserir linhas de um arquivo csv em um arquivo excel\n")
print(f"Por favor certifique-se que o excel está fechado antes de inserir qualquer informação neste programa. Caso contrário o arquivo pode corromper.\n")
print(f"Os arquivos csv e excel devem estar na pasta \"arquivos\"\n")
print(f"\n\n\n-------------------------INFORMAÇÕES DOS ARQUIVOS-------------------------")

caminho_csv = input("\nInsira o nome do arquivo csv (SEM A EXTENSÃO ex: dados):\n")
caminho_excel = input("\nInsira o nome do arquivo excel (COM A EXTENSÃO DO ARQUIVO ex: arquivo.xlsx): \n")


data = pd.read_csv(f"arquivos/{caminho_csv}.csv")
arquivo = xl.load_workbook(f"arquivos/{caminho_excel}")
backup = arquivo
aba_atual = arquivo.active
extensao_arquivo = Path(caminho_excel).suffix
nome_arquivo = Path(caminho_excel).stem

print("\n\n\n-------------------------DADOS A SEREM INSERIDOS-------------------------")
indice_inicial = int(input("\nInsira o número da linha do excel em que os dados devem começar a ser inseridos:\n")) # a linha do arquivo excel que deverá ser inserida
num_colunas = int(input("\nInsira o número de colunas que deverão ser inseridas no arquivo excel:\n")) # o número de colunas que deverão ser inseridas no arquivo excel
num_linhas = int(input("\nInsira o número de linhas que deverão ser inseridas no arquivo excel:\n")) # o número de linhas que deverão ser inseridas no arquivo excel

#Inserindo os valores no arquivo excel
linha_dataframe = 0
for i in range(indice_inicial, indice_inicial+num_linhas): # vai do índice inicial até o número de linhas que se deseja inserir
    for j in range(num_colunas-num_colunas+1,num_colunas+1):
        aba_atual.cell(row = i, column = j).value = data.iloc[linha_dataframe,j-1]
    linha_dataframe = linha_dataframe + 1

#Salvando os arquivos
data_e_hora = datetime.now()
data_hora_formatado = data_e_hora.strftime("%Y-%m-%d  %H-%M-%S")
arquivo.save(f"arquivos/{nome_arquivo}.{extensao_arquivo}")
backup.save(f"backups/backup{data_hora_formatado}.{extensao_arquivo}")
        

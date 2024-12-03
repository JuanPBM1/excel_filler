import pandas as pd
import openpyxl as xl
from datetime import datetime
from pathlib import Path

print(f"Programa para inserir linhas de um arquivo csv em um arquivo excel")
print(f"Por favor certifique-se que o excel está fechado antes de inserir qualquer informação neste programa. Caso contrário o arquivo pode corromper.")
print(f"Os arquivos csv e excel devem estar na pasta \"arquivos\"")

caminho_csv = input("Insira o nome do arquivo csv (SEM A EXTENSÃO)")
print(f"{caminho_csv}")
caminho_excel = input("Insira o nome do arquivo excel (COM A EXTENSÃO DO ARQUIVO ex: arquivo.xlsx)")

data = pd.read_csv(f"arquivos/{caminho_csv}.csv")
print(data)
arquivo = xl.load_workbook(f"arquivos/{caminho_excel}")
backup = arquivo
aba_atual = arquivo.active
extensao_arquivo = Path(caminho_excel).suffix

indice_inicial = int(input("Insira o número da linha do excel em que os dados devem começar a ser inseridos\n\n")) # a linha do arquivo excel que deverá ser inserida
num_colunas = int(input("Insira o número de colunas que deverão ser inseridas no arquivo excel")) # o número de colunas que deverão ser inseridas no arquivo excel
num_linhas = int(input("Insira o número de linhas que deverão ser inseridas no arquivo excel")) # o número de linhas que deverão ser inseridas no arquivo excel

#inserindo os valores no arquivo excel
linha_dataframe = 0
for i in range(indice_inicial, indice_inicial+num_linhas): # vai do índice inicial atpe o número de linhas que se deseja inserir
    for j in range(num_colunas-num_colunas+1,num_colunas+1):
        aba_atual.cell(row = i, column = j).value = data.iloc[linha_dataframe,j-1]
    linha_dataframe = linha_dataframe + 1

data_e_hora = datetime.now()
data_hora_formatado = data_e_hora.strftime("%Y-%m-%d  %H-%M-%S")
arquivo.save("arquivos/Employees.xlsx")
backup.save(f"backups/backup{data_hora_formatado}.{extensao_arquivo}")
        
PATH_POLITICAS = "C:\\Users\\luana\\Documents\\Exporta\\ArquivosOrigem\\"
PATH_TXT = "C:\\Users\\luana\\Documents\\Exporta\\ArquivosDestino\\"
PATH_XLSX = "C:\\Users\\luana\\Documents\\Exporta\\XLSX\\"

import os
from openpyxl import Workbook
import csv

#listando criterios da pasta
politicas = os.listdir(PATH_POLITICAS)

#convertendo os criterios em txt
for file in politicas:
    if '.txt' not in file:
        newFile = file + '.txt'
        os.rename(PATH_POLITICAS + file, PATH_TXT + newFile)
politicas = os.listdir(PATH_POLITICAS)

#ler criterios txt, selecionar variaveis e salvar novo txt
for file in politicas:
    with open(PATH_POLITICAS + file, 'r') as criterio, open(PATH_TXT + file, 'w') as formatado:
        for linha in criterio.readlines():
            if "TratNulo" in linha:
                formatado.write(linha)
        criterio.close()
        formatado.close()
arquivos = os.listdir(PATH_TXT)

#formatar novo txt
for file in arquivos:
    with open(PATH_TXT + file, 'r') as arquivo, open(PATH_TXT + 'temp', 'w') as temp:
        for linha in arquivo.readlines():
            if '"' in linha:
                first_index = linha.index('"')
                last_index = linha.index(' nulo=')
                nova_linha = linha[first_index+1:last_index-1]
                nova_linha = nova_linha.replace('"', ',')
                temp.write(nova_linha)
                temp.write("\n")
        temp.close()
        arquivo.close()
    os.replace(PATH_TXT + 'temp', PATH_TXT + file)

arquivos = os.listdir(PATH_TXT)

#enviar para xlsx
wb = Workbook()
sheet = wb.active

for arq in arquivos:
    with open(PATH_TXT + arq, 'r') as data:
        reader = csv.reader(data, delimiter=',')
        for row in reader:
            sheet.append(row)
            
wb.save(PATH_XLSX + 'Variaveis.xlsx')
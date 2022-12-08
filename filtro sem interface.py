import pandas as pd

arq = open("arquivorequerido.csv",encoding="utf-16")

linha_de_cabecalho = True
primeira_linha = True
dadosIniciais = []
cont = 0
for i in arq:
    
    if (primeira_linha):
        primeira_linha = False
        linha_de_cabecalho = False
        continue #não deixa continuar, volta para o "for". Importantíssimo.

    if (linha_de_cabecalho):
        #cabecalho = linha_aux.split("\t") #é o cabeçalho
        linha_de_cabecalho = False
        continue #não continuar (importante). Volta para o "for", pois é um cabeçalho.
    
    if (i=="\n"):
        cont+=1
        #Armazena apenas os dados das linhas relevantes no array
        if(cont == 1 or cont == 5):
            dadosIniciais.extend(linha_aux.split("\t"))

        linha_de_cabecalho = True

    linha_aux = i

#Remove Strings vazias do array
dadosIniciais = list(filter(None, dadosIniciais))

dados = []

CUSTOMERID = dadosIniciais[1]
NOMECOMPLETO = dadosIniciais[2] + " " + dadosIniciais[3]

#Armazena os dados ordenados em um novo array
dados = [CUSTOMERID, NOMECOMPLETO]
for x in range(4, 6):
    dados.extend(dadosIniciais[x].split("\t"))

for y in range(14, 48):
    dados.extend(dadosIniciais[y].split("\t"))

#Importa a planilha onde os dados obtidos serão gravados
entrada = pd.read_excel("principal.xlsx")

#Grava os dados na próxima linha limpa
entrada.loc[len(entrada)] = dados

#Salva o arquivo da planilha com os novos dados obtidos
saida = pd.ExcelWriter("principal.xlsx")
entrada.to_excel(saida,'Sheet1', index=False)
saida.save()
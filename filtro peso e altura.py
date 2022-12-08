#Importa a biblioteca
import pandas as pd

#Importa a planilha de medidas
tabela = pd.read_csv("arquivorequerido.csv", sep=';', encoding="ISO-8859-1")

#Armazena os dados em variáveis de acordo com coluna x linha
siape = (tabela['CUSTOMERID'][0])
pnome = (tabela['FIRSTNAME'][0])
unome = (tabela['LASTNAME'][0])
datanasc = (tabela['BIRTHDATE'][0])
sexo = (tabela['SEX'][0])

#Concatena o nome
nomecompleto = pnome+ ' ' +unome

#Localiza o índice de FATM e armazena na variável
indice = int(tabela.loc[tabela['FIRSTNAME']=='FATM'].index[0])

#Decrementa a variável de 2 unidades para ter acesso ao última altura gravada  
indice = indice - 2

#Armazena os dados em variáveis de acordo com coluna x linha
altura = (tabela['FIRSTNAME'][indice])
peso = (tabela['LASTNAME'][indice])

#Armazena os dados em um array
dados = [siape,nomecompleto,datanasc,sexo,altura,peso]

#Importa a planilha onde os dados obtidos serão gravados
entrada = pd.read_excel("principal.xlsx")

#Grava os dados na próxima linha limpa
entrada.loc[len(entrada)] = dados

#Salva o arquivo da planilha com os novos dados obtidos
saida = pd.ExcelWriter("principal.xlsx")
entrada.to_excel(saida,'Sheet1', index=False)
saida.save()

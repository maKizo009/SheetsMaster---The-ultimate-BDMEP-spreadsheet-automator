import pandas as pd
import numpy as np
import csv
import os

perguntaDados = input("Adicione o nome da planilha que você quer usar. Lembrando que, para o programa funcionar, tanto a planilha quanto o programa precisam estar na pasta 'Downloads': ")

caminho_para_downloads = os.path.expanduser("~/Downloads")
caminho_completo_para_arquivo = os.path.join(caminho_para_downloads, perguntaDados)


#Formatando a planilha e delimitando os dados usando o ";"
with open(caminho_completo_para_arquivo, "r") as f_entrada:
    leitor = csv.reader(f_entrada, delimiter=";")

    #Pegando o nome da estação
    primeira_linha = f_entrada.readline().rstrip()
    print(primeira_linha)

    #Tranformando os dados da planilha em um Dataframe
    leitor = pd.DataFrame(leitor)

#Ecluindo as partes inúteis da planilha, localizada nas primeiras 11 linhas
    leitor = leitor.iloc[11:]

#Renomeando e reordenando as colunas da forma que eu quero, substituindo os dados nulos por um número aleatório
    leitor = leitor.rename(columns={0: "Data", 1: "Temperatura Máxima", 2: "Temperatura Mínima", 3: "Precipitação"})
    leitor = leitor.reindex(columns=["Data", "Temperatura Mínima", "Temperatura Máxima"])

#Tranformando os dados de string para float
    leitor["Temperatura Máxima"] = pd.to_numeric(leitor["Temperatura Máxima"], errors="coerce")#.astype(float)
    leitor["Temperatura Mínima"] = pd.to_numeric(leitor["Temperatura Mínima"], errors='coerce')#.astype(float)

    
#Criando um pequeno resumo da estação
    linha = leitor["Temperatura Máxima"].idxmin()
    colunas = ["Data", "Temperatura Mínima", "Temperatura Máxima"]
    resultado = leitor.loc[linha, colunas]

    linha2 = leitor["Temperatura Mínima"].idxmin()
    colunas2 = ["Data", "Temperatura Mínima", "Temperatura Máxima"]
    resultado2 = leitor.loc[linha2, colunas2]

    linha3 = leitor["Temperatura Máxima"].idxmax()
    colunas3 = ["Data", "Temperatura Mínima", "Temperatura Máxima"]
    resultado3 = leitor.loc[linha3, colunas3]

    linha4 = leitor["Temperatura Mínima"].idxmax()
    colunas4 = ["Data", "Temperatura Mínima", "Temperatura Máxima"]
    resultado4 = leitor.loc[linha4, colunas4]

    #Imprimindo os resultados na tela
    dados1 = ("A estação de {}, registrou sua menor máxima em {}, com {} graus, e a maior em {} com {}.".format(primeira_linha, resultado["Data"], resultado["Temperatura Máxima"], resultado3["Data"], resultado3["Temperatura Máxima"]))
    dados2 = ("Nas mínimas, a menor foi {} graus em {}, e a maior foi {}, em {}.".format(resultado2["Temperatura Mínima"], resultado2["Data"], resultado4["Temperatura Mínima"], resultado4["Data"]))

    print(dados1, dados2)
    leitor = leitor.fillna("-")

    leitor = leitor.applymap(lambda x: str(x).replace('.', ','))
leitor.to_csv(primeira_linha[6:]+"_BDMEP.csv")

    
    



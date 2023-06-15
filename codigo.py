import pandas as pd
import numpy as np
import csv
with open("dadosManaus.csv", "r") as f_entrada:
    leitor = csv.reader(f_entrada, delimiter=";")

    leitor = pd.DataFrame(leitor)

    leitor = leitor.iloc[11:]

    leitor = leitor.rename(columns={0: "Data", 1: "Temperatura Máxima", 2: "Temperatura Mínima", 3: "Precipitação"})
    leitor = leitor.reindex(columns=["Data", "Temperatura Mínima", "Temperatura Máxima"]).replace("null", "70")

    leitor["Temperatura Mínima"] = leitor["Temperatura Mínima"].astype(float)
    leitor["Temperatura Máxima"] = leitor["Temperatura Máxima"].astype(float)

    leitor["Data"] = pd.to_datetime(leitor["Data"])

    pivotagem = leitor.pivot_table(index=leitor, columns=["Temperatura Mínima", "Temperatura Máxima"])

    pivotagem.to_csv("agrFoisera.csv")



    


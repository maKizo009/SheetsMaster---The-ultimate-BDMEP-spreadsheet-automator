import pandas as pd
import csv
import os
from openpyxl import load_workbook
from shutil import copyfile
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from functools import partial


class SheetsMasters:
    def __init__(self) -> None:
        self.leitor = None
        self.caminhoPlanilhaDestino = None
        self.book = None
        self.janelaUsuario = tk.Tk()
        self.larguraJanela = 500
        self.alturaJanela = 300
        self.janelaUsuario.geometry(f"{self.larguraJanela}x{self.alturaJanela}")

        self.labelArquivo =tk.Label(self.janelaUsuario, text="")
        self.labelArquivo.pack()
        self.labelTexto = tk.Label(self.janelaUsuario, text="")
        #self.mensagemErro = messagebox.showerror("Nenhum arquivo selecionado", "Por favor, selecione um arquivo com formato de planilha .csv, ou descompacte a pasta .zip que contém o arquivo")

        self.botaoPlanilhaDestino = tk.Button(self.janelaUsuario, text="Selecionar Planilha BDMEP", command=self.encontrandoArquivo)
        self.botaoPlanilhaDestino.pack()
        self.botaoPlanilhaDestino = tk.Button(self.janelaUsuario, text="Selecione a planilha de destino dos dados", command=self.planilhaDestino)
        self.botaoPlanilhaDestino.pack()
        self.labelCaminhoPlanilha = tk.Label(self.janelaUsuario, text="")
        self.labelCaminhoPlanilha.pack()
        self.botaoNaoPlanilha = tk.Button(self.janelaUsuario, text="Não tenho planilha", command=partial(self.naoTenhoPlanilha, self.leitor))
        self.botaoNaoPlanilha.pack()
        self.janelaUsuario.mainloop()


    def encontrandoArquivo(self):
        while True:
            self.botaoPlanilhaDestino = filedialog.askopenfilename(initialdir=os.path.expanduser("~/Downloads"), title="Selecione o CSV nessa pasta", filetypes=(("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")))
            if self.botaoPlanilhaDestino.endswith(".zip"):
                self.mensagemErro = ("Formato não aceito", "Selecione uma planilha com formato .csv. Caso tenha baixado o arquivo .zip, descompacte a pasta e tente novamente.")
            else:
                break
        self.processandoArquivo(self.botaoPlanilhaDestino)


    def processandoArquivo(self, botaoPlanilhaDestino):
        #if  botaoPlanilhaDestino:
    
            self.labelArquivo = ("Arquivo selecionado: ", botaoPlanilhaDestino)

            with open(botaoPlanilhaDestino, "r") as f_entrada:
                leitor = csv.reader(f_entrada, delimiter=";")
                primeira_linha = f_entrada.readline().rstrip()
                nomeDaEstacao = tk.Label(self.janelaUsuario, text=primeira_linha)
                nomeDaEstacao.pack()

                self.leitor = pd.DataFrame(leitor)
                self.leitor = self.leitor.iloc[11:]
                self.leitor = self.leitor.rename(columns={0: "Data", 1: "Max", 2: "Min", 3: "Precipitação"})
                self.leitor = self.leitor.reindex(columns=["Data", "Min", "Max"])

                self.leitor["Max"] = pd.to_numeric(self.leitor["Max"], errors="coerce")
                self.leitor["Min"] = pd.to_numeric(self.leitor["Min"], errors="coerce")
                
                linha = self.leitor["Max"].idxmin()
                colunas = ["Data", "Min", "Max"]
                resultado = self.leitor.loc[linha, colunas]

                linha2 = self.leitor["Min"].idxmin()
                colunas2 = ["Data", "Min", "Max"]
                resultado2 = self.leitor.loc[linha2, colunas2]

                linha3 = self.leitor["Max"].idxmax()
                colunas3 = ["Data", "Min", "Max"]
                resultado3 = self.leitor.loc[linha3, colunas3]
                
                linha4 = self.leitor["Min"].idxmax()
                colunas4 = ["Data", "Min", "Max"]
                resultado4 = self.leitor.loc[linha4, colunas4]
                
                dados1 = ("A estação de {}, registrou sua menor máxima em {}, com {} graus, e a maior em {} com {}.".format(primeira_linha, resultado["Data"], resultado["Max"], resultado3["Data"], resultado3["Max"]))
                dados2 = ("Nas mínimas, a menor foi {} graus em {}, e a maior foi {}, em {}.".format(resultado2["Min"], resultado2["Data"], resultado4["Min"], resultado4["Data"]))

                self.leitor["Data"] = pd.to_datetime(self.leitor["Data"])

                self.leitor["Dia"] = self.leitor["Data"].dt.day
                self.leitor["Mes"] = self.leitor["Data"].dt.month
                self.leitor["Ano"] = self.leitor["Data"].dt.year

                self.labelTexto = (dados1)
                self.labelTexto = (dados2)
            """  else:
            self.mensagemErro """


    def processandoDados(self, primeira_linha, leitor, book):
         for ano in self.leitor["Ano"].dropna().unique():
                dadosAno = self.leitor[self.leitor["Ano"]== ano]
                organizado = pd.DataFrame()

                for mes in range(1, 13):
                    dadosMes = dadosAno[dadosAno["Mes"] == mes]
                    minimas = dadosMes[["Dia", "Min"]].set_index("Dia")
                    maximas = dadosMes[["Dia", "Max"]].set_index("Dia")
                    minimas.columns = [f"Minima{mes}"]
                    maximas.columns = [f"Máxima{mes}"]
                    organizado.index = organizado.index.astype(str)
                    inicio = pd.to_datetime(f'{int(ano)}-{int(mes):02d}-01')
                    fim = inicio + pd.offsets.MonthEnd(0)
                    dias = pd.date_range(inicio, fim)
                    dias = dias.day
                    minimas = minimas.reindex(dias)
                    maximas = maximas.reindex(dias)
                    organizado = organizado.reset_index(drop=True)
                    minimas = minimas.reset_index(drop=True)
                    maximas = maximas.reset_index(drop=True)
                    organizado = pd.concat([organizado, minimas, maximas], axis=1)

                sheet_name = str(int(ano))
                if self.caminhoPlanilhaDestino and self.book:
                    self.book.save(self.caminhoPlanilhaDestino)
                else:
                    ws = self.book.create_sheet(sheet_name)

                for col, col_data in enumerate(organizado.values.T, start=2):
                    for row, value in enumerate(col_data, start=3):
                        if pd.isna(value):
                            value = "-"
                        ws.cell(row=row, column=col, value=value)

                self.book.save(self.planilhaDestino)

                
    def planilhaDestino(self):
            
        self.caminhoPlanilhaDestino = filedialog.askopenfilename(initialdir=os.path.expanduser("/~Downloads"), filetypes=(("Arquivos XLSX", "*.xlsx"), ("Todos os arquivos", "*.*")))

        if self.caminhoPlanilhaDestino:
            self.labelCaminhoPlanilha.config(text=f"Planilha encontrada: {self.caminhoPlanilhaDestino}")
            self.book = load_workbook(self.caminhoPlanilhaDestino)
        else:
            pass


    def naoTenhoPlanilha(self, leitor):
        caminhoOriginal = "Modelo para normais climatológicas 1991-2020.xlsx"
        novoCaminho = "planilhaTeste.xlsx"
        #(primeira_linha[6:]+"_BDMEP.xlsx")

        copyfile(caminhoOriginal, novoCaminho)
        self.book = load_workbook("Modelo para normais climatológicas 1991-2020.xlsx")
        self.processandoDados(self.leitor, novoCaminho, self.book)
SheetsMasters()

import csv
from openpyxl import load_workbook

planilhas = load_workbook(filename='meus_livros.xlsx')
planilha = planilhas.active
cabecalho = next(planilha.values)

serie = []
ordem = []
titulo = []
autor = []
paginas = []
situacao = []

indice_serie = cabecalho.index('SÉRIE')
serie = [linha[indice_serie] for linha in planilha.values]

indice_ordem = cabecalho.index('ORDEM')
ordem = [linha[indice_ordem] for linha in planilha.values]

indice_titulo = cabecalho.index('TÍTULO')
titulo = [linha[indice_titulo] for linha in planilha.values]

indice_autor = cabecalho.index('AUTOR')
autor = [linha[indice_autor] for linha in planilha.values]

indice_paginas = cabecalho.index('PÁGINAS')
paginas = [linha[indice_paginas] for linha in planilha.values]

indice_situacao = cabecalho.index('SITUAÇÃO')
situacao = [linha[indice_situacao] for linha in planilha.values]

with open(file='./meus_livros.csv', mode='w', encoding='utf8') as arquivo:
    escritor_csv = csv.writer(arquivo, delimiter=',')
    for valores in zip(serie, ordem, titulo, autor, paginas, situacao):
        escritor_csv.writerow(valores)
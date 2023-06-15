import csv
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
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

livros_df = pd.read_csv('meus_livros.csv')
livros_df.head()

contagem = livros_df['SITUAÇÃO'].value_counts()
contagem

situacao_df = pd.DataFrame({'Situação': contagem.index, 'Frequência': contagem.values})
situacao_df


plt.pie(contagem, autopct='%1.1f%%')
plt.axis('equal')

legenda = ['{}: {}'.format(regiao, valor) for regiao, valor in zip(contagem.index, contagem.values)]
plt.legend(legenda, loc='center left', bbox_to_anchor=(1, 0.5))

plt.title('Situação dos Livros na Estante')


with sns.axes_style('whitegrid'):
  grafico = sns.barplot(data=situacao_df, x="Situação", y="Frequência", errorbar=None, palette='magma')
  grafico.set(title='Situação dos Livros na Estante', xlabel='Situação', ylabel='Frequência');
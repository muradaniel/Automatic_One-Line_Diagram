#----------------------------------------------------------------------------------------------------------------------
#------------------------------------------------ BIBLIOTECAS ---------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------
import xlwings as xw
from openpyxl.utils import get_column_letter
import numpy as np
import pandas as pd
import tkinter as tk
from datetime import datetime
import math
import cmath
import ctypes
from tkinter.scrolledtext import ScrolledText
import networkx as nx # Grafos
import matplotlib.pyplot as plt
# from tkinter import Tk, filedialog # Biblioteca de diálogo de seleção de arquivo

#----------------------------------------------------------------------------------------------------------------------
#-------------------------------------- LEITURA DOS DADOS E TABELAS  --------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

# CAIXA DE SELEÇÃO DE ARQUIVO EXCEL
#caminho = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", ".xlsx;.xls;*.xlsm")])
# caminho = r"G:\Meu Drive\01 - Faculdade\Análise de Sistemas de Potência\ASP 1\Trabalho\Entrada de Dados.xlsm"
# caminho = r"G:\Meu Drive\01 - Faculdade\Análise de Sistemas de Potência\ASP 1\Trabalho\Video_Exemplo.xlsm"
caminho = input("Entre com seu arquivo Excel\n")
wb = xw.Book(fr"{caminho}")


# LEITURA DAS BARRAS
sheet = wb.sheets['Barra'] # Abre a planilha "Barra"
ultima_linha = sheet.range('B6').end('down').row # pega a última linha com valores
valores_barra = sheet.range(f'B6:B{ultima_linha}').value # Pega todos os números das barras
valores_barra = [int(x) for x in valores_barra] # Transforma os números das barras em números inteiros


# LEITURA DOS ELEMENTOS DO SISTEMA
Maquina = pd.read_excel(fr"{caminho}", sheet_name = "Maquina", header = 4)
Carga = pd.read_excel(fr"{caminho}", sheet_name = "Carga", header= 4)
Transformador = pd.read_excel(fr"{caminho}", sheet_name = "Transformador", header = 4)
Linha = pd.read_excel(fr"{caminho}", sheet_name= "Linha", header = 4)
Barra = pd.read_excel(fr"{caminho}", sheet_name= "Barra", header = 4)


# LEITURA DAS CONFIGURAÇÕES DA SIMULAÇÃO
sheet = wb.sheets['Configuração']
nome_caso = sheet.range('P7').value
tipo_de_curto = sheet.range('P9').value
barra_curto = int(sheet.range('P20').value)
impedancia_falta = sheet.range('P22').value
potencia_base = sheet.range('P24').value
caminho_do_arquivo_excel = wb.sheets['Dados Complementares'].range('K2').value


#----------------------------------------------------------------------------------------------------------------------
#--------------------------------------------- VARIAVEIS  -------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

# Criação das tabelas utilizadas nos cálculos e relatórios
tabela_tensoes_barras = pd.DataFrame(columns=["Barra", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Utilizado nos cálculos
tabela_tensoes_barras_corrigidos = pd.DataFrame(columns=["Barra", "Fase A", "Fase B", "Fase C", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Resultado Final

tabela_correntes_contribuicao = pd.DataFrame(columns=["Trecho", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Utilizado nos cálculos
tabela_correntes_contribuicao_corrigidos = pd.DataFrame(columns=["Trecho", "Fase A", "Fase B", "Fase C", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Resultado Final

tabela_correntes_injetadas = pd.DataFrame(columns=["Elemento", "Barra", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Utilizado nos cálculos
tabela_correntes_injetadas_corrigidos = pd.DataFrame(columns=["Elemento", "Barra", "Fase A", "Fase B", "Fase C", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Resultado Final

barras_isoladas = []

# DADOS EXTRAS/COMPLEMENTARES
data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')


# MATRIZES DE TRANSFORMAÇÕES
alfa = cmath.rect(1, math.radians(120)) # 1<120°

T012abc =  np.array([
    [1,      1,               1  ],
    [1,    alfa**2,          alfa],
    [1,     alfa,          alfa**2]
])

Tabc012 = (1/3) * np.array([
    [1,       1,            1   ],
    [1,     alfa,        alfa**2],
    [1,     alfa**2,       alfa ]
])


# Essa função tem como objetivo converter valores complexos na forma cartesiana (R + JX) para forma polar (M <teta°)
def cartesiano_polar(z):
    if abs(z) != 0:
        Z = f"{abs(z):.4f} ∠ {math.degrees(cmath.phase(z)):.2f}°"
    else:
        Z = f"{abs(0):.4f} ∠ {math.degrees(cmath.phase(0)):.2f}°"
    return Z


# Essa função tem como objetivo limpar linhas inteiramente vazias de um dataframe
def Remove_linhas_Vazias(tabela):
    tabela = tabela.dropna(how='all')
    return tabela


# Essa função tem como objetivo exportar de forma formatada as matrizes para o excel
def exporta_matriz(planilha, celula, matriz):
    sheet = wb.sheets[planilha]
    sheet.api.Unprotect(Password='DiaboPeruano')
    sheet.range("B4:XFD1048576").clear()
    Ybarra_arredondado = matriz.applymap(lambda z: complex(round(z.real, 4), round(z.imag, 4)))
    Ybarra_str = Ybarra_arredondado.applymap(lambda z: f'{z.real:.4f} + {z.imag:.4f}j')
    sheet.range(celula).options(index=True, header=True).value = Ybarra_str
    sheet.range(celula).expand().columns.autofit() # Ajustando a largura da coluna
    # sheet.range(celula).expand().rows.autofit() # Ajustando a altura da linha
    sheet.range(celula).expand().columns[0].api.Font.Bold = True # Inserindo o negrito na coluna, número das  barras
    sheet.range(celula).expand('right').api.Font.Bold = True # Inserindo o negrito na linha, número das  barras
    # for i in range(7, 13):  # De xlEdgeLeft (7) a xlInsideHorizontal (12)
    #     sheet.range(celula).expand().api.Borders(i).LineStyle = 1  # xlContinuous
    #     sheet.range(celula).expand().api.Borders(i).Weight = 2     # xlThin
    sheet.range("B4:XFD1048576").api.HorizontalAlignment = -4108  # xlCenter

    sheet.api.Protect(Password='DiaboPeruano')


# Essa função tem como objetivo corrigir o angulo, para informar o sentido real da corrente
def correcao_angulo(Z):
    if math.degrees(cmath.phase(Z)) > 0:
        Z = Z / (1j**2)
    return Z


# Essa função tem como objetivo remover barras isoladas da matriz Ybarra e inverte-las
def remocao_barra_isolada(Ybarra):
    if not isinstance(Ybarra, pd.DataFrame): # Garante que seja um DataFrame e obtém os rótulos de índice
        raise TypeError("A matriz Ybarra deve ser um pandas.DataFrame")

    linhas_nulas = np.all(Ybarra.values == 0, axis=1)
    colunas_nulas = np.all(Ybarra.values == 0, axis=0)
    indices_remover = Ybarra.index[linhas_nulas & colunas_nulas]
    barras_isoladas = Ybarra.columns[colunas_nulas].tolist()
    

    # Se não há barras isoladas, retorna a matriz original
    if len(indices_remover) == 0:
        return Ybarra, barras_isoladas

    # Remove linhas e colunas com os índices detectados
    Ybarra_reduzida = Ybarra.drop(index=indices_remover, columns=indices_remover)

    return Ybarra_reduzida, barras_isoladas

#----------------------------------------------------------------------------------------------------------------------
#-------------------------------------- TRATAMENTO DE DADOS -----------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

Maquina = Remove_linhas_Vazias(Maquina)
Carga = Remove_linhas_Vazias(Carga)
Transformador = Remove_linhas_Vazias(Transformador)
Linha = Remove_linhas_Vazias(Linha)

# CRIANDO A COLUNA DA IMPEDÂNCIA COMPLEXA Z = R + jX, PARA CADA SEQUÊNCIA
Maquina["Z1 (pu) Base"] = Maquina["R1 (pu) Base"] + 1j *  Maquina["X1 (pu) Base"]
Maquina["Z0 (pu) Base"] = Maquina["R0 (pu) Base"] + 1j * (Maquina["X0 (pu) Base"] + 3 * Maquina["XN (pu) Base"])

Transformador["Z1 (pu) Base"] = 0 + 1j * Transformador["X1 (pu) Base"]
Transformador["Z0 (pu) Base"] = 0 + 1j * (Transformador["X0 (pu) Base"] + 3 * Transformador["XN P (pu) Base"] + 3 * Transformador["XN S (pu) Base"])

Linha["Z1 (pu) Base"] = Linha["R1 (pu) Base"] + 1j * Linha["X1 (pu) Base"]
Linha["Z0 (pu) Base"] = Linha["R0 (pu) Base"] + 1j * Linha["X0 (pu) Base"]

Carga["Z1 (pu) Base"] = Carga["R1 (pu) Base"] + 1j * Carga["X1 (pu) Base"]
Carga["Z0 (pu) Base"] = Carga["R1 (pu) Base"] + 1j * Carga["X1 (pu) Base"]

#----------------------------------------------------------------------------------------------------------------------
#-------------------------------- CÁLCULOS DAS MATRIZES Y BARRA & Z BARRA  --------------------------------------------
#----------------------------------------------------------------------------------------------------------------------


# # ---------------------------- CÁLCULOS DA YBARRA (+), (-) ----------------------------

Ybarra12_ = pd.DataFrame(np.zeros((len(valores_barra), len(valores_barra))), index = valores_barra, columns = valores_barra)

for index, row in Maquina.iterrows():
    Ybarra12_.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] =  (1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra Conectada"]), int(row["Barra Conectada"])]
for index, row in Carga.iterrows():
    Ybarra12_.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] =  (1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra Conectada"]), int(row["Barra Conectada"])]
for index, row in Linha.iterrows():
    Ybarra12_.at[int(row["Barra de"]), int(row["Barra de"])] =  (1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra de"]), int(row["Barra de"])]
    Ybarra12_.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra para"]), int(row["Barra para"])]
    Ybarra12_.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra de"]), int(row["Barra para"])]
    Ybarra12_.at[int(row["Barra para"]), int(row["Barra de"])] =  (-1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra para"]), int(row["Barra de"])]
for index, row in Transformador.iterrows():
    Ybarra12_.at[int(row["Barra de"]), int(row["Barra de"])] = (1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra de"]), row["Barra de"]]
    Ybarra12_.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra para"]), row["Barra para"]]
    Ybarra12_.at[int(row["Barra para"]), int(row["Barra de"])] = (-1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra para"]), row["Barra de"]]
    Ybarra12_.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z1 (pu) Base"]) + Ybarra12_.loc[int(row["Barra de"]), row["Barra para"]]


# # ---------------------------- CÁLCULOS DA ZBARRA (+), (-) --------------------------

Zbarra12 = pd.DataFrame(np.linalg.inv(Ybarra12_.values), index = valores_barra, columns = valores_barra) # Invertendo a matriz Ybarra12_


# # ---------------------------- CÁLCULOS DA YBARRA (ZERO) ----------------------------

# As linhas que o somatório complex(0, 0) podemos apagar, deixei mais a carater ilustrativo
Ybarra0_ = pd.DataFrame(np.zeros((len(valores_barra), len(valores_barra))), index = valores_barra, columns = valores_barra)
for index, row in Maquina.iterrows():
    if row["Tipo de Conexão"] == "YT":
        Ybarra0_.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[row["Barra Conectada"], row["Barra Conectada"]]
    elif row["Tipo de Conexão"] == "Y":
        Ybarra0_.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = complex(0, 0) + Ybarra0_.loc[row["Barra Conectada"], row["Barra Conectada"]]
    elif row["Tipo de Conexão"] == "D":
        Ybarra0_.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = complex(0, 0) + Ybarra0_.loc[row["Barra Conectada"], row["Barra Conectada"]]
for index, row in Carga.iterrows():
    if row["Tipo de Conexão"] == "YT":
        Ybarra0_.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[row["Barra Conectada"], row["Barra Conectada"]]
    else:
        Ybarra0_.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = complex(0, 0) + Ybarra0_.loc[row["Barra Conectada"], row["Barra Conectada"]]
for index, row in Linha.iterrows():
    if True: # Apenas para deixar o código formatado XD
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra de"])] =  (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra de"]), row["Barra de"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra para"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra de"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra de"])] =  (-1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra para"]), row["Barra de"]]
for index, row in Transformador.iterrows():
    if row["Tipo de Conexão"] == "YT-YT":
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra de"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra de"]), row["Barra de"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra para"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra de"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra de"])] = (-1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra para"]), row["Barra de"]]
    elif (row["Tipo de Conexão"] == "YT-D"):
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra de"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra de"]), row["Barra de"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra para"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra de"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra para"]), row["Barra de"]]
    elif (row["Tipo de Conexão"] == "D-YT"):
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra de"]), row["Barra de"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0_.loc[int(row["Barra para"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra de"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra para"]), row["Barra de"]]
    else:
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra de"]), row["Barra de"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra para"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra de"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra de"]), row["Barra para"]]
        Ybarra0_.at[int(row["Barra para"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0_.loc[int(row["Barra para"]), row["Barra de"]]


# # ---------------------------- CÁLCULOS DA ZBARRA (ZERO) ----------------------------

Ybarra0_, barras_isoladas = remocao_barra_isolada(Ybarra0_)

Zbarra0 = pd.DataFrame(np.linalg.inv(Ybarra0_.values), index=Ybarra0_.index, columns=Ybarra0_.columns)


#----------------------------------------------------------------------------------------------------------------------
#-------------------------------- CÁLCULOS DAS CORRENTES DE CURTO CIRCUITO --------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

if (tipo_de_curto is not None) and (potencia_base is not None) and (potencia_base != 0):
    if tipo_de_curto == "Trifásico":
        Ifa0 = 0
        Ifa1 = (1 / (Zbarra12.loc[barra_curto, barra_curto] + 3 * impedancia_falta)) # Adicionar a Zf
        Ifa2 = 0
        IF = Ifa1
        If3f = T012abc @ np.array([[Ifa0],[Ifa1],[Ifa2]])

    elif tipo_de_curto == "Monofásico":
        Ifa0 = (1 / (2 * Zbarra12.loc[barra_curto, barra_curto] + Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)) # Adicionar a Zf
        Ifa1 = (1 / (2 * Zbarra12.loc[barra_curto, barra_curto] + Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)) # Adicionar a Zf
        Ifa2 = (1 / (2 * Zbarra12.loc[barra_curto, barra_curto] + Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)) # Adicionar a Zf
        IF = 3 * Ifa0
        If3f = T012abc @ np.array([[Ifa0], [Ifa1], [Ifa2]])

    elif tipo_de_curto == "Bifásico":
        Ifa0 = 0
        Ifa1 = (1 / 2 * Zbarra12.loc[barra_curto, barra_curto] + impedancia_falta) # Adicionar a Zf
        Ifa2 = -(1 / 2 * Zbarra12.loc[barra_curto, barra_curto] + impedancia_falta) # Adicionar a Zf
        If3f = T012abc @ np.array([[Ifa0], [Ifa1], [Ifa2]])
        IF = If3f[1][0] # Fase B

    elif tipo_de_curto == "Bifásico - Terra":
        Ifa1 = 1 / (
                Zbarra12.loc[barra_curto, barra_curto] +
                (
                        Zbarra12.loc[barra_curto, barra_curto] *
                        (Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)
                ) / (
                        Zbarra12.loc[barra_curto, barra_curto] +
                        Zbarra0.loc[barra_curto, barra_curto] +
                        3 * impedancia_falta
                )
        )
        Ifa0 = - Ifa1 * (Zbarra12.loc[barra_curto, barra_curto] / (Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta + Zbarra12.loc[barra_curto, barra_curto]))
        Ifa2 = - Ifa1 * ((Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta) / (Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta + Zbarra12.loc[barra_curto, barra_curto]))
        IF = 3 * Ifa0
        If3f = T012abc @ np.array([[Ifa0], [Ifa1], [Ifa2]])
else:
    if tipo_de_curto is None:
        ctypes.windll.user32.MessageBoxW(0, "É necessário selecionar previamente um tipo de curto circuito!", "Atenção!", 0)

    elif barra_curto is None:
        ctypes.windll.user32.MessageBoxW(0, "É necessário selecionar previamente um barramento de curto circuito!", "Atenção!", 0)

    elif potencia_base is None:
        ctypes.windll.user32.MessageBoxW(0, "É necessário selecionar previamente uma potência base em MVA.", "Atenção!", 0)

    else:
        ctypes.windll.user32.MessageBoxW(0, "Há algum erro não identificado na sua simulação, verifique todas suas informações de entrada.", "Atenção", 0)

    If = ""

Vb = Barra.loc[Barra["Número"] == barra_curto, "Tensão (kV)"].values[0]
IF *= (potencia_base / (np.sqrt(3) * Vb))
Ifa = If3f[0][0] * (potencia_base / (np.sqrt(3) * Vb))
Ifb = If3f[1][0] * (potencia_base / (np.sqrt(3) * Vb))
Ifc = If3f[2][0] * (potencia_base / (np.sqrt(3) * Vb))

tabela_curto_circuito = pd.DataFrame({
    "IF" : [cartesiano_polar(IF)],
    "Fase A" : [cartesiano_polar(Ifa)],
    "Fase B" : [cartesiano_polar(Ifb)],
    "Fase C" : [cartesiano_polar(Ifc)],
    "Seq. (0)" : [cartesiano_polar(Ifa0 * (potencia_base / (np.sqrt(3) * Vb)))],
    "Seq. (+)" : [cartesiano_polar(Ifa1 * (potencia_base / (np.sqrt(3) * Vb)))],
    "Seq. (-)" : [cartesiano_polar(Ifa2 * (potencia_base / (np.sqrt(3) * Vb)))]
})


#----------------------------------------------------------------------------------------------------------------------
#------------------------------------------ CORREÇÃO DE DEFASAGEM -----------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

G = nx.DiGraph()
for tabela in [Linha, Transformador]: # Unicos elementos que conectam os barramentos
    for index, row in tabela.iterrows():
        if tabela is Transformador: # Se tabela for transformador
            tipo = row["Tipo de Conexão"]
            if tipo in ["YT-D", "Y-D"]:
                G.add_edge(row['Barra de'], row['Barra para'], weight=-30)
                G.add_edge(row['Barra para'], row['Barra de'], weight=30)
            elif tipo in ["D-YT", "D-Y"]:
                G.add_edge(row['Barra de'], row['Barra para'], weight=30)
                G.add_edge(row['Barra para'], row['Barra de'], weight=-30)
            else:
                G.add_edge(row['Barra de'], row['Barra para'], weight=0)
                G.add_edge(row['Barra para'], row['Barra de'], weight=0)
        else: # Se a tabela for linha
            G.add_edge(row['Barra de'], row['Barra para'], weight=0)
            G.add_edge(row['Barra para'], row['Barra de'], weight=0)


#----------------------------------------------------------------------------------------------------------------------
#------------------------------------- CÁLCULOS DAS TENSÕES NAS BARRAS ------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

tabela_tensoes_barras["Barra"] = tabela_tensoes_barras["Barra"].astype(int)
tabela_tensoes_barras_corrigidos["Barra"] = tabela_tensoes_barras_corrigidos["Barra"].astype(int)

for barra in valores_barra: # Aqui ainda contempla a barra isolada, tratarei utilizando as barras que estão presentes na Zbarra0, porem perderemos informações sobre a barra isolada"
    if barra not in barras_isoladas:
        Vka0 = 0 - Zbarra0.loc[barra_curto, barra] * Ifa0
    else:
        Vka0 = 0
    Vka1 = 1 - Zbarra12.loc[barra_curto, barra] * Ifa1
    Vka2 = 0 - Zbarra12.loc[barra_curto, barra] * Ifa2

    tabela_tensoes_barras.loc[len(tabela_tensoes_barras)] = [int(barra.real), Vka0, Vka1, Vka2] # Inserção dos valores na tabela não corrigidos

    correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=barra, weight='weight') # +30
    correcao_negativa = nx.shortest_path_length(G, source=barra, target=barra_curto, weight='weight') # -30

    Vka0 = Vka0
    Vka1 = Vka1 * cmath.rect(1, math.radians(correcao_positiva))
    Vka2 = Vka2 * cmath.rect(1, math.radians(correcao_negativa))

    Vk3f = T012abc @ np.array([
        [Vka0],
        [Vka1],
        [Vka2]
    ])

    Vka = Vk3f[0][0]
    Vkb = Vk3f[1][0]
    Vkc = Vk3f[2][0]

    tabela_tensoes_barras_corrigidos.loc[len(tabela_tensoes_barras_corrigidos)] = [int(barra.real),
                                                                                   cartesiano_polar(Vka), 
                                                                                   cartesiano_polar(Vkb), 
                                                                                   cartesiano_polar(Vkc), 
                                                                                   cartesiano_polar(Vka0), 
                                                                                   cartesiano_polar(Vka1), 
                                                                                   cartesiano_polar(Vka2)] # Inserção dos valores na tabela

tabela_tensoes_barras["Barra"] = tabela_tensoes_barras["Barra"].apply(lambda x: int(x.real)) # Por algum motivo, o número da barra estava vindo complexa (???)


#----------------------------------------------------------------------------------------------------------------------
#------------------------------ CÁLCULOS DAS CORRENTES DE CONTRIBUIÇÕES -----------------------------------------------
#----------------------------------------------------------------------------------------------------------------------
 
for x in [Linha, Maquina, Carga]:
    for index, row in x.iterrows():
        if x is Linha:
            numero_barra_de = row["Barra de"]
            numero_barra_para = row["Barra para"]
            nome_elemento = row["Nome"]
            Vb = Barra.loc[Barra["Número"] == numero_barra_de, "Tensão (kV)"].values[0]
            Ib = potencia_base / (np.sqrt(3) * Vb)

            delta_vk0 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_de, "Seq. (0)"].values[0] -
                         tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_para, "Seq. (0)"].values[0])
            delta_vk1 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_de, "Seq. (+)"].values[0] -
                         tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_para, "Seq. (+)"].values[0])
            delta_vk2 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_de, "Seq. (-)"].values[0] -
                         tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_para, "Seq. (-)"].values[0])

            correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_de, weight='weight')
            correcao_negativa = nx.shortest_path_length(G, source=numero_barra_de, target=barra_curto, weight='weight')

        elif x is Maquina:  # ou Carga, se necessário depois
            barra_conectada = row["Barra Conectada"]
            nome_elemento = row["Nome"]
            Vb = Barra.loc[Barra["Número"] == barra_conectada, "Tensão (kV)"].values[0]
            Ib = potencia_base / (np.sqrt(3) * Vb)
            conexao = row["Tipo de Conexão"]

            if conexao != "YT":
                delta_vk0 = 0
            else:
                delta_vk0 = (0 - tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (0)"].values[0])
            delta_vk1 = (1 - tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (+)"].values[0])
            delta_vk2 = (0 - tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (-)"].values[0])

            correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=barra_conectada, weight='weight')
            correcao_negativa = nx.shortest_path_length(G, source=barra_conectada, target=barra_curto, weight='weight')
        
        else: # Carga Shunt
            barra_conectada = row["Barra Conectada"]
            nome_elemento = row["Nome"]
            Vb = Barra.loc[Barra["Número"] == barra_conectada, "Tensão (kV)"].values[0]
            Ib = potencia_base / (np.sqrt(3) * Vb)
            conexao = row["Tipo de Conexão"]

            if conexao != "YT":
                delta_vk0 = 0
            else:
                delta_vk0 = (0 - tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (0)"].values[0])
            delta_vk1 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (+)"].values[0])
            delta_vk2 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (-)"].values[0])

            correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=barra_conectada, weight='weight')
            correcao_negativa = nx.shortest_path_length(G, source=barra_conectada, target=barra_curto, weight='weight')

        Ic0 = delta_vk0 / row["Z0 (pu) Base"]
        Ic1 = delta_vk1 / row["Z1 (pu) Base"]
        Ic2 = delta_vk2 / row["Z1 (pu) Base"]

        tabela_correntes_contribuicao.loc[len(tabela_correntes_contribuicao)] = [nome_elemento, Ic0, Ic1, Ic2]

        # Se o gerador ou shunt, for diferente de YT ic0 = 0
        Ic0 = Ic0
        Ic1 = Ic1 * cmath.rect(1, math.radians(correcao_positiva))
        Ic2 = Ic2 * cmath.rect(1, math.radians(correcao_negativa))

        Ic3f = T012abc @ np.array([[Ic0], [Ic1], [Ic2]])
        Ica, Icb, Icc = Ic3f[0][0], Ic3f[1][0], Ic3f[2][0]

        tabela_correntes_contribuicao_corrigidos.loc[len(tabela_correntes_contribuicao_corrigidos)] = [
            nome_elemento,
            cartesiano_polar(correcao_angulo(Ica * Ib)),
            cartesiano_polar(correcao_angulo(Icb * Ib)),
            cartesiano_polar(correcao_angulo(Icc * Ib)),
            cartesiano_polar(correcao_angulo(Ic0 * Ib)),
            cartesiano_polar(correcao_angulo(Ic1 * Ib)),
            cartesiano_polar(correcao_angulo(Ic2 * Ib))
        ]

# Corrente passando pelo transformador, caso especial
for index, row in Transformador.iterrows():
    nome_elemento = row["Nome"]
    numero_barra_de = row["Barra de"]
    numero_barra_para = row["Barra para"]
    Vb1 = Barra.loc[Barra["Número"] == numero_barra_de, "Tensão (kV)"].values[0] # Tensão da Barra de
    Vb2 = Barra.loc[Barra["Número"] == numero_barra_para, "Tensão (kV)"].values[0] # Tensão da Barra para
    Ib1 = potencia_base / (np.sqrt(3) * Vb1) # Corrente Base barra de
    Ib2 = potencia_base / (np.sqrt(3) * Vb2) # Corrente Base barra para
    tipo_conexao = row["Tipo de Conexão"]

    delta_vk0 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_de, "Seq. (0)"].values[0] -
                 tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_para, "Seq. (0)"].values[0])
    delta_vk1 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_de, "Seq. (+)"].values[0] -
                 tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_para, "Seq. (+)"].values[0])
    delta_vk2 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_de, "Seq. (-)"].values[0] -
                 tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == numero_barra_para, "Seq. (-)"].values[0])

    if tipo_conexao == "YT-YT":
        Ic0 = delta_vk0 / row["Z0 (pu) Base"]
    else:
        Ic0 = 0
    Ic1 = delta_vk1 / row["Z1 (pu) Base"]
    Ic2 = delta_vk2 / row["Z1 (pu) Base"]


    if (abs(Ic1) * Ib1) > (abs(Ic1) * Ib2): # Lado de Baixa

        correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_de, weight='weight')
        correcao_negativa = nx.shortest_path_length(G, source=numero_barra_de, target=barra_curto, weight='weight')

        Ic0_corr = Ic0
        Ic1_corr = Ic1 * cmath.rect(1, math.radians(correcao_positiva))
        Ic2_corr = Ic2 * cmath.rect(1, math.radians(correcao_negativa))

        Ic3f = T012abc @ np.array([[Ic0_corr], [Ic1_corr], [Ic2_corr]])
        Ica, Icb, Icc = Ic3f[0][0], Ic3f[1][0], Ic3f[2][0]

        tabela_correntes_contribuicao_corrigidos.loc[len(tabela_correntes_contribuicao_corrigidos)] = [
            nome_elemento + " (BT)",
            cartesiano_polar(correcao_angulo(Ica * Ib1)),
            cartesiano_polar(correcao_angulo(Icb * Ib1)),
            cartesiano_polar(correcao_angulo(Icc * Ib1)),
            cartesiano_polar(correcao_angulo(Ic0_corr * Ib1)),
            cartesiano_polar(correcao_angulo(Ic1_corr * Ib1)),
            cartesiano_polar(correcao_angulo(Ic2_corr * Ib1))
        ]

        correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_para, weight='weight')
        correcao_negativa = nx.shortest_path_length(G, source=numero_barra_para, target=barra_curto, weight='weight')

        Ic0_corr = Ic0
        Ic1_corr = Ic1 * cmath.rect(1, math.radians(correcao_positiva))
        Ic2_corr = Ic2 * cmath.rect(1, math.radians(correcao_negativa))

        Ic3f = T012abc @ np.array([[Ic0_corr], [Ic1_corr], [Ic2_corr]])
        Ica, Icb, Icc = Ic3f[0][0], Ic3f[1][0], Ic3f[2][0]

        tabela_correntes_contribuicao_corrigidos.loc[len(tabela_correntes_contribuicao_corrigidos)] = [
            nome_elemento + " (AT)",
            cartesiano_polar(correcao_angulo(Ica * Ib2)),
            cartesiano_polar(correcao_angulo(Icb * Ib2)),
            cartesiano_polar(correcao_angulo(Icc * Ib2)),
            cartesiano_polar(correcao_angulo(Ic0_corr * Ib2)),
            cartesiano_polar(correcao_angulo(Ic1_corr * Ib2)),
            cartesiano_polar(correcao_angulo(Ic2_corr * Ib2))
        ]
    else:
        correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_para, weight='weight')
        correcao_negativa = nx.shortest_path_length(G, source=numero_barra_para, target=barra_curto, weight='weight')

        Ic0_corr = Ic0
        Ic1_corr = Ic1 * cmath.rect(1, math.radians(correcao_positiva))
        Ic2_corr = Ic2 * cmath.rect(1, math.radians(correcao_negativa))

        Ic3f = T012abc @ np.array([[Ic0_corr], [Ic1_corr], [Ic2_corr]])
        Ica, Icb, Icc = Ic3f[0][0], Ic3f[1][0], Ic3f[2][0]

        tabela_correntes_contribuicao_corrigidos.loc[len(tabela_correntes_contribuicao_corrigidos)] = [
            nome_elemento + " (BT)",
            cartesiano_polar(correcao_angulo(Ica * Ib2)),
            cartesiano_polar(correcao_angulo(Icb * Ib2)),
            cartesiano_polar(correcao_angulo(Icc * Ib2)),
            cartesiano_polar(correcao_angulo(Ic0_corr * Ib2)),
            cartesiano_polar(correcao_angulo(Ic1_corr * Ib2)),
            cartesiano_polar(correcao_angulo(Ic2_corr * Ib2))
        ]

        correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_de, weight='weight')
        correcao_negativa = nx.shortest_path_length(G, source=numero_barra_de, target=barra_curto, weight='weight')

        Ic0_corr = Ic0
        Ic1_corr = Ic1 * cmath.rect(1, math.radians(correcao_positiva))
        Ic2_corr = Ic2 * cmath.rect(1, math.radians(correcao_negativa))

        Ic3f = T012abc @ np.array([[Ic0_corr], [Ic1_corr], [Ic2_corr]])
        Ica, Icb, Icc = Ic3f[0][0], Ic3f[1][0], Ic3f[2][0]

        tabela_correntes_contribuicao_corrigidos.loc[len(tabela_correntes_contribuicao_corrigidos)] = [
            nome_elemento + " (AT)",
            cartesiano_polar(correcao_angulo(Ica * Ib1)),
            cartesiano_polar(correcao_angulo(Icb * Ib1)),
            cartesiano_polar(correcao_angulo(Icc * Ib1)),
            cartesiano_polar(correcao_angulo(Ic0_corr * Ib1)),
            cartesiano_polar(correcao_angulo(Ic1_corr * Ib1)),
            cartesiano_polar(correcao_angulo(Ic2_corr * Ib1))
        ]
    

#----------------------------------------------------------------------------------------------------------------------
#----------------------------------- CÁLCULOS DAS CORRENTES INJETADAS -------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

for tabela in [Maquina]: # Carga entra?
    for index, row in tabela.iterrows():
        barra_conectada = row["Barra Conectada"]
        nome_elemento = row["Nome"]
        Vb = Barra.loc[Barra["Número"] == barra_conectada, "Tensão (kV)"].values[0]
        Ib = potencia_base / (np.sqrt(3) * Vb)

        delta_vk0 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (0)"].values[0])
        delta_vk1 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (+)"].values[0])
        delta_vk2 = (tabela_tensoes_barras.loc[tabela_tensoes_barras["Barra"] == barra_conectada, "Seq. (-)"].values[0])

        correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=barra_conectada, weight='weight')
        correcao_negativa = nx.shortest_path_length(G, source=barra_conectada, target=barra_curto, weight='weight')

        if tabela is Maquina:
            Ii0 = ((0 - delta_vk0) / row["Z0 (pu) Base"])
            Ii1 = ((1 - delta_vk1) / row["Z1 (pu) Base"])
            Ii2 = ((0 - delta_vk2) / row["Z1 (pu) Base"])
        
        else: # Se for carga
            Ii0 = (delta_vk0 / row["Z0 (pu) Base"])
            Ii1 = (delta_vk1 / row["Z1 (pu) Base"])
            Ii2 = (delta_vk2 / row["Z1 (pu) Base"])

        tabela_correntes_injetadas.loc[len(tabela_correntes_injetadas)] = [nome_elemento, barra_conectada, Ii0, Ii1, Ii2]  # Inserção dos valores na tabela sem correção

        Ii0 = Ii0
        Ii1 = Ii1 * cmath.rect(1, math.radians(correcao_positiva))
        Ii2 = Ii2 * cmath.rect(1, math.radians(correcao_negativa))

        Ii3f = T012abc @ np.array([
            [Ii0],
            [Ii1],
            [Ii2]
        ])

        Iia = Ii3f[0][0]
        Iib = Ii3f[1][0]
        Iic = Ii3f[2][0]

        tabela_correntes_injetadas_corrigidos.loc[len(tabela_correntes_injetadas_corrigidos)] = [nome_elemento,
                                                                                                 barra_conectada,
                                                                                                 Iia * Ib,
                                                                                                 Iib * Ib,
                                                                                                 Iic * Ib,
                                                                                                 Ii0 * Ib,
                                                                                                 Ii1 * Ib,
                                                                                                 Ii2 * Ib]  # Inserção dos valores na tabela com correção
tabela_correntes_injetadas_corrigidos = tabela_correntes_injetadas_corrigidos.groupby("Barra")[["Fase A", "Fase B", "Fase C", "Seq. (0)", "Seq. (+)", "Seq. (-)"]].sum().reset_index() # Agrupar as correntes ijetadas do barramento
tabela_correntes_injetadas_corrigidos.loc[:, ~tabela_correntes_injetadas_corrigidos.columns.isin(["Barra"])] = tabela_correntes_injetadas_corrigidos.loc[:, ~tabela_correntes_injetadas_corrigidos.columns.isin(["Barra"])].applymap(cartesiano_polar)


#----------------------------------------------------------------------------------------------------------------------
#----------------------------------- EXPORTAR RELATÓRIO PARA O EXCEL --------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

# # ---------------------------- MATRIZES ----------------------------
exporta_matriz("Y_Barra_1_2", "B4", Ybarra12_)
exporta_matriz("Z_Barra_1_2", "B4", Zbarra12)
exporta_matriz("Y_Barra_0", "B4", Ybarra0_)
exporta_matriz("Z_Barra_0", "B4", Zbarra0)


# # --------------------- CABEÇALHO RELATÓRIO ------------------------
sheet = wb.sheets["Relatório"]
sheet.api.Unprotect(Password='DiaboPeruano')
sheet.range("A6:XFD1048576").clear()
sheet.range("H7:XFD1048576").api.HorizontalAlignment = -4108  # xlCenter

sheet.range('H6').value = f"Data: {data}" # Insere a data e hora da simulação

sheet.range('H7:N7').merge()
sheet.range('H7').value = f"Corrente de Falta (kA) - {tipo_de_curto} - Barra {barra_curto}" # Insere o tipo de curto circuito
sheet.range('H7').color = (217, 217, 217) # pintando de Cinza
sheet.range('H7').font.bold = True
sheet.range('H7').horizontal_alignment = 'center'


sheet.range('H8:N8').value = ["IF\"", "Fase A", "Fase B", "Fase C", "Sequência (0)", "Sequência (+)", "Sequência (-)"]
sheet.range('H8:N8').color = (20, 20, 20)
sheet.range('H8').font.color = (255, 255, 255)
sheet.range('I8:N8').font.color = (142, 217, 115)



sheet.range("H9").options(index=False, header=False).value = tabela_curto_circuito





sheet.range('H10:N10').merge()
sheet.range('H10').value = f"Tensões nas Barras (pu)" # Insere o tipo de curto circuito
sheet.range('H10').color = (217, 217, 217) # pintando de Cinza
sheet.range('H10').font.bold = True
sheet.range('H10').horizontal_alignment = 'center'

sheet.range('H11:N11').value = ["Barra", "Fase A", "Fase B", "Fase C", "Sequência (0)", "Sequência (+)", "Sequência (-)"]
sheet.range('H11:N11').color = (20, 20, 20) # Pintando de Preto
sheet.range('H11').font.color = (255, 255, 255) # Pintando de Branco
sheet.range('I11:N11').font.color = (166, 200, 236) # Pintando de Azul
sheet.range('H12').options(index=False, header=False).value = tabela_tensoes_barras_corrigidos




ultima_linha = sheet.range('H7').end('down').row + 1 # pega a última linha com valores
sheet.range(f'H{ultima_linha}:N{ultima_linha}').merge()
sheet.range(f'H{ultima_linha}').value = f"Correntes de Contribuição (kA)" # Insere o título da tabela
sheet.range(f'H{ultima_linha}').color = (217, 217, 217) # pintando de Cinza
sheet.range(f'H{ultima_linha}').font.bold = True
sheet.range(f'H{ultima_linha}').horizontal_alignment = 'center'

ultima_linha += 1
sheet.range(f'H{ultima_linha}:N{ultima_linha}').value = ["Elemento", "Fase A", "Fase B", "Fase C", "Sequência (0)", "Sequência (+)", "Sequência (-)"]
sheet.range(f'H{ultima_linha}:N{ultima_linha}').color = (20, 20, 20) # Pintando de Preto
sheet.range(f'H{ultima_linha}').font.color = (255, 255, 255) # Pintando de Branco
sheet.range(f'I{ultima_linha}:N{ultima_linha}').font.color = (255, 113, 79) # Pintando de Vermelho
ultima_linha += 1
sheet.range(f'H{ultima_linha}').options(index=False, header=False).value = tabela_correntes_contribuicao_corrigidos
sheet.range("H6").expand().columns[0].api.Font.Bold = True


ultima_linha = sheet.range('H7').end('down').row + 1 # pega a última linha com valores
sheet.range(f'H{ultima_linha}:N{ultima_linha}').merge()
sheet.range(f'H{ultima_linha}').value = f"Correntes Injetadas nos Barramentos (kA)" # Insere o título da tabela
sheet.range(f'H{ultima_linha}').color = (217, 217, 217) # pintando de Cinza
sheet.range(f'H{ultima_linha}').font.bold = True
sheet.range(f'H{ultima_linha}').horizontal_alignment = 'center'

ultima_linha += 1
sheet.range(f'H{ultima_linha}:N{ultima_linha}').value = ["Barra", "Fase A", "Fase B", "Fase C", "Sequência (0)", "Sequência (+)", "Sequência (-)"]
sheet.range(f'H{ultima_linha}:N{ultima_linha}').color = (20, 20, 20) # Pintando de Preto
sheet.range(f'H{ultima_linha}').font.color = (255, 255, 255) # Pintando de Branco
sheet.range(f'I{ultima_linha}:N{ultima_linha}').font.color = (255, 255, 0) # Pintando de Amarelo
ultima_linha += 1
sheet.range(f'H{ultima_linha}').options(index=False, header=False).value = tabela_correntes_injetadas_corrigidos
sheet.range("H6").expand().columns[0].api.Font.Bold = True



sheet.api.Protect(Password='DiaboPeruano')



# sheet.range('A12').value = [[v] for v in valores_barra] # Insere valores dos barramentos


# # --------------------- CORRENTES DE CURTO ------------------------

# [...]


# # ---------------------------- TENSÕES ----------------------------

# tabela_tensoes_barras_corrigidos.rename(columns={'Barra': 'Barra / Trecho'}, inplace=True)
# tabela_correntes_contribuicao_corrigidos.rename(columns={'Trecho': 'Barra / Trecho'}, inplace=True)

# relatorio = pd.concat([tabela_tensoes_barras_corrigidos, tabela_correntes_contribuicao_corrigidos], ignore_index=True)
# relatorio.fillna(0, inplace=True)

# sheet = wb.sheets["Relatório"]
# sheet.range("A11").options(index=False, header=False).value = relatorio

# # ------------------ CORRENTES DE CONTRIBUIÇÃO --------------------

# [...]


# --------------------------------APAGAR NO FINAL -----------------------------------------------

# Criar janela Tkinter para conferência das Matrizes Ybarra e Zbarra (Apagar no final)
# root = tk.Tk()
# root.title("Matriz Ybarra (completar ou inversa)")
# current_font_size = [13]
# text_widget = ScrolledText(root, width=100, height=25, font=("Courier New", current_font_size[0]))
# text_widget.pack(expand=True, fill="both")
# # Inserir o DataFrame formatado no widget
# text_widget.insert(tk.END, f"YBarra¹² = \n{Ybarra12_.to_string()}\n")
# text_widget.insert(tk.END, f"\nZBarra¹² = \n{Zbarra12.to_string()}\n")
# text_widget.insert(tk.END, f"\nYBarra° = \n{Ybarra0_.to_string()}\n")
# text_widget.insert(tk.END, f"\nZBarra° = \n{Zbarra0.to_string()}\n")
# root.mainloop()



# Desenha os nós e arestas
# matriz = nx.to_numpy_array(G, weight='weight')
# print(matriz)
# correcao = nx.shortest_path_length(G, source=barra_curto, target=4, weight='weight')
# print(correcao)
# # Desenha os rótulos de peso nas arestas
# # nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_color='red')
# #
# # plt.title("Grafo com Pesos (positivos e negativos)")
# # plt.show()
# # length = nx.shortest_path_length(G, source=1, target=4, weight='weight')
# # print("Caminho 1 → 4:", length)
# # print(G)
# # matriz = nx.to_numpy_array(G, weight='weight')  # inclui os pesos
# print(tabela_correntes_contribuicao_corrigidos)
# print(tabela_tensoes_barras_corrigidos)
# print(relatorio)
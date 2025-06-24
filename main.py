# Importação das Bibliotecas --------------------
import xlwings as xw
from openpyxl.utils import get_column_letter
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
# from tkinter import Tk, filedialog # Biblioteca de diálogo de seleção de arquivo

# NESSE TRECHO SÃO APENAS COMANDOS BÁSICOS DE LEITURA E TRATAMENTO DE TABELAS NÃO RELEVANTE PARA CÁLCULOS DE ANÁLISE DE CURTO

# Abre o arquivo do Excel -----------------------
#arquivo = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", ".xlsx;.xls;*.xlsm")])
#wb = xw.Book(fr"arquivo")

wb = xw.Book(r"G:\Meu Drive\01 - Faculdade\Análise de Sistemas de Potência\ASP 1\Trabalho\Entrada de Dados.xlsm")


# Leitura das barras ----------------------------
sheet = wb.sheets['Barra'] # Abre a planilha "Barra"
ultima_linha = sheet.range('B6').end('down').row # pega a última linha com valores
valores_barra = sheet.range(f'B6:B{ultima_linha}').value # Pega todos os números das barras
valores_barra = [int(x) for x in valores_barra] # Transforma os números das barras em números inteiros


# Leitura das Tabelas - header = 4 pois as tabelas começam na linha 5
Maquina = pd.read_excel(r"G:\Meu Drive\01 - Faculdade\Análise de Sistemas de Potência\ASP 1\Trabalho\Entrada de Dados.xlsm", sheet_name="Maquina", header=4)
Carga = pd.read_excel(r"G:\Meu Drive\01 - Faculdade\Análise de Sistemas de Potência\ASP 1\Trabalho\Entrada de Dados.xlsm", sheet_name="Carga", header=4)
Transformador = pd.read_excel(r"G:\Meu Drive\01 - Faculdade\Análise de Sistemas de Potência\ASP 1\Trabalho\Entrada de Dados.xlsm", sheet_name="Transformador", header=4)
Linha = pd.read_excel(r"G:\Meu Drive\01 - Faculdade\Análise de Sistemas de Potência\ASP 1\Trabalho\Entrada de Dados.xlsm", sheet_name="Linha", header=4)

tabelas = [Maquina, Transformador, Linha]

# Ao ler a planilha Excel, vinha no dataframe algumas linhas vazias, esta função tem como objetivo remove-las
def Remove_linhas_Vazias(tabela):
    tabela = tabela.dropna(how='all')
    return tabela

Maquina = Remove_linhas_Vazias(Maquina)
Carga = Remove_linhas_Vazias(Carga)
Transformador = Remove_linhas_Vazias(Transformador)
Linha = Remove_linhas_Vazias(Linha)


Maquina["Z1 (pu) Base"] = Maquina["R1 (pu) Base"] + 1j *  Maquina["X1 (pu) Base"]
Maquina["Z0 (pu) Base"] = Maquina["R0 (pu) Base"] + 1j * (Maquina["X0 (pu) Base"] + 3 * Maquina["XN (pu) Base"])

Transformador["Z1 (pu) Base"] = 0 + 1j * Transformador["X1 (pu) Base"]
Transformador["Z0 (pu) Base"] = 0 + 1j * (Transformador["X0 (pu) Base"] + 3 * Transformador["XN P (pu) Base"] + 3 * Transformador["XN P (pu) Base"])

Linha["Z1 (pu) Base"] = Linha["R1 (pu) Base"] + 1j * Linha["X1 (pu) Base"]
Linha["Z0 (pu) Base"] = Linha["R0 (pu) Base"] + 1j * Linha["X0 (pu) Base"]


# def trafo_seq_zero(barra_com_ligacao):
#      Trafo = Transformador[(Transformador["Barra de"] == barra_com_ligacao) | (Transformador["Barra para"] == barra_com_ligacao)]
#      return Trafo
#
# def trafo_seq_zero_cont1(nome_trafo, Trafo):
#     configuracao = Trafo[Trafo["Nome"] == nome_trafo]
#     configuracao = configuracao[["Barra de","Barra para","Tipo de Conexão", "Z0 (pu) Base"]]
#     return configuracao
#
# def logica_do_trafo(configuracao, barra_com_ligacao):
#     config = configuracao.iloc[0]  # <- extrai a linha como Series
#
#     if config["Tipo de Conexão"] == "YT-D":
#         if config["Barra de"] == barra_com_ligacao:
#             admitancia = 1 / config["Z0 (pu) Base"]
#         else:
#             admitancia = complex(0, 0)
#
#     elif config["Tipo de Conexão"] == "D-YT":
#         if config["Barra para"] == barra_com_ligacao:
#             admitancia = 1 / config["Z0 (pu) Base"]
#         else:
#             admitancia = complex(0, 0)
#
#     elif config["Tipo de Conexão"] == "YT-Y":
#         admitancia = complex(0, 0)
#
#     elif config["Tipo de Conexão"] == "Y-D":
#         admitancia = complex(0, 0)
#
#     elif config["Tipo de Conexão"] == "YT-YT":
#         admitancia = 1 / config["Z0 (pu) Base"]
#
#     elif config["Tipo de Conexão"] == "D-D":
#         admitancia = complex(0, 0)
#
#     return admitancia

def admitancia_seq_zero_total(barra_com_ligacao):
    total = 0
    # Filtra os transformadores ligados à barra
    Trafo = Transformador[
        (Transformador["Barra de"] == barra_com_ligacao) |
        (Transformador["Barra para"] == barra_com_ligacao)
    ]

    print(Trafo)

    for nome_trafo in Trafo["Nome"].tolist():
        # Filtra os dados relevantes de configuração para o trafo
        configuracao = Trafo[Trafo["Nome"] == nome_trafo]
        configuracao = configuracao[["Barra de", "Barra para", "Tipo de Conexão", "Z0 (pu) Base"]]
        print(configuracao)

        # Extrai a linha única como Series
        config = configuracao.iloc[0]

        # Aplica a lógica de cálculo da admitância de sequência zero
        tipo = config["Tipo de Conexão"]
        if tipo == "YT-D":
            admitancia = 1 / config["Z0 (pu) Base"] if config["Barra de"] == barra_com_ligacao else complex(0, 0)
        elif tipo == "D-YT":
            admitancia = 1 / config["Z0 (pu) Base"] if config["Barra para"] == barra_com_ligacao else complex(0, 0)
        elif tipo == "YT-YT":
            admitancia = 1 / config["Z0 (pu) Base"]
        else:
            admitancia = complex(0, 0)

        print(admitancia)
        total += admitancia

    print("admitancia final", total)
    return total




#----------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------


# ---------------------------- CÁLCULOS DA YBARRA (+), (-) ----------------------------

Ybarra12 = pd.DataFrame(np.zeros((len(valores_barra), len(valores_barra))), index=valores_barra, columns=valores_barra)

for index, row in Ybarra12.iterrows():
    for col in Ybarra12.columns:
        if index == col: # Admitância Própia
            Ybarra12.at[index, col] = - (
                    (1 / Maquina.loc[Maquina["Barra Conectada"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Transformador.loc[Transformador["Barra de"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Transformador.loc[Transformador["Barra para"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra de"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra para"] == index, "Z1 (pu) Base"]).sum()
            )
        else: # Admitância Mútua
            Ybarra12.at[index, col] =  (
                    (1 / Transformador.loc[(Transformador["Barra de"] == index) & (Transformador["Barra para"] == col), "Z1 (pu) Base"]).sum() +
                    (1 / Transformador.loc[(Transformador["Barra de"] == col) & (Transformador["Barra para"] == index), "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == index) & (Linha["Barra para"] == col), "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == col) & (Linha["Barra para"] == index), "Z1 (pu) Base"]).sum()
            )


# ---------------------------- CÁLCULOS DA ZBARRA (+), (-) ----------------------------

Zbarra12 = pd.DataFrame(np.linalg.inv(Ybarra12.values), index=valores_barra, columns=valores_barra) # Invertendo a matriz Ybarra12
Zbarra12 =  Zbarra12



# ---------------------------- CÁLCULOS DA YBARRA (ZERO) ----------------------------

Ybarra0 = pd.DataFrame(np.zeros((len(valores_barra), len(valores_barra))), index=valores_barra, columns=valores_barra)

for index, row in Ybarra0.iterrows():
    for col in Ybarra0.columns:
        if index == col: # Admitância Própia
            Ybarra0.at[index, col] = - (
                    (1 / Maquina.loc[(Maquina["Barra Conectada"] == index) & (Maquina["Tipo de Conexão"] == "YT"), "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra de"] == index, "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra para"] == index, "Z0 (pu) Base"]).sum() +
                    (admitancia_seq_zero_total(index))
            )
        else: # Admitância Mútua
            Ybarra0.at[index, col] =  (
                    (1 / Transformador.loc[(Transformador["Barra de"] == index) & (Transformador["Barra para"] == col), "Z0 (pu) Base"]).sum() +
                    (1 / Transformador.loc[(Transformador["Barra de"] == col) & (Transformador["Barra para"] == index), "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == index) & (Linha["Barra para"] == col), "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == col) & (Linha["Barra para"] == index), "Z0 (pu) Base"]).sum()
            )

# ---------------------------- CÁLCULOS DA ZBARRA (Zero) ----------------------------

# CÁLCULOS DAS CORRENTES DE CURTO

# [.......]

# CÁLCULO DAS TENSÕES NAS BARRAS

# [.......]

# CÁLUCLO DAS CORRENTES DAS MÁQUINAS

# [.......]

# CÁLUCLO DAS CORRENTES DE CONTRIBUIÇÃO (EXTRA)

# [.......]

# EXPORTAR OS DADOS PARA O EXCEL


# total = 0
# barrinha = 2
# Trafo = trafo_seq_zero(barrinha)
# print(Trafo)
# for x in  Trafo["Nome"].tolist():
#     configuracao = trafo_seq_zero_cont1(x, Trafo)
#     print(configuracao)
#     admitancia = logica_do_trafo(configuracao, barrinha)
#     print(admitancia)
#     total = admitancia + total
#
# print("admitancia final", total)

# print(admitancia_seq_zero_total(6))



# Criar janela Tkinter para conferência das Matrizes Ybarra e Zbarra (Apagar no final)
root = tk.Tk()
root.title("Matriz Ybarra (completar ou inversa)")
current_font_size = [12]
text_widget = ScrolledText(root, width=100, height=25, font=("Courier New", current_font_size[0]))
text_widget.pack(expand=True, fill="both")
# Inserir o DataFrame formatado no widget
text_widget.insert(tk.END, f"YBarra¹² = \n{Ybarra12.to_string()}\n")
text_widget.insert(tk.END, f"\nZBarra¹² = \n{Zbarra12.to_string()}\n")
text_widget.insert(tk.END, f"\nYBarra° = \n{Ybarra0.to_string()}\n")
#text_widget.insert(tk.END, f"\nZBarra° = \n{Zbarra0.to_string()}\n")
root.mainloop()









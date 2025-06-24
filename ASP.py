# Importação das Bibliotecas --------------------
import xlwings as xw
from openpyxl.utils import get_column_letter
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
# from tkinter import Tk, filedialog # Biblioteca de diálogo de seleção de arquivo

# NESSE TRECHO SÃO APENAS COMANDOS BÁSICOS DE LEITURA E TRATAMENTO DE TABELAS NÃO RELEVANTE PARA CÁLCULOS DE ANÁLISE DE CURTO

# Abre o arquivo do Excel -----------------------
arquivo = filedialog.askopenfilename(title="Selecione o caso", filetypes=[("Arquivos Excel", ".xlsx;.xls;*.xlsm")])
wb = xw.Book(arquivo)


# Leitura das barras ----------------------------
sheet = wb.sheets['Barra'] # Abre a planilha "Barra"
ultima_linha = sheet.range('B6').end('down').row # pega a última linha com valores
valores_barra = sheet.range(f'B6:B{ultima_linha}').value # Pega todos os números das barras
valores_barra = [int(x) for x in valores_barra] # Transforma os números das barras em números inteiros


# Leitura das Tabelas - header = 4 pois as tabelas começam na linha 5
Maquina = pd.read_excel(arquivo, sheet_name="Maquina", header=4)
Carga = pd.read_excel(arquivo, sheet_name="Carga", header=4)
Transformador = pd.read_excel(arquivo, sheet_name="Transformador", header=4)
Linha = pd.read_excel(arquivo, sheet_name="Linha", header=4)

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

def admitancia_seq_zero_total(barra_com_ligacao):
    total = 0
    # Filtra os transformadores ligados à barra
    Trafo = Transformador[
        (Transformador["Barra de"] == barra_com_ligacao) |
        (Transformador["Barra para"] == barra_com_ligacao)
    ]

    for nome_trafo in Trafo["Nome"].tolist():
        # Filtra os dados relevantes de configuração para o trafo
        configuracao = Trafo[Trafo["Nome"] == nome_trafo]
        configuracao = configuracao[["Barra de", "Barra para", "Tipo de Conexão", "Z0 (pu) Base"]]

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

        total += admitancia
    return total



#----------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------


# ---------------------------- CÁLCULOS DA YBARRA (+), (-) ----------------------------

Ybarra12 = pd.DataFrame(np.zeros((len(valores_barra), len(valores_barra))), index=valores_barra, columns=valores_barra)

for index, row in Ybarra12.iterrows():
    for col in Ybarra12.columns:
        if index == col: # Admitância Própia
            Ybarra12.at[index, col] = (
                    (1 / Maquina.loc[Maquina["Barra Conectada"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Transformador.loc[Transformador["Barra de"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Transformador.loc[Transformador["Barra para"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra de"] == index, "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra para"] == index, "Z1 (pu) Base"]).sum()
            )
        else: # Admitância Mútua
            Ybarra12.at[index, col] = - (
                    (1 / Transformador.loc[(Transformador["Barra de"] == index) & (Transformador["Barra para"] == col), "Z1 (pu) Base"]).sum() +
                    (1 / Transformador.loc[(Transformador["Barra de"] == col) & (Transformador["Barra para"] == index), "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == index) & (Linha["Barra para"] == col), "Z1 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == col) & (Linha["Barra para"] == index), "Z1 (pu) Base"]).sum()
            )


# ---------------------------- CÁLCULOS DA ZBARRA (+), (-) ----------------------------

Zbarra12 = pd.DataFrame(np.linalg.inv(Ybarra12.values), index=valores_barra, columns=valores_barra) # Invertendo a matriz Ybarra12


# ---------------------------- CÁLCULOS DA YBARRA (ZERO) ----------------------------

Ybarra0 = pd.DataFrame(np.zeros((len(valores_barra), len(valores_barra))), index=valores_barra, columns=valores_barra)

for index, row in Ybarra0.iterrows():
    for col in Ybarra0.columns:
        if index == col: # Admitância Própia
            Ybarra0.at[index, col] =  (
                    (1 / Maquina.loc[(Maquina["Barra Conectada"] == index) & (Maquina["Tipo de Conexão"] == "YT"), "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra de"] == index, "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[Linha["Barra para"] == index, "Z0 (pu) Base"]).sum() +
                    (admitancia_seq_zero_total(index))
            )
        else: # Admitância Mútua
            Ybarra0.at[index, col] = - (
                    (1 / Transformador.loc[(Transformador["Barra de"] == index) & (Transformador["Barra para"] == col) & (Transformador["Tipo de Conexão"] == "YT-YT"), "Z0 (pu) Base"]).sum() +
                    (1 / Transformador.loc[(Transformador["Barra de"] == col) & (Transformador["Barra para"] == index) & (Transformador["Tipo de Conexão"] == "YT-YT"), "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == index) & (Linha["Barra para"] == col), "Z0 (pu) Base"]).sum() +
                    (1 / Linha.loc[(Linha["Barra de"] == col) & (Linha["Barra para"] == index), "Z0 (pu) Base"]).sum() 
                    
            )

# ---------------------------- CÁLCULOS DA ZBARRA (Zero) ----------------------------

Zbarra0 = pd.DataFrame(np.linalg.inv(Ybarra0.values), index=valores_barra, columns=valores_barra) # Invertendo a matriz Ybarra12


# CÁLCULOS DAS CORRENTES DE CURTO

# [.......]

# CÁLCULO DAS TENSÕES NAS BARRAS

# [.......]

# CÁLUCLO DAS CORRENTES DAS MÁQUINAS

# [.......]

# CÁLUCLO DAS CORRENTES DE CONTRIBUIÇÃO (EXTRA)

# [.......]

# EXPORTAR OS DADOS PARA O EXCEL


def limpar_zeros_complexos(df, limiar=1e-10):
    df_formatado = df.copy()
    for i in df.index:
        for j in df.columns:
            val = df_formatado.loc[i, j]
            if isinstance(val, complex) and abs(val) < limiar:
                df_formatado.loc[i, j] = 0
            elif isinstance(val, float) and abs(val) < limiar:
                df_formatado.loc[i, j] = 0
    return df_formatado

def exibir_dataframe_em_treeview(frame, df):
    tree = ttk.Treeview(frame, show="headings")

    # Adiciona o índice como primeira coluna
    colunas = ["#"] + list(df.columns)
    tree["columns"] = colunas

    # Cabeçalhos
    tree.heading("#", text="#")
    tree.column("#", anchor="center", width=40)  # Coluna de índice

    for col in df.columns:
        tree.heading(col, text=str(col))
        tree.column(col, anchor="center", width=100)

    # Insere os dados (inclui índice no início de cada linha)
    for idx, row in df.iterrows():
        valores = [str(idx)] + [str(val) for val in row]
        tree.insert("", "end", values=valores)

    # Scrollbar vertical
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")

    tree.pack(expand=True, fill="both")

# --- Criar a Janela
root = tk.Tk()
root.title("Matriz Ybarra & Zbarra")
root.geometry("1820x980")  # Tamanho fixo (opcional)

# Estilo moderno
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview.Heading", font=("Segoe UI", 18, "bold"))
style.configure("Treeview", font=("Courier New", 16), rowheight=50)

# Notebook (Abas)
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

# Lista de matrizes e nomes
matrizes = [
    ("YBarra¹²", Ybarra12),
    ("ZBarra¹²", Zbarra12),
    ("YBarra°", Ybarra0),
    ("ZBarra°", Zbarra0)
]

# Criar uma aba para cada matriz
for nome, df in matrizes:
    frame = tk.Frame(notebook)
    notebook.add(frame, text=nome)
    df_limpo = limpar_zeros_complexos(df.round(3))
    exibir_dataframe_em_treeview(frame, df_limpo)

root.mainloop()









#----------------------------------------------------------------------------------------------------------------------
#------------------------------------------------ BIBLIOTECAS ---------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

from schemdraw.segments import SegmentText
from schemdraw.segments import Segment # Criar seus próprios elementos
import schemdraw.elements as elm  # Possui os elementos elétricos já prontos
from schemdraw.segments import *
from datetime import datetime
import seaborn as sns # Geração de cores automáticas
import networkx as nx # Organizar barramentos
import xlwings as xw
import pandas as pd  # Trabalhar com tabelas e planilhas
import schemdraw  # Realizar os desenhos elétricos
import itertools
import math # cálculos matemáticos
import fitz
import ctypes
import os


#----------------------------------------------------------------------------------------------------------------------
#--------------------------------------- FUNÇÕES & CLASSES ------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

class Tabelas:
    inicio_dados = 0
    caminho_excel = r"Arquivo_Cadastro_Elementos.xlsm"
    # caminho_excel = input("Entre com Caminho:\n")
    def __init__(self):
        self.Barra = pd.read_excel(fr"{self.caminho_excel}", sheet_name="Barra", header=self.inicio_dados)
        self.Maquina = pd.read_excel(fr"{self.caminho_excel}", sheet_name="Maquina", header=self.inicio_dados)
        self.Linha = pd.read_excel(fr"{self.caminho_excel}", sheet_name="Linha", header=self.inicio_dados)
        self.Transformador = pd.read_excel(fr"{self.caminho_excel}", sheet_name="Transformador", header=self.inicio_dados)
        self.Carga = pd.read_excel(fr"{self.caminho_excel}", sheet_name="Carga", header=self.inicio_dados)
        self.Transformador_3E = pd.read_excel(fr"{self.caminho_excel}", sheet_name="Transformador 3 Enrolamentos", header=self.inicio_dados)
tabelas = Tabelas()


# Essa função tem como objetivo resumir as ligações entre todos os componentes com o objetivo de organizar a posição das barras.
def Grafo():
    ligacao_entre_elementos = pd.DataFrame(columns=["Elemento", "Barra conectada"])
    for tabela in [tabelas.Linha, tabelas.Transformador]:
        for index, row in tabela.iterrows():
            ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra de"])]
            ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra para"])]
    for tabela in [tabelas.Maquina, tabelas.Carga]:
        for index, row in tabela.iterrows():
            ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra Conectada"])]
    for index, row in tabelas.Transformador_3E.iterrows():
        ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra P"])]
        ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra S"])]
        ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra T"])]

    # Criando o Grafo:
    G = nx.Graph()
    G.add_nodes_from(tabelas.Barra["Número"].tolist()) # Cria os nós dos Barramentos
    G.add_nodes_from(tabelas.Transformador["Nome"].tolist()) # Cria os nós dos Transformadores
    G.add_nodes_from(tabelas.Linha["Nome"].tolist()) # Cria os nós das Linhas de Transmissões
    G.add_nodes_from(tabelas.Maquina["Nome"].tolist()) # Cria os nós das Máquinas (Motores e Geradores)
    G.add_nodes_from(tabelas.Carga["Nome"].tolist()) # Cria os nós das Cargas sem Rotações
    G.add_nodes_from(tabelas.Transformador_3E["Nome"].tolist()) # Cria os nós dos Transformadores de 3 Enrolamentos

    G.add_edges_from(list(zip(ligacao_entre_elementos["Elemento"].tolist(), ligacao_entre_elementos["Barra conectada"].tolist()))) # Cria as ligações entre os nós (arestas)
    posicao_elementos = nx.kamada_kawai_layout(G, scale=60) # Esolher a tipo de Organização do Grafo, espaçamento entre os nós, e número de iterações
    return G, posicao_elementos


# Essa função tem como objetivo gerar cores para diferentes níveis de tensão
def Gera_Cores_Tensoes(niveis_de_tensoes):  # lista de Tensões
    cores_tensoes = sns.color_palette("tab10", len(niveis_de_tensoes))  # Gera as cores
    dicionario_cores = dict(
        zip(niveis_de_tensoes, cores_tensoes))  # Cria um dicionário com a tensão como chave e a cor como valor
    return dicionario_cores


# Essa função tem como objetivo, corrigir a posição central do transformador
def Ajustes_trafos(posicao_barra_de, posicao_barra_para, posicao_transformador):
        x1, y1 = posicao_barra_de
        x2, y2 = posicao_barra_para
        xm, ym = posicao_transformador
        if (x1 < x2) and (y1 < y2):
            xm = xm - 1.5
            ym = ym - 1.5
        elif (x1 > x2) and (y1 > y2):
            xm = xm + 1.5
            ym = ym + 1.5
        elif (x1 < x2) and (y1 > y2):
            xm = xm - 1.5
            ym = ym + 1.5
        else:
            xm = xm + 1.5
            ym = ym - 1.5
        return xm, ym


# Tratamento de rotação e flipagem dotrasnformador
def tratamento_trafo_3e(posicao_barra_p, posicao_barra_s, posicao_barra_t, posicao_transformador_3e, d):
    melhor_dist = float('inf')
    melhor_teta = 0
    melhor_flip = False
    melhor_reverse = False

    for teta, flip, reverse in itertools.product(range(360), [False, True], [False, True]):
        # Cria o trafo com as transformações apropriadas
        trafo = (
            Trafo_3_enrolamentos()
            .at(posicao_transformador_3e)
            .theta(teta)
        )
        if flip:
            trafo = trafo.flip()
        if reverse:
            trafo = trafo.reverse()

        # Desenho invisível
        trafo = trafo.color("#FFFFFF00").zorder(0)
        d += trafo

        # Calcula distância total entre os terminais e as barras
        ponto_p = trafo.absanchors['p']
        ponto_s = trafo.absanchors['s']
        ponto_t = trafo.absanchors['t']

        dist_atual = (
            math.dist(posicao_barra_p, ponto_p) +
            math.dist(posicao_barra_s, ponto_s) +
            math.dist(posicao_barra_t, ponto_t)
        )

        # Atualiza se for melhor
        if dist_atual < melhor_dist:
            melhor_dist = dist_atual
            melhor_teta = teta
            melhor_flip = flip
            melhor_reverse = reverse

    return melhor_teta, melhor_flip, melhor_reverse


#----------------------------------------------------------------------------------------------------------------------
#------------------------------------------------- VARIAVEIS ------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

G, posicao_elementos = Grafo()

# for key, value in posicao_elementos.items(): # Esse laço de repetição tem como objetivo deixar o desenho retângular largura x altura
#     posicao_elementos[key][0] *= 8
#     posicao_elementos[key][1] *= 6

# Caso o usuário queira observar o grafo
#import matplotlib.pyplot as plt
#nx.draw(G, pos=posicao_elementos, with_labels=True, node_color='skyblue', edge_color='black', node_size=1000, font_size=12)
#plt.title("Meu Grafo")
#plt.show()

dicionario_cores = Gera_Cores_Tensoes(tabelas.Barra["Tensão (kV)"].tolist())
caminho_diagrama = tabelas.caminho_excel.replace("xlsm", "pdf")
wb = xw.Book(fr"{Tabelas().caminho_excel}")
sheet = wb.sheets['Configuração']
barra_curto = int(sheet.range('P20').value)
nome_caso = sheet.range('P7').value
tipo_de_curto = sheet.range('P9').value
impedancia_falta = sheet.range('P22').value
data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
tensao_falta = complex(1, 0)

#----------------------------------------------------------------------------------------------------------------------
#------------------------------------------------- ELEMENTOS ------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

# Criando barramento
class Barramento(elm.Element):
    def __init__(self, length=3, **kwargs):
        super().__init__(**kwargs)
        altura = 1 # Tamanho do barramento
        self.segments.append(Segment([(0, -altura), (0, altura)], lw = 5, capstyle='square')) # Desenha a linha
        self.anchors['centro'] = (0, altura/2) # Define ponto de Conexão


# # Criando trafo de três enrolamentos
class Trafo_3_enrolamentos(elm.Element):
    def __init__(self, cor_primario=None, cor_secundario=None, cor_terciario=None, conexao=None, **kwargs):
        super().__init__(**kwargs)
        raio = 0.6
        braco = 0.5  # comprimento dos braços (ajustável)

        # Bobinas (círculos)
        self.segments.append(SegmentCircle((-raio * 0.7, 0), raio, fill=None, color=cor_primario))   # Esquerda
        self.segments.append(SegmentCircle((raio * 0.7, 0), raio, fill=None, color=cor_secundario))    # Direita
        self.segments.append(SegmentCircle((0, -raio * 1.2), raio, fill=None, color=cor_terciario))   # Inferior

        # Braços com base em 'braco'
        self.segments.append(Segment([(-raio * 0.7 - raio, 0), (-raio * 0.7 - raio - braco, 0)], color=cor_primario))  # Primário
        self.segments.append(Segment([(raio * 0.7 + raio, 0), (raio * 0.7 + raio + braco, 0)], color=cor_secundario))    # Secundário
        self.segments.append(Segment([(0, -raio * 1.2 - raio), (0, -raio * 1.2 - raio - braco)], color=cor_terciario))  # Terciário

        # Pontos de conexão finais (anchors nas pontas dos braços)
        self.anchors['p'] = (-raio * 0.7 - raio - braco, 0)
        self.anchors['s'] = (raio * 0.7 + raio + braco, 0)
        self.anchors['t'] = (0, -raio * 1.2 - raio - braco)
        
        if conexao is not None:
            if "D" in conexao:
                conexao = conexao.replace("D", "Δ")

            if "T" in conexao:
                conexao = conexao.replace("T", "t")
            # Rótulos conexao
            self.segments.append(
                SegmentText(
                    (-raio, 0),
                    conexao.split("-")[0],
                    fontsize=12,
                    align=('center', 'center'),  # CENTRALIZA no ponto (raio, 0)
                    color=cor_primario
                )
            )
            self.segments.append(
            SegmentText(
                (raio, 0),
                conexao.split("-")[1],
                fontsize=12,
                align=('center', 'center'),  # CENTRALIZA no ponto (raio, 0)
                color=cor_secundario
            )
        )
            self.segments.append(
            SegmentText(
                (0, -raio),
                conexao.split("-")[2],
                fontsize=12,
                align=('center', 'center'),  # CENTRALIZA no ponto (raio, 0)
                color=cor_terciario
            )
        )


# # Criando raio do curto
class Curto(elm.ElementImage):
    def __init__(self):
        super().__init__(
            r"raio.png",
            width=1.2, height=1.3
        )

class Transformador(elm.Element):
    def __init__(self, cor_primario=None, cor_secundario=None, conexao=None, **kwargs):
        super().__init__(**kwargs)
        raio = 0.6
        braco = 0.5  # comprimento dos braços (ajustável)

        # Bobinas (círculos)
        self.segments.append(SegmentCircle((-raio * 0.7, 0), raio, fill=None, color=cor_primario))   # Esquerda
        self.segments.append(SegmentCircle((raio * 0.7, 0), raio, fill=None, color=cor_secundario))    # Direita

        # Braços com base em 'braco'
        self.segments.append(Segment([(-raio * 0.7 - raio, 0), (-raio * 0.7 - raio - braco, 0)], color=cor_primario))  # Primário
        self.segments.append(Segment([(raio * 0.7 + raio, 0), (raio * 0.7 + raio + braco, 0)], color=cor_secundario))    # Secundário

        # Pontos de conexão finais (anchors nas pontas dos braços)
        self.anchors['p'] = (-raio * 0.7 - raio - braco, 0)
        self.anchors['s'] = (raio * 0.7 + raio + braco, 0)

        if "D" in conexao:
            conexao = conexao.replace("D", "Δ")

        if "T" in conexao:
            conexao = conexao.replace("T", "t")

        # Rótulos conexao
        self.segments.append(
            SegmentText(
                (-raio, 0),
                conexao.split("-")[0],
                fontsize=12,
                align=('center', 'center'),  # CENTRALIZA no ponto (raio, 0)
                color=cor_primario
            )
        )
        self.segments.append(
        SegmentText(
            (raio, 0),
            conexao.split("-")[1],
            fontsize=12,
            align=('center', 'center'),  # CENTRALIZA no ponto (raio, 0)
            color=cor_secundario
        )
    )

#----------------------------------------------------------------------------------------------------------------------
#-------------------------------------- INÍCIO DO DIAGRAMA ------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

with schemdraw.Drawing(file=fr"{caminho_diagrama}", dpi=150, theme='grade3', show=False) as d:
    d.config(fontsize=7)


# Desenhando o Elemento BARRA ****************************************************************************************
    for index, row in tabelas.Barra.iterrows():
        barra = row["Número"]
        tensao = row["Tensão (kV)"]
        posicao = posicao_elementos[barra]
        d += Barramento().at(posicao).label(f"Barra {barra}\n{tensao} kV").color(dicionario_cores[tensao])


# Desenhando o Elemento Maquina **************************************************************************************
    for index, row in tabelas.Maquina.iterrows():
        barra = row["Barra Conectada"]
        tensao = row["Tensão (kV)"]
        potencia = row["Potência Nominal (MVA)"]
        nome = row["Nome"]
        posicao_gerador = posicao_elementos[nome]
        posicao_barra = posicao_elementos[barra]
        conexao = row['Tipo de Conexão']
        

        teta = math.degrees(math.atan2(posicao_elementos[nome][1] - posicao_elementos[barra][1], posicao_elementos[nome][0] - posicao_elementos[barra][0]))
        d += elm.Line().at(posicao_barra).to(posicao_gerador).color(dicionario_cores[tensao])
        if row["Tipo de Máquina"] == "Gerador":
            d += elm.SourceSin().at(posicao_gerador).theta(teta).label(f"{nome}\n{tensao} kV\n{conexao}\n{potencia} MVA").color(dicionario_cores[tensao]).drop('end')
            d += elm.Line().length(0.4865).theta(teta + 180).linewidth(3).color("#FFFFFF").zorder(20) # Linha para ocultar segundo terminal do gerador (Gambiarra)
            d += elm.Line(capstyle='butt').length(0.4865).theta(teta + 180).linewidth(3).color("#FFFFFF").zorder(20)
        else:
            d += elm.Motor().at(posicao_gerador).theta(teta).label(f"{nome}\n{tensao} kV\n{potencia} MVA").color(dicionario_cores[tensao]).drop('end')
            d += elm.Line().length(0.4865).theta(teta + 180).linewidth(3).color("#FFFFFF").zorder(20) # Linha para ocultar segundo terminal do gerador (Gambiarra)
            d += elm.Line(capstyle='butt').length(0.265).theta(teta + 180).linewidth(3).color("#FFFFFF").zorder(20)

# Desenhando o Elemento Transformador ********************************************************************************
    for index, row in tabelas.Transformador.iterrows():
        nome = row["Nome"]
        posicao_barra_de = posicao_elementos[row["Barra de"]]
        posicao_barra_para = posicao_elementos[row["Barra para"]]
        posicao_transformador = posicao_elementos[row["Nome"]]
        conexao = row['Tipo de Conexão']
        potencia = row["Potência Nominal (MVA)"]
        tensao_primario = row["Tensão Primário (kV)"]
        tensao_secundario = row["Tensão Secundário (kV)"]

        teta = math.degrees(math.atan2(posicao_barra_para[1] - posicao_barra_de[1], posicao_barra_para[0] - posicao_barra_de[0]))
        xm, ym = Ajustes_trafos(posicao_barra_de, posicao_barra_para, posicao_transformador)
        # d += elm.Line().at(posicao_barra_de).to((xm, ym)).color(dicionario_cores[tensao_primario])
        trafo = Transformador(cor_primario=dicionario_cores[tensao_primario], cor_secundario=dicionario_cores[tensao_secundario], conexao=conexao).at(posicao_transformador).theta(teta).label(f"{nome}\n{tensao_primario}-{tensao_secundario} kV\n{potencia} MVA")
        d += trafo
        ponto_p = trafo.absanchors['p']
        ponto_s = trafo.absanchors['s']
        # d += elm.CurrentMirror().at((xm, ym)).theta(teta).label(f"{nome}\n{tensao_primario}-{tensao_secundario}\n{conexao}\n{potencia} MVA")
        d += elm.Line().at(ponto_s).to(posicao_barra_para).color(dicionario_cores[tensao_secundario])
        d += elm.Line().at(posicao_barra_de).to(ponto_p).color(dicionario_cores[tensao_primario])

# Desenhando o Elemento Linha ********************************************************************************
    for index, row in tabelas.Linha.iterrows():
        nome = row["Nome"]
        posicao_barra_de = posicao_elementos[row["Barra de"]]
        posicao_barra_para = posicao_elementos[row["Barra para"]]
        posicao_linha = posicao_elementos[row["Nome"]]
        tensao = row["Tensão (kV)"]
        teta = math.degrees(math.atan2(posicao_barra_para[1] - posicao_barra_de[1], posicao_barra_para[0] - posicao_barra_de[0]))
        d+= elm.Line().at(posicao_barra_de).to(posicao_linha).color(dicionario_cores[tensao]).label(f"{nome}")
        d+= elm.Line().at(posicao_linha).to(posicao_barra_para).color(dicionario_cores[tensao]).label(f"{nome}")


# Desenhando o Elemento Carga **************************************************************************************
    for index, row in tabelas.Carga.iterrows():
        barra = row["Barra Conectada"]
        tensao = row["Tensão (kV)"]
        nome = row["Nome"]
        posicao_carga = posicao_elementos[nome]
        posicao_barra = posicao_elementos[barra]
        conexao = row['Tipo de Conexão']
        potencia = f"{row['Potência Ativa (%)']} + J{row['Potência Reativa (%)']} MVA"
        teta = math.degrees(math.atan2(posicao_elementos[nome][1] - posicao_elementos[barra][1], posicao_elementos[nome][0] - posicao_elementos[barra][0]))
        d += elm.Line().at(posicao_barra).to(posicao_carga).color(dicionario_cores[tensao])
        d += elm.GroundSignal().at(posicao_carga).theta(teta+90).label(f"{nome}\n{potencia}\n{conexao}").color(dicionario_cores[tensao]).fill()


# Desenhando o Elemento Transformador de 3 Enrolamentos ***************************************************************
    
    for index, row in tabelas.Transformador_3E.iterrows():
        barra_p = row["Barra P"]
        barra_s = row["Barra S"]
        barra_t = row["Barra T"]

        tensao_p = row["Tensão Primário (kV)"]
        tensao_s = row["Tensão Secundário (kV)"]
        tensao_t = row["Tensão Terciário (kV)"]

        posicao_barra_p = posicao_elementos[barra_p]
        posicao_barra_s = posicao_elementos[barra_s]
        posicao_barra_t = posicao_elementos[barra_t]
        posicao_transformador_3e = posicao_elementos[row["Nome"]]
        potencia = row["Potência Nominal (MVA)"]
        conexao = row['Tipo de Conexão']
        nome = row["Nome"]
        
        melhor_teta, melhor_flip, melhor_reverse = tratamento_trafo_3e(posicao_barra_p, posicao_barra_s, posicao_barra_t, posicao_transformador_3e, d)
        trafo = Trafo_3_enrolamentos(cor_primario=dicionario_cores[tensao_p], cor_secundario=dicionario_cores[tensao_s], cor_terciario=dicionario_cores[tensao_t], conexao=conexao).at(posicao_transformador_3e).label(f"{nome}\n{tensao_p}-{tensao_s}-{tensao_t} kV\n{potencia} MVA").theta(melhor_teta)
        if melhor_flip:
            trafo = trafo.flip()

        if melhor_reverse:
            trafo = trafo.reverse()

        d += trafo

        # Capturar as coordenadas absolutas dos pontos de ancoragem
        ponto_p = trafo.absanchors['p']
        ponto_s = trafo.absanchors['s']
        ponto_t = trafo.absanchors['t']

        elm.Line().at(posicao_barra_p).to(ponto_p).color(dicionario_cores[tensao_p]).fill()
        elm.Line().at(posicao_barra_s).to(ponto_s).color(dicionario_cores[tensao_s]).fill()
        elm.Line().at(posicao_barra_t).to(ponto_t).color(dicionario_cores[tensao_t]).fill()


# Desenhando o Elemento Curto-Circuito ***************************************************************
    Curto().at(posicao_elementos[barra_curto])


# Essa função tem como objetivo criar uma margem branca no documento, evitando assim a interção
def adicionar_margem_pdf(arquivo_entrada, caminho_diagrama, margem_pt, nome_caso, tipo_curto, barra_curto, impedancia_falta, data, tensao_falta, dicionario_cores):
    doc_original = fitz.open(arquivo_entrada)
    novo_doc = fitz.open()

    for page in doc_original:
        largura, altura = page.rect.width, page.rect.height
        nova_largura = largura + 2 * margem_pt
        nova_altura = altura + 2 * margem_pt

        # Criar nova página no novo documento
        nova_pagina = novo_doc.new_page(width=nova_largura, height=nova_altura)

        # Inserir a página antiga com deslocamento (margem)
        nova_pagina.show_pdf_page(
            fitz.Rect(margem_pt, margem_pt, margem_pt + largura, margem_pt + altura),
            doc_original,
            page.number
        )
    doc_original.close()
    novo_doc.save(caminho_diagrama)
    novo_doc.close()

    doc = fitz.open(caminho_diagrama)
    pagina = doc[0] # Seleciona a página (ex: primeira página)

    # Pega os tamanhos do PDF
    tamanho = pagina.rect
    largura = tamanho.width
    altura = tamanho.height

    # Cria o retângulo da legenda técnica:
    distanciaH = 550 # Determina a Largaura da Legenda
    distanciaV = 125 # Determina a Altura da Legenda
    rect = fitz.Rect(largura-distanciaH, altura-distanciaV, largura-20, altura-20)
    pagina.draw_rect(rect, color=(0, 0, 0), width=3, fill=(1, 1, 1))  # Retângulo preto
    rect = fitz.Rect(20, 20, largura-20, altura-20)
    pagina.draw_rect(rect, color=(0, 0, 0), width=2)  # Retângulo preto


    # Insere dados da 1° Linha (Nome do Caso de Estudo)
    pagina.draw_line(p1=(largura-distanciaH, altura-distanciaV+35), p2=(largura-20, altura-distanciaV+35), color=(0, 0, 0)) # Referencia
    pagina.insert_text((largura-distanciaH + 5, altura-distanciaV+10), "Nome do Caso", fontsize=10, color=(0, 0, 0)) # -25
    pagina.insert_text((largura-distanciaH + 5, altura-distanciaV+32), nome_caso, fontsize=18, color=(0, 0, 0)) # -3


    # Insere dados da 2° Linha (Tipo de Curto, Barra em Curto, Impedância de Falta)
    pagina.draw_line(p1=(largura-distanciaH, altura-distanciaV+70), p2=(largura-20, altura-distanciaV+70), color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 5, altura-distanciaV+45), "Tipo de Curto Circuito", fontsize=10, color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 5, altura-distanciaV+67), f"Defeito {tipo_curto}", fontsize=18, color=(0, 0, 0))

    pagina.insert_text((largura-distanciaH + 230, altura-distanciaV+45), "Barra em Curto", fontsize=10, color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 230, altura-distanciaV+67), f"Barramento {barra_curto}", fontsize=18, color=(0, 0, 0))

    pagina.insert_text((largura-distanciaH + 390, altura-distanciaV+45), "Impedância de Falta", fontsize=10, color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 390, altura-distanciaV+67), f"{impedancia_falta} pu", fontsize=18, color=(0, 0, 0))


    # Insere dados da 3° Linha (Data, Tensão Pré Falta, "Diagram Unifilar")
    pagina.draw_line(p1=(largura-distanciaH, altura-distanciaV+105), p2=(largura-20, altura-distanciaV+105), color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 5, altura-distanciaV+80), "Data", fontsize=10, color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 5, altura-distanciaV+100), data, fontsize=18, color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 230, altura-distanciaV+80), "Tensão Pré-Falta", fontsize=10, color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 230, altura-distanciaV+100), f"{tensao_falta} pu", fontsize=18, color=(0, 0, 0))
    pagina.insert_text((largura-distanciaH + 345, altura-distanciaV+100), "Diagrama Unifilar", fontsize=18, color=(0, 0, 0))


    # Insere as divisões verticais
    pagina.draw_line(p1=(largura-300-30, altura-distanciaV+35), p2=(largura-300-30, altura-distanciaV+105), color=(0, 0, 0)) # Linha vertical
    pagina.draw_line(p1=(largura-140-30, altura-distanciaV+35), p2=(largura-140-30, altura-distanciaV+70), color=(0, 0, 0)) # Linha vertical
    pagina.draw_line(p1=(largura-170-40, altura-distanciaV+70), p2=(largura-170-40, altura-distanciaV+105), color=(0, 0, 0)) # Linha vertical

    c = 0
    inicio_x, inicio_y = largura-distanciaH-80, altura-distanciaV+60
    for chave, valor in dicionario_cores.items():
        pagina.insert_text((inicio_x - c, inicio_y+8), f"{str((chave))} kV", fontsize=18, color=(0, 0, 0))
        c = c + 20
        center = fitz.Point(inicio_x - c, inicio_y)  # coordenadas X e Y
        radius = 10
        c = c + 120
        shape = pagina.new_shape()
        shape.draw_circle(center, radius)
        shape.finish(color=valor, fill=valor, width=2)  # apenas contorno vermelho
        shape.commit()
    c = c + 80
    pagina.insert_text((inicio_x - c, inicio_y+8), "Níveis de Tensão:", fontsize=20, color=(0, 0, 0))

    # Salvar em um novo arquivo .pdf
    doc.save(fr"{caminho_diagrama.replace(".pdf","")}_Legendado.pdf")

adicionar_margem_pdf(caminho_diagrama, caminho_diagrama.replace(".pdf", "_Final.pdf"), 200, nome_caso, tipo_de_curto, barra_curto, impedancia_falta, data, tensao_falta, dicionario_cores)
ctypes.windll.user32.MessageBoxW(0, f"Diagrama gerado com sucesso!\n{caminho_diagrama.replace(".pdf","")}_Legendado.pdf", "Sucesso", 0)
if os.path.exists(caminho_diagrama):
    os.remove(caminho_diagrama)

if os.path.exists(caminho_diagrama.replace(".pdf", "_Final.pdf")):
    os.remove(caminho_diagrama.replace(".pdf", "_Final.pdf"))
os.startfile(fr"{caminho_diagrama.replace(".pdf","")}_Final_Legendado.pdf")
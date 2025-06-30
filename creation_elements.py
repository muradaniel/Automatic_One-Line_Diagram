import schemdraw
import schemdraw.elements as elm  # Possui os elementos elétricos já prontos
from schemdraw.segments import SegmentText
from schemdraw.segments import Segment # Criar seus próprios elementos
from schemdraw.segments import SegmentCircle

#from schemdraw.segments import *



# Criando barramento
class Barramento(elm.Element):
    def __init__(self, length=3, **kwargs):
        super().__init__(**kwargs)
        altura = 1 # Tamanho do barramento
        self.segments.append(Segment([(0, -altura), (0, altura)], lw = 5, capstyle='square')) # Desenha a linha
        self.anchors['centro'] = (0, altura/2) # Define ponto de Conexão



# Criando trafo de três enrolamentos
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



# Criando simbolo de curto circuito
class Curto(elm.ElementImage):
    def __init__(self):
        super().__init__(
            r"images/short_circuit_symbol.png",
            width=1.2, height=1.3
        )



# Criando trafo de 2 enrolamentos
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
        

class Maquina(elm.Element):
    def __init__(self, cor=None, type=None, **kwargs):
        super().__init__(**kwargs)  # Isso vem primeiro!
        raio = 0.6
        braco = 0.6
        self.segments.append(SegmentCircle(((2 * raio + braco)/2, 0), raio, color=cor))
        self.segments.append(Segment([(1.2+0.3, 0), (1.2+0.3+braco, 0)], color=cor))
        self.anchors["C"] = (1.2+0.3+braco, 0)
        self.segments.append(SegmentText(
            ((2 * raio + braco)/2, 0),  # posição
            type,                        # texto
            fontsize=20,
            align=('center', 'center'),
            color=cor
        ))
            


# with schemdraw.Drawing() as d:
#     d.config(fontsize=12)
#     motor = Maquina(cor="#3C00FF", type="G").theta(30)
#     d += motor
#     d += elm.Resistor().at(motor.anchors["C"]).theta(30)
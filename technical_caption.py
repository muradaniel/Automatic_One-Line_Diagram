import fitz


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
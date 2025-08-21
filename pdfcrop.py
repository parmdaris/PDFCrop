import fitz as mupdf
import sys
import os
import datetime
import shutil
import numpy as np

listaCr = [] # Armazena os finais dos códigos de rastreio para nomear os arquivos múltiplos
nova_etiqueta = mupdf.open()
nova_declaracao = mupdf.open()

coord_etiqueta_meli = [31.7, 28.7, 286.3, 449.3]
coord_etiqueta_menvio = [13, 24, 284, 400]

def prepararDiretorios():
    data_atual = getTimestamp()

    home_usuario = os.path.expanduser("~")
    raiz_meli = f"{home_usuario}/Documents/Vendas MELI/"
    dia_atual = f"{raiz_meli}/{data_atual}"

    pasta_etiquetas = f"{dia_atual}/Etiquetas"
    pasta_declaracoes = f"{dia_atual}/Declarações"
    pasta_originais = f"{dia_atual}/Originais"

    os.makedirs(dia_atual, exist_ok=True)
    os.makedirs(pasta_etiquetas, exist_ok=True)
    os.makedirs(pasta_declaracoes, exist_ok=True)
    os.makedirs(pasta_originais, exist_ok=True)

    return pasta_etiquetas, pasta_declaracoes, pasta_originais


def calcular_margem(coord_bruta, escala):
    
    x0, y0, x1, y1 = coord_bruta

    largura = x1 - x0
    altura = y1 - y0

    larg_extra = (largura * (1 - escala)) / 2
    alt_extra = (altura * (1 - escala)) / 2

    novo_x0 = x0 - larg_extra
    novo_y0 = y0 - alt_extra
    novo_x1 = x1 + larg_extra
    novo_y1 = y1 + alt_extra

    return [novo_x0, novo_y0, novo_x1, novo_y1]


def croparEtiqueta(coordenadas):
    pag_alvo = 0

    pag_etiqueta = original[0]
    crop_etiqueta = mupdf.Rect(coordenadas[0], coordenadas[1], coordenadas[2], coordenadas[3])
    
    linha_margem = pag_etiqueta.new_shape()
    linha_margem.draw_rect(crop_etiqueta)
    linha_margem.finish(width=0.5, color=(0, 0, 0))
    linha_margem.commit()
    
    coord_margem = calcular_margem(coordenadas, 0.95)

    margem = mupdf.Rect(coord_margem[0], coord_margem[1], coord_margem[2], coord_margem[3])

    pag_etiqueta.set_cropbox(margem) #Faz o tamanho da página que receberá o corte ser 5% maior que o recorte, criando uma margem.
    nova_etiqueta.insert_pdf(original, from_page=pag_alvo, to_page=pag_alvo)


def separarDeclaracao():
    pag_alvo = 1
    nova_declaracao.insert_pdf(original, from_page=pag_alvo, to_page=pag_alvo)


def getDadosMeLi(original):

    dados = original[1].get_text()
    
    linha = dados.split("\n")
    nome_destinatario = linha[5][6:]
    codigo_rastreio = linha[1][-13:]

    return nome_destinatario, codigo_rastreio


def getDadosMenvio(original):
    
    dados = original[0].get_text()
    
    linha = dados.split("\n")
    nome_destinatario = linha[11]
    codigo_rastreio = linha[6]

    return nome_destinatario, codigo_rastreio


def alterarOriginal(filepath, destino, codRastreio):
    pasta_originais = prepararDiretorios()[2]

    os.rename(os.path.abspath(filepath), f"{destino}_{codRastreio}.pdf")

    if os.path.exists(f"{pasta_originais}/{destino}_{codRastreio}.pdf"):
        os.remove(f"{pasta_originais}/{destino}_{codRastreio}.pdf")

    shutil.move(f"{destino}_{codRastreio}.pdf", pasta_originais) # Mover originais para a pasta Originais


def abrirArquivos(etiqueta_path, declaracao_path):
    os.startfile(etiqueta_path)
    if declaracao_path != False:
        os.startfile(declaracao_path)


def gerarTagsDC(destino, codRastreio, listaCr):

    finaisCr = "_".join(listaCr)

    arq_etiqueta = f"ETIQ_{destino}_{codRastreio}.pdf"
    arq_declaracao = f"DC_{destino}_{codRastreio}.pdf"

    if nova_etiqueta.page_count > 1: #Se houver mais de uma etiqueta processada
        arq_etiqueta = f"CROP_MULTIPLE_{finaisCr}.pdf"
    
    if nova_declaracao.page_count > 1: #Se houver mais de uma declaração processada
        arq_declaracao = f"DC_MULTIPLE_{finaisCr}.pdf" 

    pasta_etiquetas = prepararDiretorios()[0]
    pasta_declaracoes = prepararDiretorios()[1]

    etiqueta_path = os.path.join(pasta_etiquetas, arq_etiqueta)
    declaracao_path = os.path.join(pasta_declaracoes, arq_declaracao)

    nova_etiqueta.save(etiqueta_path)

    if nova_declaracao.page_count > 0:
        nova_declaracao.save(declaracao_path)
    else:
        declaracao_path = False
        
    nova_etiqueta.close()
    nova_declaracao.close()

    abrirArquivos(etiqueta_path, declaracao_path)


def getTimestamp():
    data = []
    timestamp = datetime.datetime.now().strftime("%Y%m%d")
    data.append(timestamp[6:]) #Dia atual
    data.append(timestamp[4:-2]) #Mês atual
    data.append(timestamp[:-4]) #Ano atual
    data_string = "-".join(data) #Criar string "dia-mes-ano"
    return data_string

if __name__ == '__main__':

    prepararDiretorios()

    for filepath in sys.argv[1:]:

        original = mupdf.open(os.path.abspath(filepath))
        qtd_pgs = original.page_count

        if qtd_pgs == 1: #Melhor envio

            dados = getDadosMenvio(original)
            
            destino = dados[0]
            codRastreio = dados[1]
            listaCr.append(codRastreio[-6:]) 

            croparEtiqueta(coord_etiqueta_menvio)

        if qtd_pgs == 3: #Mercado Livre
            
            dados = getDadosMeLi(original)
            
            destino = dados[0]
            codRastreio = dados[1]
            listaCr.append(codRastreio[-6:]) # Armazena os finais dos códigos de rastreio para nomear os arquivos múltiplos

            separarDeclaracao()
            croparEtiqueta(coord_etiqueta_meli)

        
        
        original.close()
        alterarOriginal(filepath, destino, codRastreio)


    gerarTagsDC(destino, codRastreio, listaCr)
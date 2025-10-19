import fitz as mupdf
import sys
import os
import datetime
import shutil
from PIL import Image
import pytesseract

listaCr = [] # Armazena os finais dos códigos de rastreio para nomear os arquivos múltiplos
nova_etiqueta = mupdf.open()
nova_declaracao = mupdf.open()

coord_etiqueta_meli = [31.7, 28.7, 286.3, 449.3]
coord_etiqueta_menvio = [13, 24, 284, 400]

def prepararDiretorios(): #Gera os diretórios de salvamento dos arquivos e retorna os caminhos.
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


def calcular_margem(coord_bruta, escala): #Calcula margem para redimensionamento a ser aplicada nas etiquetas.
    
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


def croparEtiqueta(coordenadas): #Faz a extração da etiqueta de dentro do PDF original.
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


def separarDeclaracao(): #Separa a declaração de conteúdo do arquivo original.
    pag_alvo = 1
    nova_declaracao.insert_pdf(original, from_page=pag_alvo, to_page=pag_alvo)


def getDadosMeLi(original): #Extrai os dados de destino de dentro da Declaração de Contrúdo do Mercado Livre.

    dados = original[1].get_text()
    
    linha = dados.split("\n")
    nome_destinatario = linha[5][6:]
    codigo_rastreio = linha[1][-13:]

    return nome_destinatario, codigo_rastreio


def getDadosMenvio(original): #Extrai os dados de destino de dentro da etiqueta do Melhor Envio.
    
    dados = original[0].get_text()
    
    if dados == "":
        nome_destinatario = "MENVIO"
        codigo_rastreio = "CR_MENVIO"
    else:
        linha = dados.split("\n")
        nome_destinatario = linha[11]
        codigo_rastreio = linha[6]

    return nome_destinatario, codigo_rastreio


def getTipoFlex(original): #Função para analisar se o documento pertence a um envio Flex ou a um documento do Melhor envio.
    dados = original[0].get_text()
    linhas = dados.split("\n")
    flex = False #Tipo padrão: Melhor envio

    for i in linhas:
        if i == f"Envio Flex":
            flex = True

    if flex == True: #Se detectado um documento de envio Flex, retornar verdadeiro.
        return True
    
    return False


def getDadosFlex(original): #Extrai os dados de destino de dentro da descrição de envio Flex.
    
    dados = original[1].get_text()
    
    if dados == "":
        nome_destinatario = "FLEX"
        codigo_id = "COD_FLEX"
    else:
        linha = dados.split("\n")
        nome_destinatario = linha[5]
        codigo_id = linha[2]

    return nome_destinatario, codigo_id
        


def alterarOriginal(filepath, destino, codRastreio): #Renomeia e move os arquivos originais.
    pasta_originais = prepararDiretorios()[2]

    os.rename(os.path.abspath(filepath), f"{destino}_{codRastreio}.pdf")

    if os.path.exists(f"{pasta_originais}/{destino}_{codRastreio}.pdf"):
        os.remove(f"{pasta_originais}/{destino}_{codRastreio}.pdf")

    shutil.move(f"{destino}_{codRastreio}.pdf", pasta_originais) # Mover originais para a pasta Originais


def abrirArquivos(etiqueta_path, declaracao_path): #Abre os arquivos gerados.
    os.startfile(etiqueta_path)
    if declaracao_path != False:
        os.startfile(declaracao_path)


def gerarTagsDC(destino, codRastreio, listaCr): #Grava as e as Declarações e etiquetas processadas em seus arquivos.

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


def getTimestamp(): #Cria uma string com a data atual.
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

        if qtd_pgs == 3: #Mercado Livre
            
            dados = getDadosMeLi(original)
            
            destino = dados[0]
            codRastreio = dados[1]
            listaCr.append(codRastreio[-6:]) # Armazena os finais dos códigos de rastreio para nomear os arquivos múltiplos

            separarDeclaracao()
            croparEtiqueta(coord_etiqueta_meli)


        if qtd_pgs < 3: #Melhor envio ou flex

            if getTipoFlex(original) == True: #Se for envio Flex:

                dados = getDadosFlex(original)
                
                destino = dados[0]
                codRastreio = dados[1]
                listaCr.append(codRastreio)

                croparEtiqueta(coord_etiqueta_meli) #Crop area da etiqueta de envio flex é a mesma da etiqueta normal do Mercado livre.
            
            else: #Se for melhor envio:
                dados = getDadosMenvio(original)
                
                destino = dados[0]
                codRastreio = dados[1]
                listaCr.append(codRastreio[-6:]) 

                croparEtiqueta(coord_etiqueta_menvio)

        
        
        original.close()
        alterarOriginal(filepath, destino, codRastreio)


    gerarTagsDC(destino, codRastreio, listaCr)
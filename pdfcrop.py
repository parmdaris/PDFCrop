import fitz as mupdf
import sys
import os
import datetime
import shutil

contador = 0
contador_timestamp = 0 # DEBUG
listaCr = []
nova_etiqueta = mupdf.open()
nova_declaracao = mupdf.open()

coord_etiqueta_meli = [31.7, 28.7, 286.3, 449.3]
coord_etiqueta_menvio = []

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

def croparEtiqueta(coordenadas):
    pag_alvo = 0

    pag_etiqueta = original[0]
    crop_etiqueta = mupdf.Rect(coordenadas[0], coordenadas[1], coordenadas[2], coordenadas[3])
    
    linha_margem = pag_etiqueta.new_shape()
    linha_margem.draw_rect(crop_etiqueta)
    linha_margem.finish(width=0.5, color=(0, 0, 0))
    linha_margem.commit()

    margem = mupdf.Rect(25, 17.632, 293, 460.368)
    pag_etiqueta.set_cropbox(margem) #Faz o tamanho da página que receberá o corte ser 5% maior que o recorte, criando uma margem.
    nova_etiqueta.insert_pdf(original, from_page=pag_alvo, to_page=pag_alvo)

def separarDeclaracao():
    pag_alvo = 1
    nova_declaracao.insert_pdf(original, from_page=pag_alvo, to_page=pag_alvo)

def getDadosDestino(original):
    dados = original[1].get_text()
    
    linha = dados.split("\n")
    nome_destinatario = linha[5][6:]
    codigo_rastreio = linha[1][-13:]

    return nome_destinatario, codigo_rastreio

def alterarOriginal(filepath, destino, codRastreio):
    pasta_originais = prepararDiretorios()[2]

    os.rename(os.path.abspath(filepath), f"{destino}_{codRastreio}.pdf")

    if os.path.exists(f"{pasta_originais}/{destino}_{codRastreio}.pdf"):
        os.remove(f"{pasta_originais}/{destino}_{codRastreio}.pdf")

    shutil.move(f"{destino}_{codRastreio}.pdf", pasta_originais) # Mover originais para a pasta Originais


def abrirArquivos(etiqueta_path, declaracao_path):
    os.startfile(etiqueta_path)
    os.startfile(declaracao_path)


def gerarTagsDC(destino, codRastreio, listaCr):
    if contador == 0:
        raise ValueError("Nenhum arquivo foi processado.")

    if contador > 1:
        finaisCr = "_".join(listaCr)
        arq_etiqueta = f"CROP_MULTIPLE_{finaisCr}.pdf"
        arq_declaracao = f"DC_MULTIPLE_{finaisCr}.pdf"
    else:
        arq_etiqueta = f"ETIQ_{destino}_{codRastreio}.pdf"
        arq_declaracao = f"DC_{destino}_{codRastreio}.pdf"

    pasta_etiquetas = prepararDiretorios()[0]
    pasta_declaracoes = prepararDiretorios()[1]

    etiqueta_path = os.path.join(pasta_etiquetas, arq_etiqueta)
    declaracao_path = os.path.join(pasta_declaracoes, arq_declaracao)

    nova_etiqueta.save(etiqueta_path)
    nova_declaracao.save(declaracao_path)

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
    #print(data_string) ######################################################### Investigar múltiplas chamadas da função
    return data_string

if __name__ == '__main__':

    prepararDiretorios()

    for filepath in sys.argv[1:]:

        original = mupdf.open(os.path.abspath(filepath))
        qtd_pgs = original.page_count
        dados = getDadosDestino(original) # Pode quebrar com uma etiqueta do melhor envio ************************************
        destino = dados[0]
        codRastreio = dados[1]

        listaCr.append(codRastreio[-6:]) # Armazena os finais dos códigos de rastreio para nomear os arquivos múltiplos

        if qtd_pgs < 3:
            coordenadas = coord_etiqueta_menvio # Melhorar tratamento do PDF do melhor envio. *****************************
        else:
            coordenadas = coord_etiqueta_meli
        croparEtiqueta(coordenadas)
        separarDeclaracao()
        original.close()
        alterarOriginal(filepath, destino, codRastreio)
        contador += 1


    gerarTagsDC(destino, codRastreio, listaCr)

    # Uma forma de diferenciar um PDF do MeLi de um do Melhor envio é a ausência de outras páginas no PDF. Pode ser utilizada como condição para uma rotina de averiguação do tipo de arquivo.
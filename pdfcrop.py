import fitz as mupdf
import sys
import os
import datetime
import shutil

contador = 0
last_filepath = None  # Variável para guardar o último nome processado
listaCr = []
nova_etiqueta = mupdf.open()
nova_declaracao = mupdf.open()

os.makedirs("Etiquetas", exist_ok=True)
os.makedirs("Declarações", exist_ok=True)
os.makedirs("Originais", exist_ok=True)

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
    dados = original[2].get_text()
    linha = dados.split("\n")
    return linha[5], linha[2]

for filepath in sys.argv[1:]:
    last_filepath = filepath

    original = mupdf.open(os.path.abspath(filepath))

    coord_etiqueta = [31.7, 28.7, 286.3, 449.3]
    croparEtiqueta(coord_etiqueta)
    separarDeclaracao()

    dados = getDadosDestino(original)

    destino = dados[0]
    codRastreio = dados[1]

    listaCr.append(codRastreio[-6:]) # Armazena os finais dos códigos de rastreio para nomear os arquivos múltiplos
    
    original.close()

    os.rename(os.path.abspath(filepath), f"{destino}_{codRastreio}.pdf")

    if os.path.exists(f"Originais/{destino}_{codRastreio}.pdf"):
        os.remove(f"Originais/{destino}_{codRastreio}.pdf")

    shutil.move(f"{destino}_{codRastreio}.pdf", "Originais") # Mover originais para a pasta Originais
    contador += 1

finaisCr = "_".join(listaCr)

if contador == 0:
    raise ValueError("Nenhum arquivo foi processado.")

if contador > 1:
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    arq_etiqueta = f"CROP_MULTIPLE_{finaisCr}.pdf"
    arq_declaracao = f"DC_MULTIPLE_{finaisCr}.pdf"
else:
    
    arq_etiqueta = f"ETIQ_{destino}_{codRastreio}.pdf"
    arq_declaracao = f"DC_{destino}_{codRastreio}.pdf"

etiqueta_path = os.path.join("Etiquetas", arq_etiqueta)
declaracao_path = os.path.join("Declarações", arq_declaracao)

nova_etiqueta.save(etiqueta_path)
nova_declaracao.save(declaracao_path)

nova_etiqueta.close()
nova_declaracao.close()

os.startfile(etiqueta_path)
os.startfile(declaracao_path)
import fitz as mupdf
import sys
import os
import datetime

counter = 0
last_filepath = None  # Variável para guardar o último nome processado

new_doc = mupdf.open()

os.makedirs("Etiquetas", exist_ok=True)

for filepath in sys.argv[1:]:
    last_filepath = filepath

    doc = mupdf.open(os.path.abspath(filepath))

    tgt_page = 0
    page = doc[0]

    crop_area = mupdf.Rect(31.7, 28.7, 286.3, 449.3) 

    margem = page.new_shape()
    margem.draw_rect(crop_area)
    margem.finish(width=0.5, color=(0, 0, 0))
    margem.commit()

    page.set_cropbox(crop_area)

    new_doc.insert_pdf(doc, from_page=tgt_page, to_page=tgt_page)

    doc.close()
    counter += 1

if counter == 0:
    raise ValueError("Nenhum arquivo foi processado. Verifique se o script está recebendo PDFs como argumento.")

if counter > 1:
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    new_name = f"CROP_MULTIPLE_{timestamp}.pdf"
else:
    new_name = f"CROP_{os.path.basename(last_filepath)}"

output_path = os.path.join("Etiquetas", new_name)
new_doc.save(output_path)
new_doc.close()

os.startfile(output_path)
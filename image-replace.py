from docx import Document
import os

# Ruta de la carpeta que contiene los documentos Word
carpeta_documentos = 'archivos'

# Ruta de la carpeta donde se guardarán los documentos modificados
carpeta_modificados = 'archivos_modificados'

# Ruta de la imagen de referencia y la nueva imagen
imagen_referencia_path = 'referencia.jpg'
nueva_imagen_path = 'image.png'

# Asegúrate de que la carpeta de destino exista, si no, créala
if not os.path.exists(carpeta_modificados):
    os.makedirs(carpeta_modificados)

# Cargar la imagen de referencia para comparar
with open(imagen_referencia_path, 'rb') as ref_img_file:
    imagen_referencia_data = ref_img_file.read()

# Función para reemplazar la imagen en un documento
def reemplazar_imagen_en_documento(doc_path):
    # Cargar el documento con python-docx
    doc = Document(doc_path)

    # Obtener todas las relaciones del documento (imágenes y otros recursos)
    relations = doc.part.rels

    imagen_reemplazada = False
    for rel in relations.values():
        if 'image' in rel.target_ref:
            imagen_doc_data = rel.target_part.blob

            # Comparar con la imagen de referencia
            if imagen_doc_data == imagen_referencia_data:
                print(f"Imagen coincidente encontrada y reemplazada en {doc_path}.")
                
                # Reemplazar la imagen
                with open(nueva_imagen_path, 'rb') as nueva_img_file:
                    rel.target_part._blob = nueva_img_file.read()
                imagen_reemplazada = True
                break

    # Guardar el documento con la imagen reemplazada
    if imagen_reemplazada:
        # Guardar el documento modificado en la carpeta de salida
        new_doc_path = os.path.join(carpeta_modificados, os.path.basename(doc_path))
        doc.save(new_doc_path)  # Guardamos el documento modificado
        print(f"Documento guardado como {new_doc_path}.")
    else:
        print(f"No se encontró ninguna imagen coincidente en {doc_path}.")

# Recorrer todos los archivos .docx en la carpeta
for archivo in os.listdir(carpeta_documentos):
    if archivo.endswith('.docx'):
        doc_path = os.path.join(carpeta_documentos, archivo)
        reemplazar_imagen_en_documento(doc_path)
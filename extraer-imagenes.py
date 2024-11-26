from docx import Document
import os

# Ruta del documento Word
doc_path = 'extractor.docx'

# Crear una carpeta para guardar las imágenes extraídas
output_folder = 'imagenes_extraidas'
os.makedirs(output_folder, exist_ok=True)

# Cargar el documento
doc = Document(doc_path)

# Obtener todas las relaciones del documento (imagenes y otros recursos)
relations = doc.part.rels

# Extraer y guardar cada imagen
for i, rel in enumerate(relations.values()):
    if 'image' in rel.target_ref:
        image_data = rel.target_part.blob
        image_path = os.path.join(output_folder, f"imagen_{i+1}.jpg")
        with open(image_path, 'wb') as img_file:
            img_file.write(image_data)
        print(f"Imagen {i+1} guardada como: {image_path}")

from docx import Document
import os
from tkinter import Tk, Label, Button, filedialog, Frame
from PIL import Image, ImageTk

# Variables globales
imagenes_referencia_paths = []  # Lista para almacenar las rutas de las imágenes de referencia
asociaciones = {}  # Diccionario para almacenar las asociaciones (imagen de referencia -> nueva imagen)

def seleccionar_imagenes_referencia():
    """Abre un cuadro de diálogo para seleccionar múltiples imágenes de referencia."""
    global imagenes_referencia_paths
    imagenes_referencia_paths = filedialog.askopenfilenames(
        title="Seleccioná las imágenes de referencia",
        filetypes=[("Archivos de imagen", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")]
    )
    if imagenes_referencia_paths:
        label_referencia.config(text=f"{len(imagenes_referencia_paths)} imágenes de referencia seleccionadas.")
        mostrar_imagenes()  # Mostrar las imágenes seleccionadas con su preview

def mostrar_imagenes():
    """Muestra las imágenes de referencia con un preview y el botón para seleccionar la nueva imagen."""
    for widget in frame_imagenes.winfo_children():
        widget.destroy()

    for imagen_referencia in imagenes_referencia_paths:
        frame_imagen = Frame(frame_imagenes)
        frame_imagen.pack(pady=5, anchor="w")

        img_ref = Image.open(imagen_referencia)
        img_ref.thumbnail((100, 100))
        img_ref_tk = ImageTk.PhotoImage(img_ref)

        label_preview_ref = Label(frame_imagen, image=img_ref_tk)
        label_preview_ref.image = img_ref_tk
        label_preview_ref.grid(row=0, column=0, padx=5)

        boton_asociar = Button(frame_imagen, text="Seleccionar nueva imagen",
                               command=lambda img_ref=imagen_referencia, parent=frame_imagen: seleccionar_nueva_imagen(img_ref, parent))
        boton_asociar.grid(row=0, column=1, padx=10)

        label_preview_new = Label(frame_imagen, text="Nueva imagen no seleccionada", fg="gray")
        label_preview_new.grid(row=0, column=2, padx=5)

def seleccionar_nueva_imagen(imagen_referencia, frame_imagen):
    """Selecciona una nueva imagen para una referencia."""
    nueva_imagen_path = filedialog.askopenfilename(
        title=f"Selecciona la nueva imagen para {os.path.basename(imagen_referencia)}",
        filetypes=[("Archivos de imagen", "*.jpg;*.jpeg;*.png;*.bmp;*.gif")]
    )
    if nueva_imagen_path:
        asociaciones[imagen_referencia] = nueva_imagen_path

        img_new = Image.open(nueva_imagen_path)
        img_new.thumbnail((100, 100))
        img_new_tk = ImageTk.PhotoImage(img_new)

        label_preview_new = frame_imagen.grid_slaves(row=0, column=2)[0]
        label_preview_new.config(image=img_new_tk, text="")
        label_preview_new.image = img_new_tk

        label_resultado.config(text=f"Asociación: {os.path.basename(imagen_referencia)} -> {os.path.basename(nueva_imagen_path)}", fg="green")
        print(f"Asociación: {os.path.basename(imagen_referencia)} -> {os.path.basename(nueva_imagen_path)}")

def reemplazar_imagenes():
    """Reemplaza las imágenes asociadas en los documentos."""
    if not asociaciones:
        label_resultado.config(text="Error: No hay imágenes asociadas para reemplazar.", fg="red")
        return

    carpeta_documentos = 'archivos'
    carpeta_modificados = 'archivos_modificados'

    if not os.path.exists(carpeta_modificados):
        os.makedirs(carpeta_modificados)

    for archivo in os.listdir(carpeta_documentos):
        if archivo.endswith('.docx'):
            doc_path = os.path.join(carpeta_documentos, archivo)
            reemplazar_en_documento(doc_path)

    label_resultado.config(text="Reemplazo completado. Revisa la carpeta de salida.", fg="green")

def reemplazar_en_documento(doc_path):
    """Aplica todas las asociaciones en un documento."""
    doc = Document(doc_path)
    relations = doc.part.rels
    reemplazo_realizado = False

    for rel in relations.values():
        if 'image' in rel.target_ref:
            imagen_doc_data = rel.target_part.blob

            for referencia_path, nueva_path in asociaciones.items():
                with open(referencia_path, 'rb') as ref_img_file:
                    referencia_data = ref_img_file.read()

                if imagen_doc_data == referencia_data:
                    with open(nueva_path, 'rb') as nueva_img_file:
                        rel.target_part._blob = nueva_img_file.read()
                        reemplazo_realizado = True
                        print(f"Reemplazada imagen en {doc_path}: {os.path.basename(referencia_path)} -> {os.path.basename(nueva_path)}")
                    break

    if reemplazo_realizado:
        carpeta_modificados = 'archivos_modificados'
        nuevo_doc_path = os.path.join(carpeta_modificados, os.path.basename(doc_path))
        doc.save(nuevo_doc_path)

# Interfaz gráfica
root = Tk()
root.title("Reemplazar Imágenes en Documentos Word")

label_instrucciones = Label(root, text="Seleccioná las imágenes para reemplazar:", font=("Arial", 12))
label_instrucciones.pack(pady=10)

boton_referencia = Button(root, text="Seleccionar Imágenes de Referencia", command=seleccionar_imagenes_referencia)
boton_referencia.pack(pady=5)

label_referencia = Label(root, text="No se han seleccionado imágenes de referencia.", fg="gray")
label_referencia.pack(pady=5)

frame_imagenes = Frame(root)
frame_imagenes.pack(pady=10)

boton_reemplazar = Button(root, text="Reemplazar Imágenes", command=reemplazar_imagenes, bg="blue", fg="white")
boton_reemplazar.pack(pady=20)

label_resultado = Label(root, text="")
label_resultado.pack(pady=10)

root.mainloop()
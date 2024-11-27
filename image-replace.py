import os
import win32com.client
import pywintypes  # Agregar esta importación
from docx import Document  # Asegúrate de que esta línea esté presente
from tkinter import Tk, Label, Button, filedialog, Frame, Listbox, Scrollbar
from PIL import Image, ImageTk

# Variables globales
imagenes_referencia_paths = []  # Lista para almacenar las rutas de las imágenes de referencia
asociaciones = {}  # Diccionario para almacenar las asociaciones (imagen de referencia -> nueva imagen)
archivos_pdf_generados = []  # Lista para almacenar los nombres de los archivos PDF generados

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
        label_preview_new.config(image=img_new_tk, text="")  # Actualizamos la imagen
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
    doc = Document(doc_path)  # Esto usará la clase Document importada
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

def convertir_documentos_a_pdf():
    """Convierte todos los documentos en 'archivos_modificados' a PDF y guarda los PDFs en una subcarpeta 'pdf'."""
    carpeta_modificados = os.path.abspath('archivos_modificados')  # Convertir a ruta absoluta
    carpeta_pdf = os.path.join(carpeta_modificados, 'pdf')

    # Crear la carpeta 'pdf' si no existe
    if not os.path.exists(carpeta_pdf):
        os.makedirs(carpeta_pdf)

    word = win32com.client.Dispatch('Word.Application')

    # Verificar si la carpeta existe
    if not os.path.exists(carpeta_modificados):
        print(f"La carpeta {carpeta_modificados} no existe.")
        return

    global archivos_pdf_generados
    archivos_pdf_generados = []  # Limpiar lista antes de agregar nuevos archivos

    for archivo in os.listdir(carpeta_modificados):
        if archivo.endswith('.docx'):
            doc_path = os.path.join(carpeta_modificados, archivo)

            # Verificar si el archivo existe
            if not os.path.exists(doc_path):
                print(f"El archivo no existe: {doc_path}")
                continue  # Saltar al siguiente archivo

            try:
                # Log de la ruta completa donde se va a intentar abrir el archivo
                print(f"Intentando abrir el archivo en la ruta: {doc_path}")
                
                doc = word.Documents.Open(doc_path)
                output_pdf = os.path.join(carpeta_pdf, os.path.splitext(archivo)[0] + '.pdf')

                # Guardar como PDF
                doc.SaveAs(output_pdf, FileFormat=17)  # 17 es el formato PDF en Word
                doc.Close()
                print(f"Documento convertido a PDF: {output_pdf}")

                # Agregar solo el nombre del archivo PDF (sin la ruta)
                archivos_pdf_generados.append(os.path.basename(output_pdf))
            except Exception as e:
                print(f"Error al convertir el archivo {doc_path}: {str(e)}")
                # Si la ruta es incorrecta o el archivo no se puede abrir, también imprimimos el error
                if isinstance(e, pywintypes.com_error):
                    print(f"Ruta intentada para abrir el archivo: {doc_path}")
                continue

    # Actualizar la lista de archivos PDF generados en la interfaz
    actualizar_lista_pdf()

    label_resultado.config(text="Conversión a PDF completada.", fg="green")

def actualizar_lista_pdf():
    """Actualiza la lista de archivos PDF generados en el Listbox."""
    listbox_pdf.delete(0, 'end')  # Limpiar la lista
    for archivo_pdf in archivos_pdf_generados:
        listbox_pdf.insert('end', archivo_pdf)

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

boton_pdf = Button(root, text="Generar PDF de Todos los Archivos Modificados", command=convertir_documentos_a_pdf, bg="green", fg="white")
boton_pdf.pack(pady=20)

label_resultado = Label(root, text="")
label_resultado.pack(pady=10)

# Frame para el Listbox de los archivos PDF generados
listbox_pdf_frame = Frame(root)
listbox_pdf_frame.pack(pady=10)

# Título para el Listbox
label_pdf_titulo = Label(listbox_pdf_frame, text="Archivos PDF generados", font=("Arial", 12))
label_pdf_titulo.pack()

# Listbox con scroll vertical y horizontal
listbox_pdf = Listbox(listbox_pdf_frame, height=10, width=50, selectmode='single')
listbox_pdf.pack(side="left", fill="y")

scrollbar_pdf_y = Scrollbar(listbox_pdf_frame, orient="vertical", command=listbox_pdf.yview)
scrollbar_pdf_y.pack(side="right", fill="y")

scrollbar_pdf_x = Scrollbar(listbox_pdf_frame, orient="horizontal", command=listbox_pdf.xview)
scrollbar_pdf_x.pack(side="bottom", fill="x")

listbox_pdf.config(yscrollcommand=scrollbar_pdf_y.set, xscrollcommand=scrollbar_pdf_x.set)

root.mainloop()

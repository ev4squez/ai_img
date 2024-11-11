import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import csv
import os
from PIL import Image, ImageTk
import pandas as pd

# Función para ejecutar el procesamiento de imágenes con barra de progreso
def ejecutar_procesamiento():
    if not carpeta_seleccionada.get():
        messagebox.showwarning("Advertencia", "Por favor, selecciona una carpeta.")
        return

    # Limpia la grilla antes de mostrar nuevos datos
    for item in tree.get_children():
        tree.delete(item)

    directorio_imagenes = carpeta_seleccionada.get()
    archivos_imagenes = [f for f in os.listdir(directorio_imagenes) if f.endswith(('.png', '.JPG',  '.jpg', '.jpeg', '.bmp', '.tiff'))]
    total_imagenes = len(archivos_imagenes)
    
    if total_imagenes == 0:
        messagebox.showinfo("Sin imágenes", "La carpeta seleccionada no contiene imágenes válidas.")
        return
    
    # Configura la barra de progreso
    barra_progreso["maximum"] = total_imagenes
    barra_progreso["value"] = 0
    ventana.update_idletasks()

    # Etiqueta para mostrar imágenes restantes y porcentaje
    etiqueta_progreso.config(text=f"Procesando... 0/{total_imagenes} imágenes procesadas")

    # Llama a la función de procesamiento y actualiza la barra de progreso en cada iteración
    archivo_csv = 'numeros_extraidos.csv'
    with open(archivo_csv, mode='w', newline='') as archivo:
        escritor = csv.writer(archivo)
        escritor.writerow(["Archivo de Imagen", "Números Extraídos"])

        for i, nombre_archivo in enumerate(archivos_imagenes, start=1):
            ruta_imagen = os.path.join(directorio_imagenes, nombre_archivo)
            numeros_extraidos = procesar_imagen_y_extraer_numeros(ruta_imagen)
            escritor.writerow([nombre_archivo, ', '.join(numeros_extraidos)])
            
            # Actualiza la barra de progreso
            barra_progreso["value"] = i
            ventana.update_idletasks()

            # Actualiza el texto con las imágenes faltantes
            imagenes_faltantes = total_imagenes - i
            etiqueta_progreso.config(text=f"Procesando... {i}/{total_imagenes} imágenes procesadas, {imagenes_faltantes} restantes")

    messagebox.showinfo("Proceso Completado", f"Extracción completada. Datos guardados en {archivo_csv}")
    mostrar_csv_en_grilla(archivo_csv)

# Función para seleccionar carpeta
def seleccionar_carpeta():
    ruta_carpeta = filedialog.askdirectory()
    carpeta_seleccionada.set(ruta_carpeta)

# Función para procesar una sola imagen y extraer números
def procesar_imagen_y_extraer_numeros(ruta_imagen):
    import easyocr
    import cv2
    
    reader = easyocr.Reader(['en'])
    imagen = cv2.imread(ruta_imagen)
    resultado = reader.readtext(imagen)

    # Filtra solo datos numéricos
    numeros_extraidos = []
    for deteccion in resultado:
        _, texto, confianza = deteccion
        if confianza > 0.5 and any(char.isdigit() for char in texto):
            numeros = ''.join([char for char in texto if char.isdigit() or char == '.'])
            numeros_extraidos.append(numeros)
    
    return numeros_extraidos

# Función para mostrar el contenido del CSV en la grilla
def mostrar_csv_en_grilla(archivo_csv):
    # Limpia la grilla si tiene datos
    for item in tree.get_children():
        tree.delete(item)

    # Carga los datos del archivo CSV en la grilla
    with open(archivo_csv, newline='') as archivo:
        lector_csv = csv.reader(archivo)
        encabezados = next(lector_csv)  # Obtén los encabezados

        # Configura los encabezados de la grilla
        tree["columns"] = encabezados
        for encabezado in encabezados:
            tree.heading(encabezado, text=encabezado, anchor="w")  # Alineación de encabezados a la izquierda
            tree.column(encabezado, anchor="w")  # Alineación de datos a la izquierda

        # Agrega los datos del CSV a la grilla
        for fila in lector_csv:
            tree.insert("", "end", values=fila)

# Función para mostrar la imagen seleccionada
def mostrar_imagen(event):
    seleccion = tree.selection()
    if seleccion:
        # Obtiene el nombre del archivo de la fila seleccionada
        item = tree.item(seleccion[0])
        nombre_archivo = item["values"][0]
        ruta_imagen = os.path.join(carpeta_seleccionada.get(), nombre_archivo)
        
        # Cargar y redimensionar la imagen
        imagen = Image.open(ruta_imagen)
        max_width = 650  # Máximo ancho deseado para la imagen
        max_height = 650  # Máxima altura deseada para la imagen
        imagen.thumbnail((max_width, max_height))  # Ajusta el tamaño proporcionalmente
        
        # Convierte la imagen a formato adecuado para Tkinter
        imagen_tk = ImageTk.PhotoImage(imagen)
        
        # Muestra la imagen en el widget Label
        etiqueta_imagen.config(image=imagen_tk)
        etiqueta_imagen.image = imagen_tk  # Guarda una referencia para evitar que sea recolectada por el GC

# Función para editar la celda seleccionada y actualizar la imagen
def editar_celda(event):
    seleccion = tree.selection()
    if seleccion:
        # Obtiene el nombre del archivo de la fila seleccionada
        item = tree.item(seleccion[0])
        nombre_archivo = item["values"][0]
        ruta_imagen = os.path.join(carpeta_seleccionada.get(), nombre_archivo)
        
        # Abre una ventana emergente para editar la segunda columna (números extraídos)
        nuevo_valor = simpledialog.askstring("Editar", f"Edita los números extraídos para {nombre_archivo}:",
                                             initialvalue=item["values"][1])
        if nuevo_valor is not None:
            # Actualiza la grilla y el archivo CSV con el nuevo valor
            tree.item(seleccion[0], values=(nombre_archivo, nuevo_valor))
            actualizar_csv()

            # Actualiza la imagen según el nuevo valor de la celda
            mostrar_imagen(event)

# Función para actualizar el archivo CSV con los nuevos valores
def actualizar_csv():
    # Guardar los datos actualizados en el archivo CSV
    with open('numeros_extraidos.csv', mode='w', newline='') as archivo:
        escritor = csv.writer(archivo)
        escritor.writerow(["Archivo de Imagen", "Números Extraídos"])
        for item in tree.get_children():
            valores = tree.item(item)["values"]
            escritor.writerow(valores)

# Función para exportar el contenido de la grilla a un archivo Excel
def exportar_a_excel():
    # Preguntar dónde guardar el archivo
    archivo_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if archivo_excel:
        # Obtener los datos de la grilla y almacenarlos en una lista de diccionarios
        datos = []
        encabezados = tree["columns"]
        for item_id in tree.get_children():
            fila = tree.item(item_id)["values"]
            datos.append(dict(zip(encabezados, fila)))
        
        # Crear un DataFrame y exportarlo a Excel
        df = pd.DataFrame(datos)
        df.to_excel(archivo_excel, index=False, engine='openpyxl')
        
        # Mensaje de éxito
        messagebox.showinfo("Exportación Exitosa", f"Los datos se han exportado a {archivo_excel} exitosamente.")

## Configuración de la ventana de la aplicación
ventana = tk.Tk()
ventana.title("Extracción de Números desde Imágenes")
ventana.geometry("800x500")

# Variable para almacenar la ruta de la carpeta seleccionada
carpeta_seleccionada = tk.StringVar()

# Elementos de la interfaz
tk.Label(ventana, text="Seleccione la carpeta con las imágenes:").pack(pady=10)
tk.Entry(ventana, textvariable=carpeta_seleccionada, width=40, state='readonly').pack(pady=5)
tk.Button(ventana, text="Seleccionar Carpeta", command=seleccionar_carpeta).pack(pady=5)
tk.Button(ventana, text="Procesar Imágenes y Guardar CSV", command=ejecutar_procesamiento).pack(pady=20)

# Barra de progreso;
barra_progreso = ttk.Progressbar(ventana, orient="horizontal", length=400, mode="determinate")
barra_progreso.pack(pady=10)

# Etiqueta para mostrar el progreso con las imágenes restantes
etiqueta_progreso = tk.Label(ventana, text="Procesando... 0/0 imágenes procesadas")
etiqueta_progreso.pack(pady=10)

# Frame principal para contener la grilla y la vista de imagen
frame_principal = tk.Frame(ventana)
frame_principal.pack(fill="both", expand=True, padx=10, pady=10)

# Crear el Treeview para mostrar la grilla en un Frame con Scrollbar
frame_grilla = tk.Frame(frame_principal)
frame_grilla.pack(side="left", fill="both", expand=True)

# Botón para exportar a Excel
tk.Button(ventana, text="Exportar a Excel", command=exportar_a_excel).pack(pady=5)

tree = ttk.Treeview(frame_grilla, show="headings", height=10)
tree.pack(side="left", fill="both", expand=True)

# Evento <<TreeviewSelect>> para detectar selección de filas y cargar la imagen
tree.bind("<<TreeviewSelect>>", mostrar_imagen)
tree.bind("<Double-1>", editar_celda)
tree.bind("<Return>", editar_celda)

# Barra de desplazamiento para la grilla
scrollbar_vertical = ttk.Scrollbar(frame_grilla, orient="vertical", command=tree.yview)
scrollbar_vertical.pack(side="right", fill="y")
tree.configure(yscroll=scrollbar_vertical.set)

# Crear un marco para la vista de imagen
marco_imagen = tk.Frame(frame_principal, width=650, height=650)
marco_imagen.pack(side="right", fill="both", expand=True)
etiqueta_imagen = tk.Label(marco_imagen)
etiqueta_imagen.pack(fill="both", expand=True)

# Inicia la aplicación
ventana.mainloop()

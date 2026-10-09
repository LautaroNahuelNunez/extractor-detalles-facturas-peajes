import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os
import re
import pandas as pd
import pdfplumber

# Limpieza y extracción

def limpiar_tag(tag_crudo):
    tag_limpio = tag_crudo.replace('S1', 'SI') # Salva error de tipeo si se pone S1 en lugar de SI
    if tag_limpio.startswith('SI9000'):
        tag_limpio = tag_limpio.replace('SI9000', 'SI90', 1) # Salva error de nomenclatura cuando rellenan con ceros
    return tag_limpio

def limpiar_importe(importe_crudo):
    
    if ',' in importe_crudo and '.' in importe_crudo: # Salva el uso de comas por punto para los decimales
        if importe_crudo.rfind(',') > importe_crudo.rfind('.'):
            return float(importe_crudo.replace('.', '').replace(',', '.'))  
        else:
            return float(importe_crudo.replace(',', ''))
    elif ',' in importe_crudo:
        return float(importe_crudo.replace(',', '.'))
    else:
        return float(importe_crudo)

def procesar_lista_pdfs(lista_pdfs, archivo_salida):
    totales_por_tag = {}
    patron_tag = re.compile(r'(SI\d+|S1\d+).*?(-?\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|-?\d+[.,]\d{2})')
    patron_importe_suelto = re.compile(r'(-?\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|-?\d+[.,]\d{2})$')

    for ruta_pdf in lista_pdfs:
        with pdfplumber.open(ruta_pdf) as pdf:
            tag_en_memoria = None
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if not texto:
                    continue

                for linea in texto.split('\n'):
                    if "SUBTOTAL" in linea.upper() or "TOTAL" in linea.upper():
                        tag_en_memoria = None
                        continue

                    matches = list(patron_tag.finditer(linea))

                    if matches:
                        for match in matches:
                            tag = limpiar_tag(match.group(1))
                            importe = limpiar_importe(match.group(2))
                            totales_por_tag[tag] = totales_por_tag.get(tag, 0.0) + importe
                            tag_en_memoria = tag
                    else:
                        if tag_en_memoria and ("PASADAS" in linea.upper() or "CAT" in linea.upper() or "PEAJE" in linea.upper()):
                            match_suelto = patron_importe_suelto.search(linea.strip())
                            if match_suelto:
                                importe_suelto = limpiar_importe(match_suelto.group(1))
                                totales_por_tag[tag_en_memoria] += importe_suelto

    df = pd.DataFrame(list(totales_por_tag.items()), columns=['TAG', 'IMPORTE TOTAL'])
    df = df.sort_values(by='IMPORTE TOTAL', ascending=False)
    
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer: # Salida a Excel
        df.to_excel(writer, index=False, sheet_name='Peajes')
        hoja = writer.sheets['Peajes']
        for celda in hoja['B']:
            if celda.row > 1:
                celda.number_format = '$ #,##0.00'

# Interfaz Gráfica

class AppPeajes:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Facturas de Peajes")
        self.root.geometry("600x450") # Ventana más grande para las miniaturas
        self.root.eval('tk::PlaceWindow . center') # Centrado de ventana
        
        self.rutas_pdfs = []
        self.imagenes_referencia = [] # Lista vital para que Tkinter no borre las imágenes de la memoria

        # Título
        tk.Label(root, text="Procesador de Peajes", font=("Arial", 14, "bold")).pack(pady=10)

        # Frame para alinear los botones
        frame_botones = tk.Frame(root)
        frame_botones.pack(pady=5)

        self.btn_cargar = tk.Button(frame_botones, text="1. Cargar facturas (PDF)", command=self.cargar_archivos, width=22, height=2)
        self.btn_cargar.pack(side=tk.LEFT, padx=10)

        self.btn_procesar = tk.Button(frame_botones, text="2. Procesar y Guardar Excel", command=self.procesar_archivos, width=22, height=2, state=tk.DISABLED)
        self.btn_procesar.pack(side=tk.LEFT, padx=10)

        self.lbl_estado = tk.Label(root, text="Ningún archivo seleccionado.", fg="gray")
        self.lbl_estado.pack(pady=5)

        # Zona de Miniaturas (Canvas con Scroll)
        marco_exterior = tk.Frame(root, bd=2, relief="sunken")
        marco_exterior.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.canvas = tk.Canvas(marco_exterior, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(marco_exterior, orient="horizontal", command=self.canvas.xview)
        self.frame_miniaturas = tk.Frame(self.canvas)

        # Configurar el scroll automático
        self.frame_miniaturas.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.frame_miniaturas, anchor="nw")
        self.canvas.configure(xscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.BOTTOM, fill=tk.X)


    def cargar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar facturas PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if archivos:
            self.rutas_pdfs = list(archivos)
            cantidad = len(self.rutas_pdfs)
            self.lbl_estado.config(text=f"{cantidad} archivo(s) seleccionado(s).", fg="green")
            self.btn_procesar.config(state=tk.NORMAL)
            
            # Generar las miniaturas visuales
            self.mostrar_miniaturas()

    def mostrar_miniaturas(self):
        # Limpiar miniaturas anteriores
        for widget in self.frame_miniaturas.winfo_children():
            widget.destroy()
        self.imagenes_referencia.clear()

        # Poner cursor de carga mientras dibuja
        self.root.config(cursor="watch")
        self.root.update()

        for ruta in self.rutas_pdfs:
            nombre_archivo = os.path.basename(ruta)
            
            # Contenedor individual para cada PDF
            item_frame = tk.Frame(self.frame_miniaturas, bd=1, relief="solid", bg="white")
            item_frame.pack(side=tk.LEFT, padx=10, pady=10)

            try:
                with pdfplumber.open(ruta) as pdf:
                    # Extraer la imagen de la primera página a baja resolución (para que cargue rápido)
                    img_cruda = pdf.pages[0].to_image(resolution=50).original
                    
                    # Recortar/Redimensionar a tamaño miniatura (100x140 píxeles aprox)
                    img_cruda.thumbnail((120, 160))
                    foto_tk = ImageTk.PhotoImage(img_cruda)
                    
                    # Guardar referencia para que el recolector de basura de Python no la borre
                    self.imagenes_referencia.append(foto_tk) 

                    lbl_img = tk.Label(item_frame, image=foto_tk, bg="white")
                    lbl_img.pack(padx=5, pady=5)
            except Exception as e:
                # Si un PDF está dañado y no se puede renderizar, muestra un cuadro gris
                lbl_img = tk.Label(item_frame, text="Sin vista\nprevia", width=15, height=8, bg="#e0e0e0")
                lbl_img.pack(padx=5, pady=5)

            # Etiqueta con el nombre del archivo (acortado si es muy largo)
            if len(nombre_archivo) > 18:
                nombre_corto = nombre_archivo[:12] + "..." + nombre_archivo[-4:]
            else:
                nombre_corto = nombre_archivo
                
            lbl_txt = tk.Label(item_frame, text=nombre_corto, font=("Arial", 8), bg="white")
            lbl_txt.pack(side=tk.BOTTOM, pady=(0, 5))

        self.root.config(cursor="") # Restaurar cursor

    def procesar_archivos(self):
        if not self.rutas_pdfs:
            return

        archivo_salida = filedialog.asksaveasfilename(
            title="Guardar Excel como...",
            defaultextension=".xlsx",
            initialfile="Resumen_Peajes.xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")]
        )

        if archivo_salida:
            try:
                self.root.config(cursor="watch")
                self.root.update()

                procesar_lista_pdfs(self.rutas_pdfs, archivo_salida)
                
                messagebox.showinfo("Éxito", f"¡Proceso completado!\n\nEl archivo se guardó en:\n{archivo_salida}")
                
                # Reiniciar estado de la app después del éxito
                self.rutas_pdfs = []
                self.lbl_estado.config(text="Ningún archivo seleccionado.", fg="gray")
                self.btn_procesar.config(state=tk.DISABLED)
                for widget in self.frame_miniaturas.winfo_children():
                    widget.destroy()
                self.imagenes_referencia.clear()

            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar las facturas:\n{str(e)}")
            finally:
                self.root.config(cursor="")

# Ejecución de la app

if __name__ == "__main__":
    ventana = tk.Tk()
    app = AppPeajes(ventana)
    ventana.mainloop()
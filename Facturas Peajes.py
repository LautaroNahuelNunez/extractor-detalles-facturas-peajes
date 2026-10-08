import tkinter as tk
from tkinter import filedialog, messagebox
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
        self.root.geometry("400x250")
        self.root.eval('tk::PlaceWindow . center') # Centrar en pantalla
        
        self.rutas_pdfs = []

        # Título
        tk.Label(root, text="Procesador de Peajes", font=("Arial", 14, "bold")).pack(pady=10)

        # Botón para cargar PDFs
        self.btn_cargar = tk.Button(root, text="1. Cargar facturas (PDF)", command=self.cargar_archivos, width=25, height=2)
        self.btn_cargar.pack(pady=10)

        # Etiqueta de estado
        self.lbl_estado = tk.Label(root, text="Ningún archivo seleccionado.", fg="gray")
        self.lbl_estado.pack(pady=5)

        # Botón para procesar y guardar
        self.btn_procesar = tk.Button(root, text="2. Procesar y Guardar Excel", command=self.procesar_archivos, width=25, height=2, state=tk.DISABLED)
        self.btn_procesar.pack(pady=10)

    def cargar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar facturas PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if archivos:
            self.rutas_pdfs = list(archivos)
            cantidad = len(self.rutas_pdfs)
            self.lbl_estado.config(text=f"{cantidad} archivo(s) seleccionado(s).", fg="green")
            self.btn_procesar.config(state=tk.NORMAL) # Habilitar el segundo botón

    def procesar_archivos(self):
        if not self.rutas_pdfs:
            return

        # Preguntar dónde guardar el archivo Excel
        archivo_salida = filedialog.asksaveasfilename(
            title="Guardar Excel como...",
            defaultextension=".xlsx",
            initialfile="Resumen_Peajes.xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")]
        )

        if archivo_salida:
            try:
                # Cambiar cursor a estado de carga
                self.root.config(cursor="watch")
                self.root.update()

                procesar_lista_pdfs(self.rutas_pdfs, archivo_salida)
                
                messagebox.showinfo("Éxito", f"¡Proceso completado!\n\nEl archivo se guardó en:\n{archivo_salida}")
                
                # Reiniciar estado
                self.rutas_pdfs = []
                self.lbl_estado.config(text="Ningún archivo seleccionado.", fg="gray")
                self.btn_procesar.config(state=tk.DISABLED)

            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al procesar las facturas:\n{str(e)}")
            finally:
                self.root.config(cursor="")

# Ejecución de la App
if __name__ == "__main__":
    ventana = tk.Tk()
    app = AppPeajes(ventana)
    ventana.mainloop()
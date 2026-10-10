"""
Extractor de facturas de peajes: lee los PDF y arma un Excel con el total por TAG,
más una hoja "Control" que compara lo extraído contra el subtotal de cada factura.

Dependencias: pdfplumber, openpyxl, Pillow, Tkinter
"""
from __future__ import annotations

import hashlib
import os
import queue
import re
import subprocess
import sys
import threading
from dataclasses import dataclass, field
from decimal import Decimal

import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import ImageTk

CENTAVO = Decimal('0.01')

# Las facturas redondean cada línea por separado, así que la suma de las líneas puede
# diferir unos centavos del subtotal impreso.
# Hasta este monto el control marca "OK (redondeo)"
TOLERANCIA_REDONDEO = Decimal('1.00') # Redondeo de $1 máximo

FORMATO_PESOS = '$ #,##0.00'


# Patrones

# Importe con 2 decimales, con o sin separador de miles ("1.234,56", "1234.56", "-98,59").
_NUM = r'-?(?:\d{1,3}(?:[.,]\d{3})+|\d+)[.,]\d{2}'

# No empieza ni termina pegado a otros dígitos/separadores: así no toma "10.93" de una
# fecha tipo 26.10.93 ni corta "1234.567" en "1234.56".
RE_IMPORTE = re.compile(rf'(?<![\d.,])({_NUM})(?!\d|[.,]\d)')

RE_TAG_PRESENTE = re.compile(r'(?<![A-Za-z0-9])(?:SI|S1)\d{5,}')

# TAG seguido (en la misma línea) de su importe. El TAG no puede ir pegado a otras letras/números,
# y entre el TAG y el importe no puede aparecer otro TAG: en las facturas de dos columnas, un TAG sin
# importe no debe quedarse con el de la columna de al lado.
_SIN_OTRO_TAG = r'(?:(?!(?<![A-Za-z0-9])(?:SI|S1)\d{5,}).)*?'
RE_TAG_IMPORTE = re.compile(rf'(?<![A-Za-z0-9])((?:SI|S1)\d+){_SIN_OTRO_TAG}(?<![\d.,])({_NUM})(?!\d|[.,]\d)')

# Formato agrupado: una línea con solo el TAG (y 0,00) y debajo sus pasadas.
RE_TAG_ENCABEZADO = re.compile(rf'\s*((?:SI|S1)\d+)(?:\s+({_NUM}))?\s*')

# Delimitan la tabla de ítems de cada página.
RE_ENCABEZADO_TABLA = re.compile(r'^\s*Cant(?:idad)?\b.*Descripci.n', re.IGNORECASE)
RE_TOTALES = re.compile(r'^\s*(?:Subtotal|Gravado|Neto)\b|\bTOTAL\b', re.IGNORECASE)
RE_PIE = re.compile(r'^\s*(?:Su saldo|Medios de pago|La presente|1er|2do|Saldo)\b', re.IGNORECASE)

RE_ETIQUETA_SUBTOTAL = re.compile(r'^\s*(?:Subtotal|Gravado)\b', re.IGNORECASE)
RE_GRAVADO_Y_NO_GRAVADO = re.compile(r'^\s*Gravado\s+No\s+Gravado', re.IGNORECASE)


# Limpieza

def limpiar_tag(tag_crudo):
    tag_limpio = tag_crudo.replace('S1', 'SI')  # Salva error de tipeo si se pone S1 en lugar de SI
    if tag_limpio.startswith('SI9000'):
        tag_limpio = tag_limpio.replace('SI9000', 'SI90', 1)  # Salva error de nomenclatura cuando rellenan con ceros
    return tag_limpio


def limpiar_importe(importe_crudo):
    # Devuelve un Decimal (con float se acumulan diferencias de centavos).
    texto = importe_crudo.strip()
    if ',' in texto and '.' in texto:  # Salva el uso de comas por punto para los decimales
        if texto.rfind(',') > texto.rfind('.'):
            texto = texto.replace('.', '').replace(',', '.')
        else:
            texto = texto.replace(',', '')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    return Decimal(texto)


def formato_ar(valor):
    # 1234.5 -> '1.234,50
    return f'{valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')



# Resultado de cada factura


@dataclass
class ResultadoPDF:
    archivo: str
    paginas: int = 0
    totales: dict = field(default_factory=dict)            # TAG -> Decimal
    sin_tag: list = field(default_factory=list)            # (descripción, Decimal): gastos, ajustes, etc.
    subtotal: Decimal | None = None                        # subtotal neto impreso en la factura
    paginas_sin_texto: list = field(default_factory=list)  # probablemente escaneadas
    ignoradas: list = field(default_factory=list)          # (página, línea) con TAG que no se pudo leer
    error: str | None = None

    @property
    def suma_tags(self):
        return sum(self.totales.values(), Decimal('0'))

    @property
    def suma_sin_tag(self):
        return sum((importe for _, importe in self.sin_tag), Decimal('0'))

    @property
    def diferencia(self):
        if self.subtotal is None:
            return None
        return self.subtotal - self.suma_tags - self.suma_sin_tag

    @property
    def estado(self):
        if self.error:
            return 'ERROR'
        if self.paginas_sin_texto or not self.totales:
            return 'REVISAR'
        if self.subtotal is None:
            return 'SIN CONTROL'
        diferencia = abs(self.diferencia)
        if diferencia <= CENTAVO:
            return 'OK'
        if diferencia <= TOLERANCIA_REDONDEO:
            return 'OK (redondeo)'
        return 'REVISAR'

    @property
    def avisos(self):
        avisos = []
        if self.error:
            avisos.append(self.error)
            return avisos
        if self.paginas_sin_texto:
            paginas = ', '.join(str(p) for p in self.paginas_sin_texto)
            avisos.append(f'Páginas sin texto (¿escaneadas?): {paginas}')
        if not self.totales:
            avisos.append('No se encontró ningún TAG')
        if self.subtotal is None:
            avisos.append('No se pudo leer el subtotal de la factura: sin control')
        elif abs(self.diferencia) > CENTAVO:
            avisos.append(f'Diferencia contra el subtotal: {formato_ar(self.diferencia)}')
        if self.sin_tag:
            avisos.append(f'Importes sin TAG: {formato_ar(self.suma_sin_tag)} en {len(self.sin_tag)} línea(s) '
                          f'(ver hoja "Sin TAG")')
        if self.ignoradas:
            ejemplo = self.ignoradas[0][1][:60]
            avisos.append(f'{len(self.ignoradas)} línea(s) con TAG que no se pudieron leer, ej.: "{ejemplo}"')
        return avisos



# Extracción


def _sumar(totales, tag, importe):
    totales[tag] = totales.get(tag, Decimal('0')) + importe


def _descripcion(texto):
    texto = re.sub(r'^\s*\d+\s+', '', texto)  # saca la cantidad del principio
    return ' '.join(texto.split())


def _procesar_linea_item(pagina, linea, res, grupo):
    """
    Procesa una línea de la tabla de ítems y devuelve el TAG "abierto" (o None).

    - Formato por línea (Oeste, AUSA, AUBASA): el TAG y su importe están en la misma línea.
      Puede haber dos ítems por línea (dos columnas).
    - Formato agrupado (Autopistas del Sol): una línea con el TAG y 0,00, y debajo sus pasadas
      sin TAG; esas se suman al último TAG abierto, también si siguen en la página siguiente.
    - Todo importe que no pertenece a un TAG (gastos administrativos, ajustes) va a res.sin_tag.
    """
    m = RE_TAG_ENCABEZADO.fullmatch(linea)
    if m:
        tag = limpiar_tag(m.group(1))
        _sumar(res.totales, tag, limpiar_importe(m.group(2)) if m.group(2) else Decimal('0'))
        return tag

    matches = list(RE_TAG_IMPORTE.finditer(linea))
    segmentos, pos = [], 0  # texto que queda fuera de los ítems con TAG
    for mt in matches:
        _sumar(res.totales, limpiar_tag(mt.group(1)), limpiar_importe(mt.group(2)))
        segmentos.append(linea[pos:mt.start()])
        pos = mt.end()
    segmentos.append(linea[pos:])

    sueltos = []
    for segmento in segmentos:
        desde = 0
        for mi in RE_IMPORTE.finditer(segmento):
            sueltos.append((_descripcion(segmento[desde:mi.start()]), limpiar_importe(mi.group(1))))
            desde = mi.end()

    if any(RE_TAG_PRESENTE.search(segmento) for segmento in segmentos):
        res.ignoradas.append((pagina, linea.strip()))  # hay un TAG al que no se le encontró importe

    if matches:
        res.sin_tag.extend(sueltos)  # ej.: "GASTOS ADMINISTRATIVOS 12.40" en la misma fila que un peaje
        return None

    if grupo and sueltos:
        _sumar(res.totales, grupo, sueltos[-1][1])  # si hubiera precio unitario y total, el total es el último
        return grupo

    res.sin_tag.extend(sueltos)
    return grupo


def leer_subtotal(textos):
    """
    Busca el subtotal neto (sin IVA ni percepciones) que imprime la factura. Las facturas
    traen una fila de rótulos (Subtotal / Gravado ...) y otra con los valores, arriba o abajo
    según la concesionaria. Devuelve None si no lo encuentra.
    """
    for texto in reversed(textos):  # los totales están en la última página
        if not texto:
            continue
        lineas = texto.split('\n')
        for i, linea in enumerate(lineas):
            if not RE_ETIQUETA_SUBTOTAL.match(linea):
                continue
            for j in (i + 1, i + 2, i - 1):
                if not 0 <= j < len(lineas):
                    continue
                valores = [limpiar_importe(x) for x in RE_IMPORTE.findall(lineas[j])]
                if len(valores) >= 4:  # la fila de valores tiene varias columnas
                    if RE_GRAVADO_Y_NO_GRAVADO.match(linea):
                        return valores[0] + valores[1]
                    return valores[0]
    return None


def analizar_textos(textos, res):
    # textos: lista con el texto de cada página (None si la página no tiene texto)
    res.paginas = len(textos)
    grupo = None  # TAG abierto (formato agrupado)

    for numero, texto in enumerate(textos, start=1):
        if not texto or not texto.strip():
            res.paginas_sin_texto.append(numero)
            continue

        en_tabla = False
        for linea in texto.split('\n'):
            if RE_ENCABEZADO_TABLA.search(linea):
                en_tabla = True
                continue

            if not en_tabla:
                if RE_TAG_PRESENTE.search(linea):
                    res.ignoradas.append((numero, linea.strip()))  # TAG fuera de la tabla: no se cuenta, se avisa
                continue

            if RE_TOTALES.search(linea) or len(RE_IMPORTE.findall(linea)) >= 4:
                en_tabla = False  # fila de totales
                grupo = None
                continue
            if RE_PIE.search(linea):
                en_tabla = False  # pie de página: el grupo sigue abierto por si continúa en la hoja siguiente
                continue

            grupo = _procesar_linea_item(numero, linea, res, grupo)

    res.subtotal = leer_subtotal(textos)


def procesar_pdf(ruta_pdf):
    # Procesa un PDF. Nunca levanta excepción: si falla, devuelve el resultado con .error
    res = ResultadoPDF(archivo=os.path.basename(ruta_pdf))
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            textos = [pagina.extract_text() for pagina in pdf.pages]
        analizar_textos(textos, res)
    except Exception as e:
        res = ResultadoPDF(archivo=res.archivo, error=f'No se pudo leer el PDF ({type(e).__name__}: {e})')
    return res



# Excel


COLORES_ESTADO = {
    'OK': 'C6EFCE',
    'OK (redondeo)': 'C6EFCE',
    'SIN CONTROL': 'FFEB9C',
    'REVISAR': 'FFC7CE',
    'ERROR': 'FFC7CE',
}


def _encabezado(hoja, anchos):
    for celda in hoja[1]:
        celda.font = Font(bold=True)
    for letra, ancho in anchos.items():
        hoja.column_dimensions[letra].width = ancho
    hoja.freeze_panes = 'A2'


def _importe_o_vacio(valor):
    return None if valor is None else float(valor.quantize(CENTAVO))


def guardar_excel(resultados, archivo_salida):
    wb = Workbook()

    # Hoja principal: total por TAG (entre todas las facturas)
    hoja = wb.active
    hoja.title = 'Peajes'
    totales = {}
    for r in resultados:
        for tag, importe in r.totales.items():
            _sumar(totales, tag, importe)
    hoja.append(['TAG', 'IMPORTE TOTAL'])
    for tag, importe in sorted(totales.items(), key=lambda kv: kv[1], reverse=True):
        hoja.append([tag, float(importe.quantize(CENTAVO))])
    for celda in hoja['B'][1:]:
        celda.number_format = FORMATO_PESOS
    _encabezado(hoja, {'A': 18, 'B': 20})

    # Control: una fila por factura
    control = wb.create_sheet('Control')
    control.append(['Archivo', 'Páginas', 'Subtotal factura', 'Suma por TAG', 'Sin TAG', 'Diferencia', 'Estado', 'Avisos'])
    for r in resultados:
        control.append([
            r.archivo,
            r.paginas or None,
            _importe_o_vacio(r.subtotal),
            _importe_o_vacio(r.suma_tags) if not r.error else None,
            _importe_o_vacio(r.suma_sin_tag) if not r.error else None,
            _importe_o_vacio(r.diferencia) if r.diferencia is not None else None,
            r.estado,
            ' | '.join(r.avisos),
        ])
        fila = control.max_row
        for columna in 'CDEF':
            control[f'{columna}{fila}'].number_format = FORMATO_PESOS
        control[f'G{fila}'].fill = PatternFill('solid', fgColor=COLORES_ESTADO[r.estado])
        control[f'H{fila}'].alignment = Alignment(wrap_text=True, vertical='top')
    _encabezado(control, {'A': 34, 'B': 9, 'C': 18, 'D': 18, 'E': 14, 'F': 14, 'G': 16, 'H': 70})

    # Detalle de lo que no tiene TAG (gastos administrativos, ajustes, bonificaciones...)
    if any(r.sin_tag for r in resultados):
        detalle = wb.create_sheet('Sin TAG')
        detalle.append(['Archivo', 'Descripción', 'Importe'])
        for r in resultados:
            for descripcion, importe in r.sin_tag:
                detalle.append([r.archivo, descripcion, float(importe)])
                detalle[f'C{detalle.max_row}'].number_format = FORMATO_PESOS
        _encabezado(detalle, {'A': 34, 'B': 50, 'C': 16})

    wb.save(archivo_salida)


def abrir_archivo(ruta):
    try:
        if sys.platform.startswith('win'):
            os.startfile(ruta)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', ruta])
        else:
            subprocess.Popen(['xdg-open', ruta])
    except Exception:
        pass



# Interfaz Gráfica


COLOR_FONDO = 'white'
COLOR_SELECCION = '#cfe3ff'
COLOR_BORDE_SELECCION = '#1a73e8'
COLOR_PLACEHOLDER = '#e0e0e0'
COLOR_AVISO = '#b45309'


def huella_archivo(ruta):
    # Hash del contenido: dos PDF con el mismo contenido son la misma factura aunque tengan distinto nombre
    try:
        h = hashlib.sha256()
        with open(ruta, 'rb') as f:
            for bloque in iter(lambda: f.read(1 << 20), b''):
                h.update(bloque)
        return h.hexdigest()
    except OSError:
        return None


@dataclass(eq=False)
class ItemPDF:
    # Un PDF de la lista: su ruta, su miniatura (widgets de Tk) y su estado
    ruta: str
    huella: str | None
    frame: object = None
    lbl_img: object = None
    lbl_aviso: object = None
    lbl_txt: object = None
    foto: object = None             # Referencia a la imagen: si se pierde, Tkinter la borra de la pantalla
    seleccionado: bool = False
    repetido: bool = False
    sin_tags: bool | None = None    # None = todavía no se revisó; True = ningún TAG en todo el PDF


class AppPeajes:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Facturas de Peajes")
        self.root.geometry("640x560")  # Ventana más grande para las miniaturas
        self.root.minsize(600, 520)
        self.root.eval('tk::PlaceWindow . center')  # Centrado de ventana

        self.items = []                # PDFs cargados, en el orden en que se cargaron
        self.cola = queue.Queue()      # El hilo de trabajo avisa por acá (Tk no se toca desde otro hilo)
        self._procesando = False
        self._aviso_carga = ''
        self._generacion_miniaturas = 0
        self._pendientes_miniaturas = []
        self._dibujando = False

        # Título
        tk.Label(root, text="Procesador de Peajes", font=("Arial", 14, "bold")).pack(pady=10)

        # Frame para alinear los botones
        frame_botones = tk.Frame(root)
        frame_botones.pack(pady=5)

        self.btn_cargar = tk.Button(frame_botones, text="1. Cargar facturas (PDF)", command=self.cargar_archivos, width=22, height=2)
        self.btn_cargar.pack(side=tk.LEFT, padx=10)

        self.btn_procesar = tk.Button(frame_botones, text="2. Procesar y Guardar Excel", command=self.procesar_archivos, width=22, height=2, state=tk.DISABLED)
        self.btn_procesar.pack(side=tk.LEFT, padx=10)

        self.lbl_estado = tk.Label(root, text="Ningún archivo seleccionado.", fg="gray", wraplength=580, justify="center")
        self.lbl_estado.pack(pady=5)

        self.barra = ttk.Progressbar(root, mode='determinate', length=360)
        self.barra.pack(pady=(0, 5))

        # Acciones sobre las miniaturas
        frame_acciones = tk.Frame(root)
        frame_acciones.pack(pady=(5, 0))

        self.btn_sel_todo = tk.Button(frame_acciones, text="Seleccionar todo", command=self.alternar_todo, width=18, state=tk.DISABLED)
        self.btn_sel_todo.pack(side=tk.LEFT, padx=5)

        self.btn_sel_rep = tk.Button(frame_acciones, text="Seleccionar repetidos", command=self.seleccionar_repetidos, width=20, state=tk.DISABLED)
        self.btn_sel_rep.pack(side=tk.LEFT, padx=5)

        self.btn_quitar = tk.Button(frame_acciones, text="Quitar seleccionados (0)", command=self.quitar_seleccionados, width=24, state=tk.DISABLED)
        self.btn_quitar.pack(side=tk.LEFT, padx=5)

        tk.Label(root, text='Clic en una miniatura para marcarla y después "Quitar seleccionados" (o tecla Supr). '
                            'Quitar solo la saca de la lista: el PDF no se borra de tu disco.',
                 fg="gray", font=("Arial", 8), wraplength=580).pack(pady=(3, 0))

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

        # Atajos de teclado
        self.root.bind('<Delete>', lambda e: self.quitar_seleccionados())
        self.root.bind('<Control-a>', lambda e: self.seleccionar_todo(True))
        self.root.bind('<Control-A>', lambda e: self.seleccionar_todo(True))

    @property
    def rutas_pdfs(self):
        return [it.ruta for it in self.items]

    #  Lista de PDF y miniaturas 

    def cargar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar facturas PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if not archivos:
            return

        # Los archivos se suman a los que ya están cargados; el mismo archivo no se agrega dos veces
        cargados = {os.path.normcase(os.path.abspath(it.ruta)) for it in self.items}
        nuevos, ya_estaban = [], 0
        for ruta in archivos:
            clave = os.path.normcase(os.path.abspath(ruta))
            if clave in cargados:
                ya_estaban += 1
                continue
            cargados.add(clave)
            nuevos.append(self._agregar_item(ruta))

        self._aviso_carga = ''
        if ya_estaban:
            self._aviso_carga = f"{ya_estaban} ya estaba en la lista" if ya_estaban == 1 else f"{ya_estaban} ya estaban en la lista"
        self.barra['value'] = 0
        self._refrescar_estado()

        self._pendientes_miniaturas.extend(nuevos)
        if nuevos and not self._dibujando:
            self._dibujando = True
            # Una miniatura por vuelta del loop de Tk: la ventana no se congela con muchos archivos
            self.root.after(10, self._siguiente_miniatura, self._generacion_miniaturas)

    def _agregar_item(self, ruta):
        item = ItemPDF(ruta=ruta, huella=huella_archivo(ruta))
        nombre_archivo = os.path.basename(ruta)

        # Contenedor individual para cada PDF (el borde grueso cambia de color al seleccionarlo)
        item.frame = tk.Frame(self.frame_miniaturas, bd=1, relief="solid", bg=COLOR_FONDO,
                              highlightthickness=3, highlightbackground=COLOR_FONDO, cursor="hand2")
        item.frame.pack(side=tk.LEFT, padx=6, pady=6)

        # Etiqueta con el nombre del archivo (acortado si es muy largo)
        if len(nombre_archivo) > 18:
            nombre_corto = nombre_archivo[:12] + "..." + nombre_archivo[-4:]
        else:
            nombre_corto = nombre_archivo

        item.lbl_txt = tk.Label(item.frame, text=nombre_corto, font=("Arial", 8), bg=COLOR_FONDO)
        item.lbl_aviso = tk.Label(item.frame, text="", height=2, font=("Arial", 8, "bold"), fg=COLOR_AVISO, bg=COLOR_FONDO)
        item.lbl_img = tk.Label(item.frame, text="Cargando...", width=15, height=8, bg=COLOR_PLACEHOLDER)

        item.lbl_txt.pack(side=tk.BOTTOM, pady=(0, 5))
        item.lbl_aviso.pack(side=tk.BOTTOM)
        item.lbl_img.pack(padx=5, pady=5)

        for widget in (item.frame, item.lbl_txt, item.lbl_aviso, item.lbl_img):
            widget.bind("<Button-1>", lambda e, it=item: self._alternar(it))
        self.items.append(item)
        return item

    def _siguiente_miniatura(self, generacion):
        if generacion != self._generacion_miniaturas:
            return  # la lista se vació mientras tanto
        if not self._pendientes_miniaturas:
            self._dibujando = False
            return
        self._dibujar_miniatura(self._pendientes_miniaturas.pop(0))
        self.root.after(10, self._siguiente_miniatura, generacion)

    def _dibujar_miniatura(self, item):
        try:
            with pdfplumber.open(item.ruta) as pdf:
                # Extraer la imagen de la primera página a baja resolución (para que cargue rápido)
                img_cruda = pdf.pages[0].to_image(resolution=50).original

                # Recortar/Redimensionar a tamaño miniatura (100x140 píxeles aprox)
                img_cruda.thumbnail((120, 160))
                item.foto = ImageTk.PhotoImage(img_cruda)
                item.lbl_img.config(image=item.foto, text="", width=0, height=0)

                # Si en todo el PDF no hay ningún TAG, seguramente no es una factura de peaje
                item.sin_tags = not any(RE_TAG_PRESENTE.search(pagina.extract_text() or '') for pagina in pdf.pages)
        except Exception:
            # Si un PDF está dañado y no se puede renderizar, muestra un cuadro gris
            item.lbl_img.config(text="Sin vista\nprevia")
        self._refrescar_estado()

    def _pintar(self, item):
        fondo = COLOR_SELECCION if item.seleccionado else COLOR_FONDO
        item.frame.config(bg=fondo, highlightbackground=COLOR_BORDE_SELECCION if item.seleccionado else COLOR_FONDO)
        item.lbl_txt.config(bg=fondo)
        avisos = []
        if item.repetido:
            avisos.append("REPETIDO")
        if item.sin_tags:
            avisos.append("NO PARECE FACTURA")
        item.lbl_aviso.config(bg=fondo, text="\n".join(avisos))
        if item.foto is not None:
            item.lbl_img.config(bg=fondo)

    def _refrescar_estado(self):
        # Recalcula repetidos, repinta las miniaturas y actualiza el texto y los botones.
        vistos = set()
        for it in self.items:
            # Repetido = mismo contenido que un PDF anterior de la lista (el primero no se marca)
            it.repetido = it.huella is not None and it.huella in vistos
            if it.huella is not None:
                vistos.add(it.huella)
            self._pintar(it)

        if self._procesando:
            return  # mientras se procesa, el texto y los botones los maneja el proceso

        n = len(self.items)
        repetidos = sum(it.repetido for it in self.items)
        no_factura = sum(bool(it.sin_tags) for it in self.items)
        seleccionados = sum(it.seleccionado for it in self.items)

        if n == 0:
            texto, color = "Ningún archivo seleccionado.", "gray"
        else:
            partes = [f"{n} archivo(s) cargado(s)"]
            if repetidos:
                partes.append(f"{repetidos} repetido(s)")
            if no_factura:
                partes.append(f"{no_factura} no parece(n) factura")
            if self._aviso_carga:
                partes.append(self._aviso_carga)
            texto = " · ".join(partes)
            color = COLOR_AVISO if (repetidos or no_factura or self._aviso_carga) else "green"
        self.lbl_estado.config(text=texto, fg=color)

        self.btn_procesar.config(state=tk.NORMAL if n else tk.DISABLED)
        self.btn_sel_todo.config(state=tk.NORMAL if n else tk.DISABLED,
                                 text="Deseleccionar todo" if n and seleccionados == n else "Seleccionar todo")
        self.btn_sel_rep.config(state=tk.NORMAL if repetidos else tk.DISABLED)
        self.btn_quitar.config(text=f"Quitar seleccionados ({seleccionados})",
                               state=tk.NORMAL if seleccionados else tk.DISABLED)

    #  Selección y borrado 

    def _alternar(self, item):
        if self._procesando:
            return
        item.seleccionado = not item.seleccionado
        self._refrescar_estado()

    def seleccionar_todo(self, valor=True):
        if self._procesando:
            return
        for it in self.items:
            it.seleccionado = valor
        self._refrescar_estado()

    def alternar_todo(self):
        self.seleccionar_todo(not all(it.seleccionado for it in self.items))

    def seleccionar_repetidos(self):
        if self._procesando:
            return
        for it in self.items:
            it.seleccionado = it.repetido
        self._refrescar_estado()

    def quitar_seleccionados(self):
        if self._procesando:
            return
        quitar = [it for it in self.items if it.seleccionado]
        if not quitar:
            return
        for it in quitar:
            it.frame.destroy()
        self.items = [it for it in self.items if not it.seleccionado]
        self._pendientes_miniaturas = [it for it in self._pendientes_miniaturas if it in self.items]
        self._aviso_carga = ''
        self._refrescar_estado()  # recalcula también los repetidos: si se quitó el original, el otro deja de serlo

    def limpiar_items(self):
        self._generacion_miniaturas += 1  # cancela las miniaturas que estaban pendientes
        self._pendientes_miniaturas = []
        self._dibujando = False
        for it in self.items:
            it.frame.destroy()
        self.items = []
        self._aviso_carga = ''

    # Procesamiento (en un hilo aparte para que la ventana no se congele) 

    def procesar_archivos(self):
        if not self.items or self._procesando:
            return

        # Un PDF repetido se sumaría dos veces: antes de seguir, el usuario tiene que confirmarlo
        repetidos = [it for it in self.items if it.repetido]
        if repetidos:
            nombres = "\n".join(f"  • {os.path.basename(it.ruta)}" for it in repetidos[:8])
            if len(repetidos) > 8:
                nombres += "\n  • ..."
            seguir = messagebox.askyesno(
                "PDF repetidos",
                f"Hay {len(repetidos)} PDF repetido(s) en la lista y sus importes se sumarían dos veces:\n\n{nombres}\n\n"
                "¿Procesar igual?",
                icon="warning", default="no")
            if not seguir:
                return

        archivo_salida = filedialog.asksaveasfilename(
            title="Guardar Excel como...",
            defaultextension=".xlsx",
            initialfile="Resumen_Peajes.xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")]
        )
        if not archivo_salida:
            return

        self._ocupado(True)
        self.barra['value'] = 0
        hilo = threading.Thread(target=self._trabajo, args=(self.rutas_pdfs, archivo_salida), daemon=True)
        hilo.start()
        self.root.after(100, self._revisar_cola)

    def _trabajo(self, rutas, archivo_salida):
        # Corre en el hilo de trabajo: no toca widgets, solo manda mensajes por la cola
        try:
            resultados = []
            for i, ruta in enumerate(rutas):
                self.cola.put(('progreso', i, len(rutas),
                               f'Procesando {os.path.basename(ruta)} ({i + 1}/{len(rutas)})...'))
                resultados.append(procesar_pdf(ruta))
            self.cola.put(('progreso', len(rutas), len(rutas), 'Guardando Excel...'))
            guardar_excel(resultados, archivo_salida)
            self.cola.put(('fin', resultados, archivo_salida))
        except PermissionError:
            self.cola.put(('error', 'No se pudo guardar el Excel. Si ese archivo está abierto, cerralo y probá de nuevo.'))
        except Exception as e:
            self.cola.put(('error', str(e)))

    def _revisar_cola(self):
        try:
            while True:
                mensaje = self.cola.get_nowait()
                if mensaje[0] == 'progreso':
                    _, hechos, total, texto = mensaje
                    self.barra['maximum'] = total
                    self.barra['value'] = hechos
                    self.lbl_estado.config(text=texto, fg="black")
                elif mensaje[0] == 'fin':
                    self._terminar_ok(mensaje[1], mensaje[2])
                    return
                elif mensaje[0] == 'error':
                    self._terminar_con_error(mensaje[1])
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._revisar_cola)

    def _ocupado(self, ocupado):
        self._procesando = ocupado
        self.root.config(cursor="watch" if ocupado else "")
        if ocupado:
            for boton in (self.btn_cargar, self.btn_procesar, self.btn_sel_todo, self.btn_sel_rep, self.btn_quitar):
                boton.config(state=tk.DISABLED)
        else:
            self.btn_cargar.config(state=tk.NORMAL)
            self._refrescar_estado()  # deja el resto de los botones según lo que haya en la lista

    def _terminar_ok(self, resultados, archivo_salida):
        self._ocupado(False)
        bien = [r for r in resultados if r.estado.startswith('OK')]
        revisar = [r for r in resultados if not r.estado.startswith('OK')]

        lineas = [f'Se procesaron {len(resultados)} factura(s): {len(bien)} coinciden con su subtotal.']
        if revisar:
            lineas += ['', 'Para revisar (el detalle está en la hoja "Control"):']
            lineas += [f'  • {r.archivo}: {r.estado}' for r in revisar]
        lineas += ['', f'El archivo se guardó en:\n{archivo_salida}', '', '¿Abrir el Excel ahora?']

        abrir = messagebox.askyesno("Proceso completado", "\n".join(lineas),
                                    icon="warning" if revisar else "info")
        if abrir:
            abrir_archivo(archivo_salida)

        # Reiniciar estado de la app después del éxito
        self.limpiar_items()
        self.barra['value'] = 0
        self._refrescar_estado()

    def _terminar_con_error(self, texto):
        self._ocupado(False)  # los archivos cargados quedan, para poder reintentar
        messagebox.showerror("Error", f"Ocurrió un error al procesar las facturas:\n{texto}")


# Ejecución de la app

if __name__ == "__main__":
    ventana = tk.Tk()
    app = AppPeajes(ventana)
    ventana.mainloop()
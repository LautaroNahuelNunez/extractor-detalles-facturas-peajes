# extractor-detalles-facturas-peajes
Extractor de detalles de importes de facturas de pasadas de peajes.

## Extractor Automático de Facturas de Peajes

Una pequeña app de escritorio desarrollada en Python para automatizar la extracción, limpieza y consolidación de datos provenientes de facturas de peajes en formato PDF (AUSA, AUBASA, Autopistas del Sol y Autopistas del Oeste). 
Este mini proyecto (desarrollado para simplificar una tarea en una empresa que, hecha a mano, demoraba aproximadamente 15/20 minutos por cada factura) elimina la necesidad de carga y control manual, resolviendo inconsistencias de formato entre diferentes concesionarias y exportando un reporte con los datos solicitados (TAG + suma de todos los importes) limpio en Excel para su futuro uso analítico.

## Características Principales

*   **Extracción de texto desde PDF:** Utiliza `pdfplumber` para leer estructuras tabulares y texto plano de facturas de múltiples páginas.
*   **Limpieza de Datos (Data Cleaning) y Regex:** 
    *   Normaliza las nomenclaturas de los TAGs (corrige errores de lectura OCR como `S1` por `SI` y unifica formatos largos que usa AUSA).
    *   Parsea importes financieros lidiando con múltiples formatos de separadores de miles y decimales (comas y puntos).
*   **Manejo de Casos Complejos:** 
    *   Suma de pasadas normales y pasadas en "hora pico".
    *   Procesamiento correcto de bonificaciones (valores negativos).
    *   Asignación de importes a TAGs mediante memoria de contexto (para facturas que agrupan varias líneas bajo un mismo TAG sin repetirlo).
*   **Interfaz Gráfica (GUI):** Interfaz simple e intuitiva construida con `tkinter` para que cualquier usuario pueda operar la herramienta sin tocar el código.
*   **Exportación a Excel:** Agrupa los importes totales por TAG y genera un archivo `.xlsx` con formato contable automático usando `pandas` y `openpyxl`.

## Tecnologías Utilizadas

*   **Python 3.12.10**
*   **Pandas:** Agrupación y estructuración de datos.
*   **pdfplumber:** Extracción de texto de documentos PDF.
*   **Re (Expresiones Regulares):** Búsqueda de patrones para capturar TAGs e importes.
*   **Tkinter:** Interfaz gráfica de usuario.
*   **Openpyxl:** Motor de escritura y formateo para Excel.

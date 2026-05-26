# Conciliador de Informes 264 vs 260

Aplicación en Streamlit para consolidar informes 264 de varias sedes y reemplazar las fechas usando los informes 260.

## ¿Qué hace?

- Carga múltiples archivos 264.
- Carga múltiples archivos 260.
- Extrae del 260 la fecha de la primera columna.
- Cruza los archivos por Doc Paciente.
- Reemplaza la fecha del 264 con la fecha encontrada en el 260.
- Agrupa cada factura / Doc Paciente en una sola fila.
- Genera un reporte final en Excel sin filas vacías.

## Archivos requeridos

La app permite cargar:

- Informes 264 en formato `.xls` o `.xlsx`
- Informes 260 en formato `.xls` o `.xlsx`

## Instalación local

```bash
pip install -r requirements.txt

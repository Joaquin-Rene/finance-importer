# Automatizacion de Finanzas Personales

Aplicacion local en Streamlit para importar, revisar y consolidar movimientos financieros en `FINANZAS.xlsx` desde dos fuentes:
- exportes Excel de Mercado Pago
- capturas bancarias con carga manual asistida y OCR opcional

El objetivo del proyecto es reducir carga manual, evitar duplicados y dejar una trazabilidad clara antes de escribir en el workbook final.

## Que resuelve

- normaliza movimientos a un esquema comun
- detecta duplicados por referencia y por clave compuesta
- genera backup antes de modificar el archivo
- muestra preview, metricas de control e insights mensuales
- permite trabajar con assets demo para mostrar el flujo sin usar datos reales

## Alcance actual

- `Mercado Pago (Excel)`: importa `.xlsx` o `.xls` exportados desde MP
- `Captura bancaria`: toma `.jpg`, `.jpeg` o `.png` y permite completar movimientos en una tabla editable
- `FINANZAS.xlsx`: escribe en hoja `Control de ingresos y gastos`, tabla `tblControlIngresosGastos`

## Requisitos

- Python 3.10+
- dependencias instaladas desde `requirements.txt`
- para `Captura bancaria`: Tesseract OCR instalado en el sistema si queres usar precarga OCR

## Instalacion

```bash
pip install -r requirements.txt
```

## Ejecutar

```bash
streamlit run app.py
```

## Flujo de uso

1. Completar `FINANZAS_PATH` con la ruta local del workbook.
2. Elegir modo de importacion.
3. Subir archivo Excel o capturas bancarias.
4. Revisar preview, duplicados y metricas.
5. Confirmar importacion.

## Privacidad y demo

El repo incluye assets seguros para demostracion:
- `mercado_pago_demo.xlsx`
- `FINANZAS_demo.xlsx`

Estos archivos estan sintetizados para mostrar el flujo completo sin exponer datos personales reales.

Las capturas bancarias reales u originales pueden servir para pruebas locales, pero no forman parte del material publicable del repo.

Si `FINANZAS_demo.xlsx` existe en la carpeta del proyecto, la app lo sugiere por defecto para reducir riesgo de escritura sobre un archivo personal.

## Regenerar assets demo

```bash
.\.venv\Scripts\python.exe scripts\generate_demo_assets.py
```

Esto recrea:
- un export demo de Mercado Pago con movimientos de enero y febrero de 2026
- un `FINANZAS_demo.xlsx` con historico previo para probar importacion, dedupe e insights

Para el flujo bancario del README o portfolio, usar capturas de la app basadas en material demo o suficientemente anonimizado.

## Capturas sugeridas

Las capturas publicables quedan en `docs/screenshots/`:
- [`01-home.png`](docs/screenshots/01-home.png)
- [`02-excel-upload.png`](docs/screenshots/02-excel-upload.png)
- [`03-review-import.png`](docs/screenshots/03-review-import.png)
- [`04-bank-capture-editor.png`](docs/screenshots/04-bank-capture-editor.png)
- [`05-bank-capture-preview.png`](docs/screenshots/05-bank-capture-preview.png)
- [`06-month-insights.png`](docs/screenshots/06-month-insights.png)

## Ejecutar con doble click (Windows)

- `run_app.bat`: inicia la app sin abrir navegador automaticamente
- `open_app.url`: abre `http://127.0.0.1:8501`
- para detener la app: cerrar la consola o presionar `Ctrl+C`

## Importacion

La importacion:
- detecta la fila de encabezado del export de Mercado Pago
- convierte los movimientos a un esquema estandar
- agrega filas nuevas al final de `tblControlIngresosGastos`
- preserva formato y formula de mes
- genera backup con nombre `FINANZAS.YYYYMMDD_HHMMSS.bak.xlsx`

## Estructura

- `app.py`
- `docs/screenshots/`
- `scripts/generate_demo_assets.py`
- `src/finanzas_importer/mp_parser.py`
- `src/finanzas_importer/bna_image_parser.py`
- `src/finanzas_importer/workbook_writer.py`
- `src/finanzas_importer/analytics.py`
- `src/finanzas_importer/ui_components.py`
- `src/finanzas_importer/utils.py`

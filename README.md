# Dashboard Encuesta de Satisfacción - Campamento 5400

Aplicación Flask lista para Render.com. Permite importar la pestaña **Encuesta de Satisfacción** desde un archivo Excel `.xlsx`, visualizar un dashboard ejecutivo y exportar los datos en PDF, Excel y CSV.

## Corrección importante de despliegue

Esta versión elimina `pandas`, `numpy`, `matplotlib` y `reportlab` para evitar el error de compilación en Render cuando usa Python 3.14. La lectura y exportación Excel se hacen con `openpyxl`; el PDF se genera internamente sin dependencias compiladas.

## Archivos principales

- `app.py`: aplicación Flask.
- `requirements.txt`: dependencias livianas compatibles con Render.
- `render.yaml`: configuración para Render.
- `.python-version` y `runtime.txt`: fuerzan Python 3.12.7.
- `static/img/logo_aramark.png`: logo Aramark.
- `static/img/logo_campamento_5400.png`: logo Escondida | BHP.

## Uso local

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abrir: http://127.0.0.1:5000

## Despliegue en Render

1. Sube todos los archivos del proyecto a la raíz del repositorio GitHub.
2. En Render crea un **New Web Service** desde el repositorio.
3. Verifica:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
4. En Render ejecuta **Manual Deploy → Clear build cache & deploy**.

## Estructura esperada del Excel

La app busca automáticamente una hoja que contenga “Encuesta” y “Satisf” en el nombre. Las columnas requeridas son:

- FECHA
- Q1_RESPUESTA, Q1_PUNTAJE
- Q2_RESPUESTA, Q2_PUNTAJE
- Q3_RESPUESTA, Q3_PUNTAJE
- Q4_RESPUESTA, Q4_PUNTAJE
- Q5_RESPUESTA, Q5_PUNTAJE
- TOTAL
- PROMEDIO
- COMENTARIOS

También puedes descargar una plantilla desde la pantalla principal.
## Corrección UTF-8

Esta versión fuerza `charset=utf-8` en las respuestas HTML/CSV/JSON para que palabras con tildes y ñ se vean correctamente en Render. También mantiene el PDF con codificación WinAnsi para preservar acentos en el exportable.


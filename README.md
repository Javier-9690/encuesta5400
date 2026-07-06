# Dashboard Encuesta de Satisfacción - Campamento 5400

Aplicación Flask lista para publicar en Render. Permite importar el archivo Excel del KPI, leer automáticamente la pestaña **Encuesta de Satisfacción**, calcular métricas ejecutivas y exportar el resultado a PDF, Excel o CSV.


## Identidad visual

Esta versión incluye los logos extraídos del reporte Aramark Campamento 5400 y una línea gráfica roja aplicada al encabezado, botones, tarjetas KPI, secciones, gráficos y PDF exportable.

## Funciones principales

- Importación de Excel `.xlsx` / `.xls`.
- Detección automática de la pestaña `Encuesta de Satisfacción`.
- Modo de importación: reemplazar histórico o agregar al histórico.
- Dashboard ejecutivo:
  - Total de evaluaciones.
  - Satisfacción global.
  - Índice de excelencia.
  - Distribución de promotores, neutros y detractores.
  - Comentarios recientes.
  - Análisis cualitativo automático.
  - Evolución semanal de las últimas 10 semanas.
  - Resumen de semanas recientes.
  - Desempeño por dimensiones Q1 a Q5.
- Exportación:
  - PDF ejecutivo de 3 páginas.
  - Excel multipestaña.
  - CSV histórico.
- Descarga de plantilla Excel.

## Columnas esperadas

La pestaña importada debe contener estas columnas:

```text
FECHA
Q1_RESPUESTA
Q1_PUNTAJE
Q2_RESPUESTA
Q2_PUNTAJE
Q3_RESPUESTA
Q3_PUNTAJE
Q4_RESPUESTA
Q4_PUNTAJE
Q5_RESPUESTA
Q5_PUNTAJE
TOTAL
PROMEDIO
COMENTARIOS
```

## Reglas de clasificación

- Promotores: promedio entre 4.8 y 5.0.
- Neutros: promedio entre 4.0 y 4.7.
- Detractores: promedio entre 1.0 y 3.9.
- Mantener estándar: promedio semanal mayor o igual a 4.80.
- Seguimiento preventivo: promedio semanal mayor o igual a 4.70 y menor a 4.80.
- Generar plan de acción: promedio semanal menor a 4.70.

## Ejecución local

```bash
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac
pip install -r requirements.txt
python app.py
```

Abrir en el navegador:

```text
http://localhost:5000
```

## Publicación en Render

1. Sube esta carpeta a un repositorio de GitHub.
2. En Render, crea un nuevo **Web Service** desde ese repositorio.
3. Render detectará el archivo `render.yaml`.
4. Comando de construcción: `pip install -r requirements.txt`.
5. Comando de inicio: `gunicorn app:app`.
6. La base SQLite queda configurada en `/var/data/encuesta_5400.db` mediante disco persistente.

## Nota operacional

El archivo PDF exportado se genera con los datos ya importados en la aplicación. Si se sube un nuevo Excel en modo "Reemplazar", se elimina el histórico anterior y el dashboard queda construido solo con el nuevo archivo.

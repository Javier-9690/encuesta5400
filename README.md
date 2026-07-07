# Dashboard Encuesta de Satisfacción - Campamento 5400

Aplicación web Flask lista para Render.com, sin pandas/numpy/matplotlib/reportlab para evitar problemas de compilación en Render.

## Funciones principales

- Importa archivos `.xlsx` desde la pestaña **Encuesta de Satisfacción**.
- Mantiene estructura semanal previamente definida: **lunes a domingo con numeración ISO**.
- Calcula reporte histórico completo.
- Permite seleccionar **fecha de inicio** y **fecha final** para mostrar un reporte por rango.
- Muestra en el dashboard el histórico y, cuando se filtra, también el reporte por fecha seleccionada.
- Exporta PDF con reporte histórico y reporte por fecha seleccionada si existe filtro activo.
- Exporta Excel con hojas históricas y hojas filtradas si existe filtro activo.
- Exporta CSV histórico.
- Incluye logos Aramark y Escondida | BHP.
- Estilo corporativo rojo.

## Regla de cumplimiento

El cumplimiento se calcula así:

- Si el promedio de notas es **mayor o igual a 4.0**, el cumplimiento es **100%**.
- Si el promedio es menor a 4.0, se calcula proporcionalmente contra el umbral 4.0.

Ejemplo: promedio 3.6 = 90% de cumplimiento.

## Despliegue en Render

1. Subir todos los archivos de esta carpeta a la raíz de un repositorio GitHub.
2. Crear un Web Service en Render.
3. Build Command:

```bash
pip install -r requirements.txt
```

4. Start Command:

```bash
gunicorn app:app
```

5. Si ya existe un despliegue anterior, usar:

**Manual Deploy → Clear build cache & deploy**

## Persistencia

La app usa SQLite. En Render se configura un disco persistente en `/var/data` mediante `render.yaml`.


## Distribución de cumplimiento operativo

La antigua tabla de promotores/neutros/detractores fue reemplazada por una medición directa de cumplimiento:

- **Cumple estándar operativo:** promedio >= 4.0.
- **No cumple estándar operativo:** promedio < 4.0.

La estructura semanal se mantiene como lunes a domingo, con numeración ISO, igual que en la hoja KPI Semanal.

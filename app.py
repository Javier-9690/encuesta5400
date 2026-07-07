# -*- coding: utf-8 -*-
import csv
import io
import json
import math
import os
import sqlite3
import zlib
from collections import defaultdict
from datetime import datetime, timedelta
from html import escape as html_escape
from statistics import mean

from flask import Flask, Response, flash, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook, load_workbook
from werkzeug.utils import secure_filename

APP_TITLE = "Dashboard Encuesta de Satisfacción - Campamento 5400"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
LOGO_ARAMARK_PATH = os.path.join(STATIC_DIR, "img", "logo_aramark.png")
LOGO_BHP_PATH = os.path.join(STATIC_DIR, "img", "logo_campamento_5400.png")
DB_PATH = os.environ.get("DATABASE_PATH", "/var/data/encuesta_5400.db")
if not os.path.exists(os.path.dirname(DB_PATH)):
    DB_PATH = os.path.join(BASE_DIR, "encuesta_5400.db")

EXPECTED_COLUMNS = [
    "FECHA",
    "Q1_RESPUESTA", "Q1_PUNTAJE",
    "Q2_RESPUESTA", "Q2_PUNTAJE",
    "Q3_RESPUESTA", "Q3_PUNTAJE",
    "Q4_RESPUESTA", "Q4_PUNTAJE",
    "Q5_RESPUESTA", "Q5_PUNTAJE",
    "TOTAL", "PROMEDIO", "COMENTARIOS",
]

DB_COLUMNS = [
    "fecha", "q1_respuesta", "q1_puntaje", "q2_respuesta", "q2_puntaje",
    "q3_respuesta", "q3_puntaje", "q4_respuesta", "q4_puntaje",
    "q5_respuesta", "q5_puntaje", "total", "promedio", "comentarios",
]

DIMENSION_LABELS = {
    "q1_puntaje": "Q1. Primera Impresión / Recepción",
    "q2_puntaje": "Q2. Calidad del Serv. Principal",
    "q3_puntaje": "Q3. Tiempos de Respuesta",
    "q4_puntaje": "Q4. Higiene y Presentación",
    "q5_puntaje": "Q5. Trato y Factor Humano",
}

RADAR_LABELS = ["Q1 Recepción", "Q2 Calidad", "Q3 Tiempos", "Q4 Higiene", "Q5 Trato"]
RED = "#ed1b2e"
RED_DARK = "#b60f1f"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cambiar-esta-clave-en-render")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS encuestas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                q1_respuesta TEXT,
                q1_puntaje REAL,
                q2_respuesta TEXT,
                q2_puntaje REAL,
                q3_respuesta TEXT,
                q3_puntaje REAL,
                q4_respuesta TEXT,
                q4_puntaje REAL,
                q5_respuesta TEXT,
                q5_puntaje REAL,
                total REAL,
                promedio REAL,
                comentarios TEXT,
                imported_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def normalize_col(value) -> str:
    return (
        str(value or "")
        .strip()
        .upper()
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ñ", "N")
        .replace(" ", "_")
    )


def to_float(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        # Fecha serial de Excel.
        if 30000 <= float(value) <= 70000:
            return datetime(1899, 12, 30) + timedelta(days=float(value))
    text = str(value).strip()
    formats = [
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d-%m-%Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def text_value(value):
    if value is None:
        return ""
    return str(value).strip()


def find_header_row(ws, max_scan=30):
    expected = {normalize_col(c) for c in EXPECTED_COLUMNS}
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        normalized = [normalize_col(cell) for cell in row]
        hits = sum(1 for col in normalized if col in expected)
        if hits >= 8:
            return row_idx, normalized
    return 1, [normalize_col(cell.value) for cell in ws[1]]


def read_uploaded_workbook(file_storage):
    filename = secure_filename(file_storage.filename or "archivo.xlsx")
    if not filename.lower().endswith(".xlsx"):
        raise ValueError("El archivo debe estar en formato Excel .xlsx.")

    file_storage.stream.seek(0)
    wb = load_workbook(filename=file_storage.stream, data_only=True, read_only=True)

    selected = None
    for name in wb.sheetnames:
        low = name.lower()
        if "encuesta" in low and "satisf" in low:
            selected = name
            break
    if selected is None:
        selected = wb.sheetnames[0]

    ws = wb[selected]
    header_row, normalized_headers = find_header_row(ws)
    header_map = {name: idx for idx, name in enumerate(normalized_headers) if name}
    expected = [normalize_col(c) for c in EXPECTED_COLUMNS]
    missing = [col for col in expected if col not in header_map]
    if missing:
        raise ValueError("Faltan columnas requeridas: " + ", ".join(missing))

    records = []
    q_score_cols = ["q1_puntaje", "q2_puntaje", "q3_puntaje", "q4_puntaje", "q5_puntaje"]
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not row or not any(value is not None and str(value).strip() != "" for value in row):
            continue

        raw = {}
        for expected_name, db_name in zip(expected, DB_COLUMNS):
            idx = header_map[expected_name]
            raw[db_name] = row[idx] if idx < len(row) else None

        fecha = parse_date(raw["fecha"])
        if fecha is None:
            continue

        record = {
            "fecha": fecha.strftime("%Y-%m-%d %H:%M:%S"),
            "q1_respuesta": text_value(raw["q1_respuesta"]),
            "q2_respuesta": text_value(raw["q2_respuesta"]),
            "q3_respuesta": text_value(raw["q3_respuesta"]),
            "q4_respuesta": text_value(raw["q4_respuesta"]),
            "q5_respuesta": text_value(raw["q5_respuesta"]),
            "comentarios": text_value(raw["comentarios"]),
        }
        for col in q_score_cols:
            record[col] = to_float(raw[col])

        valid_scores = [record[col] for col in q_score_cols if record[col] is not None]
        total = to_float(raw["total"])
        promedio = to_float(raw["promedio"])
        if total is None and valid_scores:
            total = sum(valid_scores)
        if promedio is None and valid_scores:
            promedio = sum(valid_scores) / len(valid_scores)
        if promedio is None or promedio < 1 or promedio > 5:
            continue
        record["total"] = total
        record["promedio"] = promedio
        records.append(record)

    if not records:
        raise ValueError("No se encontraron registros válidos con fecha y promedio entre 1 y 5.")
    return records


def save_records(records, mode="replace"):
    imported_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_conn() as conn:
        if mode == "replace":
            conn.execute("DELETE FROM encuestas")
        conn.executemany(
            """
            INSERT INTO encuestas (
                fecha, q1_respuesta, q1_puntaje, q2_respuesta, q2_puntaje,
                q3_respuesta, q3_puntaje, q4_respuesta, q4_puntaje,
                q5_respuesta, q5_puntaje, total, promedio, comentarios, imported_at
            ) VALUES (
                :fecha, :q1_respuesta, :q1_puntaje, :q2_respuesta, :q2_puntaje,
                :q3_respuesta, :q3_puntaje, :q4_respuesta, :q4_puntaje,
                :q5_respuesta, :q5_puntaje, :total, :promedio, :comentarios, :imported_at
            )
            """,
            [{**record, "imported_at": imported_at} for record in records],
        )
        conn.commit()
    return len(records)


def load_data():
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM encuestas ORDER BY fecha ASC, id ASC").fetchall()
    data = []
    for row in rows:
        item = dict(row)
        item["fecha_dt"] = parse_date(item.get("fecha"))
        if item["fecha_dt"] is not None and item.get("promedio") is not None:
            data.append(item)
    return data


def status_from_score(score):
    if score is None:
        return "Sin encuestas"
    if score >= 4.80:
        return "Mantener estándar"
    if score >= 4.70:
        return "Seguimiento preventivo"
    return "Generar plan de acción"


def status_class(text):
    text = (text or "").lower()
    if "mantener" in text:
        return "ok"
    if "preventivo" in text or "seguimiento" in text:
        return "warn"
    if "acción" in text or "accion" in text or "riesgo" in text:
        return "danger"
    return "neutral"


def pct_number(part, total):
    return 0.0 if total == 0 else round((part / total) * 100, 1)


def pct(part, total):
    return f"{pct_number(part, total):.1f}%"


def start_of_week(dt):
    return (dt - timedelta(days=dt.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)


def week_label(start_date):
    end_date = start_date + timedelta(days=6)
    iso = start_date.isocalendar()
    return f"Sem. {int(iso.week):02d} ({start_date:%d-%m} al {end_date:%d-%m})"


def average(values):
    clean = [v for v in values if v is not None]
    return round(mean(clean), 2) if clean else None


def build_weekly(data):
    buckets = defaultdict(list)
    for item in data:
        buckets[start_of_week(item["fecha_dt"])].append(item)
    weekly = []
    for week_start in sorted(buckets):
        rows = buckets[week_start]
        record = {
            "week_start": week_start,
            "week_end": week_start + timedelta(days=6),
            "iso_year": week_start.isocalendar().year,
            "iso_week": week_start.isocalendar().week,
            "periodo": week_label(week_start),
            "n_encuestas": len(rows),
            "promedio_general": average([r.get("promedio") for r in rows]),
        }
        for col in DIMENSION_LABELS:
            record[col] = average([r.get(col) for r in rows])
        record["estado"] = status_from_score(record["promedio_general"])
        weekly.append(record)
    return weekly


def build_metrics(data):
    if not data:
        return {
            "has_data": False,
            "total": 0,
            "global_score": 0,
            "excelencia": 0,
            "risk": [],
            "comments": [],
            "weekly": [],
            "recent_weeks": [],
            "dimensions": [],
            "analysis": [],
            "chart_labels": [],
            "chart_values": [],
            "radar_labels": RADAR_LABELS,
            "radar_values": [],
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "global_status": "Sin datos",
            "latest_label": "Sin datos",
        }

    total = len(data)
    global_score = average([item.get("promedio") for item in data]) or 0
    promoters = sum(1 for item in data if item.get("promedio", 0) >= 4.8)
    neutrals = sum(1 for item in data if 4.0 <= item.get("promedio", 0) < 4.8)
    detractors = sum(1 for item in data if item.get("promedio", 0) < 4.0)

    risk = [
        {"categoria": "Promotores (Alta Satisfacción)", "criterio": "4.8 - 5.0", "volumen": promoters, "representacion": pct(promoters, total), "kind": "ok"},
        {"categoria": "Neutros (Satisfacción Estándar)", "criterio": "4.0 - 4.7", "volumen": neutrals, "representacion": pct(neutrals, total), "kind": "neutral"},
        {"categoria": "Detractores (Riesgo Operativo)", "criterio": "1.0 - 3.9", "volumen": detractors, "representacion": pct(detractors, total), "kind": "danger"},
    ]

    comments = [item.get("comentarios", "") for item in sorted(data, key=lambda r: r["fecha_dt"], reverse=True) if item.get("comentarios", "").strip()][:6]
    weekly = build_weekly(data)
    last_10 = weekly[-10:]
    recent = weekly[-3:]

    if weekly:
        latest_start = weekly[-1]["week_start"]
        latest_rows = [item for item in data if start_of_week(item["fecha_dt"]) == latest_start]
        latest_label = weekly[-1]["periodo"]
    else:
        latest_rows = data
        latest_label = "Total histórico"

    dimensions = []
    for key, label in DIMENSION_LABELS.items():
        score = average([r.get(key) for r in latest_rows])
        dimensions.append({"codigo": key[:2].upper(), "label": label, "score": score, "foco": score is not None and score < 4.70})

    sorted_dims = sorted([d for d in dimensions if d["score"] is not None], key=lambda x: x["score"])
    analysis = []
    if sorted_dims:
        worst = sorted_dims[0]
        best = sorted_dims[-1]
        analysis.append(f"Mejor dimensión: {best['label']} con {best['score']:.2f}.")
        if worst["score"] < 4.70:
            analysis.append(f"Foco preventivo: {worst['label']} registra {worst['score']:.2f}, bajo el umbral 4.70.")
        else:
            analysis.append(f"Sin foco crítico: la dimensión más baja es {worst['label']} con {worst['score']:.2f}.")
    if detractors > 0:
        analysis.append(f"Riesgo operativo: {detractors} evaluaciones están bajo 4.0 y requieren revisión cualitativa.")
    if comments:
        analysis.append("La revisión de comentarios recientes permite identificar causas operativas específicas detrás de los puntajes.")

    return {
        "has_data": True,
        "total": total,
        "global_score": global_score,
        "global_status": status_from_score(global_score),
        "excelencia": pct_number(promoters, total),
        "risk": risk,
        "comments": comments,
        "weekly": weekly,
        "recent_weeks": [{"periodo": r["periodo"], "n_encuestas": r["n_encuestas"], "promedio": r["promedio_general"], "estado": r["estado"]} for r in recent],
        "dimensions": dimensions,
        "latest_label": latest_label,
        "analysis": analysis,
        "chart_labels": [r["periodo"] for r in last_10],
        "chart_values": [r["promedio_general"] for r in last_10],
        "radar_labels": RADAR_LABELS,
        "radar_values": [d["score"] or 0 for d in dimensions],
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def export_csv_buffer(data):
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=["id"] + DB_COLUMNS + ["imported_at"])
    writer.writeheader()
    for item in data:
        row = {key: item.get(key, "") for key in ["id"] + DB_COLUMNS + ["imported_at"]}
        row["fecha"] = item.get("fecha")
        writer.writerow(row)
    return output.getvalue().encode("utf-8-sig")


def build_excel_buffer(data, metrics):
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos Encuesta"
    ws.append(["ID"] + EXPECTED_COLUMNS + ["IMPORTADO_EN"])
    for item in data:
        ws.append([
            item.get("id"), item.get("fecha"),
            item.get("q1_respuesta"), item.get("q1_puntaje"),
            item.get("q2_respuesta"), item.get("q2_puntaje"),
            item.get("q3_respuesta"), item.get("q3_puntaje"),
            item.get("q4_respuesta"), item.get("q4_puntaje"),
            item.get("q5_respuesta"), item.get("q5_puntaje"),
            item.get("total"), item.get("promedio"), item.get("comentarios"), item.get("imported_at"),
        ])

    ws2 = wb.create_sheet("Resumen Semanal")
    ws2.append(["Desde", "Hasta", "Año", "Semana", "Periodo", "N° Encuestas", "Prom. P1", "Prom. P2", "Prom. P3", "Prom. P4", "Prom. P5", "Promedio General", "Estado / Acción"])
    for row in metrics["weekly"]:
        ws2.append([
            row["week_start"].strftime("%Y-%m-%d"), row["week_end"].strftime("%Y-%m-%d"), row["iso_year"], row["iso_week"], row["periodo"], row["n_encuestas"],
            row.get("q1_puntaje"), row.get("q2_puntaje"), row.get("q3_puntaje"), row.get("q4_puntaje"), row.get("q5_puntaje"), row.get("promedio_general"), row.get("estado"),
        ])

    ws3 = wb.create_sheet("Distribución Riesgo")
    ws3.append(["Categoría", "Criterio", "Volumen", "Representación"])
    for item in metrics["risk"]:
        ws3.append([item["categoria"], item["criterio"], item["volumen"], item["representacion"]])

    ws4 = wb.create_sheet("Dimensiones")
    ws4.append(["Dimensión", "Promedio", "Foco"])
    for item in metrics["dimensions"]:
        ws4.append([item["label"], item["score"], "Sí" if item["foco"] else "No"])

    for wsx in wb.worksheets:
        for column_cells in wsx.columns:
            length = max(len(str(cell.value or "")) for cell in column_cells)
            wsx.column_dimensions[column_cells[0].column_letter].width = min(max(length + 2, 12), 48)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def build_template_buffer():
    wb = Workbook()
    ws = wb.active
    ws.title = "Encuesta de Satisfacción"
    ws.append(EXPECTED_COLUMNS)
    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Muy satisfecho", 5,
        "Satisfecho", 4,
        "Muy satisfecho", 5,
        "Muy satisfecho", 5,
        "Muy satisfecho", 5,
        24, 4.8, "Comentario de ejemplo",
    ])
    for column_cells in ws.columns:
        length = max(len(str(cell.value or "")) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = min(max(length + 2, 12), 36)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# Generador PDF minimalista sin pandas, numpy, matplotlib ni reportlab.
# -----------------------------------------------------------------------------

def pdf_escape(text):
    """
    Codifica texto para strings PDF usando bytes WinAnsi/CP1252 con escapes octales.
    Esto evita que lectores PDF de navegador muestren tildes y eñes como caracteres rotos.
    Ejemplo: ó se escribe como \363 dentro del PDF, no como byte ambiguo.
    """
    replacements = {
        "•": "-",
        "–": "-",
        "—": "-",
        "“": '"',
        "”": '"',
        "‘": "'",
        "’": "'",
        "…": "...",
    }
    value = str(text or "")
    for src, dst in replacements.items():
        value = value.replace(src, dst)

    raw = value.encode("cp1252", "replace")
    escaped = []
    for byte in raw:
        # Paréntesis y backslash deben escaparse en strings literales PDF.
        # Bytes no ASCII se escriben como octal para conservar tildes/ñ.
        if byte in (0x28, 0x29, 0x5C):
            escaped.append(f"\\{byte:03o}")
        elif 32 <= byte <= 126:
            escaped.append(chr(byte))
        elif byte in (9, 10, 13):
            escaped.append(" ")
        else:
            escaped.append(f"\\{byte:03o}")
    return "".join(escaped)


def hex_to_rgb(hex_color):
    h = hex_color.strip().lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))


def wrap_text(text, max_chars):
    words = str(text or "").split()
    lines = []
    current = ""
    for word in words:
        if len(current) + len(word) + 1 <= max_chars:
            current = (current + " " + word).strip()
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def parse_png_rgb(path):
    with open(path, "rb") as fh:
        data = fh.read()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("No es PNG")
    pos = 8
    width = height = bit_depth = color_type = None
    compressed = b""
    while pos < len(data):
        length = int.from_bytes(data[pos:pos+4], "big")
        chunk_type = data[pos+4:pos+8]
        chunk_data = data[pos+8:pos+8+length]
        pos += 12 + length
        if chunk_type == b"IHDR":
            width = int.from_bytes(chunk_data[0:4], "big")
            height = int.from_bytes(chunk_data[4:8], "big")
            bit_depth = chunk_data[8]
            color_type = chunk_data[9]
        elif chunk_type == b"IDAT":
            compressed += chunk_data
        elif chunk_type == b"IEND":
            break
    if bit_depth != 8 or color_type not in (2, 6):
        raise ValueError("PNG no soportado para incrustación")
    raw = zlib.decompress(compressed)
    bpp = 4 if color_type == 6 else 3
    stride = width * bpp
    rows = []
    i = 0
    prev = bytearray(stride)

    def paeth(a, b, c):
        p = a + b - c
        pa = abs(p - a)
        pb = abs(p - b)
        pc = abs(p - c)
        if pa <= pb and pa <= pc:
            return a
        if pb <= pc:
            return b
        return c

    for _ in range(height):
        filt = raw[i]
        i += 1
        scan = bytearray(raw[i:i+stride])
        i += stride
        recon = bytearray(stride)
        for x in range(stride):
            left = recon[x-bpp] if x >= bpp else 0
            up = prev[x]
            up_left = prev[x-bpp] if x >= bpp else 0
            val = scan[x]
            if filt == 0:
                recon[x] = val
            elif filt == 1:
                recon[x] = (val + left) & 0xFF
            elif filt == 2:
                recon[x] = (val + up) & 0xFF
            elif filt == 3:
                recon[x] = (val + ((left + up) // 2)) & 0xFF
            elif filt == 4:
                recon[x] = (val + paeth(left, up, up_left)) & 0xFF
            else:
                raise ValueError("Filtro PNG no soportado")
        rows.append(recon)
        prev = recon
    rgb = bytearray()
    if color_type == 6:
        for row in rows:
            for x in range(0, len(row), 4):
                r, g, b, a = row[x], row[x+1], row[x+2], row[x+3]
                alpha = a / 255
                # Se aplana sobre fondo blanco para mantener compatibilidad PDF simple.
                rgb.extend([int(r * alpha + 255 * (1 - alpha)), int(g * alpha + 255 * (1 - alpha)), int(b * alpha + 255 * (1 - alpha))])
    else:
        for row in rows:
            rgb.extend(row)
    return width, height, bytes(rgb)


def build_winansi_tounicode_stream():
    """CMap ToUnicode para que copiar/visualizar tildes en PDF sea estable."""
    pairs = []
    for code in range(32, 256):
        try:
            char = bytes([code]).decode("cp1252")
        except UnicodeDecodeError:
            continue
        if not char:
            continue
        pairs.append((code, ord(char)))

    lines = [
        "/CIDInit /ProcSet findresource begin",
        "12 dict begin",
        "begincmap",
        "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
        "/CMapName /WinAnsiToUnicode def",
        "/CMapType 2 def",
        "1 begincodespacerange",
        "<00> <FF>",
        "endcodespacerange",
    ]
    for start in range(0, len(pairs), 100):
        chunk = pairs[start:start + 100]
        lines.append(f"{len(chunk)} beginbfchar")
        for code, uni in chunk:
            lines.append(f"<{code:02X}> <{uni:04X}>")
        lines.append("endbfchar")
    lines.extend([
        "endcmap",
        "CMapName currentdict /CMap defineresource pop",
        "end",
        "end",
    ])
    stream = "\n".join(lines).encode("ascii")
    return f"<< /Length {len(stream)} >>\nstream\n".encode("latin-1") + stream + b"\nendstream"


class SimplePDF:
    def __init__(self):
        self.objects = [None]
        self.catalog_id = self.reserve()
        self.pages_id = self.reserve()
        self.tounicode_id = self.add_object(build_winansi_tounicode_stream())
        self.font_id = self.add_object(
            f"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding /ToUnicode {self.tounicode_id} 0 R >>"
        )
        self.font_bold_id = self.add_object(
            f"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding /ToUnicode {self.tounicode_id} 0 R >>"
        )
        self.pages = []
        self.images = {}

    def reserve(self):
        self.objects.append(None)
        return len(self.objects) - 1

    def set_object(self, obj_id, data):
        self.objects[obj_id] = data if isinstance(data, bytes) else data.encode("latin-1")

    def add_object(self, data):
        self.objects.append(data if isinstance(data, bytes) else data.encode("latin-1"))
        return len(self.objects) - 1

    def add_image(self, name, path):
        try:
            width, height, rgb = parse_png_rgb(path)
            comp = zlib.compress(rgb)
            obj = (
                f"<< /Type /XObject /Subtype /Image /Width {width} /Height {height} "
                f"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length {len(comp)} >>\nstream\n"
            ).encode("latin-1") + comp + b"\nendstream"
            obj_id = self.add_object(obj)
            self.images[name] = {"id": obj_id, "width": width, "height": height}
        except Exception:
            pass

    def add_page(self, commands):
        content_bytes = "\n".join(commands).encode("latin-1")
        content_id = self.add_object(f"<< /Length {len(content_bytes)} >>\nstream\n".encode("latin-1") + content_bytes + b"\nendstream")
        page_id = self.reserve()
        xobjects = " ".join(f"/{name} {info['id']} 0 R" for name, info in self.images.items())
        xobject_part = f" /XObject << {xobjects} >>" if xobjects else ""
        page = (
            f"<< /Type /Page /Parent {self.pages_id} 0 R /MediaBox [0 0 595 842] "
            f"/Resources << /Font << /F1 {self.font_id} 0 R /F2 {self.font_bold_id} 0 R >>{xobject_part} >> "
            f"/Contents {content_id} 0 R >>"
        )
        self.set_object(page_id, page)
        self.pages.append(page_id)

    def finish(self):
        kids = " ".join(f"{pid} 0 R" for pid in self.pages)
        self.set_object(self.pages_id, f"<< /Type /Pages /Kids [ {kids} ] /Count {len(self.pages)} >>")
        self.set_object(self.catalog_id, f"<< /Type /Catalog /Pages {self.pages_id} 0 R >>")
        out = io.BytesIO()
        out.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for idx in range(1, len(self.objects)):
            offsets.append(out.tell())
            out.write(f"{idx} 0 obj\n".encode("latin-1"))
            out.write(self.objects[idx] or b"")
            out.write(b"\nendobj\n")
        xref = out.tell()
        out.write(f"xref\n0 {len(self.objects)}\n".encode("latin-1"))
        out.write(b"0000000000 65535 f \n")
        for off in offsets[1:]:
            out.write(f"{off:010d} 00000 n \n".encode("latin-1"))
        out.write(f"trailer\n<< /Size {len(self.objects)} /Root {self.catalog_id} 0 R >>\nstartxref\n{xref}\n%%EOF".encode("latin-1"))
        out.seek(0)
        return out


def cmd_set_fill(color):
    r, g, b = hex_to_rgb(color)
    return f"{r:.3f} {g:.3f} {b:.3f} rg"


def cmd_set_stroke(color):
    r, g, b = hex_to_rgb(color)
    return f"{r:.3f} {g:.3f} {b:.3f} RG"


def rect(x, y, w, h, color):
    return [cmd_set_fill(color), f"{x:.1f} {y:.1f} {w:.1f} {h:.1f} re f"]


def line(x1, y1, x2, y2, color="#dddddd", width=1):
    return [cmd_set_stroke(color), f"{width:.1f} w", f"{x1:.1f} {y1:.1f} m {x2:.1f} {y2:.1f} l S"]


def text(x, y, value, size=10, color="#303030", bold=False):
    r, g, b = hex_to_rgb(color)
    font = "/F2" if bold else "/F1"
    return f"BT {r:.3f} {g:.3f} {b:.3f} rg {font} {size:.1f} Tf {x:.1f} {y:.1f} Td ({pdf_escape(value)}) Tj ET"


def image_cmd(name, x, y, w, h):
    return f"q {w:.1f} 0 0 {h:.1f} {x:.1f} {y:.1f} cm /{name} Do Q"


def pdf_header(cmds, pdf):
    if "LogoA" in pdf.images:
        cmds.append(image_cmd("LogoA", 42, 790, 115, 29))
    else:
        cmds.append(text(42, 804, "aramark", 16, RED, True))
    if "LogoB" in pdf.images:
        cmds.append(image_cmd("LogoB", 315, 790, 235, 21))
    else:
        cmds.append(text(360, 804, "ESCONDIDA | BHP", 15, "#f58220", True))
    cmds.extend(line(42, 778, 553, 778, RED, 2.5))


def draw_section(cmds, y, title):
    cmds.extend(rect(42, y - 4, 511, 24, RED))
    cmds.append(text(50, y + 3, title, 11, "#ffffff", True))
    return y - 36


def draw_kpi_cards(cmds, y, metrics):
    cards = [
        ("TOTAL DE EVALUACIONES", str(metrics["total"]), "Muestra Histórica"),
        ("SATISFACCIÓN GLOBAL", f"{metrics['global_score']:.2f} / 5.0", metrics["global_status"]),
        ("ÍNDICE DE EXCELENCIA", f"{metrics['excelencia']:.1f}%", "Usuarios Promotores"),
    ]
    x = 42
    w = 158
    for label, value, note in cards:
        cmds.extend(rect(x, y - 72, w, 62, "#fff8f9"))
        cmds.extend(line(x, y - 10, x + w, y - 10, RED, 2))
        cmds.append(text(x + 12, y - 29, label, 7.5, "#666666", True))
        cmds.append(text(x + 46, y - 52, value, 18, "#222222", True))
        cmds.append(text(x + 20, y - 66, note, 7, "#2f7d32", True))
        x += 176
    return y - 88


def draw_table(cmds, x, y, headers, rows, col_widths, row_h=18, max_rows=None):
    max_rows = max_rows or len(rows)
    total_w = sum(col_widths)
    cmds.extend(rect(x, y - row_h, total_w, row_h, "#f1f1f1"))
    cx = x
    for header, cw in zip(headers, col_widths):
        cmds.append(text(cx + 4, y - 13, header, 7.2, "#444444", True))
        cx += cw
    yy = y - row_h
    for row in rows[:max_rows]:
        yy -= row_h
        cmds.extend(line(x, yy + row_h, x + total_w, yy + row_h, "#dddddd", 0.4))
        cx = x
        for val, cw in zip(row, col_widths):
            shown = str(val or "")
            if len(shown) > int(cw / 4.4):
                shown = shown[: int(cw / 4.4) - 1] + "…"
            cmds.append(text(cx + 4, yy + 5, shown, 7.4, "#303030", False))
            cx += cw
    cmds.extend(line(x, yy, x + total_w, yy, "#dddddd", 0.4))
    return yy - 18


def draw_comments_and_analysis(cmds, y, metrics):
    cmds.append(text(42, y, "Comentarios Recientes Destacados", 10.5, RED, True))
    cmds.append(text(308, y, "Análisis Cualitativo", 10.5, RED, True))
    y0 = y - 18
    left_y = y0
    for comment in metrics["comments"][:4]:
        cmds.extend(line(42, left_y + 8, 42, left_y - 34, RED, 2.5))
        lines = wrap_text(comment, 45)[:3]
        ty = left_y
        for ln in lines:
            cmds.append(text(50, ty, f"\"{ln}", 7.6, "#555555"))
            ty -= 10
        left_y -= 46
    if not metrics["comments"]:
        cmds.append(text(42, left_y, "Sin comentarios registrados.", 8, "#777777"))
    right_y = y0
    for item in metrics["analysis"][:5]:
        cmds.extend(rect(308, right_y - 2, 3, 3, RED))
        for ln in wrap_text(item, 52)[:3]:
            cmds.append(text(316, right_y, ln, 7.6, "#303030"))
            right_y -= 10
        right_y -= 8
    return min(left_y, right_y) - 4


def draw_trend_chart(cmds, x, y, w, h, labels, values):
    cmds.extend(line(x, y, x, y + h, "#999999", 0.8))
    cmds.extend(line(x, y, x + w, y, "#999999", 0.8))
    for i in range(1, 5):
        gy = y + h * i / 4
        cmds.extend(line(x, gy, x + w, gy, "#e5e5e5", 0.3))
    if values:
        pts = []
        for idx, val in enumerate(values):
            px = x + (w * idx / max(len(values) - 1, 1))
            py = y + ((float(val) - 1) / 4) * h
            pts.append((px, py))
        cmds.append(cmd_set_stroke(RED))
        cmds.append("2 w")
        path = f"{pts[0][0]:.1f} {pts[0][1]:.1f} m " + " ".join(f"{px:.1f} {py:.1f} l" for px, py in pts[1:]) + " S"
        cmds.append(path)
        for idx, (px, py) in enumerate(pts):
            cmds.append(cmd_set_fill(RED))
            cmds.append(f"{px:.1f} {py:.1f} 3.2 0 360 arc f")
            cmds.append(text(px - 8, py + 10, f"{values[idx]:.2f}", 7, "#303030"))
    cmds.append(text(x - 25, y + h + 2, "5.0", 7, "#666666"))
    cmds.append(text(x - 25, y - 2, "1.0", 7, "#666666"))


def draw_radar_chart(cmds, cx, cy, radius, labels, values):
    n = len(values) or 5
    angles = [-math.pi / 2 + i * 2 * math.pi / n for i in range(n)]
    for level in range(1, 6):
        r = radius * level / 5
        pts = [(cx + math.cos(a) * r, cy + math.sin(a) * r) for a in angles]
        cmds.append(cmd_set_stroke("#dddddd"))
        cmds.append("0.5 w")
        cmds.append(f"{pts[0][0]:.1f} {pts[0][1]:.1f} m " + " ".join(f"{x:.1f} {y:.1f} l" for x, y in pts[1:]) + " h S")
    pts = [(cx + math.cos(a) * radius * (float(v) / 5), cy + math.sin(a) * radius * (float(v) / 5)) for a, v in zip(angles, values)]
    cmds.append(cmd_set_stroke(RED))
    cmds.append("1.8 w")
    if pts:
        cmds.append(f"{pts[0][0]:.1f} {pts[0][1]:.1f} m " + " ".join(f"{x:.1f} {y:.1f} l" for x, y in pts[1:]) + " h S")
    for a, label in zip(angles, labels):
        lx = cx + math.cos(a) * (radius + 20)
        ly = cy + math.sin(a) * (radius + 14)
        cmds.append(text(lx - 25, ly, label[:14], 6.5, "#555555"))


def generate_pdf_report(metrics):
    pdf = SimplePDF()
    pdf.add_image("LogoA", LOGO_ARAMARK_PATH)
    pdf.add_image("LogoB", LOGO_BHP_PATH)

    cmds = []
    pdf_header(cmds, pdf)
    cmds.append(text(42, 742, "REPORTE EJECUTIVO DE CALIDAD", 18, "#202020", True))
    cmds.append(text(42, 724, "Evaluación de Satisfacción e Impacto del Factor Humano en la Operación", 10, "#666666"))
    y = draw_section(cmds, 690, "1. Resumen Ejecutivo (Macrométricas)")
    y = draw_kpi_cards(cmds, y, metrics)
    y = draw_section(cmds, y, "2. Distribución del Riesgo Operativo")
    risk_rows = [[r["categoria"], r["criterio"], r["volumen"], r["representacion"]] for r in metrics["risk"]]
    y = draw_table(cmds, 42, y, ["CATEGORÍA DE USUARIO", "CRITERIO", "VOLUMEN", "REP."], risk_rows, [235, 95, 80, 100], row_h=22)
    y = draw_section(cmds, y, "3. La Voz del Usuario y Observaciones")
    draw_comments_and_analysis(cmds, y, metrics)
    cmds.extend(line(42, 46, 553, 46, "#dddddd", 0.5))
    cmds.append(text(126, 30, f"Generado automáticamente a partir de datos operacionales • Evaluación Interna • {datetime.now():%Y-%m-%d}", 7, "#999999"))
    pdf.add_page(cmds)

    cmds = []
    pdf_header(cmds, pdf)
    cmds.append(text(42, 742, "ANÁLISIS DE TENDENCIAS Y DESGLOSE", 18, "#202020", True))
    cmds.append(text(42, 724, "Monitoreo semanal de resultados operativos", 10, "#666666"))
    y = draw_section(cmds, 690, "4. Evolución Histórica (Últimas 10 Semanas)")
    draw_trend_chart(cmds, 70, y - 230, 455, 190, metrics["chart_labels"], metrics["chart_values"])
    y = y - 260
    y = draw_section(cmds, y, "5. Resumen Semanas Recientes")
    recent_rows = [[r["periodo"], r["n_encuestas"], f"{r['promedio']:.2f}", r["estado"]] for r in metrics["recent_weeks"]]
    y = draw_table(cmds, 42, y, ["SEMANA", "N° ENC.", "PROM.", "ESTADO / ACCIÓN"], recent_rows, [210, 75, 75, 150], row_h=22)
    cmds.extend(line(42, 46, 553, 46, "#dddddd", 0.5))
    cmds.append(text(126, 30, f"Generado automáticamente a partir de datos operacionales • Evaluación Interna • {datetime.now():%Y-%m-%d}", 7, "#999999"))
    pdf.add_page(cmds)

    # Página exclusiva para la sección 6, así se evita que el gráfico radar se mezcle con el pie de página.
    cmds = []
    pdf_header(cmds, pdf)
    cmds.append(text(42, 742, "ANÁLISIS DE TENDENCIAS Y DESGLOSE", 18, "#202020", True))
    cmds.append(text(42, 724, "Monitoreo semanal de resultados operativos", 10, "#666666"))
    y = draw_section(cmds, 690, f"6. Desempeño Específico por Dimensiones ({metrics['latest_label']})")
    cmds.append(text(42, y + 10, "Una calificación por debajo de 4.70 se considera foco de atención preventivo.", 8.5, "#555555"))
    draw_radar_chart(cmds, 175, y - 150, 80, metrics["radar_labels"], metrics["radar_values"])
    dy = y - 60
    for item in metrics["dimensions"]:
        color = RED if item["foco"] else "#2f7d32"
        score = "-" if item["score"] is None else f"{item['score']:.2f}" + (" (Foco)" if item["foco"] else "")
        cmds.extend(rect(330, dy - 1, 4, 4, RED))
        cmds.append(text(340, dy, item["label"], 8.5, "#303030", True))
        cmds.append(text(505, dy, score, 8.5, color, True))
        dy -= 20
    cmds.extend(line(42, 46, 553, 46, "#dddddd", 0.5))
    cmds.append(text(126, 30, f"Generado automáticamente a partir de datos operacionales • Evaluación Interna • {datetime.now():%Y-%m-%d}", 7, "#999999"))
    pdf.add_page(cmds)
    return pdf.finish()


@app.after_request
def force_utf8_headers(response):
    """Evita caracteres rotos en Render/navegador para tildes, ñ y símbolos."""
    if response.mimetype in {"text/html", "text/css", "text/javascript", "application/javascript", "text/csv", "application/json"}:
        response.headers["Content-Type"] = f"{response.mimetype}; charset=utf-8"
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response

@app.context_processor
def inject_helpers():
    return {"json_dumps": json.dumps, "status_class": status_class}


@app.route("/")
def dashboard():
    data = load_data()
    metrics = build_metrics(data)
    return render_template("dashboard.html", title=APP_TITLE, metrics=metrics)


@app.route("/importar", methods=["POST"])
def importar():
    archivo = request.files.get("archivo")
    modo = request.form.get("modo", "replace")
    if not archivo or archivo.filename == "":
        flash("Debes seleccionar un archivo Excel.", "error")
        return redirect(url_for("dashboard"))
    try:
        records = read_uploaded_workbook(archivo)
        count = save_records(records, mode=modo)
        flash(f"Importación exitosa: {count} registros procesados.", "success")
    except Exception as exc:
        flash(f"Error al importar: {exc}", "error")
    return redirect(url_for("dashboard"))


@app.route("/plantilla")
def plantilla():
    buffer = build_template_buffer()
    return send_file(buffer, as_attachment=True, download_name="plantilla_encuesta_satisfaccion_5400.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/exportar/csv")
def exportar_csv():
    data = load_data()
    content = export_csv_buffer(data)
    return Response(content, content_type="text/csv; charset=utf-8", headers={"Content-Disposition": "attachment; filename=encuesta_satisfaccion_5400.csv"})


@app.route("/exportar/excel")
def exportar_excel():
    data = load_data()
    metrics = build_metrics(data)
    buffer = build_excel_buffer(data, metrics)
    return send_file(buffer, as_attachment=True, download_name="reporte_encuesta_satisfaccion_5400.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/exportar/pdf")
def exportar_pdf():
    data = load_data()
    metrics = build_metrics(data)
    if not metrics["has_data"]:
        flash("No hay datos para exportar.", "error")
        return redirect(url_for("dashboard"))
    buffer = generate_pdf_report(metrics)
    return send_file(buffer, as_attachment=True, download_name="reporte_encuesta_satisfaccion_5400.pdf", mimetype="application/pdf")


@app.route("/health")
def health():
    return {"status": "ok", "app": APP_TITLE}


init_db()

if __name__ == "__main__":
    app.run(debug=True)

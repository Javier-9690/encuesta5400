import csv
import io
import json
import os
import sqlite3
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from flask import Flask, Response, flash, redirect, render_template, request, send_file, url_for
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from werkzeug.utils import secure_filename

APP_TITLE = "Dashboard Encuesta de Satisfacción - Campamento 5400"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
LOGO_ARAMARK_PATH = os.path.join(STATIC_DIR, "img", "logo_aramark.png")
LOGO_CAMPAMENTO_PATH = os.path.join(STATIC_DIR, "img", "logo_campamento_5400.png")
DB_PATH = os.environ.get("DATABASE_PATH", "/var/data/encuesta_5400.db")
if not os.path.exists(os.path.dirname(DB_PATH)):
    DB_PATH = os.path.join(os.path.dirname(__file__), "encuesta_5400.db")

EXPECTED_COLUMNS = [
    "FECHA",
    "Q1_RESPUESTA", "Q1_PUNTAJE",
    "Q2_RESPUESTA", "Q2_PUNTAJE",
    "Q3_RESPUESTA", "Q3_PUNTAJE",
    "Q4_RESPUESTA", "Q4_PUNTAJE",
    "Q5_RESPUESTA", "Q5_PUNTAJE",
    "TOTAL", "PROMEDIO", "COMENTARIOS",
]

DIMENSION_LABELS = {
    "q1_puntaje": "Q1. Primera Impresión / Recepción",
    "q2_puntaje": "Q2. Calidad del Serv. Principal",
    "q3_puntaje": "Q3. Tiempos de Respuesta",
    "q4_puntaje": "Q4. Higiene y Presentación",
    "q5_puntaje": "Q5. Trato y Factor Humano",
}

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


def normalize_col(value: str) -> str:
    return (
        str(value)
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


def parse_excel_date(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    # Los números de serie de Excel para fechas recientes suelen estar entre 30000 y 70000.
    # Si se intentan convertir como fecha normal, pandas puede interpretarlos como nanosegundos de 1970.
    excel_mask = numeric.between(30000, 70000)
    parsed_numeric = pd.to_datetime(numeric.where(excel_mask), unit="D", origin="1899-12-30", errors="coerce")
    parsed_regular = pd.to_datetime(series.where(~excel_mask), errors="coerce", dayfirst=True)
    return parsed_regular.fillna(parsed_numeric)


def clean_score(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",", ".", regex=False), errors="coerce")


def read_uploaded_workbook(file_storage) -> pd.DataFrame:
    filename = secure_filename(file_storage.filename or "archivo.xlsx")
    if not filename.lower().endswith((".xlsx", ".xls")):
        raise ValueError("El archivo debe ser Excel (.xlsx o .xls).")

    # Se leen todas las pestañas para encontrar automáticamente "Encuesta de Satisfacción".
    file_storage.stream.seek(0)
    sheets = pd.read_excel(file_storage.stream, sheet_name=None)
    selected_name = None
    for name in sheets.keys():
        if "encuesta" in name.lower() and "satisf" in name.lower():
            selected_name = name
            break
    if selected_name is None:
        selected_name = next(iter(sheets.keys()))

    df = sheets[selected_name].copy()
    df = df.dropna(how="all")
    if df.empty:
        raise ValueError(f"La pestaña '{selected_name}' no contiene datos.")

    df.columns = [normalize_col(c) for c in df.columns]
    normalized_expected = [normalize_col(c) for c in EXPECTED_COLUMNS]
    missing = [col for col in normalized_expected if col not in df.columns]
    if missing:
        raise ValueError(
            "Faltan columnas requeridas en la pestaña de encuesta: " + ", ".join(missing)
        )

    df = df[normalized_expected].copy()
    df.columns = [
        "fecha",
        "q1_respuesta", "q1_puntaje",
        "q2_respuesta", "q2_puntaje",
        "q3_respuesta", "q3_puntaje",
        "q4_respuesta", "q4_puntaje",
        "q5_respuesta", "q5_puntaje",
        "total", "promedio", "comentarios",
    ]

    df["fecha"] = parse_excel_date(df["fecha"])
    for col in ["q1_puntaje", "q2_puntaje", "q3_puntaje", "q4_puntaje", "q5_puntaje", "total", "promedio"]:
        df[col] = clean_score(df[col])

    q_cols = ["q1_puntaje", "q2_puntaje", "q3_puntaje", "q4_puntaje", "q5_puntaje"]
    df["total"] = df["total"].fillna(df[q_cols].sum(axis=1))
    df["promedio"] = df["promedio"].fillna(df[q_cols].mean(axis=1))

    df = df.dropna(subset=["fecha", "promedio"])
    df = df[(df["promedio"] >= 1) & (df["promedio"] <= 5)]
    if df.empty:
        raise ValueError("No se encontraron registros válidos con fecha y promedio entre 1 y 5.")

    for col in ["q1_respuesta", "q2_respuesta", "q3_respuesta", "q4_respuesta", "q5_respuesta", "comentarios"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["fecha"] = df["fecha"].dt.strftime("%Y-%m-%d %H:%M:%S")
    return df


def save_dataframe(df: pd.DataFrame, mode: str = "replace") -> int:
    records = df.to_dict(orient="records")
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
            [{**row, "imported_at": imported_at} for row in records],
        )
        conn.commit()
    return len(records)


def load_data() -> pd.DataFrame:
    with get_conn() as conn:
        df = pd.read_sql_query("SELECT * FROM encuestas ORDER BY fecha ASC, id ASC", conn)
    if df.empty:
        return df
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    numeric_cols = ["q1_puntaje", "q2_puntaje", "q3_puntaje", "q4_puntaje", "q5_puntaje", "total", "promedio"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df.dropna(subset=["fecha", "promedio"])


def status_from_score(score):
    if pd.isna(score):
        return "Sin encuestas"
    if score >= 4.80:
        return "Mantener estándar"
    if score >= 4.70:
        return "Seguimiento preventivo"
    return "Generar plan de acción"


def week_label(start_date):
    end_date = start_date + pd.Timedelta(days=6)
    iso = start_date.isocalendar()
    return f"Sem. {int(iso.week):02d} ({start_date:%d-%m} al {end_date:%d-%m})"


def add_week_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["week_start"] = out["fecha"].dt.normalize() - pd.to_timedelta(out["fecha"].dt.weekday, unit="D")
    out["week_end"] = out["week_start"] + pd.Timedelta(days=6)
    out["iso_year"] = out["fecha"].dt.isocalendar().year.astype(int)
    out["iso_week"] = out["fecha"].dt.isocalendar().week.astype(int)
    out["periodo"] = out["week_start"].apply(week_label)
    return out


def build_weekly(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    dfw = add_week_columns(df)
    q_cols = list(DIMENSION_LABELS.keys())
    agg_dict = {col: "mean" for col in q_cols + ["promedio"]}
    agg_dict["id"] = "count"
    weekly = dfw.groupby(["week_start", "week_end", "iso_year", "iso_week", "periodo"], as_index=False).agg(agg_dict)
    weekly = weekly.rename(columns={"id": "n_encuestas", "promedio": "promedio_general"})
    weekly["estado"] = weekly["promedio_general"].apply(status_from_score)
    weekly = weekly.sort_values("week_start")
    return weekly


def build_metrics(df: pd.DataFrame):
    if df.empty:
        return {
            "has_data": False,
            "total": 0,
            "global": 0,
            "excelencia": 0,
            "risk": [],
            "comments": [],
            "weekly": [],
            "recent_weeks": [],
            "dimensions": [],
            "analysis": [],
            "chart_labels": [],
            "chart_values": [],
            "radar_labels": [],
            "radar_values": [],
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "global_status": "Sin datos",
        }

    total = int(len(df))
    global_score = float(df["promedio"].mean())
    promoters = int((df["promedio"] >= 4.8).sum())
    neutrals = int(((df["promedio"] >= 4.0) & (df["promedio"] < 4.8)).sum())
    detractors = int((df["promedio"] < 4.0).sum())

    risk = [
        {"categoria": "Promotores (Alta Satisfacción)", "criterio": "4.8 - 5.0", "volumen": promoters, "representacion": pct(promoters, total), "kind": "ok"},
        {"categoria": "Neutros (Satisfacción Estándar)", "criterio": "4.0 - 4.7", "volumen": neutrals, "representacion": pct(neutrals, total), "kind": "neutral"},
        {"categoria": "Detractores (Riesgo Operativo)", "criterio": "1.0 - 3.9", "volumen": detractors, "representacion": pct(detractors, total), "kind": "danger"},
    ]

    comments_df = df[df["comentarios"].astype(str).str.len() > 0].sort_values("fecha", ascending=False)
    comments = comments_df["comentarios"].head(6).tolist()

    weekly = build_weekly(df)
    last_10 = weekly.tail(10).copy()
    recent = weekly.tail(3).copy()

    latest_week = weekly.tail(1)
    if latest_week.empty:
        dim_scores = df[list(DIMENSION_LABELS.keys())].mean()
        latest_label = "Total histórico"
    else:
        latest_start = latest_week.iloc[0]["week_start"]
        latest_label = latest_week.iloc[0]["periodo"]
        week_df = add_week_columns(df)
        week_df = week_df[week_df["week_start"] == latest_start]
        dim_scores = week_df[list(DIMENSION_LABELS.keys())].mean()

    dimensions = []
    for key, label in DIMENSION_LABELS.items():
        score = float(dim_scores.get(key, np.nan)) if not pd.isna(dim_scores.get(key, np.nan)) else None
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
        "global": global_score,
        "global_status": status_from_score(global_score),
        "excelencia": pct_number(promoters, total),
        "risk": risk,
        "comments": comments,
        "weekly": weekly,
        "recent_weeks": records_for_recent(recent),
        "dimensions": dimensions,
        "latest_label": latest_label,
        "analysis": analysis,
        "chart_labels": last_10["periodo"].tolist(),
        "chart_values": [round(float(v), 2) for v in last_10["promedio_general"].tolist()],
        "radar_labels": ["Q1 Recepción", "Q2 Calidad", "Q3 Tiempos", "Q4 Higiene", "Q5 Trato"],
        "radar_values": [round(d["score"], 2) if d["score"] is not None else 0 for d in dimensions],
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def pct_number(part, total):
    return 0 if total == 0 else round((part / total) * 100, 1)


def pct(part, total):
    return f"{pct_number(part, total):.1f}%"


def records_for_recent(recent: pd.DataFrame):
    if recent.empty:
        return []
    out = []
    for row in recent.to_dict(orient="records"):
        out.append({
            "periodo": row["periodo"],
            "n_encuestas": int(row["n_encuestas"]),
            "promedio": round(float(row["promedio_general"]), 2),
            "estado": row["estado"],
        })
    return out


def make_line_chart(metrics):
    labels = metrics["chart_labels"]
    values = metrics["chart_values"]
    fig, ax = plt.subplots(figsize=(8.0, 4.0), dpi=140)
    if values:
        ax.plot(range(len(values)), values, marker="o", linewidth=2.5, color="#d71920")
        ax.set_xticks(range(len(values)))
        ax.set_xticklabels([label.split(" ")[1] if " " in label else label for label in labels], rotation=0)
        for idx, value in enumerate(values):
            ax.annotate(f"{value:.2f}", (idx, value), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8)
    ax.set_ylim(1, 5.05)
    ax.set_ylabel("Promedio (sobre 5.0)")
    ax.set_xlabel("Semana")
    ax.grid(True, linestyle="--", alpha=0.35)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def make_radar_chart(metrics):
    labels = metrics["radar_labels"]
    values = metrics["radar_values"]
    if not values:
        values = [0, 0, 0, 0, 0]
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values_closed = values + values[:1]
    angles_closed = angles + angles[:1]
    fig = plt.figure(figsize=(5.2, 4.2), dpi=140)
    ax = plt.subplot(111, polar=True)
    ax.plot(angles_closed, values_closed, color="#d71920", linewidth=2)
    ax.fill(angles_closed, values_closed, color="#d71920", alpha=0.18)
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=7)
    ax.grid(True, alpha=0.35)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def table_style(header_bg="#303030"):
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#dddddd")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7f7f7")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ])


def section_title(text):
    return Table([[text]], colWidths=[17.0 * cm], style=TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#d71920")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))


def header_footer(canvas, doc):
    canvas.saveState()
    width, height = A4
    canvas.setStrokeColor(colors.HexColor("#d71920"))
    canvas.setLineWidth(2.0)
    canvas.line(2 * cm, height - 2.15 * cm, width - 2 * cm, height - 2.15 * cm)

    if os.path.exists(LOGO_ARAMARK_PATH):
        canvas.drawImage(LOGO_ARAMARK_PATH, 2 * cm, height - 1.75 * cm, width=3.0 * cm, height=0.55 * cm, preserveAspectRatio=True, mask="auto")
    else:
        canvas.setFont("Helvetica-Bold", 13)
        canvas.setFillColor(colors.black)
        canvas.drawString(2 * cm, height - 1.55 * cm, "★ aramark")

    if os.path.exists(LOGO_CAMPAMENTO_PATH):
        canvas.drawImage(LOGO_CAMPAMENTO_PATH, width - 7.8 * cm, height - 1.72 * cm, width=5.8 * cm, height=0.50 * cm, preserveAspectRatio=True, mask="auto")
    else:
        canvas.setFont("Helvetica-Bold", 13)
        canvas.setFillColor(colors.HexColor("#d71920"))
        canvas.drawRightString(width - 2 * cm, height - 1.55 * cm, "Aramark Campamento 5400")

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#999999"))
    canvas.line(2 * cm, 1.35 * cm, width - 2 * cm, 1.35 * cm)
    canvas.drawCentredString(width / 2, 1.0 * cm, f"Generado automáticamente a partir de datos operacionales • Evaluación Interna • {datetime.now():%Y-%m-%d}")
    canvas.restoreState()


def generate_pdf_report(df: pd.DataFrame) -> io.BytesIO:
    metrics = build_metrics(df)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=2 * cm, leftMargin=2 * cm, topMargin=2.7 * cm, bottomMargin=1.7 * cm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("Title5400", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18, alignment=0, spaceAfter=4)
    subtitle = ParagraphStyle("Subtitle5400", parent=styles["Normal"], fontSize=10, textColor=colors.HexColor("#666666"), spaceAfter=16)
    normal = ParagraphStyle("Normal5400", parent=styles["Normal"], fontSize=9, leading=11)
    quote = ParagraphStyle("Quote5400", parent=styles["Normal"], fontSize=8.5, leading=10, leftIndent=8, textColor=colors.HexColor("#555555"))

    story = []
    story.append(Paragraph("REPORTE EJECUTIVO DE CALIDAD", title))
    story.append(Paragraph("Evaluación de Satisfacción e Impacto del Factor Humano en la Operación", subtitle))
    story.append(section_title("1. Resumen Ejecutivo (Macrométricas)"))
    story.append(Spacer(1, 0.25 * cm))

    card_label = ParagraphStyle("CardLabel", parent=styles["Normal"], fontSize=7, alignment=1, textColor=colors.HexColor("#666666"), fontName="Helvetica-Bold")
    card_value = ParagraphStyle("CardValue", parent=styles["Normal"], fontSize=18, alignment=1, textColor=colors.HexColor("#222222"), fontName="Helvetica-Bold", leading=22)
    card_note = ParagraphStyle("CardNote", parent=styles["Normal"], fontSize=7, alignment=1, textColor=colors.HexColor("#3a8a43"), fontName="Helvetica-Bold")
    cards = [[
        [Paragraph("TOTAL DE EVALUACIONES", card_label), Paragraph(f"{metrics['total']}", card_value), Paragraph("Muestra Histórica", card_note)],
        [Paragraph("SATISFACCIÓN GLOBAL", card_label), Paragraph(f"{metrics['global']:.2f} / 5.0", card_value), Paragraph(status_from_score(metrics["global"]), card_note)],
        [Paragraph("ÍNDICE DE EXCELENCIA", card_label), Paragraph(f"{metrics['excelencia']:.1f}%", card_value), Paragraph("Usuarios Promotores", card_note)],
    ]]
    card_table = Table(cards, colWidths=[5.4 * cm, 5.4 * cm, 5.4 * cm])
    card_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f7f7f7")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LINEABOVE", (0, 0), (-1, 0), 2, colors.HexColor("#d71920")),
        ("BOX", (0, 0), (-1, -1), 0.25, colors.HexColor("#eeeeee")),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(card_table)
    story.append(Spacer(1, 0.45 * cm))

    story.append(section_title("2. Distribución del Riesgo Operativo"))
    risk_rows = [["CATEGORÍA DE USUARIO", "CRITERIO (NOTA)", "VOLUMEN", "REPRESENTACIÓN"]]
    for item in metrics["risk"]:
        risk_rows.append([item["categoria"], item["criterio"], item["volumen"], item["representacion"]])
    risk_table = Table(risk_rows, colWidths=[7.0 * cm, 3.4 * cm, 3.0 * cm, 3.2 * cm])
    risk_table.setStyle(table_style("#eeeeee"))
    risk_table.setStyle(TableStyle([("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#444444"))]))
    story.append(Spacer(1, 0.25 * cm))
    story.append(risk_table)
    story.append(Spacer(1, 0.45 * cm))

    story.append(section_title("3. La Voz del Usuario y Observaciones"))
    story.append(Spacer(1, 0.25 * cm))
    comments = metrics["comments"][:4] or ["Sin comentarios registrados."]
    comments_text = []
    for c in comments:
        comments_text.append(Paragraph(f'“{c}”', quote))
    analysis_text = []
    for a in metrics["analysis"] or ["Sin datos suficientes para generar análisis cualitativo."]:
        analysis_text.append(Paragraph("■ " + a, normal))
    story.append(Table([[comments_text, analysis_text]], colWidths=[8.0 * cm, 8.5 * cm], style=TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBEFORE", (0, 0), (0, 0), 2, colors.HexColor("#d71920")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ])))

    story.append(PageBreak())
    story.append(Paragraph("ANÁLISIS DE TENDENCIAS Y DESGLOSE", title))
    story.append(Paragraph("Monitoreo Semanal de Resultados Operativos", subtitle))
    story.append(section_title("4. Evolución Histórica (Últimas 10 Semanas)"))
    story.append(Spacer(1, 0.25 * cm))
    line_buf = make_line_chart(metrics)
    story.append(Image(line_buf, width=16.5 * cm, height=8.1 * cm))
    story.append(Spacer(1, 0.35 * cm))

    story.append(section_title("5. Resumen Semanas Recientes"))
    recent_rows = [["SEMANA", "N° ENCUESTAS", "NOTA PROMEDIO", "ESTADO / ACCIÓN"]]
    for item in metrics["recent_weeks"]:
        recent_rows.append([item["periodo"], item["n_encuestas"], f"{item['promedio']:.2f}", item["estado"]])
    if len(recent_rows) == 1:
        recent_rows.append(["Sin datos", 0, "-", "Sin encuestas"])
    recent_table = Table(recent_rows, colWidths=[6.2 * cm, 3.4 * cm, 3.4 * cm, 3.8 * cm])
    recent_table.setStyle(table_style("#eeeeee"))
    recent_table.setStyle(TableStyle([("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#444444"))]))
    story.append(Spacer(1, 0.25 * cm))
    story.append(recent_table)
    story.append(Spacer(1, 0.35 * cm))

    story.append(section_title(f"6. Desempeño Específico por Dimensiones ({metrics.get('latest_label', 'Última medición')})"))
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph("Una calificación por debajo de 4.70 se considera un foco de atención preventivo.", normal))

    story.append(PageBreak())
    radar_buf = make_radar_chart(metrics)
    dim_rows = []
    for item in metrics["dimensions"]:
        score_text = "-" if item["score"] is None else f"{item['score']:.2f}" + (" (Foco)" if item["foco"] else "")
        dim_rows.append([item["label"], score_text])
    dim_table = Table(dim_rows, colWidths=[9.0 * cm, 4.0 * cm], style=TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("TEXTCOLOR", (1, 0), (1, -1), colors.HexColor("#d71920")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(Table([[Image(radar_buf, width=7.5 * cm, height=6.2 * cm), dim_table]], colWidths=[7.8 * cm, 8.7 * cm], style=TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")])) )

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
    buffer.seek(0)
    return buffer


def build_export_workbook(df: pd.DataFrame) -> io.BytesIO:
    metrics = build_metrics(df)
    output = io.BytesIO()
    raw = df.copy()
    if not raw.empty:
        raw["fecha"] = raw["fecha"].dt.strftime("%Y-%m-%d %H:%M:%S")
    weekly = build_weekly(df)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        raw.to_excel(writer, index=False, sheet_name="Encuesta de Satisfacción")
        if not weekly.empty:
            weekly_export = weekly.copy()
            for col in ["week_start", "week_end"]:
                weekly_export[col] = weekly_export[col].dt.strftime("%Y-%m-%d")
            weekly_export.to_excel(writer, index=False, sheet_name="KPI Semanal")
        resumen = pd.DataFrame([
            ["Total de evaluaciones", metrics["total"]],
            ["Satisfacción global", round(metrics["global"], 2)],
            ["Índice de excelencia", f"{metrics['excelencia']:.1f}%"],
            ["Fecha de generación", metrics["generated_at"]],
        ], columns=["Métrica", "Valor"])
        resumen.to_excel(writer, index=False, sheet_name="Resumen Ejecutivo")
        dims = pd.DataFrame(metrics["dimensions"])
        dims.to_excel(writer, index=False, sheet_name="Dimensiones")
    output.seek(0)
    return output


@app.route("/", methods=["GET"])
def dashboard():
    df = load_data()
    metrics = build_metrics(df)
    return render_template("dashboard.html", title=APP_TITLE, metrics=metrics)


@app.route("/importar", methods=["POST"])
def importar():
    uploaded = request.files.get("archivo")
    mode = request.form.get("modo", "replace")
    if not uploaded or not uploaded.filename:
        flash("Selecciona un archivo Excel para importar.", "error")
        return redirect(url_for("dashboard"))
    try:
        df = read_uploaded_workbook(uploaded)
        count = save_dataframe(df, mode=mode)
        accion = "reemplazados" if mode == "replace" else "agregados"
        flash(f"Importación exitosa: {count} registros {accion}.", "success")
    except Exception as exc:
        flash(f"Error al importar: {exc}", "error")
    return redirect(url_for("dashboard"))


@app.route("/exportar/pdf")
def exportar_pdf():
    df = load_data()
    if df.empty:
        flash("No hay datos para exportar.", "error")
        return redirect(url_for("dashboard"))
    pdf = generate_pdf_report(df)
    return send_file(pdf, as_attachment=True, download_name="reporte_encuesta_satisfaccion_5400.pdf", mimetype="application/pdf")


@app.route("/exportar/excel")
def exportar_excel():
    df = load_data()
    if df.empty:
        flash("No hay datos para exportar.", "error")
        return redirect(url_for("dashboard"))
    output = build_export_workbook(df)
    return send_file(output, as_attachment=True, download_name="export_encuesta_satisfaccion_5400.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/exportar/csv")
def exportar_csv():
    df = load_data()
    if df.empty:
        flash("No hay datos para exportar.", "error")
        return redirect(url_for("dashboard"))
    raw = df.copy()
    raw["fecha"] = raw["fecha"].dt.strftime("%Y-%m-%d %H:%M:%S")
    output = io.StringIO()
    raw.to_csv(output, index=False, quoting=csv.QUOTE_MINIMAL)
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=export_encuesta_satisfaccion_5400.csv"})


@app.route("/plantilla")
def plantilla():
    output = io.BytesIO()
    sample = pd.DataFrame(columns=EXPECTED_COLUMNS)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Encuesta de Satisfacción")
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="plantilla_encuesta_satisfaccion_5400.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.context_processor
def inject_helpers():
    return {
        "json_dumps": lambda obj: json.dumps(obj, ensure_ascii=False),
        "status_class": lambda s: "ok" if s == "Mantener estándar" else ("warn" if s == "Seguimiento preventivo" else "danger"),
    }


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)

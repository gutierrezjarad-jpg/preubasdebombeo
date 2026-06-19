
from __future__ import annotations

from pathlib import Path
from datetime import datetime
from io import BytesIO
import textwrap

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    SimpleDocTemplate, BaseDocTemplate, PageTemplate, Frame,
    Paragraph, Spacer, Table, TableStyle, Image as RLImage,
    PageBreak, KeepTogether
)

# =============================================================================
# CONFIGURACIÓN GENERAL
# =============================================================================

APP_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = APP_DIR / "assets"
EXPORTS_DIR = APP_DIR / "exports"
ASSETS_DIR.mkdir(exist_ok=True)
EXPORTS_DIR.mkdir(exist_ok=True)

LOGO_PATH = ASSETS_DIR / "logo_irrisal.jpg"

COMPANY_DEFAULTS = {
    "empresa": "Irrisal Consulting Ltda.",
    "direccion": "San Martín 553, oficina 901, Concepción, Región del Biobío, Chile.",
    "celular": "+56 9 6796 0884",
    "correo": "irrisalconsulting@gmail.com",
}

st.set_page_config(
    page_title="Pruebas de Bombeo - Irrisal Consulting",
    layout="wide"
)

# =============================================================================
# UTILIDADES
# =============================================================================

def clean_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def safe_text(value, default: str = "Dato no informado") -> str:
    if value is None:
        return default
    if isinstance(value, float) and pd.isna(value):
        return default
    value = str(value).strip()
    return value if value else default


def fmt(value, suffix: str = "", decimals: int = 2, empty: str = "No evaluable") -> str:
    if value is None:
        return empty
    try:
        if pd.isna(value):
            return empty
    except Exception:
        pass
    if isinstance(value, (int, float, np.number)):
        return f"{float(value):.{decimals}f}{suffix}"
    return f"{value}{suffix}"


def df_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


def add_warning(warnings: list[str], condition: bool, message: str):
    if condition:
        warnings.append(message)


# =============================================================================
# CÁLCULOS
# =============================================================================

def calculate_flow_stats(df: pd.DataFrame) -> dict:
    if df.empty or "Caudal_L_s" not in df.columns:
        return {"q_mean": None, "q_min": None, "q_max": None, "q_std": None, "q_var_pct": None, "is_constant": None}

    q = pd.to_numeric(df["Caudal_L_s"], errors="coerce").dropna()
    if q.empty:
        return {"q_mean": None, "q_min": None, "q_max": None, "q_std": None, "q_var_pct": None, "is_constant": None}

    q_mean = float(q.mean())
    q_min = float(q.min())
    q_max = float(q.max())
    q_std = float(q.std(ddof=0)) if len(q) > 1 else 0.0
    q_var_pct = ((q_max - q_min) / q_mean) * 100 if q_mean > 0 else None

    # Criterio simple: diferencia max-min <= 10% del caudal medio.
    is_constant = q_var_pct is not None and q_var_pct <= 10

    return {
        "q_mean": q_mean,
        "q_min": q_min,
        "q_max": q_max,
        "q_std": q_std,
        "q_var_pct": q_var_pct,
        "is_constant": is_constant,
    }


def calculate_pumped_volume(df: pd.DataFrame) -> float | None:
    if df.empty or not {"Tiempo_min", "Caudal_L_s"}.issubset(df.columns):
        return None

    data = df_numeric(df, ["Tiempo_min", "Caudal_L_s"]).dropna(subset=["Tiempo_min", "Caudal_L_s"])
    data = data.sort_values("Tiempo_min")
    if len(data) < 2:
        return None

    total = 0.0
    times = data["Tiempo_min"].to_numpy()
    flows = data["Caudal_L_s"].to_numpy()

    for i in range(1, len(data)):
        dt = times[i] - times[i - 1]
        if dt < 0:
            continue
        q_avg = (flows[i] + flows[i - 1]) / 2
        total += q_avg * dt * 60 / 1000

    return float(total)


def evaluate_stabilization(df: pd.DataFrame) -> dict:
    result = {
        "evaluable": False,
        "slope_cm_h": None,
        "meets": False,
        "message": "No evaluable: faltan datos suficientes.",
    }

    if df.empty or not {"Tiempo_min", "Nivel_m"}.issubset(df.columns):
        return result

    data = df_numeric(df, ["Tiempo_min", "Nivel_m"]).dropna(subset=["Tiempo_min", "Nivel_m"])
    data = data.sort_values("Tiempo_min")
    if len(data) < 2:
        return result

    duration = float(data["Tiempo_min"].max() - data["Tiempo_min"].min())
    if duration < 180:
        result["message"] = "No evaluable: se requieren al menos 180 minutos de datos para evaluar tendencia de estabilización."
        return result

    tmax = data["Tiempo_min"].max()
    last = data[data["Tiempo_min"] >= tmax - 180].copy()
    if len(last) < 2:
        result["message"] = "No evaluable: no hay suficientes puntos en los últimos 180 minutos."
        return result

    x_h = last["Tiempo_min"].to_numpy() / 60.0
    y_m = last["Nivel_m"].to_numpy()
    slope_m_h = np.polyfit(x_h, y_m, 1)[0]
    slope_cm_h = float(slope_m_h * 100)

    result["evaluable"] = True
    result["slope_cm_h"] = slope_cm_h
    result["meets"] = slope_cm_h <= 2

    if result["meets"]:
        result["message"] = "Presenta estabilización o franca tendencia según criterio ≤ 2 cm/h en los últimos 180 minutos."
    else:
        result["message"] = "No presenta estabilización según criterio ≤ 2 cm/h en los últimos 180 minutos."

    return result


def calculate_recovery(recovery_df: pd.DataFrame, static_level: float, final_dynamic_level: float) -> pd.DataFrame:
    data = recovery_df.copy()
    if data.empty or not {"Tiempo_min", "Nivel_m"}.issubset(data.columns):
        data["Recuperacion_pct"] = np.nan
        return data

    data = df_numeric(data, ["Tiempo_min", "Nivel_m"])
    denom = final_dynamic_level - static_level
    if denom <= 0:
        data["Recuperacion_pct"] = np.nan
        return data

    data["Recuperacion_pct"] = ((final_dynamic_level - data["Nivel_m"]) / denom) * 100
    data["Recuperacion_pct"] = data["Recuperacion_pct"].clip(lower=0, upper=100)
    return data


def time_to_recovery(df: pd.DataFrame, target_pct: float) -> float | None:
    if df.empty or "Recuperacion_pct" not in df.columns:
        return None
    valid = df_numeric(df, ["Tiempo_min", "Recuperacion_pct"]).dropna(subset=["Tiempo_min", "Recuperacion_pct"])
    reached = valid[valid["Recuperacion_pct"] >= target_pct]
    if reached.empty:
        return None
    return float(reached["Tiempo_min"].iloc[0])


def build_calculations(pumping_df: pd.DataFrame, recovery_df: pd.DataFrame, static_level: float) -> tuple[dict, pd.DataFrame]:
    pumping = df_numeric(pumping_df, ["Tiempo_min", "Nivel_m", "Caudal_L_s"])
    pumping_valid = pumping.dropna(subset=["Tiempo_min", "Nivel_m"]).sort_values("Tiempo_min")

    duration_min = None
    final_dynamic = None
    drawdown = None

    if not pumping_valid.empty:
        duration_min = float(pumping_valid["Tiempo_min"].max() - pumping_valid["Tiempo_min"].min())
        final_dynamic = float(pumping_valid["Nivel_m"].iloc[-1])
        drawdown = final_dynamic - static_level if final_dynamic >= static_level else None

    flow_stats = calculate_flow_stats(pumping)
    specific_capacity = flow_stats["q_mean"] / drawdown if drawdown and drawdown > 0 and flow_stats["q_mean"] is not None else None
    volume_m3 = calculate_pumped_volume(pumping)
    stabilization = evaluate_stabilization(pumping)

    recovery_with_pct = calculate_recovery(recovery_df, static_level, final_dynamic) if final_dynamic is not None else recovery_df.copy()
    if "Recuperacion_pct" not in recovery_with_pct.columns:
        recovery_with_pct["Recuperacion_pct"] = np.nan

    rec_valid = pd.to_numeric(recovery_with_pct["Recuperacion_pct"], errors="coerce").dropna()
    recovery_max = float(rec_valid.max()) if not rec_valid.empty else None

    calculations = {
        **flow_stats,
        "duration_min": duration_min,
        "final_dynamic": final_dynamic,
        "drawdown": drawdown,
        "specific_capacity": specific_capacity,
        "volume_m3": volume_m3,
        "stabilization_evaluable": stabilization["evaluable"],
        "stabilization_meets": stabilization["meets"],
        "slope_cm_h": stabilization["slope_cm_h"],
        "stabilization_message": stabilization["message"],
        "recovery_max": recovery_max,
        "t75": time_to_recovery(recovery_with_pct, 75),
        "t90": time_to_recovery(recovery_with_pct, 90),
        "t100": time_to_recovery(recovery_with_pct, 100),
    }

    return calculations, recovery_with_pct


# =============================================================================
# GRÁFICOS
# =============================================================================

def make_line_chart_image(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    title: str,
    x_label: str,
    y_label: str,
    invert_y: bool = False,
) -> BytesIO | None:
    if df is None or df.empty or not {x_col, y_col}.issubset(df.columns):
        return None

    data = df_numeric(df, [x_col, y_col]).dropna(subset=[x_col, y_col]).sort_values(x_col)
    if len(data) < 2:
        return None

    fig, ax = plt.subplots(figsize=(8.0, 4.1))
    ax.plot(data[x_col], data[y_col], marker="o", linewidth=1.5, markersize=3.4)
    ax.set_title(title, fontsize=11)
    ax.set_xlabel(x_label, fontsize=9)
    ax.set_ylabel(y_label, fontsize=9)
    ax.grid(True, alpha=0.32)

    if invert_y:
        ax.invert_yaxis()

    fig.tight_layout()
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=180)
    plt.close(fig)
    buffer.seek(0)
    return buffer


def image_flowable(image_buffer: BytesIO | None, width_cm: float = 16.5, height_cm: float = 8.4):
    if image_buffer is None:
        return Paragraph("Gráfico no disponible: datos insuficientes.", get_styles()["Body"])
    image_buffer.seek(0)
    img = rl_image_preserve_aspect(image_buffer, max_width_cm=width_cm, max_height_cm=height_cm)
    if img is None:
        return Paragraph("Gráfico no disponible: error al procesar imagen.", get_styles()["Body"])
    return img



# =============================================================================
# IMÁGENES CON PROPORCIÓN CONSERVADA
# =============================================================================

def _image_source_from_uploaded(uploaded_file):
    if uploaded_file is None:
        return None
    try:
        return BytesIO(uploaded_file.getvalue())
    except Exception:
        return None


def _scaled_dimensions(img_width_px, img_height_px, max_width_cm: float, max_height_cm: float):
    """
    Calcula dimensiones para insertar imagen sin deformarla.
    """
    max_w = max_width_cm * cm
    max_h = max_height_cm * cm

    if img_width_px <= 0 or img_height_px <= 0:
        return max_w, max_h

    scale = min(max_w / img_width_px, max_h / img_height_px)
    return img_width_px * scale, img_height_px * scale


def rl_image_preserve_aspect(source, max_width_cm: float, max_height_cm: float):
    """
    Crea un Image flowable de ReportLab preservando proporción.
    Acepta Path, str, BytesIO o UploadedFile convertido a BytesIO.
    """
    if source is None:
        return None

    try:
        if isinstance(source, Path):
            source_for_reader = str(source)
            source_for_image = str(source)
        elif isinstance(source, str):
            source_for_reader = source
            source_for_image = source
        elif isinstance(source, BytesIO):
            data = source.getvalue()
            source_for_reader = BytesIO(data)
            source_for_image = BytesIO(data)
        else:
            return None

        reader = ImageReader(source_for_reader)
        img_w, img_h = reader.getSize()
        width, height = _scaled_dimensions(img_w, img_h, max_width_cm, max_height_cm)
        return RLImage(source_for_image, width=width, height=height)
    except Exception:
        return None


def draw_header_logo(canvas, x, y, max_w, max_h):
    """
    Dibuja el logo en canvas sin deformarlo.
    """
    if not LOGO_PATH.exists():
        return False
    try:
        canvas.drawImage(
            str(LOGO_PATH),
            x,
            y,
            width=max_w,
            height=max_h,
            preserveAspectRatio=True,
            anchor="w",
            mask="auto",
        )
        return True
    except Exception:
        return False


# =============================================================================
# PDF: ESTILOS Y COMPONENTES
# =============================================================================

def get_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="CoverTitle", parent=styles["Title"], fontSize=18, leading=22,
        alignment=TA_CENTER, textColor=colors.HexColor("#006b2e"), spaceAfter=14
    ))
    styles.add(ParagraphStyle(
        name="CoverSubtitle", parent=styles["Normal"], fontSize=12, leading=15,
        alignment=TA_CENTER, spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        name="SectionTitle", parent=styles["Heading2"], fontSize=12.5, leading=15,
        textColor=colors.HexColor("#006b2e"), spaceBefore=10, spaceAfter=7
    ))
    styles.add(ParagraphStyle(
        name="Body", parent=styles["Normal"], fontSize=8.7, leading=11.2, alignment=TA_LEFT
    ))
    styles.add(ParagraphStyle(
        name="Small", parent=styles["Normal"], fontSize=7.2, leading=9.2
    ))
    styles.add(ParagraphStyle(
        name="FigureCaption", parent=styles["Normal"], fontSize=7.5, leading=9.4,
        alignment=TA_CENTER, textColor=colors.HexColor("#555555"), spaceAfter=6
    ))
    return styles


def report_header_footer(canvas, doc, company: dict):
    canvas.saveState()
    width, height = A4

    green = colors.HexColor("#006b2e")
    gray = colors.HexColor("#444444")

    # Header: logo pequeño + nombre empresa
    logo_drawn = draw_header_logo(
        canvas,
        x=1.45 * cm,
        y=height - 1.05 * cm,
        max_w=1.25 * cm,
        max_h=0.72 * cm,
    )

    text_x = 2.85 * cm if logo_drawn else 1.5 * cm

    canvas.setFont("Helvetica-Bold", 7.4)
    canvas.setFillColor(green)
    canvas.drawString(text_x, height - 0.72 * cm, safe_text(company.get("empresa"), ""))

    canvas.setStrokeColor(green)
    canvas.setLineWidth(0.65)
    canvas.line(1.4 * cm, height - 1.22 * cm, width - 1.4 * cm, height - 1.22 * cm)

    # Footer: contacto + número de página
    canvas.setStrokeColor(colors.HexColor("#bbbbbb"))
    canvas.setLineWidth(0.35)
    canvas.line(1.4 * cm, 1.02 * cm, width - 1.4 * cm, 1.02 * cm)

    canvas.setFont("Helvetica", 6.5)
    canvas.setFillColor(gray)
    footer = f"{safe_text(company.get('direccion'), '')} | {safe_text(company.get('celular'), '')} | {safe_text(company.get('correo'), '')}"
    canvas.drawString(1.5 * cm, 0.68 * cm, footer[:130])
    canvas.drawRightString(width - 1.5 * cm, 0.68 * cm, f"Página {doc.page}")

    canvas.restoreState()

def make_table(data, col_widths=None, header=False, font_size=7.2, first_col_bold=True):
    wrapped = []
    styles = get_styles()
    for row in data:
        wrapped.append([Paragraph(safe_text(cell, ""), styles["Small"]) for cell in row])

    table = Table(wrapped, colWidths=col_widths)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.28, colors.HexColor("#999999")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if first_col_bold:
        style += [
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#edf7ef")),
        ]
    if header:
        style += [
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9ead3")),
        ]
    table.setStyle(TableStyle(style))
    return table


def df_to_pdf_table(df: pd.DataFrame, max_rows: int = 55, font_size: float = 6.0):
    styles = get_styles()
    if df is None or df.empty:
        return Paragraph("Sin datos ingresados.", styles["Body"])

    show = df.copy().head(max_rows).fillna("")
    data = [list(show.columns)] + show.astype(str).values.tolist()

    # Ancho flexible para A4 normal
    ncols = max(1, len(data[0]))
    col_width = min(16.5 / ncols, 4.5) * cm

    table = Table(data, colWidths=[col_width] * ncols, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#999999")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9ead3")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return table


def uploaded_image_flowable(uploaded_file, width_cm: float = 15.5, height_cm: float = 8.5):
    """
    Inserta imágenes subidas por el usuario sin deformarlas.
    """
    src = _image_source_from_uploaded(uploaded_file)
    if src is None:
        return None
    return rl_image_preserve_aspect(src, width_cm, height_cm)

def make_simple_well_scheme(capture: dict):
    """
    Esquema constructivo simple generado automáticamente.
    No inventa cribas ni profundidad de bomba si esos datos no fueron ingresados.
    Si falta profundidad total, no genera esquema.
    """
    styles = get_styles()

    total_depth = capture.get("profundidad_total")
    static_level = capture.get("nivel_estatico")
    pump_depth = capture.get("profundidad_bomba")
    screen_from = capture.get("criba_desde")
    screen_to = capture.get("criba_hasta")

    try:
        total_depth = float(total_depth)
        if total_depth <= 0:
            return None
    except Exception:
        return None

    try:
        static_level = float(static_level)
    except Exception:
        static_level = None

    try:
        pump_depth = float(pump_depth)
        has_pump = pump_depth > 0
    except Exception:
        pump_depth = None
        has_pump = False

    try:
        screen_from = float(screen_from)
        screen_to = float(screen_to)
        has_screen = screen_from > 0 and screen_to > screen_from
    except Exception:
        screen_from = None
        screen_to = None
        has_screen = False

    fig, ax = plt.subplots(figsize=(3.2, 7.0))
    ax.set_xlim(0, 10)
    ax.set_ylim(total_depth + 2, -2)
    ax.axis("off")

    # Terreno
    ax.plot([0.5, 9.5], [0, 0], color="saddlebrown", linewidth=2)
    ax.text(6.2, -0.35, "Nivel de terreno", fontsize=7)

    # Tubería
    ax.plot([4, 4], [0, total_depth], color="black", linewidth=2)
    ax.plot([6, 6], [0, total_depth], color="black", linewidth=2)
    ax.plot([4, 6], [total_depth, total_depth], color="black", linewidth=2)

    # Nivel estático
    if static_level is not None:
        ax.plot([3.7, 6.3], [static_level, static_level], color="blue", linewidth=1.8)
        ax.text(6.5, static_level, f"Nivel estático {static_level:.2f} m", fontsize=7, va="center")
    else:
        ax.text(6.5, total_depth * 0.2, "Nivel estático: no informado", fontsize=7, va="center")

    # Cribas / tramo ranurado
    if has_screen:
        ax.fill_between([4, 6], screen_from, screen_to, color="#d9ead3", alpha=0.75)
        for y in np.linspace(screen_from, screen_to, 12):
            ax.plot([4.05, 5.95], [y, y], color="gray", linewidth=0.7)
        ax.text(6.5, (screen_from + screen_to) / 2, f"Cribas {screen_from:.1f}-{screen_to:.1f} m", fontsize=7, va="center")
    else:
        ax.text(6.5, total_depth * 0.65, "Cribas/tramo filtrante: no informado", fontsize=7, va="center")

    # Bomba
    if has_pump:
        ax.scatter([5], [pump_depth], marker="s", s=42, color="black")
        ax.text(6.5, pump_depth, f"Bomba {pump_depth:.1f} m", fontsize=7, va="center")
    else:
        ax.text(6.5, total_depth * 0.82, "Profundidad bomba: no informada", fontsize=7, va="center")

    ax.text(6.5, total_depth, f"Profundidad total {total_depth:.1f} m", fontsize=7, va="center")
    ax.set_title("Esquema constructivo referencial", fontsize=9)

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf

# =============================================================================
# GENERACIÓN PDF
# =============================================================================

def generate_conclusions(capture: dict, calculations: dict, warnings: list[str], test_mode: str) -> list[str]:
    tipo = safe_text(capture.get("tipo"), "")
    duration = calculations.get("duration_min")
    q_mean = calculations.get("q_mean")
    stab_meets = calculations.get("stabilization_meets")
    rec_max = calculations.get("recovery_max")

    conclusions = []

    if q_mean is not None:
        conclusions.append(
            f"Durante el periodo medido, la prueba se desarrolló con un caudal promedio de {fmt(q_mean, ' L/s')}."
        )

    if calculations.get("drawdown") is not None:
        conclusions.append(
            f"El abatimiento final calculado fue de {fmt(calculations.get('drawdown'), ' m')}, con un caudal específico de {fmt(calculations.get('specific_capacity'), ' L/s/m')}."
        )

    if calculations.get("stabilization_evaluable"):
        if stab_meets:
            conclusions.append(
                "La captación presenta estabilización o franca tendencia a estabilización bajo el criterio de variación ≤ 2 cm/h en los últimos 180 minutos evaluados."
            )
        else:
            conclusions.append(
                "La captación no presenta estabilización bajo el criterio de variación ≤ 2 cm/h en los últimos 180 minutos evaluados."
            )
    else:
        conclusions.append(
            "La estabilización no es evaluable con los datos disponibles."
        )

    if rec_max is not None:
        conclusions.append(
            f"La recuperación máxima observada alcanzó {fmt(rec_max, ' %')}. La interpretación de recuperación debe considerar la duración real del seguimiento posterior al bombeo."
        )

    if tipo == "Pozo profundo" and duration is not None and duration < 1440:
        conclusions.append(
            "La prueba corresponde a un ensayo de duración menor a 24 horas para pozo profundo. Sus resultados permiten una evaluación operativa del comportamiento de la captación durante el periodo medido, pero no reemplazan una prueba estándar de 24 horas cuando esta sea exigida."
        )

    if "abreviado" in test_mode.lower() or (duration is not None and duration <= 180):
        conclusions.append(
            "El ensayo abreviado debe interpretarse como antecedente técnico preliminar y sus conclusiones deben restringirse al periodo efectivamente medido."
        )

    conclusions.append(
        "El informe se basa exclusivamente en datos ingresados o importados por el usuario. El sistema no rellena mediciones faltantes ni presenta datos estimados como medidos."
    )

    return conclusions


def generate_recommendations(capture: dict, calculations: dict, warnings: list[str]) -> list[str]:
    recs = []
    tipo = safe_text(capture.get("tipo"), "")
    rec_max = calculations.get("recovery_max")

    if not calculations.get("stabilization_evaluable"):
        recs.append("Registrar al menos 180 minutos de mediciones continuas para evaluar tendencia de estabilización.")
    elif calculations.get("stabilization_meets"):
        recs.append("Mantener como referencia el caudal ensayado, siempre que las condiciones de operación y recuperación se mantengan similares.")
    else:
        recs.append("Evaluar una reducción del caudal de operación o repetir la prueba con mayor duración para definir un caudal más conservador.")

    if rec_max is None or rec_max < 75:
        recs.append("Extender la medición de recuperación hasta alcanzar al menos 75% de recuperación, idealmente hasta recuperación completa o estabilización clara.")

    if tipo == "Puntera":
        recs.append("Para punteras, habilitar piezómetro de control para medición de niveles, evitando interpretar niveles directamente desde la tubería de extracción.")

    if calculations.get("is_constant") is False:
        recs.append("Mejorar el control del caudal durante la prueba para asegurar condiciones de gasto constante.")

    recs.append("Conservar respaldo de planillas, fotografías, ubicación y equipos utilizados como anexos del expediente técnico.")

    return recs


def make_pdf(
    company: dict,
    project: dict,
    capture: dict,
    stratigraphy_df: pd.DataFrame,
    equipment: dict,
    methodology: dict,
    pumping_df: pd.DataFrame,
    recovery_df: pd.DataFrame,
    calculations: dict,
    warnings: list[str],
    location_image=None,
    scheme_image=None,
) -> bytes:
    buffer = BytesIO()
    styles = get_styles()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.85 * cm,
        bottomMargin=1.45 * cm,
    )

    def later_pages(canvas, doc_obj):
        report_header_footer(canvas, doc_obj, company)

    story = []

    # -------------------------------------------------------------------------
    # PORTADA
    # -------------------------------------------------------------------------
    if LOGO_PATH.exists():
        logo_flow = rl_image_preserve_aspect(LOGO_PATH, max_width_cm=6.2, max_height_cm=3.4)
        if logo_flow:
            story.append(logo_flow)
            story.append(Spacer(1, 0.35 * cm))

    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("INFORME DE PRUEBA DE BOMBEO", styles["CoverTitle"]))
    story.append(Paragraph(f"<b>{safe_text(project.get('identificacion'), 'Captación subterránea')}</b>", styles["CoverSubtitle"]))
    story.append(Spacer(1, 0.5 * cm))

    cover_data = [
        ["Cliente / Beneficiario", project.get("cliente")],
        ["Proyecto", project.get("nombre_proyecto")],
        ["Sector / Predio", project.get("sector")],
        ["Comuna", project.get("comuna")],
        ["Región", project.get("region")],
        ["Fecha de prueba", methodology.get("fecha_prueba")],
        ["Fecha de emisión", datetime.now().strftime("%d-%m-%Y")],
        ["Consultor responsable", project.get("consultor")],
    ]
    story.append(make_table(cover_data, col_widths=[5.2 * cm, 11.2 * cm]))
    story.append(Spacer(1, 1.0 * cm))
    story.append(Paragraph(f"<b>{safe_text(company.get('empresa'))}</b>", styles["CoverSubtitle"]))
    story.append(Paragraph(safe_text(company.get("direccion")), styles["CoverSubtitle"]))
    story.append(Paragraph(f"Celular: {safe_text(company.get('celular'))} | Correo: {safe_text(company.get('correo'))}", styles["CoverSubtitle"]))
    story.append(PageBreak())

    # -------------------------------------------------------------------------
    # 1. INTRODUCCIÓN
    # -------------------------------------------------------------------------
    story.append(Paragraph("1. Introducción", styles["SectionTitle"]))
    intro = (
        "El presente informe resume los antecedentes de la captación, su habilitación, "
        "la metodología de prueba de bombeo, los registros de nivel y caudal, la recuperación "
        "posterior y los resultados calculados a partir de los datos ingresados. "
        "La interpretación se limita al periodo efectivamente medido y a la calidad de los datos disponibles."
    )
    story.append(Paragraph(intro, styles["Body"]))

    # -------------------------------------------------------------------------
    # 2. ANTECEDENTES GENERALES
    # -------------------------------------------------------------------------
    story.append(Paragraph("2. Antecedentes generales", styles["SectionTitle"]))
    general_data = [
        ["Cliente", project.get("cliente")],
        ["Proyecto", project.get("nombre_proyecto")],
        ["Identificación de captación", project.get("identificacion")],
        ["Sector / Predio", project.get("sector")],
        ["Comuna", project.get("comuna")],
        ["Región", project.get("region")],
        ["Consultor responsable", project.get("consultor")],
        ["Observaciones generales", project.get("observaciones")],
    ]
    story.append(make_table(general_data, col_widths=[5.2 * cm, 11.2 * cm]))

    # -------------------------------------------------------------------------
    # 3. UBICACIÓN Y HABILITACIÓN
    # -------------------------------------------------------------------------
    story.append(Paragraph("3. Ubicación y habilitación de la captación", styles["SectionTitle"]))
    cap_data = [
        ["Tipo de captación", capture.get("tipo")],
        ["Coordenada UTM Norte", capture.get("utm_norte")],
        ["Coordenada UTM Este", capture.get("utm_este")],
        ["Datum / Huso", f"{safe_text(capture.get('datum'))} / {safe_text(capture.get('huso'))}"],
        ["Condición", capture.get("condicion")],
        ["Profundidad total", fmt(capture.get("profundidad_total"), " m")],
        ["Diámetro perforación", capture.get("diametro_perforacion")],
        ["Diámetro entubación", capture.get("diametro_entubacion")],
        ["Material / espesor tubería", capture.get("material_tuberia")],
        ["Altura sobre terreno", capture.get("altura_sobre_terreno")],
        ["Nivel estático inicial", fmt(capture.get("nivel_estatico"), " m")],
        ["Cribas / ranuras", (f"Desde {capture.get('criba_desde')} m hasta {capture.get('criba_hasta')} m" if capture.get("criba_desde") and capture.get("criba_hasta") else "No informado")],
        ["Tubería ciega", capture.get("tuberia_ciega")],
        ["Profundidad de bomba", fmt(capture.get("profundidad_bomba"), " m")],
        ["Tubería extracción/succión", capture.get("tuberia_extraccion")],
        ["Observaciones", capture.get("observaciones")],
    ]
    story.append(make_table(cap_data, col_widths=[5.2 * cm, 11.2 * cm]))

    if location_image is not None:
        loc_flow = uploaded_image_flowable(location_image, width_cm=16.0, height_cm=9.0)
        if loc_flow:
            story.append(Spacer(1, 0.35 * cm))
            story.append(loc_flow)
            story.append(Paragraph("Figura 1. Croquis o imagen de ubicación de la captación.", styles["FigureCaption"]))

    story.append(Paragraph("4. Esquema constructivo", styles["SectionTitle"]))
    scheme_flow = uploaded_image_flowable(scheme_image, width_cm=9.0, height_cm=12.0) if scheme_image else None
    if scheme_flow is None:
        scheme_buf = make_simple_well_scheme(capture)
        if scheme_buf is not None:
            scheme_flow = rl_image_preserve_aspect(scheme_buf, max_width_cm=8.2, max_height_cm=12.0)
            story.append(Paragraph("Esquema referencial generado automáticamente a partir de los datos ingresados. No reemplaza plano constructivo real.", styles["Small"]))
        else:
            scheme_flow = Paragraph("Esquema constructivo no generado: falta profundidad total o datos mínimos de captación.", styles["Body"])
    story.append(scheme_flow)
    story.append(Paragraph("Figura 2. Esquema constructivo referencial de la captación.", styles["FigureCaption"]))

    # -------------------------------------------------------------------------
    # 5. ESTRATIGRAFÍA
    # -------------------------------------------------------------------------
    story.append(Paragraph("5. Estratigrafía", styles["SectionTitle"]))
    story.append(Paragraph(
        "La estratigrafía ingresada se presenta como antecedente descriptivo del material perforado o reconocido durante la habilitación.",
        styles["Body"]
    ))
    story.append(df_to_pdf_table(stratigraphy_df, max_rows=35, font_size=6.4))

    # -------------------------------------------------------------------------
    # 6. EQUIPOS Y METODOLOGÍA
    # -------------------------------------------------------------------------
    story.append(Paragraph("6. Equipos utilizados", styles["SectionTitle"]))
    eq_data = [
        ["Bomba", equipment.get("bomba")],
        ["Potencia", equipment.get("potencia")],
        ["Tubería de extracción", capture.get("tuberia_extraccion")],
        ["Medidor de caudal", equipment.get("medidor_caudal")],
        ["Instrumento de nivel", equipment.get("instrumento_nivel")],
        ["Generador", equipment.get("generador")],
        ["Observaciones", equipment.get("observaciones")],
    ]
    story.append(make_table(eq_data, col_widths=[5.2 * cm, 11.2 * cm]))

    story.append(Paragraph("7. Metodología de prueba de bombeo", styles["SectionTitle"]))
    met_data = [
        ["Modo de prueba", methodology.get("modo_prueba")],
        ["Fecha de prueba", methodology.get("fecha_prueba")],
        ["Hora inicio bombeo", methodology.get("hora_inicio")],
        ["Hora término bombeo", methodology.get("hora_termino")],
        ["Duración registrada", fmt(calculations.get("duration_min"), " min", decimals=0)],
        ["Caudal objetivo", methodology.get("caudal_objetivo")],
        ["Método de medición de caudal", methodology.get("metodo_caudal")],
        ["Método de medición de niveles", methodology.get("metodo_nivel")],
        ["Frecuencia de medición", methodology.get("frecuencia")],
        ["Observaciones metodológicas", methodology.get("observaciones")],
    ]
    story.append(make_table(met_data, col_widths=[5.2 * cm, 11.2 * cm]))

    # -------------------------------------------------------------------------
    # 8. RESULTADOS
    # -------------------------------------------------------------------------
    story.append(Paragraph("8. Resultados calculados", styles["SectionTitle"]))
    results_data = [
        ["Caudal promedio", fmt(calculations.get("q_mean"), " L/s")],
        ["Caudal mínimo", fmt(calculations.get("q_min"), " L/s")],
        ["Caudal máximo", fmt(calculations.get("q_max"), " L/s")],
        ["Variación relativa de caudal", fmt(calculations.get("q_var_pct"), " %")],
        ["Nivel estático inicial", fmt(capture.get("nivel_estatico"), " m")],
        ["Nivel dinámico final", fmt(calculations.get("final_dynamic"), " m")],
        ["Abatimiento final", fmt(calculations.get("drawdown"), " m")],
        ["Caudal específico", fmt(calculations.get("specific_capacity"), " L/s/m")],
        ["Volumen bombeado", fmt(calculations.get("volume_m3"), " m³")],
        ["Pendiente final", fmt(calculations.get("slope_cm_h"), " cm/h")],
        ["Evaluación estabilización", calculations.get("stabilization_message")],
        ["Recuperación máxima", fmt(calculations.get("recovery_max"), " %")],
        ["Tiempo a 75% recuperación", fmt(calculations.get("t75"), " min", decimals=0)],
        ["Tiempo a 90% recuperación", fmt(calculations.get("t90"), " min", decimals=0)],
        ["Tiempo a 100% recuperación", fmt(calculations.get("t100"), " min", decimals=0)],
    ]
    story.append(make_table(results_data, col_widths=[5.2 * cm, 11.2 * cm]))

    # -------------------------------------------------------------------------
    # 9. GRÁFICOS
    # -------------------------------------------------------------------------
    story.append(Paragraph("9. Gráficos", styles["SectionTitle"]))

    pump_chart = make_line_chart_image(
        pumping_df, "Tiempo_min", "Nivel_m",
        "Prueba de gasto constante",
        "Tiempo (min)", "Nivel/profundidad (m)",
        invert_y=True
    )
    story.append(image_flowable(pump_chart))
    story.append(Paragraph("Figura 3. Gráfico de prueba a caudal constante.", styles["FigureCaption"]))

    rec_chart = make_line_chart_image(
        recovery_df, "Tiempo_min", "Nivel_m",
        "Prueba de recuperación",
        "Tiempo (min)", "Nivel/profundidad (m)",
        invert_y=True
    )
    story.append(image_flowable(rec_chart))
    story.append(Paragraph("Figura 4. Gráfico de recuperación de nivel.", styles["FigureCaption"]))

    if recovery_df is not None and "Recuperacion_pct" in recovery_df.columns:
        rec_pct_chart = make_line_chart_image(
            recovery_df, "Tiempo_min", "Recuperacion_pct",
            "Porcentaje de recuperación",
            "Tiempo (min)", "Recuperación (%)",
            invert_y=False
        )
        story.append(image_flowable(rec_pct_chart))
        story.append(Paragraph("Figura 5. Porcentaje de recuperación acumulada.", styles["FigureCaption"]))

    # -------------------------------------------------------------------------
    # 10. TABLAS
    # -------------------------------------------------------------------------
    story.append(Paragraph("10. Tabla de prueba de gasto constante", styles["SectionTitle"]))
    story.append(df_to_pdf_table(pumping_df, max_rows=70, font_size=5.8))

    story.append(Paragraph("11. Tabla de recuperación", styles["SectionTitle"]))
    story.append(df_to_pdf_table(recovery_df, max_rows=70, font_size=5.8))

    # -------------------------------------------------------------------------
    # 12. ADVERTENCIAS, CONCLUSIONES Y RECOMENDACIONES
    # -------------------------------------------------------------------------
    story.append(Paragraph("12. Advertencias técnicas", styles["SectionTitle"]))
    if warnings:
        for warning in warnings:
            story.append(Paragraph(f"• {warning}", styles["Body"]))
    else:
        story.append(Paragraph("No se registran advertencias técnicas críticas con los datos ingresados.", styles["Body"]))

    story.append(Paragraph("13. Conclusiones", styles["SectionTitle"]))
    conclusions = generate_conclusions(capture, calculations, warnings, methodology.get("modo_prueba", ""))
    for c in conclusions:
        story.append(Paragraph(f"• {c}", styles["Body"]))

    story.append(Paragraph("14. Recomendaciones", styles["SectionTitle"]))
    recs = generate_recommendations(capture, calculations, warnings)
    for r in recs:
        story.append(Paragraph(f"• {r}", styles["Body"]))

    story.append(Spacer(1, 1.0 * cm))
    firmas = [
        ["____________________________", "____________________________"],
        ["Firma consultor", "Firma beneficiario / cliente"],
    ]
    story.append(make_table(firmas, col_widths=[8 * cm, 8 * cm], first_col_bold=False))

    doc.build(story, onFirstPage=later_pages, onLaterPages=later_pages)
    return buffer.getvalue()


# =============================================================================
# INTERFAZ STREAMLIT
# =============================================================================

st.title("Sistema de Pruebas de Bombeo")
st.caption("Irrisal Consulting Ltda. | Informe técnico profesional v2.2")

if LOGO_PATH.exists():
    st.image(str(LOGO_PATH), width=260)

with st.sidebar:
    st.header("Configuración empresa")
    empresa = st.text_input("Empresa", COMPANY_DEFAULTS["empresa"])
    direccion = st.text_area("Dirección", COMPANY_DEFAULTS["direccion"])
    celular = st.text_input("Celular", COMPANY_DEFAULTS["celular"])
    correo = st.text_input("Correo", COMPANY_DEFAULTS["correo"])
    st.caption("Estos datos se insertan automáticamente en portada y pie de página.")

company = {"empresa": empresa, "direccion": direccion, "celular": celular, "correo": correo}

tabs = st.tabs([
    "1. Proyecto",
    "2. Captación",
    "3. Estratigrafía",
    "4. Equipos y metodología",
    "5. Bombeo",
    "6. Recuperación",
    "7. Resultados e informe",
])

with tabs[0]:
    st.subheader("Datos del proyecto")
    col1, col2 = st.columns(2)
    with col1:
        nombre_proyecto = st.text_input("Nombre del proyecto", "Prueba de bombeo")
        identificacion = st.text_input("Identificación de captación", "Captación subterránea")
        cliente = st.text_input("Cliente / beneficiario")
        sector = st.text_input("Sector / predio")
    with col2:
        comuna = st.text_input("Comuna")
        region = st.text_input("Región", "Región del Biobío")
        consultor = st.text_input("Consultor responsable")
        observaciones_proyecto = st.text_area("Observaciones generales")

    project = {
        "nombre_proyecto": nombre_proyecto,
        "identificacion": identificacion,
        "cliente": cliente,
        "sector": sector,
        "comuna": comuna,
        "region": region,
        "consultor": consultor,
        "observaciones": observaciones_proyecto,
    }

with tabs[1]:
    st.subheader("Captación, ubicación y habilitación")
    col1, col2, col3 = st.columns(3)

    with col1:
        tipo = st.selectbox("Tipo de captación", ["Pozo profundo", "Noria / pozo de gran diámetro", "Puntera", "Dren", "Otro"])
        condicion = st.selectbox("Condición", ["No informado", "No surgente", "Surgente"])
        utm_norte = st.text_input("UTM Norte")
        utm_este = st.text_input("UTM Este")
        datum = st.text_input("Datum", "SIRGAS WGS84 / WGS84")
        huso = st.text_input("Huso", "18S")

    with col2:
        profundidad_total = st.number_input("Profundidad total (m)", min_value=0.0, step=0.1, value=0.0)
        diametro_perforacion = st.text_input("Diámetro perforación")
        diametro_entubacion = st.text_input("Diámetro entubación")
        material_tuberia = st.text_input("Material / espesor tubería")
        altura_sobre_terreno = st.text_input("Altura tubería sobre terreno")
        nivel_estatico = st.number_input("Nivel estático inicial (m)", min_value=-50.0, step=0.01, value=0.0)

    with col3:
        st.caption("Si no tienes información de cribas o tramo filtrante, deja estos campos en 0.")
        criba_desde = st.number_input("Criba desde (m)", min_value=0.0, step=0.1, value=0.0)
        criba_hasta = st.number_input("Criba hasta (m)", min_value=0.0, step=0.1, value=0.0)
        tuberia_ciega = st.text_input("Tramos tubería ciega")
        profundidad_bomba = st.number_input("Profundidad bomba (m)", min_value=0.0, step=0.1, value=0.0)
        tuberia_extraccion = st.text_input("Tubería extracción/succión")
        observaciones_captacion = st.text_area("Observaciones de habilitación")

    st.write("#### Imágenes opcionales")
    location_image = st.file_uploader("Cargar croquis o imagen de ubicación", type=["jpg", "jpeg", "png"], key="location_image")
    scheme_image = st.file_uploader("Cargar esquema constructivo del pozo/captación", type=["jpg", "jpeg", "png"], key="scheme_image")

    capture = {
        "tipo": tipo,
        "condicion": condicion,
        "utm_norte": utm_norte,
        "utm_este": utm_este,
        "datum": datum,
        "huso": huso,
        "profundidad_total": profundidad_total if profundidad_total > 0 else None,
        "diametro_perforacion": diametro_perforacion,
        "diametro_entubacion": diametro_entubacion,
        "material_tuberia": material_tuberia,
        "altura_sobre_terreno": altura_sobre_terreno,
        "nivel_estatico": nivel_estatico,
        "criba_desde": criba_desde if criba_desde > 0 else "",
        "criba_hasta": criba_hasta if criba_hasta > 0 else "",
        "tuberia_ciega": tuberia_ciega,
        "profundidad_bomba": profundidad_bomba if profundidad_bomba > 0 else None,
        "tuberia_extraccion": tuberia_extraccion,
        "observaciones": observaciones_captacion,
    }

with tabs[2]:
    st.subheader("Estratigrafía")
    st.info("Ingresa los tramos reconocidos/perforados. Esta tabla se incluirá en el informe.")
    default_strat = pd.DataFrame({
        "Desde_m": [0.0, 1.0, 5.0],
        "Hasta_m": [1.0, 5.0, 12.0],
        "Descripcion": ["Suelo vegetal", "Material fino / arcilloso", "Arena / grava / material permeable"],
        "Observacion": ["", "", ""],
    })
    stratigraphy_df = st.data_editor(default_strat, num_rows="dynamic", use_container_width=True, key="strat_editor")

with tabs[3]:
    st.subheader("Equipos utilizados y metodología")
    col1, col2 = st.columns(2)
    with col1:
        bomba = st.text_input("Bomba: tipo/marca/modelo")
        potencia = st.text_input("Potencia")
        medidor_caudal = st.text_input("Medidor de caudal")
        instrumento_nivel = st.text_input("Instrumento de medición de nivel")
        generador = st.text_input("Generador / fuente eléctrica")
        observaciones_equipos = st.text_area("Observaciones de equipos")

    with col2:
        modo_prueba = st.selectbox("Modo de prueba", ["Ensayo abreviado 180 min + recuperación", "DGA estándar 24 h", "Otro"])
        fecha_prueba = st.text_input("Fecha de prueba", datetime.now().strftime("%d-%m-%Y"))
        hora_inicio = st.text_input("Hora inicio bombeo")
        hora_termino = st.text_input("Hora término bombeo")
        caudal_objetivo = st.text_input("Caudal objetivo")
        metodo_caudal = st.text_input("Método medición caudal", "Caudalímetro")
        metodo_nivel = st.text_input("Método medición niveles", "Pozómetro")
        frecuencia = st.text_input("Frecuencia de medición")
        observaciones_metodologia = st.text_area("Observaciones metodológicas")

    equipment = {
        "bomba": bomba,
        "potencia": potencia,
        "medidor_caudal": medidor_caudal,
        "instrumento_nivel": instrumento_nivel,
        "generador": generador,
        "observaciones": observaciones_equipos,
    }

    methodology = {
        "modo_prueba": modo_prueba,
        "fecha_prueba": fecha_prueba,
        "hora_inicio": hora_inicio,
        "hora_termino": hora_termino,
        "caudal_objetivo": caudal_objetivo,
        "metodo_caudal": metodo_caudal,
        "metodo_nivel": metodo_nivel,
        "frecuencia": frecuencia,
        "observaciones": observaciones_metodologia,
    }

with tabs[4]:
    st.subheader("Prueba de gasto constante")
    st.warning("El sistema no rellena datos faltantes. Solo calcula con datos ingresados/importados.")

    if "24 h" in locals().get("modo_prueba", ""):
        default_times = [0,1,2,3,4,5,10,20,30,60,120,180,240,300,360,420,480,540,600,660,720,780,840,900,960,1020,1080,1140,1200,1260,1320,1380,1440]
    else:
        default_times = [0,1,2,3,4,5,10,15,20,30,45,60,90,120,150,180]

    default_pumping = pd.DataFrame({
        "Fecha": [""] * len(default_times),
        "Hora": [""] * len(default_times),
        "Tiempo_min": default_times,
        "Nivel_m": [np.nan] * len(default_times),
        "Caudal_L_s": [np.nan] * len(default_times),
        "Observacion": [""] * len(default_times),
    })

    pumping_df = st.data_editor(default_pumping, num_rows="dynamic", use_container_width=True, key="pumping_editor")

with tabs[5]:
    st.subheader("Prueba de recuperación")
    default_rec_times = [0,1,2,3,4,5,10,15,20,30,45,60,90,120,150,180]
    default_recovery = pd.DataFrame({
        "Fecha": [""] * len(default_rec_times),
        "Hora": [""] * len(default_rec_times),
        "Tiempo_min": default_rec_times,
        "Nivel_m": [np.nan] * len(default_rec_times),
        "Observacion": [""] * len(default_rec_times),
    })

    recovery_df = st.data_editor(default_recovery, num_rows="dynamic", use_container_width=True, key="recovery_editor")

with tabs[6]:
    st.subheader("Resultados, gráficos e informe")

    calculations, recovery_with_pct = build_calculations(pumping_df, recovery_df, nivel_estatico)

    warnings = []
    add_warning(warnings, tipo == "Pozo profundo" and calculations.get("duration_min") is not None and calculations.get("duration_min") < 1440,
                "Pozo profundo con duración menor a 24 horas: no declarar cumplimiento formal de prueba estándar de 24 h.")
    add_warning(warnings, tipo == "Puntera" and "piez" not in safe_text(instrumento_nivel, "").lower(),
                "En punteras, el control de niveles debe efectuarse en piezómetro habilitado.")
    add_warning(warnings, not utm_norte or not utm_este, "Faltan coordenadas UTM.")
    add_warning(warnings, calculations.get("stabilization_evaluable") is False, calculations.get("stabilization_message"))
    add_warning(warnings, calculations.get("is_constant") is False, "El caudal no se mantuvo constante dentro de la tolerancia definida.")
    add_warning(warnings, calculations.get("recovery_max") is not None and calculations.get("recovery_max") < 75, "Recuperación inferior a 75%; interpretación limitada.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Duración", fmt(calculations.get("duration_min"), " min", decimals=0))
    c2.metric("Caudal promedio", fmt(calculations.get("q_mean"), " L/s"))
    c3.metric("Abatimiento", fmt(calculations.get("drawdown"), " m"))
    c4.metric("Caudal específico", fmt(calculations.get("specific_capacity"), " L/s/m"))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Volumen bombeado", fmt(calculations.get("volume_m3"), " m³"))
    c6.metric("Pendiente final", fmt(calculations.get("slope_cm_h"), " cm/h"))
    c7.metric("Recuperación máxima", fmt(calculations.get("recovery_max"), " %"))
    c8.metric("Tiempo 90%", fmt(calculations.get("t90"), " min", decimals=0))

    st.write("### Evaluación de estabilización")
    if calculations.get("stabilization_meets"):
        st.success(calculations.get("stabilization_message"))
    elif calculations.get("stabilization_evaluable"):
        st.error(calculations.get("stabilization_message"))
    else:
        st.warning(calculations.get("stabilization_message"))

    st.write("### Advertencias técnicas")
    if warnings:
        for w in warnings:
            st.warning(w)
    else:
        st.success("Sin advertencias críticas con los datos ingresados.")

    st.write("### Gráficos")
    g1, g2 = st.columns(2)

    pump_valid = df_numeric(pumping_df, ["Tiempo_min", "Nivel_m"]).dropna(subset=["Tiempo_min", "Nivel_m"])
    rec_valid = df_numeric(recovery_with_pct, ["Tiempo_min", "Nivel_m"]).dropna(subset=["Tiempo_min", "Nivel_m"])

    with g1:
        if len(pump_valid) >= 2:
            fig1 = px.line(
                pump_valid.sort_values("Tiempo_min"), x="Tiempo_min", y="Nivel_m", markers=True,
                title="Nivel/profundidad vs tiempo de bombeo",
                labels={"Tiempo_min": "Tiempo (min)", "Nivel_m": "Nivel/profundidad (m)"}
            )
            fig1.update_yaxes(autorange="reversed")
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("No hay suficientes datos para gráfico de bombeo.")

    with g2:
        if len(rec_valid) >= 2:
            fig2 = px.line(
                rec_valid.sort_values("Tiempo_min"), x="Tiempo_min", y="Nivel_m", markers=True,
                title="Nivel/profundidad vs tiempo de recuperación",
                labels={"Tiempo_min": "Tiempo (min)", "Nivel_m": "Nivel/profundidad (m)"}
            )
            fig2.update_yaxes(autorange="reversed")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No hay suficientes datos para gráfico de recuperación.")

    st.write("### Exportar")

    # Excel
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        pd.DataFrame([company]).to_excel(writer, sheet_name="Empresa", index=False)
        pd.DataFrame([project]).to_excel(writer, sheet_name="Proyecto", index=False)
        pd.DataFrame([capture]).to_excel(writer, sheet_name="Captacion", index=False)
        stratigraphy_df.to_excel(writer, sheet_name="Estratigrafia", index=False)
        pd.DataFrame([equipment]).to_excel(writer, sheet_name="Equipos", index=False)
        pd.DataFrame([methodology]).to_excel(writer, sheet_name="Metodologia", index=False)
        pumping_df.to_excel(writer, sheet_name="Bombeo", index=False)
        recovery_with_pct.to_excel(writer, sheet_name="Recuperacion", index=False)
        pd.DataFrame([calculations]).to_excel(writer, sheet_name="Calculos", index=False)
        pd.DataFrame({"Advertencias": warnings}).to_excel(writer, sheet_name="Advertencias", index=False)

    st.download_button(
        "Descargar Excel",
        data=excel_buffer.getvalue(),
        file_name="prueba_bombeo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    pdf_bytes = make_pdf(
        company=company,
        project=project,
        capture=capture,
        stratigraphy_df=stratigraphy_df,
        equipment=equipment,
        methodology=methodology,
        pumping_df=pumping_df,
        recovery_df=recovery_with_pct,
        calculations=calculations,
        warnings=warnings,
        location_image=location_image,
        scheme_image=scheme_image,
    )

    st.download_button(
        "Generar y descargar PDF profesional",
        data=pdf_bytes,
        file_name="informe_prueba_bombeo_profesional.pdf",
        mime="application/pdf",
    )

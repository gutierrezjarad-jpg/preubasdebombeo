
from __future__ import annotations

from pathlib import Path
from datetime import datetime
from io import BytesIO
import math

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, PageBreak
)

# =========================
# Configuración general
# =========================

APP_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = APP_DIR / "assets"
EXPORTS_DIR = APP_DIR / "exports"
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


# =========================
# Funciones de cálculo
# =========================

def clean_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def calculate_drawdown(static_level: float, dynamic_level: float) -> float | None:
    if static_level is None or dynamic_level is None:
        return None
    drawdown = dynamic_level - static_level
    return drawdown if drawdown >= 0 else None


def calculate_specific_capacity(q_l_s: float, drawdown_m: float) -> float | None:
    if drawdown_m is None or drawdown_m <= 0:
        return None
    return q_l_s / drawdown_m


def calculate_flow_stats(df: pd.DataFrame) -> dict:
    if df.empty or "Caudal_L_s" not in df:
        return {
            "q_mean": None, "q_min": None, "q_max": None,
            "q_std": None, "q_var_pct": None, "is_constant": None
        }

    q = clean_numeric_series(df["Caudal_L_s"]).dropna()
    if q.empty:
        return {
            "q_mean": None, "q_min": None, "q_max": None,
            "q_std": None, "q_var_pct": None, "is_constant": None
        }

    q_mean = float(q.mean())
    q_min = float(q.min())
    q_max = float(q.max())
    q_std = float(q.std(ddof=0)) if len(q) > 1 else 0.0

    if q_mean > 0:
        q_var_pct = ((q_max - q_min) / q_mean) * 100
    else:
        q_var_pct = None

    is_constant = q_var_pct is not None and q_var_pct <= 10  # rango max-min <=10% del promedio

    return {
        "q_mean": q_mean,
        "q_min": q_min,
        "q_max": q_max,
        "q_std": q_std,
        "q_var_pct": q_var_pct,
        "is_constant": is_constant,
    }


def calculate_pumped_volume(df: pd.DataFrame) -> float | None:
    """
    Integra volumen bombeado usando intervalos:
    volumen_m3 = sum(Q_prom_intervalo_L_s * delta_t_min * 60 / 1000)
    """
    if df.empty or not {"Tiempo_min", "Caudal_L_s"}.issubset(df.columns):
        return None

    data = df.copy()
    data["Tiempo_min"] = clean_numeric_series(data["Tiempo_min"])
    data["Caudal_L_s"] = clean_numeric_series(data["Caudal_L_s"])
    data = data.dropna(subset=["Tiempo_min", "Caudal_L_s"]).sort_values("Tiempo_min")

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
    """
    Evalúa pendiente en los últimos 180 minutos disponibles.
    Nivel_m se asume como profundidad medida desde punto de referencia.
    Pendiente positiva = mayor profundidad = descenso del nivel del agua.
    """
    result = {
        "evaluable": False,
        "slope_cm_h": None,
        "meets": False,
        "message": "No evaluable: faltan datos suficientes.",
    }

    if df.empty or not {"Tiempo_min", "Nivel_m"}.issubset(df.columns):
        return result

    data = df.copy()
    data["Tiempo_min"] = clean_numeric_series(data["Tiempo_min"])
    data["Nivel_m"] = clean_numeric_series(data["Nivel_m"])
    data = data.dropna(subset=["Tiempo_min", "Nivel_m"]).sort_values("Tiempo_min")

    if len(data) < 2:
        return result

    t_max = data["Tiempo_min"].max()
    t_min = data["Tiempo_min"].min()
    duration = t_max - t_min

    if duration < 180:
        result["message"] = "No evaluable: se requieren al menos 180 minutos de datos."
        return result

    last = data[data["Tiempo_min"] >= t_max - 180].copy()
    if len(last) < 2:
        result["message"] = "No evaluable: no hay suficientes puntos en los últimos 180 minutos."
        return result

    x_h = last["Tiempo_min"].to_numpy() / 60.0
    y_m = last["Nivel_m"].to_numpy()

    slope_m_h = np.polyfit(x_h, y_m, 1)[0]
    slope_cm_h = slope_m_h * 100

    result["evaluable"] = True
    result["slope_cm_h"] = float(slope_cm_h)
    result["meets"] = slope_cm_h <= 2

    if result["meets"]:
        result["message"] = "Presenta estabilización o franca tendencia según criterio ≤ 2 cm/h en los últimos 180 minutos."
    else:
        result["message"] = "No presenta estabilización según criterio ≤ 2 cm/h en los últimos 180 minutos."

    return result


def calculate_recovery(recovery_df: pd.DataFrame, static_level: float, final_dynamic_level: float) -> pd.DataFrame:
    if recovery_df.empty or not {"Tiempo_min", "Nivel_m"}.issubset(recovery_df.columns):
        return recovery_df

    data = recovery_df.copy()
    data["Tiempo_min"] = clean_numeric_series(data["Tiempo_min"])
    data["Nivel_m"] = clean_numeric_series(data["Nivel_m"])

    denom = final_dynamic_level - static_level
    if denom <= 0:
        data["Recuperacion_pct"] = np.nan
        return data

    data["Recuperacion_pct"] = ((final_dynamic_level - data["Nivel_m"]) / denom) * 100
    data["Recuperacion_pct"] = data["Recuperacion_pct"].clip(lower=0, upper=100)
    return data


def time_to_recovery(data: pd.DataFrame, target_pct: float) -> float | None:
    if data.empty or "Recuperacion_pct" not in data.columns:
        return None
    valid = data.dropna(subset=["Tiempo_min", "Recuperacion_pct"])
    reached = valid[valid["Recuperacion_pct"] >= target_pct]
    if reached.empty:
        return None
    return float(reached["Tiempo_min"].iloc[0])


def format_value(value, suffix="", decimals=2):
    if value is None:
        return "No evaluable"
    try:
        if pd.isna(value):
            return "No evaluable"
    except Exception:
        pass
    if isinstance(value, (float, int, np.number)):
        return f"{value:.{decimals}f}{suffix}"
    return str(value)


def add_warning(warnings: list[str], condition: bool, message: str):
    if condition:
        warnings.append(message)



# =========================
# Gráficos para PDF
# =========================

def make_line_chart_image(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    title: str,
    x_label: str,
    y_label: str,
    invert_y: bool = False,
) -> BytesIO | None:
    """
    Crea un gráfico simple en PNG para insertar en PDF.
    Usa matplotlib para evitar depender de navegadores o conversiones externas.
    """
    if df is None or df.empty or not {x_col, y_col}.issubset(df.columns):
        return None

    data = df.copy()
    data[x_col] = clean_numeric_series(data[x_col])
    data[y_col] = clean_numeric_series(data[y_col])
    data = data.dropna(subset=[x_col, y_col]).sort_values(x_col)

    if len(data) < 2:
        return None

    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ax.plot(data[x_col], data[y_col], marker="o", linewidth=1.6, markersize=3.5)
    ax.set_title(title)
    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    ax.grid(True, alpha=0.3)

    if invert_y:
        ax.invert_yaxis()

    fig.tight_layout()
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=180)
    plt.close(fig)
    buffer.seek(0)
    return buffer


def chart_image_flowable(
    image_buffer: BytesIO | None,
    width_cm: float = 16.0,
    height_cm: float = 8.0,
):
    if image_buffer is None:
        return Paragraph("Gráfico no disponible: datos insuficientes.", getSampleStyleSheet()["Normal"])
    image_buffer.seek(0)
    return RLImage(image_buffer, width=width_cm * cm, height=height_cm * cm)


# =========================
# Generación de PDF
# =========================

def make_pdf(
    company: dict,
    project: dict,
    capture: dict,
    equipment: dict,
    pumping_df: pd.DataFrame,
    recovery_df: pd.DataFrame,
    calculations: dict,
    warnings: list[str],
) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.4 * cm,
        bottomMargin=1.4 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", fontSize=8, leading=10))
    styles.add(ParagraphStyle(name="TitleCenter", fontSize=18, leading=22, alignment=1, spaceAfter=14))
    styles.add(ParagraphStyle(name="Section", fontSize=13, leading=16, spaceBefore=10, spaceAfter=6, textColor=colors.HexColor("#006b2e")))

    story = []

    if LOGO_PATH.exists():
        try:
            story.append(RLImage(str(LOGO_PATH), width=6.5 * cm, height=3.8 * cm))
            story.append(Spacer(1, 0.2 * cm))
        except Exception:
            pass

    story.append(Paragraph("INFORME DE PRUEBA DE BOMBEO", styles["TitleCenter"]))
    story.append(Paragraph(f"<b>{company.get('empresa', 'Dato no informado')}</b>", styles["Normal"]))
    story.append(Paragraph(company.get("direccion", "Dato no informado"), styles["Small"]))
    story.append(Paragraph(f"Celular: {company.get('celular', 'Dato no informado')} | Correo: {company.get('correo', 'Dato no informado')}", styles["Small"]))
    story.append(Spacer(1, 0.8 * cm))

    portada_data = [
        ["Proyecto", project.get("nombre_proyecto") or "Dato no informado"],
        ["Cliente", project.get("cliente") or "Dato no informado"],
        ["Sector", project.get("sector") or "Dato no informado"],
        ["Comuna", project.get("comuna") or "Dato no informado"],
        ["Región", project.get("region") or "Dato no informado"],
        ["Fecha informe", datetime.now().strftime("%d-%m-%Y")],
    ]
    story.append(make_table(portada_data, col_widths=[4 * cm, 12 * cm]))
    story.append(PageBreak())

    # Antecedentes
    story.append(Paragraph("1. Antecedentes generales", styles["Section"]))
    story.append(Paragraph(
        "El presente informe resume los antecedentes de la captación, la prueba de gasto constante, "
        "la recuperación posterior y los cálculos técnicos derivados de los datos ingresados por el usuario.",
        styles["Normal"]
    ))
    story.append(Spacer(1, 0.2 * cm))

    # Captación
    story.append(Paragraph("2. Habilitación y ubicación de la captación", styles["Section"]))
    cap_data = [
        ["Tipo de captación", capture.get("tipo") or "Dato no informado"],
        ["Coordenada UTM Norte", capture.get("utm_norte") or "Dato no informado"],
        ["Coordenada UTM Este", capture.get("utm_este") or "Dato no informado"],
        ["Datum / Huso", f"{capture.get('datum') or 'Dato no informado'} / {capture.get('huso') or 'Dato no informado'}"],
        ["Profundidad total (m)", capture.get("profundidad_total") or "Dato no informado"],
        ["Diámetro", capture.get("diametro") or "Dato no informado"],
        ["Nivel estático inicial (m)", capture.get("nivel_estatico") or "Dato no informado"],
        ["Observaciones", capture.get("observaciones") or "Dato no informado"],
    ]
    story.append(make_table(cap_data, col_widths=[5 * cm, 11 * cm]))

    # Equipos
    story.append(Paragraph("3. Equipos utilizados", styles["Section"]))
    eq_data = [
        ["Bomba", equipment.get("bomba") or "Dato no informado"],
        ["Potencia", equipment.get("potencia") or "Dato no informado"],
        ["Medidor de caudal", equipment.get("medidor_caudal") or "Dato no informado"],
        ["Instrumento de nivel", equipment.get("instrumento_nivel") or "Dato no informado"],
        ["Observaciones", equipment.get("observaciones") or "Dato no informado"],
    ]
    story.append(make_table(eq_data, col_widths=[5 * cm, 11 * cm]))

    # Resultados
    story.append(Paragraph("4. Resultados calculados", styles["Section"]))
    calc_data = [
        ["Caudal promedio", format_value(calculations.get("q_mean"), " L/s")],
        ["Caudal mínimo", format_value(calculations.get("q_min"), " L/s")],
        ["Caudal máximo", format_value(calculations.get("q_max"), " L/s")],
        ["Variación relativa de caudal", format_value(calculations.get("q_var_pct"), " %")],
        ["Abatimiento final", format_value(calculations.get("drawdown"), " m")],
        ["Caudal específico", format_value(calculations.get("specific_capacity"), " L/s/m")],
        ["Volumen bombeado", format_value(calculations.get("volume_m3"), " m³")],
        ["Pendiente final", format_value(calculations.get("slope_cm_h"), " cm/h")],
        ["Evaluación estabilización", calculations.get("stabilization_message", "No evaluable")],
        ["Recuperación máxima", format_value(calculations.get("recovery_max"), " %")],
        ["Tiempo a 75% recuperación", format_value(calculations.get("t75"), " min")],
        ["Tiempo a 90% recuperación", format_value(calculations.get("t90"), " min")],
        ["Tiempo a 100% recuperación", format_value(calculations.get("t100"), " min")],
    ]
    story.append(make_table(calc_data, col_widths=[5.5 * cm, 10.5 * cm]))

    # Gráficos técnicos
    story.append(Paragraph("5. Gráficos técnicos", styles["Section"]))

    pumping_chart = make_line_chart_image(
        pumping_df,
        x_col="Tiempo_min",
        y_col="Nivel_m",
        title="Nivel/profundidad vs tiempo de bombeo",
        x_label="Tiempo (min)",
        y_label="Nivel/profundidad (m)",
        invert_y=True,
    )
    story.append(chart_image_flowable(pumping_chart))
    story.append(Paragraph("Figura 1. Gráfico de prueba a caudal constante.", styles["Small"]))
    story.append(Spacer(1, 0.4 * cm))

    recovery_chart = make_line_chart_image(
        recovery_df,
        x_col="Tiempo_min",
        y_col="Nivel_m",
        title="Nivel/profundidad vs tiempo de recuperación",
        x_label="Tiempo (min)",
        y_label="Nivel/profundidad (m)",
        invert_y=True,
    )
    story.append(chart_image_flowable(recovery_chart))
    story.append(Paragraph("Figura 2. Gráfico de recuperación.", styles["Small"]))
    story.append(Spacer(1, 0.4 * cm))

    if recovery_df is not None and "Recuperacion_pct" in recovery_df.columns:
        recovery_pct_chart = make_line_chart_image(
            recovery_df,
            x_col="Tiempo_min",
            y_col="Recuperacion_pct",
            title="Porcentaje de recuperación vs tiempo",
            x_label="Tiempo (min)",
            y_label="Recuperación (%)",
            invert_y=False,
        )
        story.append(chart_image_flowable(recovery_pct_chart))
        story.append(Paragraph("Figura 3. Porcentaje de recuperación acumulada.", styles["Small"]))
        story.append(Spacer(1, 0.4 * cm))

    # Advertencias
    story.append(Paragraph("6. Advertencias técnicas", styles["Section"]))
    if warnings:
        for w in warnings:
            story.append(Paragraph(f"• {w}", styles["Normal"]))
    else:
        story.append(Paragraph("No se registran advertencias técnicas críticas con los datos ingresados.", styles["Normal"]))

    # Datos bombeo
    story.append(Paragraph("7. Tabla de prueba de gasto constante", styles["Section"]))
    story.append(df_to_reportlab_table(pumping_df, max_rows=35))

    # Datos recuperación
    story.append(Paragraph("8. Tabla de recuperación", styles["Section"]))
    story.append(df_to_reportlab_table(recovery_df, max_rows=35))

    # Conclusiones
    story.append(Paragraph("9. Conclusiones", styles["Section"]))

    tipo = capture.get("tipo", "")
    dur = calculations.get("duration_min")
    if tipo == "Pozo profundo" and dur is not None and dur < 1440:
        story.append(Paragraph(
            "La prueba corresponde a un ensayo abreviado respecto de una prueba estándar de 24 horas para pozo profundo. "
            "Los resultados permiten una evaluación preliminar del comportamiento hidráulico durante el periodo medido, "
            "pero no deben interpretarse como cumplimiento formal de una prueba de 24 horas.",
            styles["Normal"]
        ))

    story.append(Paragraph(
        "Las conclusiones se basan exclusivamente en los datos ingresados por el usuario. "
        "No se han rellenado ni inventado mediciones faltantes.",
        styles["Normal"]
    ))

    story.append(Spacer(1, 1.5 * cm))
    firmas = [
        ["____________________________", "____________________________"],
        ["Firma consultor", "Firma beneficiario/cliente"],
    ]
    story.append(make_table(firmas, col_widths=[8 * cm, 8 * cm], header=False))

    doc.build(story)
    return buffer.getvalue()


def make_table(data, col_widths=None, header=False):
    table = Table(data, colWidths=col_widths)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#edf7ef")),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9ead3")))
        style.append(("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"))
    table.setStyle(TableStyle(style))
    return table


def df_to_reportlab_table(df: pd.DataFrame, max_rows: int = 30):
    if df is None or df.empty:
        return Paragraph("Sin datos ingresados.", getSampleStyleSheet()["Normal"])

    show = df.head(max_rows).copy()
    show = show.fillna("")
    data = [list(show.columns)] + show.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9ead3")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return table


# =========================
# Interfaz
# =========================

st.title("Sistema de Pruebas de Bombeo")
st.caption("Irrisal Consulting Ltda. | Versión mínima funcional")

if LOGO_PATH.exists():
    st.image(str(LOGO_PATH), width=260)

with st.sidebar:
    st.header("Configuración empresa")
    empresa = st.text_input("Empresa", COMPANY_DEFAULTS["empresa"])
    direccion = st.text_area("Dirección", COMPANY_DEFAULTS["direccion"])
    celular = st.text_input("Celular", COMPANY_DEFAULTS["celular"])
    correo = st.text_input("Correo", COMPANY_DEFAULTS["correo"])

company = {
    "empresa": empresa,
    "direccion": direccion,
    "celular": celular,
    "correo": correo,
}

tabs = st.tabs([
    "1. Proyecto",
    "2. Captación",
    "3. Equipos",
    "4. Bombeo",
    "5. Recuperación",
    "6. Resultados e informe"
])

with tabs[0]:
    st.subheader("Datos del proyecto")
    col1, col2 = st.columns(2)
    with col1:
        nombre_proyecto = st.text_input("Nombre del proyecto", "Prueba de bombeo")
        cliente = st.text_input("Cliente / beneficiario")
        sector = st.text_input("Sector / predio")
    with col2:
        comuna = st.text_input("Comuna")
        region = st.text_input("Región", "Región del Biobío")
        consultor = st.text_input("Consultor responsable")

    project = {
        "nombre_proyecto": nombre_proyecto,
        "cliente": cliente,
        "sector": sector,
        "comuna": comuna,
        "region": region,
        "consultor": consultor,
    }

with tabs[1]:
    st.subheader("Captación y habilitación")
    col1, col2, col3 = st.columns(3)
    with col1:
        tipo = st.selectbox("Tipo de captación", ["Pozo profundo", "Noria / pozo de gran diámetro", "Puntera", "Dren", "Otro"])
        utm_norte = st.text_input("UTM Norte")
        utm_este = st.text_input("UTM Este")
        datum = st.text_input("Datum", "SIRGAS WGS84 / WGS84")
    with col2:
        huso = st.text_input("Huso", "18S")
        profundidad_total = st.number_input("Profundidad total (m)", min_value=0.0, step=0.1)
        diametro = st.text_input("Diámetro")
        nivel_estatico = st.number_input("Nivel estático inicial (m)", min_value=-50.0, step=0.01, value=0.0)
    with col3:
        observaciones_captacion = st.text_area("Observaciones de captación/habilitación")

    capture = {
        "tipo": tipo,
        "utm_norte": utm_norte,
        "utm_este": utm_este,
        "datum": datum,
        "huso": huso,
        "profundidad_total": profundidad_total if profundidad_total > 0 else "",
        "diametro": diametro,
        "nivel_estatico": nivel_estatico,
        "observaciones": observaciones_captacion,
    }

with tabs[2]:
    st.subheader("Equipos utilizados")
    col1, col2 = st.columns(2)
    with col1:
        bomba = st.text_input("Bomba: tipo/marca/modelo")
        potencia = st.text_input("Potencia")
        medidor_caudal = st.text_input("Medidor de caudal")
    with col2:
        instrumento_nivel = st.text_input("Instrumento de medición de nivel")
        observaciones_equipos = st.text_area("Observaciones de equipos")

    equipment = {
        "bomba": bomba,
        "potencia": potencia,
        "medidor_caudal": medidor_caudal,
        "instrumento_nivel": instrumento_nivel,
        "observaciones": observaciones_equipos,
    }

with tabs[3]:
    st.subheader("Prueba de gasto constante")

    st.info("Puedes editar la tabla directamente. No se rellenan datos faltantes.")

    default_pumping = pd.DataFrame({
        "Tiempo_min": [0, 1, 2, 3, 4, 5, 10, 20, 30, 60, 90, 120, 150, 180],
        "Nivel_m": [0.0, 2.5, 3.2, 3.8, 4.2, 4.6, 5.2, 5.8, 6.1, 6.6, 6.8, 6.9, 6.95, 7.0],
        "Caudal_L_s": [1.0] * 14,
        "Observacion": [""] * 14,
    })

    pumping_df = st.data_editor(
        default_pumping,
        num_rows="dynamic",
        use_container_width=True,
        key="pumping_editor"
    )

with tabs[4]:
    st.subheader("Prueba de recuperación")

    default_recovery = pd.DataFrame({
        "Tiempo_min": [0, 1, 2, 3, 4, 5, 10, 20, 30, 60, 90, 120, 150, 180],
        "Nivel_m": [7.0, 6.4, 5.9, 5.4, 5.0, 4.6, 3.6, 2.5, 1.8, 0.8, 0.35, 0.1, 0.0, 0.0],
        "Observacion": [""] * 14,
    })

    recovery_df = st.data_editor(
        default_recovery,
        num_rows="dynamic",
        use_container_width=True,
        key="recovery_editor"
    )

with tabs[5]:
    st.subheader("Resultados, advertencias y exportaciones")

    # Limpieza para cálculos
    pumping_calc = pumping_df.copy()
    pumping_calc["Tiempo_min"] = clean_numeric_series(pumping_calc["Tiempo_min"])
    pumping_calc["Nivel_m"] = clean_numeric_series(pumping_calc["Nivel_m"])
    pumping_calc["Caudal_L_s"] = clean_numeric_series(pumping_calc["Caudal_L_s"])

    recovery_calc = recovery_df.copy()
    recovery_calc["Tiempo_min"] = clean_numeric_series(recovery_calc["Tiempo_min"])
    recovery_calc["Nivel_m"] = clean_numeric_series(recovery_calc["Nivel_m"])

    valid_pumping = pumping_calc.dropna(subset=["Tiempo_min", "Nivel_m"]).sort_values("Tiempo_min")
    duration_min = None
    final_dynamic = None

    if not valid_pumping.empty:
        duration_min = float(valid_pumping["Tiempo_min"].max() - valid_pumping["Tiempo_min"].min())
        final_dynamic = float(valid_pumping.sort_values("Tiempo_min")["Nivel_m"].iloc[-1])

    flow_stats = calculate_flow_stats(pumping_calc)
    drawdown = calculate_drawdown(nivel_estatico, final_dynamic) if final_dynamic is not None else None
    specific_capacity = calculate_specific_capacity(flow_stats["q_mean"], drawdown) if flow_stats["q_mean"] is not None else None
    volume_m3 = calculate_pumped_volume(pumping_calc)
    stab = evaluate_stabilization(pumping_calc)

    if final_dynamic is not None:
        recovery_with_pct = calculate_recovery(recovery_calc, nivel_estatico, final_dynamic)
    else:
        recovery_with_pct = recovery_calc.copy()
        recovery_with_pct["Recuperacion_pct"] = np.nan

    recovery_max = None
    if "Recuperacion_pct" in recovery_with_pct.columns:
        rec_valid = recovery_with_pct["Recuperacion_pct"].dropna()
        recovery_max = float(rec_valid.max()) if not rec_valid.empty else None

    t75 = time_to_recovery(recovery_with_pct, 75)
    t90 = time_to_recovery(recovery_with_pct, 90)
    t100 = time_to_recovery(recovery_with_pct, 100)

    warnings = []
    add_warning(warnings, tipo == "Pozo profundo" and duration_min is not None and duration_min < 1440,
                "Pozo profundo con duración menor a 24 horas: no declarar cumplimiento formal de prueba estándar de 24 h.")
    add_warning(warnings, stab["evaluable"] is False,
                stab["message"])
    add_warning(warnings, flow_stats["is_constant"] is False,
                "El caudal no se mantuvo constante dentro de la tolerancia definida.")
    add_warning(warnings, recovery_max is not None and recovery_max < 75,
                "Recuperación inferior a 75%; interpretación limitada.")
    add_warning(warnings, not utm_norte or not utm_este,
                "Faltan coordenadas UTM.")
    add_warning(warnings, tipo == "Puntera" and not instrumento_nivel.lower().strip().startswith("piez"),
                "En punteras, el control de niveles debe efectuarse en piezómetro habilitado.")

    calculations = {
        **flow_stats,
        "duration_min": duration_min,
        "final_dynamic": final_dynamic,
        "drawdown": drawdown,
        "specific_capacity": specific_capacity,
        "volume_m3": volume_m3,
        "slope_cm_h": stab["slope_cm_h"],
        "stabilization_message": stab["message"],
        "recovery_max": recovery_max,
        "t75": t75,
        "t90": t90,
        "t100": t100,
    }

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Duración", format_value(duration_min, " min", 0))
    c2.metric("Caudal promedio", format_value(flow_stats["q_mean"], " L/s"))
    c3.metric("Abatimiento", format_value(drawdown, " m"))
    c4.metric("Caudal específico", format_value(specific_capacity, " L/s/m"))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Volumen bombeado", format_value(volume_m3, " m³"))
    c6.metric("Pendiente final", format_value(stab["slope_cm_h"], " cm/h"))
    c7.metric("Recuperación máxima", format_value(recovery_max, " %"))
    c8.metric("Tiempo 90%", format_value(t90, " min", 0))

    st.write("### Evaluación de estabilización")
    if stab["meets"]:
        st.success(stab["message"])
    elif stab["evaluable"]:
        st.error(stab["message"])
    else:
        st.warning(stab["message"])

    st.write("### Advertencias técnicas")
    if warnings:
        for w in warnings:
            st.warning(w)
    else:
        st.success("Sin advertencias críticas con los datos ingresados.")

    st.write("### Gráficos")
    g1, g2 = st.columns(2)

    with g1:
        if not valid_pumping.empty:
            fig1 = px.line(
                valid_pumping,
                x="Tiempo_min",
                y="Nivel_m",
                markers=True,
                title="Nivel/profundidad vs tiempo de bombeo",
                labels={"Tiempo_min": "Tiempo (min)", "Nivel_m": "Nivel/profundidad (m)"}
            )
            fig1.update_yaxes(autorange="reversed", title_text="Nivel/profundidad (m)")
            fig1.update_xaxes(title_text="Tiempo (min)")
            st.plotly_chart(fig1, use_container_width=True)

    with g2:
        rec_valid = recovery_with_pct.dropna(subset=["Tiempo_min", "Nivel_m"]) if not recovery_with_pct.empty else pd.DataFrame()
        if not rec_valid.empty:
            fig2 = px.line(
                rec_valid,
                x="Tiempo_min",
                y="Nivel_m",
                markers=True,
                title="Nivel/profundidad vs tiempo de recuperación",
                labels={"Tiempo_min": "Tiempo (min)", "Nivel_m": "Nivel/profundidad (m)"}
            )
            fig2.update_yaxes(autorange="reversed", title_text="Nivel/profundidad (m)")
            fig2.update_xaxes(title_text="Tiempo (min)")
            st.plotly_chart(fig2, use_container_width=True)

    st.write("### Exportar")

    # Excel
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        pd.DataFrame([company]).to_excel(writer, sheet_name="Empresa", index=False)
        pd.DataFrame([project]).to_excel(writer, sheet_name="Proyecto", index=False)
        pd.DataFrame([capture]).to_excel(writer, sheet_name="Captacion", index=False)
        pd.DataFrame([equipment]).to_excel(writer, sheet_name="Equipos", index=False)
        pumping_calc.to_excel(writer, sheet_name="Bombeo", index=False)
        recovery_with_pct.to_excel(writer, sheet_name="Recuperacion", index=False)
        pd.DataFrame([calculations]).to_excel(writer, sheet_name="Calculos", index=False)
        pd.DataFrame({"Advertencias": warnings}).to_excel(writer, sheet_name="Advertencias", index=False)

    st.download_button(
        "Descargar Excel",
        data=excel_buffer.getvalue(),
        file_name="prueba_bombeo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF
    pdf_bytes = make_pdf(
        company=company,
        project=project,
        capture=capture,
        equipment=equipment,
        pumping_df=pumping_calc,
        recovery_df=recovery_with_pct,
        calculations=calculations,
        warnings=warnings,
    )

    st.download_button(
        "Generar y descargar PDF",
        data=pdf_bytes,
        file_name="informe_prueba_bombeo.pdf",
        mime="application/pdf"
    )

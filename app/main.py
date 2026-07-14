from __future__ import annotations

from io import BytesIO
from pathlib import Path
from datetime import datetime
import json
import re
import tempfile
from typing import Any

import pandas as pd
import streamlit as st
from PIL import Image
import pytesseract
import fitz

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image as RLImage,
    KeepTogether,
)

# =============================================================================
# CONFIGURACION
# =============================================================================

st.set_page_config(page_title="Memoria Explicativa DGA - Riego", layout="wide")

APP_VERSION = "v2.9 - corrección aplicada punto 4, UTM y alineación"

PETICIONARIO_EMPRESA = {
    "tipo_persona": "Persona jurídica",
    "nombre": "Irrisal Consulting Ltda.",
    "domicilio": "San Martín 553 oficina 901",
    "rut": "78.271.963-7",
    "fono": "+56 9 6796 0884",
    "correo": "Irrisalconsulting@gmail.com",
}

DEFAULTS = {
    "tipo_persona": "Persona jurídica",
    "sexo": "",
    "nombre": PETICIONARIO_EMPRESA["nombre"],
    "rut": PETICIONARIO_EMPRESA["rut"],
    "domicilio": PETICIONARIO_EMPRESA["domicilio"],
    "fono": PETICIONARIO_EMPRESA["fono"],
    "correo": PETICIONARIO_EMPRESA["correo"],
    "naturaleza": "Subterránea",
    "tipo_derecho": "Consuntivo",
    "ejercicio_1": "Permanente",
    "ejercicio_2": "Continuo",
    "caudal_l_s": 0.0,
    "volumen_anual_m3": 0.0,
    "utm_norte": "",
    "utm_este": "",
    "datum_huso": "WGS84 / Huso 18",
    "descripcion_ubicacion": "",
    "region": "Región del Biobío",
    "provincia": "Biobío",
    "comuna": "",
    "descripcion_proyecto": "",
    "porcentaje_riego": 95.0,
    "porcentaje_subsistencia": 5.0,
    "personas_beneficiadas": 1,
    "predio": "",
    "rol_sii": "",
    "hectareas_riego": 0.0,
    "fojas": "",
    "numero_inscripcion": "",
    "anio_inscripcion": "",
    "conservador": "",
    "informacion_adicional": "Se adjuntan antecedentes técnicos, cartográficos y legales de respaldo al expediente.",
    "firmante_nombre": "",
    "firmante_rut": "",
}

USOS_OPCIONES = ["Riego + Uso doméstico de subsistencia", "Solo riego"]

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
TEMPLATE_DGA_PDF = ASSETS_DIR / "MEMORIA-EXPLICATIVA-PARA-DIFERENTES-USOS.pdf"

# =============================================================================
# UTILIDADES
# =============================================================================


def clean_text(value: Any, default: str = "") -> str:
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    value = str(value).strip()
    return value if value else default


def to_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None:
            return default
        if isinstance(value, str):
            value = value.strip().replace(".", "").replace(",", ".") if re.search(r"\d+\.\d{3}", value) else value.strip().replace(",", ".")
            if not value:
                return default
        out = float(value)
        if pd.isna(out):
            return default
        return out
    except Exception:
        return default


def fmt_num(value: Any, decimals: int = 2) -> str:
    n = to_float(value, 0.0)
    if decimals == 0:
        return f"{n:.0f}"
    text = f"{n:.{decimals}f}"
    return text.replace(".", ",")


def fmt_ca(w: str) -> str:
    return clean_text(w, "No informado")


def checkbox(checked: bool) -> str:
    return "X" if checked else ""


def update_volume_from_flow(q_l_s: float) -> float:
    # L/s -> m3/anio para uso continuo: q * segundos/anio / 1000
    return q_l_s * 31_536_000 / 1000


def init_state():
    for k, v in DEFAULTS.items():
        st.session_state.setdefault(k, v)
    st.session_state.setdefault("uso_agua", USOS_OPCIONES[0])
    st.session_state.setdefault("firma_png", None)
    st.session_state.setdefault("croquis_png", None)


def get_data() -> dict[str, Any]:
    data = {k: st.session_state.get(k, v) for k, v in DEFAULTS.items()}
    data["uso_agua"] = st.session_state.get("uso_agua", USOS_OPCIONES[0])
    return data


def normalize_payload_keys(payload: dict[str, Any]) -> dict[str, Any]:
    """
    Mapea nombres antiguos o detectados a los nombres reales de campos de la app.
    """
    if not payload:
        return {}

    mapping = {
        "predio_rol": "rol_sii",
        "predio_fojas": "fojas",
        "predio_numero": "numero_inscripcion",
        "predio_anio": "anio_inscripcion",
        "predio_conservador": "conservador",
        "utm_n": "utm_norte",
        "utm_e": "utm_este",
        "norte": "utm_norte",
        "este": "utm_este",
        "utm_north": "utm_norte",
        "utm_east": "utm_este",
        "caudal": "caudal_l_s",
        "caudal_solicitado": "caudal_l_s",
        "caudal_lts_seg": "caudal_l_s",
        "caudal_l/s": "caudal_l_s",
        "volumen_anual": "volumen_anual_m3",
        "volumen_m3_anual": "volumen_anual_m3",
        "volumen": "volumen_anual_m3",
    }

    out = {}
    for k, v in payload.items():
        out[mapping.get(k, k)] = v
    return out


def apply_data(payload: dict[str, Any]):
    payload = normalize_payload_keys(payload)
    for k in DEFAULTS:
        if k in payload and payload[k] is not None:
            st.session_state[k] = payload[k]
    if "uso_agua" in payload:
        st.session_state["uso_agua"] = payload["uso_agua"]


# =============================================================================
# AUTOCOMPLETAR DESDE INFORME WORD BLA
# =============================================================================


def docx_to_lines(uploaded_file) -> list[str]:
    """
    Extrae texto desde párrafos y tablas de un .docx, manteniendo líneas separadas.
    Esto permite distinguir campos estructurados como 'Solicitante : Nombre'
    de frases narrativas como 'La solicitante es propietaria...'.
    """
    doc = Document(uploaded_file)
    parts: list[str] = []

    for p in doc.paragraphs:
        txt = re.sub(r"\s+", " ", p.text or "").strip()
        if txt:
            parts.append(txt)

    for table in doc.tables:
        for row in table.rows:
            cells = []
            for cell in row.cells:
                txt = re.sub(r"\s+", " ", cell.text or "").strip()
                if txt:
                    cells.append(txt)
            if cells:
                parts.append(" | ".join(cells))

    return parts


def docx_to_text(uploaded_file) -> str:
    return "\n".join(docx_to_lines(uploaded_file))


def _norm_label(value: str) -> str:
    value = (value or "").strip().lower()
    value = value.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    value = value.replace(":", "").strip()
    return value


KNOWN_BLA_LABELS = [
    "Solicitante", "RUT", "Domicilio", "Comuna", "Provincia", "Correo electrónico",
    "Teléfono", "Telefono", "Naturaleza del agua", "Tipo de derecho", "Ejercicio del derecho",
    "Modo de extracción", "Tipo de obra", "Uso del agua", "Caudal solicitado",
    "Coordenadas, huso", "Datum", "Cuenca", "Acuífero", "Acuifero", "Sector acuífero",
    "Sector acuifero"
]


def _is_known_label(token: str) -> bool:
    n = _norm_label(token)
    return any(n == _norm_label(label) for label in KNOWN_BLA_LABELS)


def clean_detected_value(value: str) -> str:
    value = re.sub(r"\s+", " ", value or "").strip(" :|")
    if value.lower() in ["no registra", "no informado", "sin información", "sin informacion"]:
        return ""
    return value


def get_structured_field(lines: list[str], label: str) -> str:
    """
    Lee campos de tablas o líneas tipo:
    Solicitante | : | María Contreras Acuña
    Uso del agua | : Riego | Caudal solicitado | : 0,95 l/s
    Evita capturar frases narrativas como 'La solicitante es propietaria...'.
    """
    label_norm = _norm_label(label)

    for line in lines:
        # 1) Tabla con separadores |.
        tokens = [t.strip() for t in line.split("|") if t.strip()]
        token_norms = [_norm_label(t) for t in tokens]

        if label_norm in token_norms:
            idx = token_norms.index(label_norm)
            collected = []
            for t in tokens[idx + 1:]:
                nt = _norm_label(t)
                if nt in [":", ""] or t.strip() == ":":
                    continue
                if _is_known_label(t):
                    break
                value = clean_detected_value(t)
                if value and value not in collected:
                    collected.append(value)
            if collected:
                return collected[0]

        # 2) En algunos docx la celda trae 'Etiqueta : valor' completa.
        # Cortar si luego aparece otra etiqueta conocida.
        m = re.search(rf"(?:^|\|)\s*{re.escape(label)}\s*:??\s*([^\n|]+)", line, flags=re.I)
        if m:
            value = clean_detected_value(m.group(1))
            if value:
                value = re.split(r"\s+(RUT|Domicilio|Comuna|Provincia|Correo electrónico|Tel[eé]fono|Caudal solicitado|Datum|Huso|Coordenadas)\s*:??", value, flags=re.I)[0]
                return clean_detected_value(value)

    return ""

def first_match(pattern: str, text: str, flags=re.I) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


def parse_coordinates(text_value: str) -> tuple[str, str]:
    """
    Busca coordenadas UTM en formatos frecuentes de informes BLA.
    Devuelve (utm_norte, utm_este).

    Versión robusta: acepta puntos de miles, espacios, E/N, Este/Norte,
    coordenadas en tablas y coordenadas en párrafos narrativos.
    """
    text_value = text_value or ""
    flat = re.sub(r"\s+", " ", text_value)
    flat = flat.replace(";", ",")
    flat = flat.replace("UTM E", "UTM Este").replace("UTM N", "UTM Norte")
    flat = flat.replace("m.E", "m E").replace("m.N", "m N").replace("m.S", "m S")

    def clean_coord(value: str) -> str:
        return re.sub(r"[^0-9]", "", value or "")

    def valid_pair(norte: str, este: str) -> bool:
        return len(este) == 6 and len(norte) == 7

    patterns = [
        # Huso 18, coordenadas 753.811 m E, 5.821.082 m S/N
        (r"coordenadas?.{0,80}?([0-9]{3}[\.\s]?[0-9]{3})\s*m?\s*(?:E|Este)\s*,?\s*(?:y\s*)?([0-9]{1}[\.\s]?[0-9]{3}[\.\s]?[0-9]{3})\s*m?\s*(?:S|N|Norte)?", "E_N"),
        # coordenadas UTM Datum WGS84, Huso 18, correspondientes a Este 753.811 m y Norte 5.821.082 m
        (r"(?:Este|UTM Este|\bE\b)\s*[:\-]?\s*([0-9]{3}[\.\s]?[0-9]{3})\s*m?.{0,80}?(?:Norte|UTM Norte|\bN\b|\bS\b)\s*[:\-]?\s*([0-9]{1}[\.\s]?[0-9]{3}[\.\s]?[0-9]{3})", "E_N"),
        # Norte primero
        (r"(?:Norte|UTM Norte|\bN\b|\bS\b)\s*[:\-]?\s*([0-9]{1}[\.\s]?[0-9]{3}[\.\s]?[0-9]{3})\s*m?.{0,80}?(?:Este|UTM Este|\bE\b)\s*[:\-]?\s*([0-9]{3}[\.\s]?[0-9]{3})", "N_E"),
        # Coordenadas: 753811 5821082
        (r"coordenadas?.{0,90}?([0-9]{6})\D+([0-9]{7})", "E_N"),
    ]

    for pat, order in patterns:
        m = re.search(pat, flat, flags=re.I)
        if not m:
            continue
        a = clean_coord(m.group(1))
        b = clean_coord(m.group(2))
        if order == "N_E":
            norte, este = a, b
        else:
            este, norte = a, b
        if valid_pair(norte, este):
            return norte, este

    # Último recurso: buscar candidatos numéricos cerca de palabras clave.
    # Este: 6 dígitos típicos; Norte: 7 dígitos típicos.
    candidate_text = flat
    nums = re.findall(r"\b[0-9]{1,2}(?:\.[0-9]{3}){1,2}\b|\b[0-9]{6,7}\b", candidate_text)
    cleaned = [clean_coord(n) for n in nums]
    estes = [n for n in cleaned if len(n) == 6]
    nortes = [n for n in cleaned if len(n) == 7]
    if estes and nortes:
        return nortes[0], estes[0]

    return "", ""


def parse_direct_bla_text(raw_text: str) -> dict[str, Any]:
    """
    Extrae datos críticos directamente desde texto plano de informe BLA.
    Sirve como refuerzo cuando el Word viene con tablas complejas o cuando el
    usuario sube PDF/JPG/PNG del informe.
    """
    raw_text = raw_text or ""
    compact = re.sub(r"[ \t]+", " ", raw_text)
    compact_one = re.sub(r"\s+", " ", raw_text)
    data: dict[str, Any] = {}

    north, east = parse_coordinates(compact_one)
    if north and east:
        data["utm_norte"] = north
        data["utm_este"] = east

    huso = first_match(r"Huso\s*([0-9]+)", compact_one)
    datum = first_match(r"Datum\s*:?\s*([A-Za-z0-9]+)", compact_one)
    if huso or datum:
        data["datum_huso"] = f"{datum or 'WGS84'} / Huso {huso or '18'}"

    # Caudal solicitado: prioriza frases que incluyan 'solicitado'.
    caudal = first_match(r"caudal\s+solicitado(?:\s+de|\s*:)?\s*([0-9]+(?:[\.,][0-9]+)?)\s*l/?s", compact_one)
    if not caudal:
        caudal = first_match(r"por\s+un\s+caudal(?:\s+de|\s*:)?\s*([0-9]+(?:[\.,][0-9]+)?)\s*l/?s", compact_one)
    if caudal:
        q = to_float(caudal)
        if q > 0:
            data["caudal_l_s"] = q
            data["volumen_anual_m3"] = round(update_volume_from_flow(q), 1)

    # Superficie bajo riego.
    sup = first_match(r"(?:incorporar|riego\s+de|bajo\s+riego\s+de)\s*([0-9]+(?:[\.,][0-9]+)?)\s*ha", compact_one)
    if sup:
        data["hectareas_riego"] = to_float(sup)

    comuna = first_match(r"comuna\s+de\s+([A-Za-zÁÉÍÓÚÑáéíóúñüÜ\s]+?)(?:,|\.|\s+con\s+|\s+por\s+|$)", compact_one)
    if comuna:
        data["comuna"] = clean_detected_value(comuna).strip(" .,:;")

    if data.get("utm_este") and data.get("utm_norte"):
        data["descripcion_ubicacion"] = (
            f"La captación se ubica en la comuna de {data.get('comuna', '')}, "
            f"individualizada mediante coordenadas UTM Este {data.get('utm_este')} m y Norte {data.get('utm_norte')} m, "
            f"{data.get('datum_huso', 'WGS84 / Huso 18')}."
        )

    return {k: v for k, v in data.items() if v not in [None, ""]}


def extract_memoria_explicativa_summary(lines: list[str]) -> str:
    """
    Extrae un resumen directo desde la sección 'Memoria explicativa' del informe BLA.
    Usa los primeros párrafos técnicos y evita tablas/encabezados.
    """
    collecting = False
    selected = []
    stop_words = [
        "Justificación matemática", "Antecedentes Legales", "Análisis de interferencia",
        "Recomendación del consultor", "4.1", "4.2", "5.1"
    ]
    for line in lines:
        clean = re.sub(r"\s+", " ", line).strip()
        if not clean:
            continue
        if re.search(r"^Memoria explicativa$", clean, flags=re.I):
            collecting = True
            continue
        if collecting and any(sw.lower() in clean.lower() for sw in stop_words):
            break
        if collecting:
            # evitar tablas cortas o títulos
            if len(clean) > 80 and not re.search(r"^\d+[\.-]", clean):
                selected.append(clean)
        if len(selected) >= 3:
            break

    if selected:
        text_joined = " ".join(selected)
        # Limitar para que quepa en el formulario.
        if len(text_joined) > 820:
            text_joined = text_joined[:817].rsplit(" ", 1)[0] + "..."
        return text_joined
    return ""

def parse_informe_bla(uploaded_file) -> dict[str, Any]:
    lines = docx_to_lines(uploaded_file)
    text = "\n".join(lines)
    compact = re.sub(r"[ \t]+", " ", text)

    data: dict[str, Any] = {}
    # Refuerzo: extraer coordenadas/caudal/superficie desde todo el texto antes de las tablas.
    data.update(parse_direct_bla_text(text))
    memoria_summary = extract_memoria_explicativa_summary(lines)
    if memoria_summary:
        data["descripcion_proyecto"] = memoria_summary

    # Campos estructurados obligatorios: solo desde tablas/líneas con etiqueta exacta.
    data["beneficiario_nombre"] = get_structured_field(lines, "Solicitante")
    data["beneficiario_rut"] = get_structured_field(lines, "RUT")
    data["beneficiario_domicilio"] = get_structured_field(lines, "Domicilio")
    data["beneficiario_fono"] = get_structured_field(lines, "Teléfono") or get_structured_field(lines, "Telefono")
    data["beneficiario_correo"] = get_structured_field(lines, "Correo electrónico")

    comuna = get_structured_field(lines, "Comuna")
    provincia = get_structured_field(lines, "Provincia")
    if comuna:
        data["comuna"] = comuna
    if provincia:
        data["provincia"] = provincia

    # Derecho solicitado
    caudal_txt = get_structured_field(lines, "Caudal solicitado")
    if not caudal_txt:
        caudal_txt = first_match(r"caudal solicitado(?:\s+de|\s*:)?\s*([0-9\.,]+)\s*l/?s", compact)
    if not caudal_txt:
        caudal_txt = first_match(r"por un caudal(?:\s+solicitado)?(?:\s+de|\s*:)?\s*([0-9\.,]+)\s*l/?s", compact)
    caudal_num = first_match(r"([0-9]+(?:[\.,][0-9]+)?)", caudal_txt)
    if caudal_num:
        data["caudal_l_s"] = to_float(caudal_num)
        data["volumen_anual_m3"] = round(update_volume_from_flow(to_float(caudal_num)), 1)
    else:
        # Último intento: buscar cualquier expresión de caudal en el informe.
        m_caudal = re.search(r"caudal(?:\s+solicitado)?(?:\s+de|\s*:)?\s*([0-9]+(?:[\.,][0-9]+)?)\s*l/?s", compact, flags=re.I)
        if m_caudal:
            data["caudal_l_s"] = to_float(m_caudal.group(1))
            data["volumen_anual_m3"] = round(update_volume_from_flow(data["caudal_l_s"]), 1)

    coord_line = get_structured_field(lines, "Coordenadas, huso")
    north, east = parse_coordinates(coord_line or compact)
    if not (north and east):
        # Segundo intento: buscar en todo el informe, porque a veces las coordenadas están en texto narrativo.
        north, east = parse_coordinates(compact)

    if north and east:
        data["utm_norte"] = north
        data["utm_este"] = east
        data["descripcion_ubicacion"] = (
            f"La captación se ubica en la comuna de {data.get('comuna', '')}, individualizada mediante coordenadas UTM Este {east} m y Norte {north} m, {data.get('datum_huso', 'WGS84 / Huso 18')}."
        )

    huso = first_match(r"Huso\s*([0-9]+)", coord_line or compact)
    datum = get_structured_field(lines, "Datum") or first_match(r"Datum\s*:?\s*([A-Za-z0-9]+)", compact)
    if huso or datum:
        data["datum_huso"] = f"{datum or 'WGS84'} / Huso {huso or '18'}"

    uso = get_structured_field(lines, "Uso del agua")
    if uso and not data.get("descripcion_proyecto"):
        data["descripcion_proyecto"] = (
            f"El proyecto considera la regularización de un derecho de aprovechamiento de aguas subterráneas "
            f"destinado principalmente a {uso.lower()}, asociado al predio del solicitante."
        )

    # Superficie y hectáreas a regar, desde la memoria explicativa.
    sup_match = re.search(
        r"superficie total de\s*([0-9\.,]+)\s*ha.*?incorporar\s*([0-9\.,]+)\s*ha",
        compact,
        re.I | re.S
    )
    if sup_match:
        data["hectareas_riego"] = to_float(sup_match.group(2))
    else:
        ha_match = re.search(r"superficie nueva bajo riego de\s*([0-9\.,]+)\s*ha", compact, re.I)
        if ha_match:
            data["hectareas_riego"] = to_float(ha_match.group(1))

    # Predio: usar domicilio si no hay nombre de predio específico.
    if data.get("beneficiario_domicilio"):
        data["predio"] = data["beneficiario_domicilio"]

    # Descripción de ubicación.
    if data.get("utm_este") and data.get("utm_norte"):
        data["descripcion_ubicacion"] = (
            f"El punto de captación se ubica en la comuna de {data.get('comuna', '')}, "
            f"coordenadas UTM Este {data.get('utm_este')} m y Norte {data.get('utm_norte')} m, "
            f"{data.get('datum_huso', 'WGS84 / Huso 18')}."
        )

    if not data.get("descripcion_proyecto"):
        ha = data.get("hectareas_riego", "")
        data["descripcion_proyecto"] = (
            f"El proyecto corresponde a la regularización de un derecho de aprovechamiento de aguas subterráneas "
            f"destinado al riego predial. El recurso será utilizado para abastecer una superficie agrícola bajo riego"
            f"{f' de {fmt_num(ha, 2)} ha' if ha else ''}, fortaleciendo la actividad productiva familiar y permitiendo "
            f"mejorar la eficiencia y seguridad hídrica del predio."
        )

    # Valores fijos para tus casos BLA: subterránea, consuntivo, permanente, continuo, riego + subsistencia.
    data["tipo_persona"] = "Persona jurídica"
    data["sexo"] = ""
    data["naturaleza"] = "Subterránea"
    data["tipo_derecho"] = "Consuntivo"
    data["ejercicio_1"] = "Permanente"
    data["ejercicio_2"] = "Continuo"
    data["uso_agua"] = USOS_OPCIONES[0]
    data["porcentaje_riego"] = 95.0
    data["porcentaje_subsistencia"] = 5.0
    data["personas_beneficiadas"] = 1

    # Firmante: solo usar nombre/rut estructurado, nunca frases narrativas.
    data["firmante_nombre"] = PETICIONARIO_EMPRESA["nombre"]
    data["firmante_rut"] = PETICIONARIO_EMPRESA["rut"]

    # Peticionario del formulario: empresa Irrisal, no beneficiario BLA.
    data.update(PETICIONARIO_EMPRESA)
    data["firmante_nombre"] = PETICIONARIO_EMPRESA["nombre"]
    data["firmante_rut"] = PETICIONARIO_EMPRESA["rut"]

    return {k: v for k, v in data.items() if v not in [None, ""]}

# =============================================================================
# AUTOCOMPLETAR DESDE AVALÚO FISCAL Y CONSERVADOR
# =============================================================================

def uploaded_file_to_text_any(uploaded_file) -> str:
    """
    Extrae texto desde PDF, JPG o PNG.
    PDF con texto: usa PyMuPDF. PDF escaneado o imagen: usa OCR con Tesseract.
    """
    if uploaded_file is None:
        return ""
    name = (uploaded_file.name or "").lower()
    raw = uploaded_file.getvalue()
    try:
        if name.endswith(".pdf"):
            doc_pdf = fitz.open(stream=raw, filetype="pdf")
            texts = []
            for page in doc_pdf:
                page_text = page.get_text("text") or ""
                if page_text.strip():
                    texts.append(page_text)
            text_joined = "\n".join(texts).strip()
            if len(text_joined) < 80:
                ocr_texts = []
                for page in list(doc_pdf)[:3]:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_texts.append(pytesseract.image_to_string(img, lang="spa"))
                text_joined = "\n".join(ocr_texts)
            return re.sub(r"[ \t]+", " ", text_joined)
        img = Image.open(BytesIO(raw)).convert("RGB")
        return re.sub(r"[ \t]+", " ", pytesseract.image_to_string(img, lang="spa"))
    except Exception:
        return ""


def normalize_doc_text(value: str) -> str:
    value = value or ""
    value = value.replace("\xa0", " ")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{2,}", "\n", value)
    return value.strip()


def find_rol_sii(text_value: str) -> str:
    text_value = normalize_doc_text(text_value)
    patterns = [
        r"(?:Rol|ROL|Rol de Aval[uú]o|Rol SII|N[°º]\s*Rol)\s*[:\-]?\s*([0-9]{1,6}\s*[-–]\s*[0-9]{1,6})",
        r"\b([0-9]{1,6}\s*[-–]\s*[0-9]{1,6})\b",
    ]
    for pat in patterns:
        m = re.search(pat, text_value, flags=re.I)
        if m:
            return m.group(1).replace(" ", "").replace("–", "-")
    return ""


def find_surface_ha(text_value: str):
    text_value = normalize_doc_text(text_value)
    patterns = [
        r"Superficie(?:\s+total)?\s*[:\-]?\s*([0-9]+(?:[\.,][0-9]+)?)\s*(?:ha|h[aá]s|hect[aá]reas)",
        r"([0-9]+(?:[\.,][0-9]+)?)\s*(?:ha|h[aá]s|hect[aá]reas)",
    ]
    for pat in patterns:
        m = re.search(pat, text_value, flags=re.I)
        if m:
            try: return to_float(m.group(1))
            except Exception: pass
    return None


def find_comuna_from_text(text_value: str) -> str:
    text_value = normalize_doc_text(text_value)
    patterns = [
        r"Comuna\s*[:\-]?\s*([A-Za-zÁÉÍÓÚÑáéíóúñüÜ\s]+?)(?:\n|Provincia|Regi[oó]n|Rol|$)",
        r"Comuna de\s+([A-Za-zÁÉÍÓÚÑáéíóúñüÜ\s]+?)(?:\.|,|\n|$)",
    ]
    for pat in patterns:
        m = re.search(pat, text_value, flags=re.I)
        if m:
            return clean_detected_value(m.group(1)).strip(" .,:;")
    return ""


def find_predio_from_avaluo(text_value: str) -> str:
    text_value = normalize_doc_text(text_value)
    patterns = [
        r"(?:Nombre del predio|Predio|Direcci[oó]n predial|Ubicaci[oó]n)\s*[:\-]?\s*([^\n]{5,90})",
        r"(?:Direcci[oó]n|Ubicaci[oó]n)\s*[:\-]?\s*([^\n]{5,90})",
    ]
    for pat in patterns:
        m = re.search(pat, text_value, flags=re.I)
        if m:
            value = clean_detected_value(m.group(1))
            if value and not re.search(r"servicio|impuestos|internos|certificado|aval[uú]o", value, re.I):
                return value.strip(" .,:;")
    return ""


def find_cbr_data(text_value: str) -> dict[str, str]:
    """
    Extrae fojas, número, año y conservador desde una inscripción CBR.
    Es más tolerante a OCR y formatos notariales/conservatorios.
    """
    original = normalize_doc_text(text_value)
    flat = re.sub(r"\s+", " ", original)

    # Normalizar errores OCR frecuentes sin perder el texto original
    flat_norm = flat
    flat_norm = flat_norm.replace("Nº", "N°").replace("Nro", "N°").replace("No ", "N° ")
    flat_norm = re.sub(r"\bfoja\b", "fojas", flat_norm, flags=re.I)

    data: dict[str, str] = {}

    patterns_combo = [
        # "fojas 123 número 456 año 2020"
        r"fojas?\s*([0-9]+)\s*(?:vta\.?|vuelta)?\s*(?:,|\.|\s)*\s*(?:N[°º]|n[uú]mero|numero|nro\.?)\s*([0-9]+).*?(?:año|ano|del año|del ano)\s*([12][0-9]{3})",
        # "a fojas 123 ... Nº 456 ... Registro ... 2020"
        r"a\s+fojas?\s*([0-9]+).*?(?:N[°º]|n[uú]mero|numero|nro\.?)\s*([0-9]+).*?([12][0-9]{3})",
        # "Fojas: 123 / Número: 456 / Año: 2020"
        r"Fojas?\s*[:\-]?\s*([0-9]+).*?(?:N[°º]|n[uú]mero|numero|nro\.?)\s*[:\-]?\s*([0-9]+).*?Año\s*[:\-]?\s*([12][0-9]{3})",
    ]

    for pat in patterns_combo:
        m = re.search(pat, flat_norm, flags=re.I)
        if m:
            data["fojas"] = m.group(1)
            data["numero_inscripcion"] = m.group(2)
            data["anio_inscripcion"] = m.group(3)
            break

    if "fojas" not in data:
        fojas = first_match(r"fojas?\s*[:\-]?\s*([0-9]+)", flat_norm)
        if fojas:
            data["fojas"] = fojas

    if "numero_inscripcion" not in data:
        numero = first_match(r"(?:N[°º]|n[uú]mero|numero|nro\.?)\s*[:\-]?\s*([0-9]+)", flat_norm)
        if numero:
            data["numero_inscripcion"] = numero

    if "anio_inscripcion" not in data:
        anio = first_match(r"(?:año|ano|del año|del ano|Registro de Propiedad del año)\s*[:\-]?\s*([12][0-9]{3})", flat_norm)
        if not anio:
            # Último recurso: primer año razonable cerca de "Registro de Propiedad"
            m = re.search(r"Registro de Propiedad.{0,80}?([12][0-9]{3})", flat_norm, flags=re.I)
            anio = m.group(1) if m else ""
        if anio:
            data["anio_inscripcion"] = anio

    # Conservador
    conservador_patterns = [
        r"Conservador(?: de Bienes Ra[ií]ces)?\s+de\s+([A-Za-zÁÉÍÓÚÑáéíóúñüÜ\s]+?)(?:,|\.|\n| certifica|$)",
        r"Conservador(?: de Bienes Ra[ií]ces)?\s*[:\-]\s*([A-Za-zÁÉÍÓÚÑáéíóúñüÜ\s]+?)(?:,|\.|\n|$)",
        r"Registro de Propiedad.*?Conservador(?: de Bienes Ra[ií]ces)?\s+de\s+([A-Za-zÁÉÍÓÚÑáéíóúñüÜ\s]+?)(?:,|\.|\n|$)",
    ]
    for pat in conservador_patterns:
        m = re.search(pat, flat_norm, flags=re.I)
        if m:
            cons = clean_detected_value(m.group(1)).strip(" .,:;")
            cons = re.sub(r"\s+", " ", cons)
            if cons and not re.search(r"fojas|numero|año|registro", cons, re.I):
                data["conservador"] = _clean_conservador_value(cons)
                break

    return data

def parse_avaluo_fiscal(uploaded_file) -> dict[str, Any]:
    text_value = uploaded_file_to_text_any(uploaded_file)
    data: dict[str, Any] = {}
    rol = find_rol_sii(text_value)
    if rol: data['rol_sii'] = rol
    superficie = find_surface_ha(text_value)
    if superficie is not None and superficie > 0:
        data['hectareas_riego'] = superficie
    comuna = find_comuna_from_text(text_value)
    if comuna: data['comuna'] = comuna
    predio = find_predio_from_avaluo(text_value)
    if predio: data['predio'] = predio
    return {k: v for k, v in data.items() if v not in [None, '']}


def parse_conservador(uploaded_file) -> dict[str, Any]:
    text_value = uploaded_file_to_text_any(uploaded_file)
    data = find_cbr_data(text_value)
    predio = first_match(r"(?:inmueble|predio|propiedad)\s+(?:ubicado|situado|denominado)?\s*(?:en)?\s*([^\n]{8,120})", text_value)
    if predio:
        predio = clean_detected_value(predio)
        if predio and not re.search(r"inscripci[oó]n|fojas|registro|conservador", predio, re.I):
            data.setdefault('predio', predio.strip(' .,:;'))
    return {k: v for k, v in data.items() if v not in [None, '']}


# =============================================================================
# WORD HELPERS
# =============================================================================


def set_cell_shading(cell, fill: str = "D9D9D9"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_border(cell, color="000000", sz="8"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), sz)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def set_cell_text(cell, text: str = "", bold: bool = False, size: float = 7.0, align=WD_ALIGN_PARAGRAPH.LEFT, shade: bool = False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(clean_text(text))
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.bold = bold
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_border(cell)
    if shade:
        set_cell_shading(cell)


def merge_cells(row, start: int, end: int):
    if end <= start:
        return row.cells[start]
    cell = row.cells[start]
    for i in range(start + 1, end + 1):
        cell = cell.merge(row.cells[i])
    return cell


def add_doc_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(11 if level == 1 else 10)
    run.bold = True
    return p


def add_doc_paragraph(doc: Document, text: str, size: float = 9.0, justify: bool = True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_after = Pt(5)
    run = p.add_run(clean_text(text))
    run.font.name = "Arial"
    run.font.size = Pt(size)
    return p


def add_form_table(doc: Document, rows: int, cols: int, widths: list[float] | None = None):
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = True
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return table


def add_signature_docx(doc: Document, data: dict, signature_file):
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if signature_file is not None:
        try:
            p.add_run().add_picture(BytesIO(signature_file.getvalue()), width=Inches(2.2))
        except Exception:
            pass
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p2.add_run("_____________________________________\nFirma y RUT del Solicitante o Representante Legal")
    r.font.name = "Arial"
    r.font.size = Pt(9)
    r.bold = True
    if data.get("firmante_nombre") or data.get("firmante_rut"):
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = p3.add_run(f"{clean_text(data.get('firmante_nombre'))} - RUT {clean_text(data.get('firmante_rut'))}")
        r3.font.name = "Arial"
        r3.font.size = Pt(9)


# =============================================================================
# WORD GENERATION
# =============================================================================


def make_docx(data: dict[str, Any], signature_file=None, croquis_file=None) -> bytes:
    data = force_peticionario_empresa(data)
    data = _normalize_export_data(data)
    """
    Exporta Word con formato visual idéntico al formulario oficial.
    Técnica usada: se genera primero el PDF sobre la plantilla oficial y luego se inserta
    cada página renderizada como imagen a página completa en un DOCX tamaño oficio.
    Esto mantiene la apariencia exacta, aunque el contenido no queda editable como texto.
    """
    pdf_bytes = make_pdf(data, signature_file=signature_file, croquis_file=croquis_file)
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    docx = Document()
    sec = docx.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(13)
    sec.top_margin = Inches(0)
    sec.bottom_margin = Inches(0)
    sec.left_margin = Inches(0)
    sec.right_margin = Inches(0)

    for i, page in enumerate(pdf_doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        img_bytes = pix.tobytes("png")
        if i > 0:
            docx.add_page_break()
        p = docx.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(BytesIO(img_bytes), width=Inches(8.5), height=Inches(13))

    out = BytesIO()
    docx.save(out)
    out.seek(0)
    return out.getvalue()


# =============================================================================
# PDF GENERATION
# =============================================================================


def pdf_styles():
    s = getSampleStyleSheet()
    s.add(ParagraphStyle("DGAHeader", parent=s["Normal"], fontSize=7, leading=8, fontName="Helvetica-Bold"))
    s.add(ParagraphStyle("DGATitle", parent=s["Normal"], fontSize=16, leading=18, fontName="Helvetica-Bold", alignment=TA_CENTER))
    s.add(ParagraphStyle("Section", parent=s["Normal"], fontSize=10, leading=12, fontName="Helvetica-Bold", spaceBefore=8, spaceAfter=4, keepWithNext=True))
    s.add(ParagraphStyle("Body", parent=s["Normal"], fontSize=8.5, leading=11, fontName="Helvetica", alignment=TA_JUSTIFY))
    s.add(ParagraphStyle("Cell", parent=s["Normal"], fontSize=7, leading=8.2, fontName="Helvetica"))
    s.add(ParagraphStyle("CellBold", parent=s["Cell"], fontName="Helvetica-Bold"))
    s.add(ParagraphStyle("Small", parent=s["Normal"], fontSize=6.6, leading=8, fontName="Helvetica"))
    return s


def P(text: Any, style):
    return Paragraph(clean_text(text), style)


def pdf_table(data, widths=None, shade_first_col=False):
    styles = pdf_styles()
    wrapped = []
    for row in data:
        wrapped.append([P(cell, styles["Cell"]) for cell in row])
    t = Table(wrapped, colWidths=widths, repeatRows=0)
    ts = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if shade_first_col:
        ts.append(("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#D9D9D9")))
    t.setStyle(TableStyle(ts))
    return t


def pdf_shaded_table(data, widths=None, shaded_cells: list[tuple[int, int]] | None = None):
    styles = pdf_styles()
    wrapped = []
    for row in data:
        wrapped.append([P(cell, styles["Cell"]) for cell in row])
    t = Table(wrapped, colWidths=widths)
    ts = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if shaded_cells:
        for r, c in shaded_cells:
            ts.append(("BACKGROUND", (c, r), (c, r), colors.HexColor("#D9D9D9")))
    t.setStyle(TableStyle(ts))
    return t


def add_pdf_signature(story, data, signature_file, styles):
    story.append(Spacer(1, 1.0 * cm))
    if signature_file is not None:
        try:
            img = RLImage(BytesIO(signature_file.getvalue()), width=4.8 * cm, height=2.0 * cm)
            tbl = Table([[img]], colWidths=[8 * cm])
            tbl.hAlign = "CENTER"
            tbl.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("GRID", (0, 0), (-1, -1), 0, colors.white)]))
            story.append(tbl)
        except Exception:
            pass
    story.append(P("_____________________________________", ParagraphStyle("Sig", parent=styles["Body"], alignment=TA_CENTER, fontSize=9, leading=11, fontName="Helvetica-Bold")))
    story.append(P("Firma y RUT del Solicitante o Representante Legal", ParagraphStyle("Sig2", parent=styles["Body"], alignment=TA_CENTER, fontSize=8, leading=10, fontName="Helvetica-Bold")))
    if data.get("firmante_nombre") or data.get("firmante_rut"):
        story.append(P(f"{clean_text(data.get('firmante_nombre'))} - RUT {clean_text(data.get('firmante_rut'))}", ParagraphStyle("Sig3", parent=styles["Body"], alignment=TA_CENTER, fontSize=8, leading=10)))


def fmt_num_blank(value: Any, decimals: int = 2) -> str:
    if value is None or clean_text(value) == "":
        return ""
    n = to_float(value, 0.0)
    if n == 0:
        return ""
    return fmt_num(n, decimals)


def _pdf_text(page, x, y, w, h, value, size=8.5, bold=False, align=0):
    """
    Escribe texto sobre la plantilla oficial en coordenadas PDF.
    """
    value = clean_text(value)
    if not value:
        return
    fontname = "helv"
    page.insert_textbox(
        fitz.Rect(x, y, x + w, y + h),
        value,
        fontsize=size,
        fontname=fontname,
        color=(0, 0, 0),
        align=align,
    )


def _coord_digits(value: Any, n_boxes: int) -> str:
    """
    Normaliza coordenadas UTM para escribir un dígito por casilla.
    Corrige casos como 5869412.0, 5.869.412, 747711.0 o 747.711.
    """
    if value is None:
        return ""

    if isinstance(value, (int, float)):
        try:
            raw = str(int(float(value)))
        except Exception:
            raw = ""
    else:
        s = clean_text(value)
        # caso: "5869412.0"
        m = re.match(r"^\s*([0-9]+)(?:\.0+)?\s*$", s)
        if m:
            raw = m.group(1)
        else:
            raw = re.sub(r"[^0-9]", "", s)

    if not raw:
        return ""

    # Si por un decimal .0 quedó un dígito extra, eliminarlo.
    if len(raw) == n_boxes + 1 and raw.endswith("0"):
        raw = raw[:-1]

    if len(raw) > n_boxes:
        raw = raw[-n_boxes:]

    return raw.rjust(n_boxes)


def _pdf_digits_in_boxes(page, x0, y0, cell_w, h, value, n_boxes, size=9.0):
    """
    Escribe coordenadas UTM con un dígito por casilla.
    """
    raw = _coord_digits(value, n_boxes)
    if not raw:
        return
    for i, ch in enumerate(raw):
        if not ch.strip():
            continue
        page.insert_textbox(
            fitz.Rect(x0 + i * cell_w, y0, x0 + (i + 1) * cell_w, y0 + h),
            ch,
            fontsize=size,
            fontname="helv",
            color=(0, 0, 0),
            align=1,
        )


def _pdf_cell_text(page, x0, y0, x1, y1, value, size=8.0, bold=False, align=1):
    value = clean_text(value)
    if not value:
        return
    page.insert_textbox(
        fitz.Rect(x0, y0, x1, y1),
        value,
        fontsize=size,
        fontname="helv",
        color=(0, 0, 0),
        align=align,
    )


def _pdf_draw_cell(page, x0, y0, x1, y1, fill=None, width=0.7):
    if fill is not None:
        page.draw_rect(fitz.Rect(x0, y0, x1, y1), color=(0, 0, 0), fill=fill, width=width, overlay=True)
    else:
        page.draw_rect(fitz.Rect(x0, y0, x1, y1), color=(0, 0, 0), width=width, overlay=True)


def _pdf_label(page, x0, y0, x1, y1, text_value, size=8.0, bold=False, align=1):
    page.insert_textbox(
        fitz.Rect(x0, y0, x1, y1),
        clean_text(text_value),
        fontsize=size,
        fontname="helv",
        color=(0, 0, 0),
        align=align,
    )


def _draw_custom_42_43(page, data: dict[str, Any]):
    """
    Redibuja la página 5 del formulario dejando solo los usos aplicables:
    4.2 uso doméstico de subsistencia y 4.3 riego.
    El bloque 4.1 Agua Potable queda eliminado visualmente.
    """
    # Limpiar toda la zona del punto 4 original, incluyendo 4.1.
    page.draw_rect(fitz.Rect(0, 45, 612, 875), color=(1, 1, 1), fill=(1, 1, 1), overlay=True)

    region = _region_text(data.get("region"))
    provincia = clean_text(data.get("provincia")) or "Biobío"
    comuna = clean_text(data.get("comuna")).title()
    predio = _clean_predio_value(data.get("predio")).title()
    rol = clean_text(data.get("rol_sii"))
    has = fmt_num_blank(data.get("hectareas_riego"), 2)
    personas = fmt_num(data.get("personas_beneficiadas"), 0)
    fojas = clean_text(data.get("fojas"))
    numero = clean_text(data.get("numero_inscripcion"))
    anio = clean_text(data.get("anio_inscripcion"))
    conservador = _clean_conservador_value(data.get("conservador"))

    gray = (0.86, 0.86, 0.86)

    # Título general del punto 4
    _pdf_label(page, 46, 58, 560, 78, "4. Antecedentes complementarios del proyecto", size=13, bold=True, align=0)
    _pdf_label(page, 46, 76, 560, 94, "(complete según el o los usos del derecho de agua requerido por el proyecto, de acuerdo a lo señalado en el numeral 3.3.)", size=7.4, align=0)

    # ===== 4.2 =====
    _pdf_label(page, 58, 116, 560, 138, "4.2. Antecedentes requeridos para uso Doméstico de Subsistencia", size=12, bold=True, align=0)

    x0, y0, x1 = 70, 160, 545
    left_w = 93
    col_w = (x1 - x0 - left_w) / 3
    y_header = y0 + 22
    y_values = y0 + 55
    y_bottom = y0 + 105

    _pdf_draw_cell(page, x0, y0, x0 + left_w, y_values, fill=None)
    _pdf_label(page, x0 + 3, y0 + 14, x0 + left_w - 3, y_values - 6, "Antecedentes\nde ubicación", size=8.3, bold=True, align=1)

    for i, lab in enumerate(["Región", "Provincia", "Comuna"]):
        xa = x0 + left_w + i * col_w
        xb = xa + col_w
        _pdf_draw_cell(page, xa, y0, xb, y_header, fill=gray)
        _pdf_label(page, xa, y0 + 4, xb, y_header - 2, lab, size=8.2, bold=True)
        _pdf_draw_cell(page, xa, y_header, xb, y_values, fill=None)

    _pdf_cell_text(page, x0 + left_w, y_header + 9, x0 + left_w + col_w, y_values - 3, region, size=7.4)
    _pdf_cell_text(page, x0 + left_w + col_w, y_header + 9, x0 + left_w + 2 * col_w, y_values - 3, provincia, size=7.4)
    _pdf_cell_text(page, x0 + left_w + 2 * col_w, y_header + 9, x1, y_values - 3, comuna, size=7.4)

    _pdf_draw_cell(page, x0, y_values, x0 + left_w, y_bottom, fill=None)
    _pdf_label(page, x0 + 3, y_values + 12, x0 + left_w - 3, y_bottom - 6, "Antecedentes\nrequeridos", size=8.3, bold=True)
    _pdf_draw_cell(page, x0 + left_w, y_values, x1 - 125, y_bottom, fill=gray)
    _pdf_label(page, x0 + left_w + 5, y_values + 21, x1 - 125, y_bottom - 10, "N° de personas beneficiadas (valor en números)", size=7.8, bold=True, align=0)
    _pdf_draw_cell(page, x1 - 125, y_values, x1, y_bottom, fill=None)
    _pdf_cell_text(page, x1 - 125, y_values + 18, x1, y_bottom - 4, personas, size=8.2)

    # ===== 4.3 =====
    _pdf_label(page, 58, 306, 560, 328, "4.3. Antecedentes requeridos para uso en Riego", size=12, bold=True, align=0)

    x0, y0, x1 = 70, 352, 545
    left_w = 93
    col_w = (x1 - x0 - left_w) / 3
    y_header = y0 + 22
    y_values = y0 + 55
    y_predio_label = y_values
    y_predio_value = y_values + 30
    y_rol = y_predio_value + 34
    y_bottom = y_rol + 34

    _pdf_draw_cell(page, x0, y0, x0 + left_w, y_bottom, fill=None)
    _pdf_label(page, x0 + 3, y0 + 48, x0 + left_w - 3, y_bottom - 30, "Antecedentes\nde ubicación", size=8.3, bold=True)

    for i, lab in enumerate(["Región", "Provincia", "Comuna"]):
        xa = x0 + left_w + i * col_w
        xb = xa + col_w
        _pdf_draw_cell(page, xa, y0, xb, y_header, fill=gray)
        _pdf_label(page, xa, y0 + 4, xb, y_header - 2, lab, size=8.2, bold=True)
        _pdf_draw_cell(page, xa, y_header, xb, y_values, fill=None)

    _pdf_cell_text(page, x0 + left_w, y_header + 9, x0 + left_w + col_w, y_values - 3, region, size=7.4)
    _pdf_cell_text(page, x0 + left_w + col_w, y_header + 9, x0 + left_w + 2 * col_w, y_values - 3, provincia, size=7.4)
    _pdf_cell_text(page, x0 + left_w + 2 * col_w, y_header + 9, x1, y_values - 3, comuna, size=7.4)

    _pdf_draw_cell(page, x0 + left_w, y_predio_label, x1, y_predio_value, fill=gray)
    _pdf_label(page, x0 + left_w + 5, y_predio_label + 5, x1 - 5, y_predio_value - 3, "Nombre del predio beneficiado", size=7.8, bold=True, align=0)
    _pdf_draw_cell(page, x0 + left_w, y_predio_value, x1, y_rol, fill=None)
    _pdf_cell_text(page, x0 + left_w + 4, y_predio_value + 9, x1 - 4, y_rol - 3, predio, size=7.4, align=0)

    _pdf_draw_cell(page, x0 + left_w, y_rol, x0 + left_w + 72, y_bottom, fill=gray)
    _pdf_label(page, x0 + left_w + 4, y_rol + 9, x0 + left_w + 72, y_bottom - 4, "N° Rol SII", size=7.8, bold=True)
    _pdf_draw_cell(page, x0 + left_w + 72, y_rol, x0 + left_w + 230, y_bottom, fill=None)
    _pdf_cell_text(page, x0 + left_w + 72, y_rol + 9, x0 + left_w + 230, y_bottom - 4, rol, size=7.6)

    _pdf_draw_cell(page, x0 + left_w + 230, y_rol, x0 + left_w + 330, y_bottom, fill=gray)
    _pdf_label(page, x0 + left_w + 230, y_rol + 9, x0 + left_w + 330, y_bottom - 4, "N° Hás a regar", size=7.8, bold=True)
    _pdf_draw_cell(page, x0 + left_w + 330, y_rol, x1, y_bottom, fill=None)
    _pdf_cell_text(page, x0 + left_w + 330, y_rol + 9, x1, y_bottom - 4, has, size=7.6)

    # Antecedentes legales
    y0 = y_bottom + 26
    y1 = y0 + 25
    y2 = y1 + 34
    y3 = y2 + 34
    y4 = y3 + 34

    x_left = 70
    x_label = 230
    x_val = 330
    x_cons = 545

    _pdf_draw_cell(page, x_left, y0, x_label, y4, fill=None)
    _pdf_label(page, x_left + 3, y0 + 42, x_label - 3, y4 - 35, "Antecedentes legales del\npredio", size=8.3, bold=True)

    _pdf_draw_cell(page, x_label, y0, x_cons, y1, fill=gray)
    _pdf_label(page, x_label, y0 + 6, x_cons, y1 - 2, "Inscripción del predio en Conservador de Bienes Raíces", size=8.2, bold=True)

    _pdf_draw_cell(page, x_label, y1, x_val, y2, fill=gray)
    _pdf_label(page, x_label + 4, y1 + 10, x_val, y2 - 4, "Fojas", size=8.1, bold=True, align=0)
    _pdf_draw_cell(page, x_val, y1, 405, y2, fill=None)
    _pdf_cell_text(page, x_val, y1 + 10, 405, y2 - 4, fojas, size=7.8)

    _pdf_draw_cell(page, x_label, y2, x_val, y3, fill=gray)
    _pdf_label(page, x_label + 4, y2 + 10, x_val, y3 - 4, "Número", size=8.1, bold=True, align=0)
    _pdf_draw_cell(page, x_val, y2, 405, y3, fill=None)
    _pdf_cell_text(page, x_val, y2 + 10, 405, y3 - 4, numero, size=7.8)

    _pdf_draw_cell(page, x_label, y3, x_val, y4, fill=gray)
    _pdf_label(page, x_label + 4, y3 + 10, x_val, y4 - 4, "Año", size=8.1, bold=True, align=0)
    _pdf_draw_cell(page, x_val, y3, 405, y4, fill=None)
    _pdf_cell_text(page, x_val, y3 + 10, 405, y4 - 4, anio, size=7.8)

    _pdf_draw_cell(page, 405, y1, x_cons, y2, fill=gray)
    _pdf_label(page, 405, y1 + 10, x_cons, y2 - 4, "Conservador", size=8.1, bold=True)
    _pdf_draw_cell(page, 405, y2, x_cons, y4, fill=None)
    _pdf_cell_text(page, 405, y2 + 25, x_cons, y4 - 4, conservador, size=7.8)

    _pdf_text(page, 70, y4 + 20, 475, 30, "Se adjunta copia de la inscripción conservatoria vigente del predio y antecedentes legales de respaldo para la tramitación.", size=7.0)


def _build_location_description(data: dict[str, Any]) -> str:
    # Si hay descripción manual, usarla solo si no es genérica.
    desc = clean_text(data.get("descripcion_ubicacion"))
    if desc and "antecedentes técnicos" not in desc.lower():
        return desc
    comuna = clean_text(data.get("comuna"))
    predio = _clean_predio_value(data.get("predio"))
    este = clean_text(data.get("utm_este"))
    norte = clean_text(data.get("utm_norte"))
    datum = clean_text(data.get("datum_huso"), "WGS84 / Huso 18")
    partes = []
    if predio:
        partes.append(f"La captación se ubica en el predio {predio}")
    else:
        partes.append("La captación se ubica en el predio individualizado en los antecedentes del expediente")
    if comuna:
        partes.append(f"comuna de {comuna}")
    if este and norte:
        partes.append(f"coordenadas UTM Este {este} m y Norte {norte} m, {datum}")
    else:
        partes.append(f"con ubicación respaldada por los antecedentes técnicos y cartográficos adjuntos, {datum}")
    return ", ".join(partes) + ". La obra corresponde a una captación de aguas subterráneas destinada al abastecimiento de riego predial y uso doméstico de subsistencia."


def _build_project_description(data: dict[str, Any]) -> str:
    desc = clean_text(data.get("descripcion_proyecto"))
    # Evitar texto excesivamente largo de informe, usarlo si parece bien formado y no se superpone.
    if desc and len(desc) <= 650 and "caudal de l/s" not in desc.lower():
        return desc
    ha_val = to_float(data.get("hectareas_riego"), 0)
    ha_txt = f"aproximadamente {fmt_num(ha_val, 2)} ha" if ha_val > 0 else "la superficie predial indicada en los antecedentes"
    comuna = clean_text(data.get("comuna"))
    predio = _clean_predio_value(data.get("predio"))
    caudal = fmt_num_blank(data.get("caudal_l_s"), 2)
    caudal_txt = f"por un caudal de {caudal} l/s" if caudal else "por el caudal indicado en la solicitud"
    return (
        f"El proyecto corresponde a la solicitud de un derecho de aprovechamiento de aguas subterráneas, "
        f"de carácter consuntivo, permanente y continuo, {caudal_txt}. "
        f"El recurso será destinado principalmente al riego predial y, en menor proporción, al uso doméstico de subsistencia. "
        f"La captación abastecerá el predio {predio}, ubicado en la comuna de {comuna}, permitiendo regar {ha_txt}, "
        f"mejorar la seguridad hídrica del sistema productivo y respaldar la actividad agrícola familiar."
    )


def _build_additional_info(data: dict[str, Any]) -> str:
    caudal = fmt_num_blank(data.get("caudal_l_s"), 2)
    ha_val = to_float(data.get("hectareas_riego"), 0)
    ha_txt = f"una superficie aproximada de {fmt_num(ha_val, 2)} ha bajo riego" if ha_val > 0 else "la superficie agrícola declarada en los antecedentes"
    caudal_txt = f"por un caudal solicitado de {caudal} l/s" if caudal else "por el caudal indicado en la solicitud"
    comuna = clean_text(data.get("comuna"))
    predio = _clean_predio_value(data.get("predio"))
    return (
        f"La solicitud se fundamenta en la regularización de una captación de aguas subterráneas destinada principalmente al riego del predio {predio}, "
        f"ubicado en la comuna de {comuna}, {caudal_txt} y {ha_txt}. "
        f"El uso doméstico de subsistencia se considera como uso complementario y minoritario del recurso. "
        f"Se adjuntan antecedentes técnicos, cartográficos y legales de respaldo, incluyendo informe técnico, avalúo fiscal, inscripción conservatoria vigente y croquis de ubicación."
    )


def _add_signature_to_pdf_page(page, signature_file, data):
    """
    Firma en página final del formulario oficial.

    Ajuste v2.5: la firma PNG se inserta centrada en el espacio blanco
    inmediatamente sobre la línea oficial de firma. No se agrega texto bajo
    la línea, porque el formulario ya contiene la leyenda "Firma y RUT del
    Solicitante o Representante Legal". Esto evita que la firma o el texto
    aparezcan fuera del espacio asignado.
    """
    if signature_file is not None:
        try:
            # Espacio útil sobre la línea de firma de la página 10 del formulario.
            # Coordenadas calibradas para tamaño carta/oficio del PDF oficial.
            sig_rect = fitz.Rect(205, 455, 405, 535)
            page.insert_image(sig_rect, stream=signature_file.getvalue(), keep_proportion=True, overlay=True)
        except Exception:
            pass


def _normalize_export_data(data: dict[str, Any]) -> dict[str, Any]:
    """Normaliza campos antes de exportar y calcula valores derivados."""
    out = dict(data or {})

    # Si el caudal viene como texto en otra clave, rescatarlo.
    if to_float(out.get("caudal_l_s"), 0) <= 0:
        for key in ["caudal", "caudal_solicitado", "caudal_lts_seg", "caudal_l/s"]:
            if out.get(key):
                out["caudal_l_s"] = to_float(out.get(key), 0)
                break

    # Calcular volumen anual si hay caudal y el volumen está vacío/cero.
    if to_float(out.get("caudal_l_s"), 0) > 0 and to_float(out.get("volumen_anual_m3"), 0) <= 0:
        out["volumen_anual_m3"] = round(update_volume_from_flow(out.get("caudal_l_s")), 1)

    # Intentar extraer coordenadas desde la descripción si los campos directos están vacíos.
    if not clean_text(out.get("utm_norte")) or not clean_text(out.get("utm_este")):
        text_blob = " ".join([clean_text(out.get("descripcion_ubicacion")), clean_text(out.get("descripcion_proyecto")), clean_text(out.get("informacion_adicional"))])
        north, east = parse_coordinates(text_blob)
        if north and east:
            out["utm_norte"] = north
            out["utm_este"] = east

    # Normalizar coordenadas si el usuario las ingresó con puntos de miles.
    if clean_text(out.get("utm_norte")):
        out["utm_norte"] = re.sub(r"[^0-9]", "", clean_text(out.get("utm_norte")))
    if clean_text(out.get("utm_este")):
        out["utm_este"] = re.sub(r"[^0-9]", "", clean_text(out.get("utm_este")))

    out["region"] = _region_text(out.get("region"))
    out["predio"] = _clean_predio_value(out.get("predio"))
    out["conservador"] = _clean_conservador_value(out.get("conservador"))
    return out


def force_peticionario_empresa(data: dict[str, Any]) -> dict[str, Any]:
    """
    Fuerza que el punto 1 del formulario DGA identifique a Irrisal Consulting Ltda.
    como peticionario. No altera los datos prediales/productivos del beneficiario.
    """
    out = dict(data or {})
    out.update(PETICIONARIO_EMPRESA)
    out["firmante_nombre"] = PETICIONARIO_EMPRESA["nombre"]
    out["firmante_rut"] = PETICIONARIO_EMPRESA["rut"]
    return out


def make_pdf(data: dict[str, Any], signature_file=None, croquis_file=None) -> bytes:
    data = force_peticionario_empresa(data)
    data = _normalize_export_data(data)
    """
    Exporta PDF usando el formulario oficial DGA como plantilla.
    Versión 2.3: coordenadas recalibradas y páginas aplicables a riego + subsistencia.
    """
    if not TEMPLATE_DGA_PDF.exists():
        raise FileNotFoundError("No se encontró la plantilla PDF oficial en assets.")

    # Limpieza y normalización previa
    data = dict(data or {})
    data["region"] = _region_text(data.get("region"))
    data["predio"] = _clean_predio_value(data.get("predio"))
    data["conservador"] = _clean_conservador_value(data.get("conservador"))
    if not clean_text(data.get("descripcion_proyecto")):
        data["descripcion_proyecto"] = _build_project_description(data)
    if not clean_text(data.get("descripcion_ubicacion")) or "antecedentes técnicos" in clean_text(data.get("descripcion_ubicacion")).lower():
        data["descripcion_ubicacion"] = _build_location_description(data)
    if not clean_text(data.get("informacion_adicional")) or clean_text(data.get("informacion_adicional")).startswith("Se adjuntan antecedentes"):
        data["informacion_adicional"] = _build_additional_info(data)

    src = fitz.open(str(TEMPLATE_DGA_PDF))
    out = fitz.open()
    # Páginas: 1, 2, 3, 4, 5 y 10. Se omiten páginas de usos no aplicables.
    out.insert_pdf(src, from_page=0, to_page=4)
    out.insert_pdf(src, from_page=9, to_page=9)

    # =========================
    # PAGE 1
    # =========================
    p = out[0]

    # 1. Identificación peticionario: Irrisal Consulting Ltda.
    _mark_box(p, 528, 444, 556, 467, 12)  # Persona jurídica
    _pdf_text(p, 165, 486, 385, 16, PETICIONARIO_EMPRESA["nombre"], size=7.8)
    _pdf_text(p, 225, 508, 325, 22, PETICIONARIO_EMPRESA["domicilio"], size=7.6)
    _pdf_text(p, 122, 537, 150, 15, PETICIONARIO_EMPRESA["rut"], size=7.8)
    _pdf_text(p, 316, 537, 235, 15, PETICIONARIO_EMPRESA["fono"], size=7.8)
    _pdf_text(p, 170, 558, 380, 15, PETICIONARIO_EMPRESA["correo"], size=7.8)

    # 2.1 Naturaleza: subterránea
    _mark_box(p, 379, 653, 422, 674, 12)

    # =========================
    # PAGE 2
    # =========================
    p = out[1]

    # 2.2 Tipo y ejercicio: consuntivo, permanente, continuo
    _mark_box(p, 216, 119, 245, 142, 12)   # Consuntivo
    _mark_box(p, 336, 119, 366, 142, 12)   # Permanente
    _mark_box(p, 529, 119, 556, 142, 12)   # Continuo

    # 2.3 Caudal solicitado
    _pdf_text(p, 50, 340, 142, 18, fmt_num_blank(data.get("caudal_l_s"), 2), size=8.2, align=1)
    _mark_box(p, 343, 316, 373, 338, 12)  # Lts/seg
    _pdf_text(p, 210, 399, 130, 16, fmt_num_blank(data.get("volumen_anual_m3"), 1), size=8.2, align=1)

    # 2.4 Captación - coordenadas con un dígito por casilla.
    _pdf_digits_in_boxes(p, 110.5, 576.5, 19.1, 22.0, data.get("utm_norte"), 7, size=9.0)
    _pdf_digits_in_boxes(p, 244.0, 576.5, 19.2, 22.0, data.get("utm_este"), 6, size=9.0)
    _pdf_text(p, 362, 576.5, 190, 22, data.get("datum_huso"), size=7.4, align=1)
    _pdf_lines(p, 116, 708, 438, 132, _build_location_description(data), size=7.2)

    # =========================
    # PAGE 3
    # =========================
    p = out[2]

    # 3.1 Breve descripción del proyecto: dentro de las líneas, sin invadir instrucciones.
    _pdf_lines(p, 106, 245, 430, 135, _build_project_description(data), size=7.1)

    # 3.2 Derechos asociados: se marca NO en ambos casos.
    _mark_box(p, 287, 432, 316, 461, 12)  # NO constituidos
    _mark_box(p, 528, 432, 557, 461, 12)  # NO en trámite

    # =========================
    # PAGE 4
    # =========================
    p = out[3]

    # 3.3 Uso del agua: uso doméstico de subsistencia y riego.
    uso = clean_text(data.get("uso_agua"), USOS_OPCIONES[0]).lower()
    if "subsistencia" in uso or to_float(data.get("porcentaje_subsistencia"), 0) > 0:
        _mark_box(p, 514, 328, 571, 356, 12)
    _mark_box(p, 514, 392, 571, 420, 12)  # Riego
    # No se escribe nota aquí para no alterar el formato; va en información adicional.

    # =========================
    # PAGE 5
    # =========================
    p = out[4]

    # Redibujar 4.2 y 4.3 con tablas calibradas.
    _draw_custom_42_43(p, data)

    # =========================
    # PAGE 6 = original PAGE 10
    # =========================
    p = out[5]
    _pdf_write_on_ruled_lines(p, 62, 96.5, 488, _build_additional_info(data), size=7.0, max_lines=10, chars_per_line=118, line_gap=18.0)
    _add_signature_to_pdf_page(p, signature_file, data)

    pdf_bytes = out.tobytes(deflate=True, garbage=4)
    out.close()
    src.close()
    return pdf_bytes


# =============================================================================
# UI
# =============================================================================

init_state()

st.title("Memoria Explicativa DGA - Diferentes usos")
st.caption(APP_VERSION)
st.info("Aplicación enfocada en solicitudes de aguas subterráneas para riego, con porcentaje opcional para uso doméstico de subsistencia. No considera derechos constituidos ni solicitudes en trámite asociadas al proyecto.")
st.success("Versión activa: v2.9 - punto 4 limpio, UTM por casillas e información adicional alineada")

with st.sidebar:
    st.header("Autocompletar")
    if st.session_state.get("informe_autocomplete_msg"):
        st.success("Datos autocompletados desde el informe técnico. Revisa antes de exportar.")
        with st.expander("Ver datos detectados del informe"):
            st.json(st.session_state.get("informe_autocomplete_msg"))
        if st.button("Ocultar datos del informe"):
            st.session_state["informe_autocomplete_msg"] = None
            st.rerun()

    informe_upload = st.file_uploader(
        "Subir informe técnico BLA (Word, PDF, JPG o PNG)",
        type=["docx", "pdf", "jpg", "jpeg", "png"],
        key="informe_bla_upload",
    )
    manual_informe_text = st.text_area(
        "Texto pegado del informe BLA, opcional",
        key="manual_informe_text",
        help="Úsalo si las coordenadas UTM o el caudal no se detectan desde el archivo.",
    )

    if st.button("Autocompletar desde informe técnico"):
        try:
            extracted = {}
            debug_text = ""
            if informe_upload is not None:
                name = (informe_upload.name or "").lower()
                if name.endswith(".docx"):
                    extracted.update(parse_informe_bla(informe_upload))
                    # segundo pase directo sobre el texto completo del Word
                    try:
                        lines_tmp = docx_to_lines(informe_upload)
                        debug_text = "\n".join(lines_tmp)
                        extracted.update(parse_direct_bla_text(debug_text))
                    except Exception:
                        pass
                else:
                    debug_text = uploaded_file_to_text_any(informe_upload)
                    extracted.update(parse_direct_bla_text(debug_text))

            if manual_informe_text.strip():
                extracted.update(parse_direct_bla_text(manual_informe_text.strip()))

            if extracted:
                apply_data(extracted)
                st.session_state["informe_autocomplete_msg"] = extracted
                st.rerun()
            else:
                st.warning("No se detectaron datos útiles. Pega el texto del informe o completa manualmente UTM Norte, UTM Este y caudal.")
                if debug_text:
                    with st.expander("Ver texto leído del informe"):
                        st.text(debug_text[:5000])
        except Exception as e:
            st.error(f"No se pudo leer el informe: {e}")

    st.header("Autocompletar predio")
    if st.session_state.get("predio_autocomplete_msg"):
        st.success("Datos del predio autocompletados. Revisa y corrige antes de exportar.")
        with st.expander("Ver datos detectados del predio"):
            st.json(st.session_state.get("predio_autocomplete_msg"))
        if st.button("Ocultar datos detectados"):
            st.session_state["predio_autocomplete_msg"] = None
            st.rerun()
    avaluo_upload = st.file_uploader(
        "Subir avalúo fiscal SII (PDF, JPG o PNG)",
        type=["pdf", "jpg", "jpeg", "png"],
        key="avaluo_upload",
    )
    cbr_upload = st.file_uploader(
        "Subir inscripción Conservador Bienes Raíces (PDF, JPG o PNG)",
        type=["pdf", "jpg", "jpeg", "png"],
        key="cbr_upload",
    )

    manual_avaluo_text = st.text_area(
        "Texto pegado del avalúo fiscal, opcional",
        key="manual_avaluo_text",
        help="Úsalo si el PDF o imagen no se lee bien. Puedes copiar y pegar el texto del certificado.",
    )
    manual_cbr_text = st.text_area(
        "Texto pegado de la inscripción CBR, opcional",
        key="manual_cbr_text",
        help="Úsalo si el OCR no detecta fojas, número, año o conservador.",
    )

    if st.button("Autocompletar datos del predio"):
        detected_predio = {}
        debug_texts = {}
        try:
            if avaluo_upload is not None:
                txt_avaluo = uploaded_file_to_text_any(avaluo_upload)
                debug_texts["texto_avaluo_detectado"] = txt_avaluo[:3000]
                detected_predio.update(parse_avaluo_fiscal(avaluo_upload))

            if cbr_upload is not None:
                txt_cbr = uploaded_file_to_text_any(cbr_upload)
                debug_texts["texto_cbr_detectado"] = txt_cbr[:3000]
                detected_predio.update(parse_conservador(cbr_upload))

            if manual_avaluo_text.strip():
                txt = manual_avaluo_text.strip()
                detected_predio.update({
                    k: v for k, v in {
                        "rol_sii": find_rol_sii(txt),
                        "comuna": find_comuna_from_text(txt),
                        "predio": find_predio_from_avaluo(txt),
                        "hectareas_riego": find_surface_ha(txt),
                    }.items() if v not in [None, ""]
                })

            if manual_cbr_text.strip():
                detected_predio.update(find_cbr_data(manual_cbr_text.strip()))

            if detected_predio:
                apply_data(detected_predio)
                st.session_state["predio_autocomplete_msg"] = detected_predio
                st.rerun()
                if debug_texts:
                    with st.expander("Ver texto leído por OCR / PDF"):
                        st.write(debug_texts)
            else:
                st.warning("No se detectaron datos útiles. Prueba pegando el texto en los campos opcionales o ingrésalos manualmente.")
                if debug_texts:
                    with st.expander("Ver texto leído por OCR / PDF"):
                        st.write(debug_texts)
        except Exception as e:
            st.error(f"No se pudo leer el avalúo o inscripción: {e}")

    st.header("Firma y croquis")
    firma = st.file_uploader("Subir firma en PNG", type=["png"], key="firma_file")
    croquis = st.file_uploader("Subir croquis/mapa opcional", type=["png", "jpg", "jpeg"], key="croquis_file")

    st.header("Guardar/Cargar")
    ficha_upload = st.file_uploader("Cargar ficha JSON", type=["json"])
    if ficha_upload is not None and st.button("Aplicar ficha JSON"):
        try:
            payload = json.loads(ficha_upload.getvalue().decode("utf-8"))
            apply_data(payload)
            st.success("Ficha cargada.")
        except Exception as e:
            st.error(f"No se pudo cargar la ficha: {e}")

# Top controls
with st.expander("Criterios fijos de esta versión", expanded=False):
    st.write("- Naturaleza: aguas subterráneas.")
    st.write("- Tipo de derecho: consuntivo.")
    st.write("- Ejercicio: permanente y continuo.")
    st.write("- No existen otros derechos constituidos asociados al proyecto.")
    st.write("- No existen otros derechos en trámite asociados al proyecto.")
    st.write("- Usos habilitados: riego y uso doméstico de subsistencia.")

# UI tabs
tabs = st.tabs(["1. Peticionario", "2. Derecho solicitado", "3. Proyecto y usos", "4. Riego/Subsistencia", "5. Exportar"])

with tabs[0]:
    st.subheader("1. Identificación del peticionario")
    st.info("Este bloque se completa con los datos de Irrisal Consulting Ltda. El beneficiario del BLA se usa para los antecedentes del proyecto/predio.")
    c1, c2 = st.columns(2)
    with c1:
        st.selectbox("Tipo de persona", ["Persona natural", "Persona jurídica"], key="tipo_persona")
        if st.session_state.tipo_persona == "Persona natural":
            st.selectbox("Sexo", ["F", "M"], key="sexo")
        st.text_input("Nombre o Razón Social", key="nombre")
        st.text_input("RUT", key="rut")
    with c2:
        st.text_area("Domicilio", key="domicilio")
        st.text_input("Fono", key="fono")
        st.text_input("Correo electrónico", key="correo")
        st.text_input("Nombre firmante", key="firmante_nombre")
        st.text_input("RUT firmante", key="firmante_rut")

with tabs[1]:
    st.subheader("2. Derecho de aprovechamiento solicitado")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.text_input("Naturaleza", key="naturaleza", disabled=True)
        st.text_input("Tipo de derecho", key="tipo_derecho", disabled=True)
        st.text_input("Ejercicio 1", key="ejercicio_1", disabled=True)
        st.text_input("Ejercicio 2", key="ejercicio_2", disabled=True)
    with c2:
        st.number_input("Caudal solicitado (L/s)", min_value=0.0, step=0.01, key="caudal_l_s")
        if st.button("Calcular volumen anual desde caudal continuo"):
            st.session_state.volumen_anual_m3 = round(update_volume_from_flow(st.session_state.caudal_l_s), 1)
        st.number_input("Volumen anual (m3/año)", min_value=0.0, step=100.0, key="volumen_anual_m3")
    with c3:
        st.text_input("UTM Norte", key="utm_norte")
        st.text_input("UTM Este", key="utm_este")
        st.text_input("Datum y Huso", key="datum_huso")
    st.text_area("Descripción complementaria de ubicación", key="descripcion_ubicacion", height=110)

with tabs[2]:
    st.subheader("3. Proyecto y usos del agua")
    c1, c2 = st.columns(2)
    with c1:
        st.text_area("Breve descripción del proyecto", key="descripcion_proyecto", height=180)
        st.selectbox("Uso del agua", USOS_OPCIONES, key="uso_agua")
    with c2:
        st.number_input("Porcentaje riego (%)", min_value=0.0, max_value=100.0, step=1.0, key="porcentaje_riego")
        st.number_input("Porcentaje subsistencia (%)", min_value=0.0, max_value=100.0, step=1.0, key="porcentaje_subsistencia")
        total_pct = st.session_state.porcentaje_riego + st.session_state.porcentaje_subsistencia
        if abs(total_pct - 100) > 0.01:
            st.warning(f"Los porcentajes suman {total_pct:.1f}%. Para DGA conviene que sumen 100%.")
        else:
            st.success("Los porcentajes suman 100%.")
    st.info("Los derechos asociados al proyecto se marcarán automáticamente como NO: sin derechos constituidos y sin derechos en trámite.")

with tabs[3]:
    st.subheader("4.2 Uso doméstico de subsistencia y 4.3 riego")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.text_input("Región", key="region")
        st.text_input("Provincia", key="provincia")
        st.text_input("Comuna", key="comuna")
        st.number_input("N° personas beneficiadas", min_value=0, step=1, key="personas_beneficiadas")
    with c2:
        st.text_input("Nombre del predio beneficiado", key="predio")
        st.text_input("N° Rol SII", key="rol_sii")
        st.number_input("N° hectáreas a regar", min_value=0.0, step=0.1, key="hectareas_riego")
    with c3:
        st.text_input("Fojas", key="fojas")
        st.text_input("Número inscripción", key="numero_inscripcion")
        st.text_input("Año inscripción", key="anio_inscripcion")
        st.text_input("Conservador", key="conservador")
    st.text_area("Información adicional", key="informacion_adicional", height=100)

with tabs[4]:
    st.subheader("Exportar")
    st.warning("Antes de exportar, confirme que arriba diga Versión activa: v2.9. Si no aparece, Cloud Run sigue usando una versión antigua.")
    st.info("La exportación usa como fondo la plantilla oficial DGA. El PDF mantiene el formato idéntico; el Word se genera como páginas-imagen para conservar la apariencia exacta.")
    data = get_data()
    st.write("Revisa la vista de datos antes de exportar.")
    with st.expander("Vista de datos"):
        st.json(data)

    ficha_bytes = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button("Guardar ficha JSON", data=ficha_bytes, file_name="ficha_memoria_dga.json", mime="application/json")

    c1, c2 = st.columns(2)
    with c1:
        try:
            docx_bytes = make_docx(data, signature_file=firma, croquis_file=croquis)
            st.download_button("Descargar Word (.docx)", data=docx_bytes, file_name="memoria_explicativa_dga.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Error generando Word: {e}")
    with c2:
        try:
            pdf_bytes = make_pdf(data, signature_file=firma, croquis_file=croquis)
            st.download_button("Descargar PDF (.pdf)", data=pdf_bytes, file_name="memoria_explicativa_dga.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"Error generando PDF: {e}")

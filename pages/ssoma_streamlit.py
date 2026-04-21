"""
SSOMA - Registro de Inducción, Capacitación y Entrenamientos
Versión con ReportLab (sin Playwright) — Compatible con Streamlit Cloud
"""

import streamlit as st
import pandas as pd
import requests
import os
import tempfile
import base64
from io import BytesIO
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
    PageBreak, KeepTogether, Image
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════

LOGO_URL = "https://github.com/CCozd/BI_SIG_CAMPO/blob/main/logo.png?raw=true"
SHAREPOINT_URL = (
    "https://aquanqape-my.sharepoint.com/:x:/g/personal/soporteti_aquanqa_pe/"
    "IQAKudFYiQ-bTIzlTJeXm5HbAVAc35k94KZWvES9s6BvOWc?download=1"
)

LOGO_TMP = os.path.join(tempfile.gettempdir(), "ssoma_logo.png")

# Colores corporativos
C_NAVY = colors.HexColor("#0D2340")
C_BLUE_LITE = colors.HexColor("#E8EDF4")
C_GRAY_D = colors.HexColor("#334155")
C_GRAY_M = colors.HexColor("#64748B")
C_GRAY_L = colors.HexColor("#CBD5E1")
C_GRAY_BG = colors.HexColor("#F8FAFC")
C_WHITE = colors.white

# Página
PW, PH = A4
ML, MR = 1.8*cm, 1.8*cm
MT, MB = 2.5*cm, 2.0*cm
CW = PW - ML - MR

st.set_page_config(page_title="SSOMA - Registro de Inducción", layout="wide")
st.markdown(
    '<style>[data-testid="stSidebarNav"],[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none}</style>',
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def get_logo_b64() -> str:
    """Descarga logo desde GitHub y lo convierte a base64."""
    try:
        if not os.path.exists(LOGO_TMP):
            r = requests.get(LOGO_URL, timeout=15)
            r.raise_for_status()
            with open(LOGO_TMP, "wb") as f:
                f.write(r.content)
        with open(LOGO_TMP, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

def descargar_imagen_temp(url: str) -> str | None:
    """Descarga imagen desde URL y la guarda temporalmente."""
    try:
        r = requests.get(str(url).strip(), timeout=15)
        r.raise_for_status()
        ext = ".png" if "png" in r.headers.get("Content-Type", "") else ".jpg"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        tmp.write(r.content)
        tmp.close()
        return tmp.name
    except Exception:
        return None

def safe(v) -> str:
    """Sanitiza valor."""
    return "" if pd.isna(v) else str(v).strip()

def fmt_fecha(v) -> str:
    """Formatea fecha a dd/mm/yyyy."""
    try:
        return pd.to_datetime(v).strftime("%d/%m/%Y")
    except Exception:
        return str(v)

# ══════════════════════════════════════════════════════════════════════════════
# CARGAR DATOS
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(ttl=300)
def cargar_excel(url: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Carga hojas de Excel desde SharePoint."""
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    
    df_cap = pd.read_excel(BytesIO(r.content), sheet_name="Capacitaciones", dtype=str)
    df_cap.columns = df_cap.columns.str.strip()
    
    df_par = pd.read_excel(BytesIO(r.content), sheet_name="Participantes", dtype=str)
    df_par.columns = df_par.columns.str.strip()
    
    return df_cap, df_par

with st.spinner("Cargando datos..."):
    try:
        df_cap, df_par = cargar_excel(SHAREPOINT_URL)
    except Exception as e:
        st.error(f"Error al cargar Excel: {e}")
        st.stop()

# Procesar fechas
df_cap["_FECHA_DT"] = pd.to_datetime(df_cap["FECHA"], errors="coerce")
df_cap["_FECHA_LABEL"] = df_cap["_FECHA_DT"].dt.strftime("%d/%m/%Y").fillna(df_cap["FECHA"].fillna(""))

# ══════════════════════════════════════════════════════════════════════════════
# FILTROS EN CASCADA
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("### 📋 Selecciona un Registro")

col1, col2, col3 = st.columns(3)

# Fecha
fechas_unicas = df_cap["_FECHA_LABEL"].replace("", pd.NA).dropna().sort_values().unique().tolist()
fecha_sel = col1.selectbox("📅 Fecha", fechas_unicas)
df_by_fecha = df_cap[df_cap["_FECHA_LABEL"] == fecha_sel].copy()

# Tema
temas_unicas = df_by_fecha["TEMA"].dropna().map(str.strip).replace("", pd.NA).dropna().unique().tolist()
tema_sel = col2.selectbox("📚 Tema", temas_unicas)
df_by_tema = df_by_fecha[df_by_fecha["TEMA"].str.strip() == tema_sel].copy()

# Expositor
expositores_unicas = df_by_tema["EXPOSITOR"].dropna().map(str.strip).replace("", pd.NA).dropna().unique().tolist()
expositor_sel = col3.selectbox("👤 Expositor", expositores_unicas)
df_filtrado = df_by_tema[df_by_tema["EXPOSITOR"].str.strip() == expositor_sel].copy()

if df_filtrado.empty:
    st.warning("No hay registros para esa combinación.")
    st.stop()

row = df_filtrado.iloc[0]
id_sel = row["ID_CAPACITACION"]

# ── Extraer datos ─────────────────────────────────────────────────────────────

def g(col: str) -> str:
    try:
        v = row.get(col, "")
        return "" if pd.isna(v) or str(v).strip() in ("nan", "NaT", "None") else str(v).strip().upper()
    except Exception:
        return ""

def g_url(col: str) -> str:
    try:
        v = row.get(col, "")
        return "" if pd.isna(v) or str(v).strip() in ("nan", "NaT", "None") else str(v).strip()
    except Exception:
        return ""

empresa = g("EMPRESA")
sede = g("FUNDO")
ubicacion = g("UBICACIÓN")
semana = g("SEMANA")
duracion = g("DURACIÓN")
procedencia = g("PROCEDENCIA")
tipo = g("TIPO")
tipo_otro = g("TIPO_OTRO(OPCIONAL)")
tema = g("TEMA")
tema_otro = g("TEMA_OTRO(OPCIONAL)")
objetivo = g("OBJETIVO")
fecha_raw = g("FECHA")
try:
    fecha = pd.to_datetime(fecha_raw).strftime("%d/%m/%Y")
except Exception:
    fecha = fecha_raw
observ = g("OBSERVACIONES")
expositor = g("EXPOSITOR")
responsable = g("RESPONSABLE")

firma_exp_url = g_url("FIRMA_EXPOSITOR_URL")
firma_resp_url = g_url("FIRMA_RESPONSABLE_URL")

tema_display = tema_otro if tema == "OTRO" and tema_otro else tema

# Participantes
part_filtrados = df_par[df_par["ID_CAPACITACION"] == id_sel].reset_index(drop=True)

# ══════════════════════════════════════════════════════════════════════════════
# GENERAR PDF CON REPORTLAB
# ══════════════════════════════════════════════════════════════════════════════

def generar_pdf_reportlab(row, participantes_df) -> BytesIO:
    """Genera PDF usando ReportLab."""
    buf = BytesIO()
    tmp_files = []
    
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=ML, rightMargin=MR,
        topMargin=MT, bottomMargin=MB
    )
    elements = []
    
    # Estilos
    S_TITULO = ParagraphStyle(
        "titulo", fontSize=12, fontName="Helvetica-Bold",
        textColor=C_NAVY, alignment=TA_CENTER, spaceAfter=6
    )
    S_SUBTITULO = ParagraphStyle(
        "subtit", fontSize=9, textColor=C_GRAY_M,
        alignment=TA_CENTER, spaceAfter=10
    )
    S_LABEL = ParagraphStyle(
        "lbl", fontSize=9, fontName="Helvetica-Bold",
        textColor=C_NAVY, leading=12
    )
    S_VAL = ParagraphStyle(
        "val", fontSize=9, leading=12, textColor=C_GRAY_D
    )
    S_CELL_HDR = ParagraphStyle(
        "chdr", fontSize=8, fontName="Helvetica-Bold",
        textColor=C_WHITE, alignment=TA_CENTER, leading=10
    )
    S_CELL = ParagraphStyle(
        "ccell", fontSize=8, alignment=TA_CENTER,
        leading=11, textColor=C_GRAY_D
    )
    
    # ── TÍTULO ────────────────────────────────────────────────────────────────
    elements.append(Paragraph(
        "REGISTRO DE INDUCCIÓN, CAPACITACIÓN,<br/>ENTRENAMIENTO Y SIMULACROS DE EMERGENCIA",
        S_TITULO
    ))
    elements.append(Paragraph("Sistema de Gestión de Seguridad y Salud Ocupacional", S_SUBTITULO))
    elements.append(Spacer(1, 0.3*cm))
    
    # ── DATOS HEADER ──────────────────────────────────────────────────────────
    header_data = [
        [Paragraph("Empresa:", S_LABEL), Paragraph(empresa, S_VAL)],
        [Paragraph("Sede:", S_LABEL), Paragraph(sede, S_VAL)],
        [Paragraph("Ubicación:", S_LABEL), Paragraph(ubicacion, S_VAL)],
        [Paragraph("Semana:", S_LABEL), Paragraph(semana, S_VAL)],
        [Paragraph("Duración:", S_LABEL), Paragraph(duracion, S_VAL)],
        [Paragraph("Tema:", S_LABEL), Paragraph(tema_display, S_VAL)],
        [Paragraph("Objetivo:", S_LABEL), Paragraph(objetivo, S_VAL)],
        [Paragraph("Fecha:", S_LABEL), Paragraph(fecha, S_VAL)],
    ]
    
    t_header = Table(header_data, colWidths=[2.5*cm, CW-2.5*cm])
    t_header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, -1), (-1, -1), 0.6, C_GRAY_L),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 0.3*cm))
    
    # ── TABLA DE PARTICIPANTES ────────────────────────────────────────────────
    elements.append(Paragraph("Participantes:", ParagraphStyle(
        "ptit", fontSize=10, fontName="Helvetica-Bold", textColor=C_NAVY
    )))
    elements.append(Spacer(1, 0.15*cm))
    
    # Headers
    table_data = [[
        Paragraph("N°", S_CELL_HDR),
        Paragraph("Apellidos y Nombres", S_CELL_HDR),
        Paragraph("DNI", S_CELL_HDR),
        Paragraph("Puesto", S_CELL_HDR),
        Paragraph("Área", S_CELL_HDR),
        Paragraph("Firma", S_CELL_HDR),
        Paragraph("Fecha", S_CELL_HDR),
    ]]
    
    # Filas
    for idx, par in participantes_df.iterrows():
        nombre = safe(par.get("APELLIDOS_NOMBRES", "")).upper()
        dni = safe(par.get("DNI", "")).upper()
        puesto = safe(par.get("PUESTO", "")).upper()
        area = safe(par.get("AREA", "")).upper()
        firma_url = safe(par.get("FIRMA_URL", "")).lower()
        
        # Descargar firma si existe
        firma_cell = ""
        if firma_url and (firma_url.startswith("http://") or firma_url.startswith("https://")):
            img_path = descargar_imagen_temp(firma_url)
            if img_path:
                tmp_files.append(img_path)
                try:
                    firma_img = Image(img_path, width=1.2*cm, height=0.6*cm)
                    firma_cell = firma_img
                except Exception:
                    firma_cell = ""
        
        table_data.append([
            Paragraph(str(idx + 1), S_CELL),
            Paragraph(nombre, S_CELL),
            Paragraph(dni, S_CELL),
            Paragraph(puesto, S_CELL),
            Paragraph(area, S_CELL),
            firma_cell if firma_cell else "",
            Paragraph(fecha, S_CELL),
        ])
    
    t_part = Table(table_data, colWidths=[0.6*cm, 3.2*cm, 1.5*cm, 2.2*cm, 2.2*cm, 1.5*cm, 1.2*cm])
    t_part.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C_NAVY),
        ("BACKGROUND", (0, 1), (-1, -1), C_BLUE_LITE),
        ("BOX", (0, 0), (-1, -1), 0.8, C_NAVY),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, C_GRAY_L),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    elements.append(t_part)
    elements.append(Spacer(1, 0.3*cm))
    
    # ── FIRMAS ────────────────────────────────────────────────────────────────
    elements.append(PageBreak())
    elements.append(Paragraph("Firmas Autorizadas", ParagraphStyle(
        "ftit", fontSize=10, fontName="Helvetica-Bold", textColor=C_NAVY
    )))
    elements.append(Spacer(1, 0.3*cm))
    
    firmas_data = [
        [
            Paragraph("<b>Expositor:</b>", S_LABEL),
            Paragraph(expositor, S_VAL),
        ],
        [
            Paragraph("<b>Responsable del Registro:</b>", S_LABEL),
            Paragraph(responsable, S_VAL),
        ],
    ]
    
    t_firmas = Table(firmas_data, colWidths=[3*cm, CW-3*cm])
    t_firmas.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    elements.append(t_firmas)
    
    # Build
    doc.build(elements)
    buf.seek(0)
    
    # Limpiar archivos temporales
    for f in tmp_files:
        try:
            os.unlink(f)
        except Exception:
            pass
    
    return buf

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

st.divider()
col_b1, col_b2 = st.columns(2)

with col_b1:
    if st.button("🔍 Generar Vista Previa", use_container_width=True, type="secondary"):
        with st.spinner("Generando PDF..."):
            pdf_buf = generar_pdf_reportlab(row, part_filtrados)
            pdf_bytes = pdf_buf.read()
        st.session_state["pdf_bytes"] = pdf_bytes
        st.session_state["pdf_listo"] = True
        st.session_state["pdf_id"] = id_sel
        st.rerun()

with col_b2:
    if st.session_state.get("pdf_listo"):
        nombre_archivo = f"registro_{id_sel}_{fmt_fecha(fecha)}.pdf"
        st.download_button(
            label="⬇ Descargar PDF",
            data=st.session_state["pdf_bytes"],
            file_name=nombre_archivo,
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )
    else:
        st.button("⬇ Descargar PDF", disabled=True, use_container_width=True)

# Vista previa
if st.session_state.get("pdf_listo"):
    st.success("✅ PDF generado correctamente.")
    st.markdown("---")
    st.markdown("### 🔍 Vista Previa")
    
    pdf_bytes = st.session_state["pdf_bytes"]
    
    # Intentar mostrar con pdf2image
    try:
        from pdf2image import convert_from_bytes
        paginas = convert_from_bytes(pdf_bytes, dpi=150)
        for i, img in enumerate(paginas, 1):
            st.image(img, use_container_width=True, caption=f"Página {i}")
    except Exception:
        # Fallback: iframe con base64
        import base64
        b64 = base64.b64encode(pdf_bytes).decode()
        st.markdown(
            f'<iframe src="data:application/pdf;base64,{b64}" '
            f'width="100%" height="900px" style="border:1px solid #ccc;"></iframe>',
            unsafe_allow_html=True,
        )
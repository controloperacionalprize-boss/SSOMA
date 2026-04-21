import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
import base64
import os
import tempfile

# logo — descarga desde URL y cachea en disco
LOGO_URL  = "https://github.com/CCozd/BI_SIG_CAMPO/blob/main/logo.png?raw=true"
_LOGO_TMP = os.path.join(tempfile.gettempdir(), "ssoma_logo.png")

def _get_logo_b64() -> str:
    try:
        if not os.path.exists(_LOGO_TMP):
            r = requests.get(LOGO_URL, timeout=15)
            r.raise_for_status()
            with open(_LOGO_TMP, "wb") as f:
                f.write(r.content)
        with open(_LOGO_TMP, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

_LOGO_B64 = _get_logo_b64()
LOGO_HTML = (
    f'<img src="data:image/png;base64,{_LOGO_B64}" style="max-width:85px;max-height:60px;object-fit:contain;">'
    if _LOGO_B64 else ""
)

SHAREPOINT_URL = (
    "https://aquanqape-my.sharepoint.com/:x:/g/personal/soporteti_aquanqa_pe/"
    "IQAKudFYiQ-bTIzlTJeXm5HbAVAc35k94KZWvES9s6BvOWc?download=1"
)

st.set_page_config(page_title="SSOMA - Registro de Inducción", layout="wide")
st.markdown(
    '<style>[data-testid="stSidebarNav"],[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none}</style>',
    unsafe_allow_html=True,
)


@st.cache_data(ttl=300)
def cargar_excel(url: str) -> bytes:
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.content


with st.spinner("Cargando datos..."):
    try:
        contenido = cargar_excel(SHAREPOINT_URL)
        df_cap = pd.read_excel(BytesIO(contenido), sheet_name="Capacitaciones", dtype=str)
        df_cap.columns = df_cap.columns.str.strip()
        df_par = pd.read_excel(BytesIO(contenido), sheet_name="Participantes", dtype=str)
        df_par.columns = df_par.columns.str.strip()
    except Exception as e:
        st.error(f"Error al cargar el archivo: {e}")
        st.stop()

df_cap["_FECHA_DT"]    = pd.to_datetime(df_cap["FECHA"], errors="coerce")
df_cap["_FECHA_LABEL"] = df_cap["_FECHA_DT"].dt.strftime("%d/%m/%Y").fillna(df_cap["FECHA"].fillna(""))

fechas_disponibles = (
    df_cap["_FECHA_LABEL"].replace("", pd.NA).dropna()
    .sort_values().unique().tolist()
)
col1, col2, col3 = st.columns(3)
with col1:
    fecha_sel = st.selectbox("Fecha", fechas_disponibles)

df_por_fecha = df_cap[df_cap["_FECHA_LABEL"] == fecha_sel].copy()

temas_disponibles = (
    df_por_fecha["TEMA"].dropna().map(str.strip)
    .replace("", pd.NA).dropna().unique().tolist()
)
with col2:
    tema_sel = st.selectbox("Tema", temas_disponibles)

df_por_tema = df_por_fecha[df_por_fecha["TEMA"].str.strip() == tema_sel].copy()

capacitadores_disponibles = (
    df_por_tema["EXPOSITOR"].dropna().map(str.strip)
    .replace("", pd.NA).dropna().unique().tolist()
)
with col3:
    capacitador_sel = st.selectbox("Capacitador", capacitadores_disponibles)

df_filtrado = df_por_tema[df_por_tema["EXPOSITOR"].str.strip() == capacitador_sel].copy()

if df_filtrado.empty:
    st.warning("No hay registros para esa combinación de filtros.")
    st.stop()

row    = df_filtrado.iloc[0]
id_sel = row["ID_CAPACITACION"]


def g(col: str) -> str:
    try:
        v = row.get(col, "")
        return "" if pd.isna(v) or str(v).strip() in ("nan", "NaT", "None") else str(v).strip().upper()
    except Exception:
        return ""


empresa     = g("EMPRESA").upper()
sede        = g("FUNDO")
ubicacion   = g("UBICACIÓN")
semana      = g("SEMANA")
duracion    = g("DURACIÓN")
procedencia = g("PROCEDENCIA").upper()
tipo        = g("TIPO").upper()
tipo_otro   = g("TIPO_OTRO(OPCIONAL)")
tema        = g("TEMA")
tema_otro   = g("TEMA_OTRO(OPCIONAL)")
objetivo    = g("OBJETIVO")
_fecha_raw  = g("FECHA")
try:
    fecha = pd.to_datetime(_fecha_raw).strftime("%d/%m/%Y")
except Exception:
    fecha = _fecha_raw
observ      = g("OBSERVACIONES")
expositor   = g("EXPOSITOR")
responsable = g("RESPONSABLE")


def g_url(col: str) -> str:
    try:
        v = row.get(col, "")
        return "" if pd.isna(v) or str(v).strip() in ("nan", "NaT", "None") else str(v).strip()
    except Exception:
        return ""


firma_exp_url  = g_url("FIRMA_EXPOSITOR_URL")
firma_resp_url = g_url("FIRMA_RESPONSABLE_URL")

tema_display = tema_otro if tema.upper() == "OTRO" and tema_otro else tema

part_filtrados = df_par[df_par["ID_CAPACITACION"] == id_sel].reset_index(drop=True)


def gp(df_row, col: str) -> str:
    try:
        v = df_row.get(col, "")
        return "" if pd.isna(v) or str(v).strip() in ("nan", "NaT", "None") else str(v).strip().upper()
    except Exception:
        return ""


ROWS_PER_PAGE = 10
total_part = len(part_filtrados)
chunks = [
    part_filtrados.iloc[i:i + ROWS_PER_PAGE]
    for i in range(0, max(total_part, 1), ROWS_PER_PAGE)
]


def build_rows(chunk, offset: int) -> str:
    rows = ""
    for i in range(ROWS_PER_PAGE):
        n  = offset + i + 1
        bg = "#D6E0F0" if i % 2 == 0 else "#ffffff"
        if i < len(chunk):
            p         = chunk.iloc[i]
            nombre    = gp(p, "APELLIDOS_NOMBRES")
            dni       = gp(p, "DNI")
            puesto    = gp(p, "PUESTO")
            area      = gp(p, "AREA")
            firma_url = gp(p, "FIRMA_URL").lower()
            p_fecha   = fecha
        else:
            nombre = dni = puesto = area = firma_url = p_fecha = ""
        firma_img = (
            f'<img src="{firma_url}" style="max-height:36px;max-width:100%;display:block;margin:auto;">'
            if firma_url else ""
        )
        rows += f"""
      <tr style="background:{bg}; height:40px;">
        <td style="text-align:center;color:#555;width:28px;">{n}</td>
        <td colspan="2">{nombre}</td>
        <td style="text-align:center;width:75px;">{dni}</td>
        <td colspan="2">{puesto}</td>
        <td colspan="2">{area}</td>
        <td style="text-align:center;padding:2px;">{firma_img}</td>
        <td style="text-align:center;width:70px;">{p_fecha}</td>
      </tr>"""
    return rows


def cb(marcado: bool = False) -> str:
    if marcado:
        return '<span style="display:inline-block;width:12px;height:12px;border:1px solid #0D2340;background:#0D2340;margin-right:3px;vertical-align:middle;text-align:center;color:white;font-size:10px;line-height:12px;">✓</span>'
    return '<span style="display:inline-block;width:12px;height:12px;border:1px solid #444;margin-right:3px;vertical-align:middle;"></span>'


emp1_bg = "#fff9c4" if "AQU ANQA II" not in empresa and "AQU ANQA" in empresa else "white"
emp2_bg = "#fff9c4" if "AQU ANQA II" in empresa else "white"


def header_table() -> str:
    return f"""
  <tr>
    <td rowspan="2" style="width:95px;text-align:center;padding:4px;background:white;border:1px solid #777;">{LOGO_HTML}</td>
    <td colspan="7" class="gh" style="font-size:14px;padding:10px;">
      REGISTRO DE INDUCCIÓN, CAPACITACIÓN,<br>ENTRENAMIENTO Y SIMULACROS DE EMERGENCIA
    </td>
    <td colspan="2" style="background:#D6E0F0;font-size:11px;padding:5px;vertical-align:top;">
      <b>Código:</b> PAQ-DO-FT-001<br><b>Versión:</b> 03 &nbsp;|&nbsp; Set - 2025<br><b>Rev:</b> 02
    </td>
  </tr>
  <tr>
    <td class="lb" style="width:80px;">Elaborado por:</td>
    <td colspan="2">Desarrollo Organizacional</td>
    <td class="lb">Revisado por:</td>
    <td>SIG &amp; Certificaciones</td>
    <td class="lb">Aprobado por:</td>
    <td colspan="3" style="font-size:11px;">Jefatura de Gestión de Talento Humano</td>
  </tr>
  <tr>
    <td class="lh">MARCA</td><td class="lh" colspan="2">RAZÓN SOCIAL</td>
    <td class="lh">RUC</td><td class="lh" colspan="2">DOMICILIO</td>
    <td class="lh" colspan="2">ACTIVIDAD ECONÓMICA</td>
    <td class="lh">N° TRAB.</td><td class="lh"></td>
  </tr>
  <tr style="background:{emp1_bg}">
    <td style="text-align:center;font-size:10px;line-height:1.2;">AQUA<br>NOA</td>
    <td colspan="2">AQU ANQA S.A.C.</td><td>20608345770</td>
    <td colspan="2" style="font-size:10px;">Car. Panamericana Km. 625 Sec. Las Dos Rayas - La Arenita - La Libertad - Ascope - Razuri</td>
    <td colspan="2" style="font-size:10px;">0122 - Cultivo de frutas tropicales y subtropicales</td>
    <td></td><td></td>
  </tr>
  <tr style="background:{emp2_bg}">
    <td style="text-align:center;font-size:10px;line-height:1.2;">AQUA<br>NOA II</td>
    <td colspan="2">AQU ANQA II S.A.C.</td><td>20610068767</td>
    <td colspan="2" style="font-size:10px;">Car. Panamericana Km. 639 Sec C.P. Men - La Arenita - La Libertad - Ascope - Paijan</td>
    <td colspan="2" style="font-size:10px;">0122 - Cultivo de frutas tropicales y subtropicales</td>
    <td></td><td></td>
  </tr>
  <tr>
    <td class="lb">SEDE</td><td colspan="2">{sede}</td>
    <td class="lb">UBICACIÓN</td><td colspan="2">{ubicacion}</td>
    <td class="lb" colspan="2">SEMANA (S) &nbsp; {semana}</td>
    <td class="lb" colspan="2">DURACIÓN &nbsp; {duracion}</td>
  </tr>
  <tr>
    <td class="lb">PROCEDENCIA</td>
    <td colspan="2">{cb(procedencia=="INTERNA")} Interna &nbsp;&nbsp; {cb(procedencia=="EXTERNA")} Externa</td>
    <td class="lb" colspan="2">TIPO:</td>
    <td>{cb("INDUCC" in tipo)} Inducción</td>
    <td>{cb("CAPACIT" in tipo)} Capacitación</td>
    <td colspan="2">{cb("ENTREN" in tipo)} Entrenamiento</td>
    <td>{cb("SIMUL" in tipo)} Simulacro</td>
  </tr>
  <tr>
    <td colspan="4" style="border-top:none;"></td><td></td>
    <td>{cb("CHARLA" in tipo)} Charla 5 min</td>
    <td>{cb("CURSO" in tipo)} Curso</td>
    <td colspan="2">{cb("TALLER" in tipo)} Taller</td>
    <td>{cb("OTRO" in tipo)} Otro: {tipo_otro}</td>
  </tr>
  <tr><td class="lb">TEMA(S):</td><td colspan="9" style="height:24px;">{tema_display}</td></tr>
  <tr><td class="lb">OBJETIVO(S):</td><td colspan="9" style="height:24px;">{objetivo}</td></tr>
  <tr>
    <td class="lh" style="text-align:center;">N°</td>
    <td class="lh" colspan="2">APELLIDOS Y NOMBRES</td>
    <td class="lh" style="text-align:center;">DNI</td>
    <td class="lh" colspan="2">PUESTO</td>
    <td class="lh" colspan="2">ÁREA</td>
    <td class="lh">FIRMA</td>
    <td class="lh" style="text-align:center;">FECHA</td>
  </tr>"""


def footer_table() -> str:
    return f"""
  <tr><td class="lb">OBSERVACIONES:</td><td colspan="9" style="height:28px;">{observ}</td></tr>
  <tr><td colspan="10" class="nb" style="height:8px;"></td></tr>
  <tr>
    <td class="lh" colspan="4">EXPOSITOR 1:</td>
    <td class="lh" colspan="3">EXPOSITOR 2:</td>
    <td class="lh" colspan="3">EXPOSITOR 3:</td>
  </tr>
  <tr>
    <td class="lb">CARGO:</td><td></td><td class="lb">EMPRESA:</td><td>{empresa}</td>
    <td class="lb">CARGO:</td><td class="lb">EMPRESA:</td><td></td>
    <td class="lb">CARGO:</td><td class="lb">EMPRESA:</td><td></td>
  </tr>
  <tr>
    <td class="lb">NOMBRE:</td><td colspan="3">{expositor}</td>
    <td class="lb">NOMBRE:</td><td colspan="2"></td>
    <td class="lb">NOMBRE:</td><td colspan="2"></td>
  </tr>
  <tr style="height:50px;">
    <td class="lb">FIRMA:</td>
    <td colspan="3" style="text-align:center;padding:2px;">{"<img src='" + firma_exp_url + "' style='max-height:44px;max-width:100%;display:block;margin:auto;'>" if firma_exp_url else ""}</td>
    <td class="lb">FIRMA:</td><td colspan="2"></td>
    <td class="lb">FIRMA:</td><td colspan="2"></td>
  </tr>
  <tr><td colspan="10" class="nb" style="height:8px;"></td></tr>
  <tr><td class="lh" colspan="10">RESPONSABLE DEL REGISTRO</td></tr>
  <tr>
    <td class="lh" colspan="4">APELLIDOS Y NOMBRES</td>
    <td class="lh" colspan="3">CARGO</td>
    <td class="lh" colspan="2">FIRMA</td>
    <td class="lh">FECHA</td>
  </tr>
  <tr style="height:50px;">
    <td colspan="4">{responsable}</td><td colspan="3"></td>
    <td colspan="2" style="text-align:center;padding:2px;">{"<img src='" + firma_resp_url + "' style='max-height:44px;max-width:100%;display:block;margin:auto;'>" if firma_resp_url else ""}</td>
    <td>{fecha}</td>
  </tr>"""


gen_fecha = datetime.now().strftime("%d/%m/%Y %H:%M")


def build_pages_html() -> str:
    pages = ""
    for idx, chunk in enumerate(chunks):
        is_first = idx == 0
        is_last  = idx == len(chunks) - 1
        page_break = "" if is_last else 'style="page-break-after:always;"'
        if is_first:
            pages += f"""
<div class="page-wrap" {page_break}>
  <table class="ssf">
    {header_table()}
    {build_rows(chunk, idx * ROWS_PER_PAGE)}
    {footer_table() if is_last else ""}
  </table>
</div>"""
        else:
            pages += f"""
<div class="page-wrap" {page_break}>
  <table class="ssf">
    {build_rows(chunk, idx * ROWS_PER_PAGE)}
    {footer_table() if is_last else ""}
  </table>
</div>"""
    return pages


CSS_PREVIEW = """
  body { margin:0; padding:12px; font-family:Arial,sans-serif; background:#F8FAFC; }
  .page-wrap { background:white; box-shadow:0 2px 8px rgba(0,0,0,0.12); padding:8px; margin-bottom:28px; }
  .ssf { width:100%; border-collapse:collapse; font-size:12px; color:#111; }
  .ssf td { border:1px solid #777; padding:3px 6px; vertical-align:middle; }
  .gh { background:#0D2340; color:white; font-weight:bold; text-align:center; padding:6px; }
  .lh { background:#0D2340; color:white; font-weight:bold; text-align:center; padding:4px; }
  .lb { background:#D6E0F0; font-weight:bold; }
  .nb { border:none !important; }
"""


def build_html_preview() -> str:
    pages_html = build_pages_html()
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>{CSS_PREVIEW}</style>
</head><body>
{pages_html}
</body></html>
"""


# ── Generar PDF con ReportLab puro (sin dependencias del sistema) ─────────────
def generar_pdf() -> bytes:
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle,
        Paragraph, Spacer, Image as RLImage
    )
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    C_NAVY    = colors.HexColor("#0D2340")
    C_BLUE_LT = colors.HexColor("#D6E0F0")
    C_GRAY    = colors.HexColor("#777777")
    C_WHITE   = colors.white
    C_YELLOW  = colors.HexColor("#fff9c4")

    PW, PH = A4
    ML = MR = 0.5 * cm
    MT = 0.8 * cm
    MB = 1.2 * cm
    CW = PW - ML - MR

    buf       = BytesIO()
    tmp_files = []

    def make_footer(canv, doc):
        canv.saveState()
        canv.setFont("Helvetica", 6)
        canv.setFillColor(colors.HexColor("#64748B"))
        canv.drawString(ML, 0.5 * cm,
                        "Documento de uso interno — Prohibida su reproducción sin autorización")
        canv.drawCentredString(PW / 2, 0.5 * cm, f"— {doc.page} —")
        canv.drawRightString(PW - MR, 0.5 * cm, f"Generado: {gen_fecha}")
        canv.restoreState()

    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=ML, rightMargin=MR,
                            topMargin=MT, bottomMargin=MB)

    def PS(name, **kw):
        return ParagraphStyle(name, **kw)

    S_HDR  = PS("hdr",  fontSize=6,   fontName="Helvetica-Bold",
                textColor=C_WHITE,  alignment=TA_CENTER, leading=8)
    S_LBL  = PS("lbl",  fontSize=6,   fontName="Helvetica-Bold",
                textColor=C_NAVY,   alignment=TA_LEFT,   leading=8)
    S_VAL  = PS("val",  fontSize=6,   fontName="Helvetica",
                textColor=colors.black, alignment=TA_LEFT,   leading=8)
    S_VALC = PS("valc", fontSize=6,   fontName="Helvetica",
                textColor=colors.black, alignment=TA_CENTER, leading=8)
    S_TITL = PS("titl", fontSize=7.5, fontName="Helvetica-Bold",
                textColor=C_WHITE,  alignment=TA_CENTER, leading=10)

    el = []

    # ── LOGO ──────────────────────────────────────────────────────────────────
    logo_cell = ""
    if _LOGO_B64:
        try:
            logo_data = base64.b64decode(_LOGO_B64)
            ltmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            ltmp.write(logo_data)
            ltmp.close()
            tmp_files.append(ltmp.name)
            logo_cell = RLImage(ltmp.name, width=2.2 * cm, height=1.4 * cm)
        except Exception:
            pass

    # ── CABECERA ──────────────────────────────────────────────────────────────
    cod_s = PS("cod", fontSize=5.5, fontName="Helvetica",
               textColor=colors.black, leading=8)
    sub_s = PS("sub", fontSize=5,   fontName="Helvetica",
               textColor=colors.black, alignment=TA_CENTER, leading=7)

    head_data = [
        [logo_cell,
         Paragraph("REGISTRO DE INDUCCIÓN, CAPACITACIÓN,<br/>ENTRENAMIENTO Y SIMULACROS DE EMERGENCIA", S_TITL),
         Paragraph("<b>Código:</b> PAQ-DO-FT-001<br/><b>Versión:</b> 03 | Set-2025<br/><b>Rev:</b> 02", cod_s)],
        ["",
         Paragraph("Elaborado por: Desarrollo Organizacional  |  Revisado por: SIG & Certificaciones  |  "
                   "Aprobado por: Jefatura de Gestión de Talento Humano", sub_s),
         ""],
    ]
    t_head = Table(head_data, colWidths=[2.4 * cm, CW - 5.2 * cm, 2.8 * cm],
                   rowHeights=[1.4 * cm, 0.4 * cm])
    t_head.setStyle(TableStyle([
        ("SPAN",         (0, 0), (0, 1)),
        ("SPAN",         (1, 1), (2, 1)),
        ("BACKGROUND",   (1, 0), (1, 0), C_NAVY),
        ("BACKGROUND",   (2, 0), (2, 0), C_BLUE_LT),
        ("BACKGROUND",   (1, 1), (2, 1), C_BLUE_LT),
        ("BOX",          (0, 0), (-1, -1), 0.5, C_GRAY),
        ("INNERGRID",    (0, 0), (-1, -1), 0.3, C_GRAY),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",        (0, 0), (0, -1), "CENTER"),
    ]))
    el.append(t_head)

    # ── EMPRESAS ──────────────────────────────────────────────────────────────
    ec1 = C_YELLOW if "AQU ANQA II" not in empresa and "AQU ANQA" in empresa else C_WHITE
    ec2 = C_YELLOW if "AQU ANQA II" in empresa else C_WHITE

    emp_data = [
        [Paragraph("MARCA", S_HDR), Paragraph("RAZÓN SOCIAL", S_HDR),
         Paragraph("RUC", S_HDR), Paragraph("DOMICILIO", S_HDR),
         Paragraph("ACTIVIDAD ECONÓMICA", S_HDR), Paragraph("N° TRAB.", S_HDR), ""],
        [Paragraph("AQUA\nNOA", S_VALC), Paragraph("AQU ANQA S.A.C.", S_VAL),
         Paragraph("20608345770", S_VALC),
         Paragraph("Car. Panamericana Km. 625 - La Arenita - Razuri", S_VAL),
         Paragraph("0122 - Cultivo de frutas tropicales y subtropicales", S_VAL), "", ""],
        [Paragraph("AQUA\nNOA II", S_VALC), Paragraph("AQU ANQA II S.A.C.", S_VAL),
         Paragraph("20610068767", S_VALC),
         Paragraph("Car. Panamericana Km. 639 - La Arenita - Paijan", S_VAL),
         Paragraph("0122 - Cultivo de frutas tropicales y subtropicales", S_VAL), "", ""],
    ]
    t_emp = Table(emp_data, colWidths=[1.2*cm, 3.2*cm, 1.8*cm, 4.5*cm, 4.5*cm, 1.2*cm, 1.2*cm])
    t_emp.setStyle(TableStyle([
        ("BACKGROUND",      (0, 0), (-1, 0), C_NAVY),
        ("BACKGROUND",      (0, 1), (-1, 1), ec1),
        ("BACKGROUND",      (0, 2), (-1, 2), ec2),
        ("BOX",             (0, 0), (-1, -1), 0.5, C_GRAY),
        ("INNERGRID",       (0, 0), (-1, -1), 0.3, C_GRAY),
        ("VALIGN",          (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",      (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",   (0, 0), (-1, -1), 2),
    ]))
    el.append(t_emp)

    # ── DATOS GENERALES ───────────────────────────────────────────────────────
    def cb_rl(m): return "■" if m else "□"

    dat_data = [
        [Paragraph("SEDE", S_LBL), Paragraph(sede, S_VAL),
         Paragraph("UBICACIÓN", S_LBL), Paragraph(ubicacion, S_VAL),
         Paragraph(f"SEMANA: {semana}", S_LBL), Paragraph(f"DURACIÓN: {duracion}", S_VAL)],
        [Paragraph("PROCEDENCIA", S_LBL),
         Paragraph(f"{cb_rl(procedencia=='INTERNA')} Interna  {cb_rl(procedencia=='EXTERNA')} Externa", S_VAL),
         Paragraph("TIPO:", S_LBL),
         Paragraph(f"{cb_rl('INDUCC' in tipo)} Inducción  {cb_rl('CAPACIT' in tipo)} Capacitación  "
                   f"{cb_rl('ENTREN' in tipo)} Entrenamiento  {cb_rl('SIMUL' in tipo)} Simulacro", S_VAL),
         Paragraph(f"{cb_rl('CHARLA' in tipo)} Charla 5min  {cb_rl('CURSO' in tipo)} Curso  "
                   f"{cb_rl('TALLER' in tipo)} Taller  {cb_rl('OTRO' in tipo)} Otro: {tipo_otro}", S_VAL), ""],
        [Paragraph("TEMA(S):", S_LBL), Paragraph(tema_display, S_VAL), "", "", "", ""],
        [Paragraph("OBJETIVO(S):", S_LBL), Paragraph(objetivo, S_VAL),  "", "", "", ""],
    ]
    t_dat = Table(dat_data, colWidths=[1.8*cm, 3.2*cm, 1.5*cm, 4.5*cm, 4.5*cm, 2.1*cm])
    t_dat.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (0, -1), C_BLUE_LT),
        ("BACKGROUND",    (2, 0), (2, -1), C_BLUE_LT),
        ("SPAN",          (1, 2), (5, 2)),
        ("SPAN",          (1, 3), (5, 3)),
        ("BOX",           (0, 0), (-1, -1), 0.5, C_GRAY),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, C_GRAY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    el.append(t_dat)

    # ── TABLA PARTICIPANTES ───────────────────────────────────────────────────
    def fetch_img(url, w=1.8*cm, h=0.7*cm):
        if not url or not url.startswith("http"): return ""
        try:
            r = requests.get(url, timeout=8)
            r.raise_for_status()
            ext  = ".png" if "png" in r.headers.get("Content-Type", "") else ".jpg"
            ftmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
            ftmp.write(r.content); ftmp.close()
            tmp_files.append(ftmp.name)
            return RLImage(ftmp.name, width=w, height=h)
        except Exception:
            return ""

    part_header = [
        Paragraph("N°", S_HDR), Paragraph("APELLIDOS Y NOMBRES", S_HDR),
        Paragraph("DNI", S_HDR), Paragraph("PUESTO", S_HDR),
        Paragraph("ÁREA", S_HDR), Paragraph("FIRMA", S_HDR), Paragraph("FECHA", S_HDR),
    ]
    part_rows   = [part_header]
    row_heights = [0.5 * cm]

    for pg, chunk in enumerate(chunks):
        for i in range(ROWS_PER_PAGE):
            n = pg * ROWS_PER_PAGE + i + 1
            if i < len(chunk):
                p = chunk.iloc[i]
                part_rows.append([
                    Paragraph(str(n), S_VALC),
                    Paragraph(gp(p, "APELLIDOS_NOMBRES"), S_VAL),
                    Paragraph(gp(p, "DNI"), S_VALC),
                    Paragraph(gp(p, "PUESTO"), S_VAL),
                    Paragraph(gp(p, "AREA"), S_VAL),
                    fetch_img(gp(p, "FIRMA_URL").lower()),
                    Paragraph(fecha, S_VALC),
                ])
            else:
                part_rows.append([Paragraph(str(n), S_VALC), "", "", "", "", "", ""])
            row_heights.append(0.9 * cm)

    cw_part = [0.6*cm, 4.5*cm, 1.6*cm, 3.0*cm, 3.0*cm, 2.2*cm, 1.7*cm]
    sty_part = [
        ("BACKGROUND",   (0, 0), (-1, 0), C_NAVY),
        ("BOX",          (0, 0), (-1, -1), 0.5, C_GRAY),
        ("INNERGRID",    (0, 0), (-1, -1), 0.3, C_GRAY),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",        (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING",   (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 1),
    ]
    for i in range(1, len(part_rows)):
        bg = C_BLUE_LT if i % 2 == 1 else C_WHITE
        sty_part.append(("BACKGROUND", (0, i), (-1, i), bg))

    t_part = Table(part_rows, colWidths=cw_part, rowHeights=row_heights)
    t_part.setStyle(TableStyle(sty_part))
    el.append(t_part)

    # ── OBSERVACIONES ─────────────────────────────────────────────────────────
    t_obs = Table(
        [[Paragraph("OBSERVACIONES:", S_LBL), Paragraph(observ, S_VAL)]],
        colWidths=[2.2 * cm, CW - 2.2 * cm]
    )
    t_obs.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (0, 0), C_BLUE_LT),
        ("BOX",           (0, 0), (-1, -1), 0.5, C_GRAY),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, C_GRAY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    el.append(t_obs)

    # ── FIRMAS EXPOSITORES ────────────────────────────────────────────────────
    exp_firma  = fetch_img(firma_exp_url,  w=2.5*cm, h=1.0*cm)
    resp_firma = fetch_img(firma_resp_url, w=2.5*cm, h=1.0*cm)

    t_exp = Table([
        [Paragraph("EXPOSITOR 1:", S_HDR), "", "",
         Paragraph("EXPOSITOR 2:", S_HDR), "",
         Paragraph("EXPOSITOR 3:", S_HDR), ""],
        [Paragraph("NOMBRE:", S_LBL), Paragraph(expositor, S_VAL), "",
         Paragraph("NOMBRE:", S_LBL), "",
         Paragraph("NOMBRE:", S_LBL), ""],
        [Paragraph("FIRMA:",  S_LBL), exp_firma or "", "",
         Paragraph("FIRMA:",  S_LBL), "",
         Paragraph("FIRMA:",  S_LBL), ""],
    ], colWidths=[1.2*cm, 3.8*cm, 0.5*cm, 1.2*cm, 3.8*cm, 1.2*cm, 3.8*cm],
       rowHeights=[0.4*cm, 0.4*cm, 1.2*cm])
    t_exp.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (2, 0), C_NAVY),
        ("BACKGROUND",    (3, 0), (4, 0), C_NAVY),
        ("BACKGROUND",    (5, 0), (6, 0), C_NAVY),
        ("BACKGROUND",    (0, 1), (0, -1), C_BLUE_LT),
        ("BACKGROUND",    (3, 1), (3, -1), C_BLUE_LT),
        ("BACKGROUND",    (5, 1), (5, -1), C_BLUE_LT),
        ("BOX",           (0, 0), (-1, -1), 0.5, C_GRAY),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, C_GRAY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (1, 2), (1, 2), "CENTER"),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    el.append(t_exp)

    # ── RESPONSABLE DEL REGISTRO ──────────────────────────────────────────────
    t_resp = Table([
        [Paragraph("RESPONSABLE DEL REGISTRO", S_HDR), "", "", ""],
        [Paragraph("APELLIDOS Y NOMBRES", S_HDR), Paragraph("CARGO", S_HDR),
         Paragraph("FIRMA", S_HDR), Paragraph("FECHA", S_HDR)],
        [Paragraph(responsable, S_VAL), "", resp_firma or "", Paragraph(fecha, S_VALC)],
    ], colWidths=[5*cm, 4*cm, 3.5*cm, 2*cm],
       rowHeights=[0.4*cm, 0.4*cm, 1.2*cm])
    t_resp.setStyle(TableStyle([
        ("SPAN",          (0, 0), (-1, 0)),
        ("BACKGROUND",    (0, 0), (-1, 1), C_NAVY),
        ("BOX",           (0, 0), (-1, -1), 0.5, C_GRAY),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, C_GRAY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (2, 2), (2, 2), "CENTER"),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    el.append(t_resp)

    doc.build(el, onFirstPage=make_footer, onLaterPages=make_footer)
    buf.seek(0)

    for f in tmp_files:
        try:
            os.unlink(f)
        except Exception:
            pass

    return buf.getvalue()


# ── Invalidar PDF si cambian los filtros ──────────────────────────────────────
if st.session_state.get("pdf_id") and st.session_state.get("pdf_id") != id_sel:
    st.session_state.pop("pdf_bytes", None)
    st.session_state.pop("pdf_listo", None)
    st.session_state.pop("pdf_id", None)

# ── UI: botones ───────────────────────────────────────────────────────────────
st.divider()
col_b1, col_b2 = st.columns(2)

with col_b1:
    if st.button("🔍 Generar Vista Previa", use_container_width=True, type="secondary"):
        with st.spinner("Generando informe..."):
            pdf_bytes = generar_pdf()
        st.session_state["pdf_bytes"] = pdf_bytes
        st.session_state["pdf_listo"] = True
        st.session_state["pdf_id"]    = id_sel
        st.rerun()

with col_b2:
    if st.session_state.get("pdf_listo"):
        st.download_button(
            label="⬇ Descargar PDF",
            data=st.session_state["pdf_bytes"],
            file_name=f"registro_{st.session_state.get('pdf_id', id_sel)}.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )
    else:
        st.button("⬇ Descargar PDF", disabled=True, use_container_width=True)

if st.session_state.get("pdf_listo"):
    st.success("✅ Registro generado. Revisa la vista previa antes de descargar.")
    st.markdown("---")
    st.markdown("### 🔍 Vista Previa")
    altura = 1600 + (len(chunks) - 1) * 1400
    components.html(build_html_preview(), height=altura, scrolling=True)
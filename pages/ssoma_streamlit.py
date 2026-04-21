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

# fecha sin ISO
df_cap["_FECHA_DT"]    = pd.to_datetime(df_cap["FECHA"], errors="coerce")
df_cap["_FECHA_LABEL"] = df_cap["_FECHA_DT"].dt.strftime("%d/%m/%Y").fillna(df_cap["FECHA"].fillna(""))

# ── Filtros ───────────────────────────────────────────────────────────────────
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

# ── HTML para la nueva pestaña (PDF idéntico a vista previa) ──────────────────
# Mismo HTML, mismo CSS, solo cambia @page para A4 y font-size a 7px
CSS_PDF_TAB = """
  body  { margin:0; padding:4px; font-family:Arial,sans-serif; background:#F8FAFC;
          -webkit-print-color-adjust:exact; print-color-adjust:exact; }
  .page-wrap { background:white; box-shadow:0 2px 8px rgba(0,0,0,0.12);
               padding:2px; margin-bottom:16px; }
  .ssf { width:100%; border-collapse:collapse; font-size:7px;
         color:#111; table-layout:fixed; }
  .ssf td { border:1px solid #777; padding:1px 2px; vertical-align:middle;
            overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
            line-height:1.3; }
  .gh { background:#0D2340 !important; color:white !important; font-weight:bold;
        text-align:center; font-size:7.5px; padding:3px;
        white-space:normal; word-break:break-word; }
  .lh { background:#0D2340 !important; color:white !important; font-weight:bold;
        text-align:center; padding:2px; white-space:normal; word-break:break-word; }
  .lb { background:#D6E0F0 !important; font-weight:bold; white-space:normal; }
  .nb { border:none !important; }
  img { max-width:100% !important; }
"""

COLGROUP = (
    '<colgroup>'
    '<col style="width:28px"><col style="width:14%"><col style="width:8%">'
    '<col style="width:7%"><col style="width:14%"><col style="width:12%">'
    '<col style="width:12%"><col style="width:11%"><col style="width:10%">'
    '<col style="width:70px">'
    '</colgroup>'
)


def build_html_pdf_tab() -> str:
    """HTML completo con html2pdf.js para descargar PDF identico a vista previa."""
    pages_html = build_pages_html()
    pages_html = pages_html.replace(
        '<table class="ssf">',
        f'<table class="ssf">{COLGROUP}'
    )
    nombre_archivo = f"registro_{id_sel}.pdf"
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Registro {id_sel}</title>
<style>{CSS_PDF_TAB}</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
</head><body>
<div id="contenido">
{pages_html}
</div>
<script>
  window.onload = function() {{
    var opt = {{
      margin:       [8, 5, 14, 5],
      filename:     '{nombre_archivo}',
      image:        {{ type: 'jpeg', quality: 0.98 }},
      html2canvas:  {{ scale: 2, useCORS: true, letterRendering: true }},
      jsPDF:        {{ unit: 'mm', format: 'a4', orientation: 'portrait' }},
      pagebreak:    {{ mode: ['avoid-all', 'css', 'legacy'] }}
    }};
    html2pdf().set(opt).from(document.getElementById('contenido')).save();
  }};
</script>
</body></html>
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


# ── Invalidar si cambian filtros ──────────────────────────────────────────────
if st.session_state.get("pdf_id") and st.session_state.get("pdf_id") != id_sel:
    st.session_state.pop("pdf_listo", None)
    st.session_state.pop("pdf_id", None)

# ── UI ────────────────────────────────────────────────────────────────────────
st.divider()
col_b1, col_b2 = st.columns(2)

with col_b1:
    if st.button("🔍 Generar Vista Previa", use_container_width=True, type="secondary"):
        st.session_state["pdf_listo"] = True
        st.session_state["pdf_id"]    = id_sel
        st.rerun()

with col_b2:
    if st.session_state.get("pdf_listo"):
        html_pdf = build_html_pdf_tab()
        b64_pdf  = base64.b64encode(html_pdf.encode("utf-8")).decode()
        # Abre nueva pestaña con html2pdf.js que descarga el PDF automáticamente
        components.html(f"""
        <script>
          function descargarPDF() {{
            var win = window.open('', '_blank');
            var html = atob('{b64_pdf}');
            win.document.open();
            win.document.write(html);
            win.document.close();
          }}
        </script>
        <button onclick="descargarPDF()"
          style="width:100%;padding:10px;background:#0D2340;color:white;
                 border:none;border-radius:8px;font-size:15px;font-weight:700;
                 cursor:pointer;letter-spacing:.03em;">
          ⬇ Descargar PDF
        </button>
        """, height=55)
    else:
        st.button("⬇ Descargar PDF", disabled=True, use_container_width=True)

if st.session_state.get("pdf_listo"):
    st.success("✅ Registro generado. Presiona **⬇ Descargar PDF** — el archivo se descargará automáticamente.")
    st.markdown("---")
    st.markdown("### 🔍 Vista Previa")
    altura = 1600 + (len(chunks) - 1) * 1400
    components.html(build_html_preview(), height=altura, scrolling=True)
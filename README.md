# SSOMA — Registro de Inducción, Capacitación, Entrenamiento y Simulacros

Aplicación Streamlit que genera y descarga el formulario oficial **PAQ-DO-FT-001** en PDF, leyendo datos en tiempo real desde un Excel público en SharePoint.

---

## Estructura de archivos

```
SSOMA/
├── ssoma_streamlit.py       # App principal
├── logo_prize_nuevo.png     # Logo embebido en el formulario
└── README.md
```

---

## Fuente de datos

El Excel en SharePoint tiene **dos hojas** que usa la app:

| Hoja | Descripción |
|---|---|
| `Capacitaciones` | Una fila por sesión: fecha, sede, tipo, tema, expositor, responsable, etc. |
| `Participantes` | Una fila por asistente, vinculada por `ID_CAPACITACION` |

```
SharePoint (Excel público)
├── Hoja: Capacitaciones
│   └── ID_CAPACITACION, EMPRESA, FUNDO, UBICACIÓN, SEMANA, DURACIÓN,
│       PROCEDENCIA, TIPO, TEMA, OBJETIVO, FECHA, OBSERVACIONES,
│       EXPOSITOR, RESPONSABLE, FIRMA_EXPOSITOR_URL, ...
└── Hoja: Participantes
    └── ID_PARTICIPANTE, ID_CAPACITACION, DNI, APELLIDOS_NOMBRES,
        PUESTO, AREA, FIRMA_URL, ...
```

---

## Flujo de la aplicación

```
1. CARGA DE DATOS
   └── GET SharePoint URL → descarga el .xlsx en memoria
       ├── Lee hoja "Capacitaciones"  → df_cap
       └── Lee hoja "Participantes"   → df_par
       (cache de 5 minutos con @st.cache_data)

2. FILTROS EN CASCADA (3 selectboxes en una fila)
   ├── Filtro 1: FECHA        → lista única DD/MM/YYYY
   ├── Filtro 2: TEMA         → filtrado por la fecha seleccionada
   └── Filtro 3: CAPACITADOR  → filtrado por fecha + tema

3. RESOLUCIÓN DEL REGISTRO
   └── Se toma la primera fila de df_cap que coincide con los 3 filtros
       → obtiene: sede, ubicación, semana, duración, procedencia,
                  tipo, tema, objetivo, fecha, observaciones,
                  expositor, responsable

4. PARTICIPANTES
   └── df_par filtrado por ID_CAPACITACION del registro encontrado
       → campos: N°, APELLIDOS_NOMBRES, DNI, PUESTO, AREA, FIRMA_URL, FECHA
       → divididos en páginas de 20 filas

5. GENERACIÓN DEL HTML
   ├── header_table()   → logo, encabezado, empresas, sede, checkboxes, tema
   ├── build_rows()     → filas de participantes (con imagen de firma)
   └── footer_table()   → observaciones, expositores, responsable
   Cada grupo de 20 participantes = 1 página con page-break-after

6. VISTA PREVIA
   └── st.components.v1.html → iframe con el formulario renderizado

7. DESCARGA PDF
   └── Playwright (Chromium headless) renderiza el HTML
       → page.pdf(A4, landscape, print_background=True)
       → asyncio.new_event_loop() con WindowsProactorEventLoopPolicy
       → st.download_button → descarga directa sin diálogo de impresión
```

---

## Checkboxes automáticos

| Campo Excel | Marca en el formulario |
|---|---|
| `PROCEDENCIA = INTERNA` | ✓ Interna |
| `PROCEDENCIA = EXTERNA` | ✓ Externa |
| `TIPO = INDUCCIÓN` | ✓ Inducción |
| `TIPO = CAPACITACIÓN` | ✓ Capacitación |
| `TIPO = ENTRENAMIENTO` | ✓ Entrenamiento |
| `TIPO = SIMULACRO` | ✓ Simulacro |
| `TIPO = CHARLA 5 MIN` | ✓ Charla 5 min |
| `TIPO = CURSO` | ✓ Curso |
| `TIPO = TALLER` | ✓ Taller |
| `TIPO = OTRO` | ✓ Otro: (valor de TIPO_OTRO) |

---

## Instalación

```bash
pip install streamlit pandas requests openpyxl playwright
python -m playwright install chromium
```

## Ejecución

```bash
streamlit run ssoma_streamlit.py
```

---

## Colores del formulario

| Uso | Color |
|---|---|
| Encabezados oscuros | `#0D2340` |
| Etiquetas de campo | `#D6E0F0` |
| Filas alternas participantes | `#D6E0F0` / `#ffffff` |

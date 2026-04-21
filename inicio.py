import streamlit as st

st.set_page_config(page_title="SSOMA - Inicio", layout="wide")

st.markdown("""
<style>
    [data-testid="stSidebarNav"] { display: none; }
    [data-testid="stSidebar"] { display: none; }
    [data-testid="collapsedControl"] { display: none; }
    .block-container { padding-top: 3rem; max-width: 700px; }
    .stButton > button {
        width: 100%; height: 90px;
        font-size: 1.05rem; font-weight: 700;
        border-radius: 10px; letter-spacing: .03em;
    }
    .card-desc {
        font-size: 0.82rem; color: #64748B;
        margin-top: -10px; margin-bottom: 20px;
        padding-left: 4px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("## 📋 Sistema SSOMA")
st.caption("Selecciona la vista que deseas usar.")
st.divider()

col1, col2 = st.columns(2)

with col1:
    if st.button("📝 Registro de Inducción", type="secondary", use_container_width=True):
        st.switch_page("pages/ssoma_streamlit.py")
    st.markdown(
        '<p class="card-desc">Visualiza y descarga el registro de inducción, '
        'capacitación, entrenamiento y simulacros.</p>',
        unsafe_allow_html=True,
    )

with col2:
    if st.button("📊 Informe de Capacitaciones", type="primary", use_container_width=True):
        st.switch_page("pages/v2ssoma_streamlit.py")
    st.markdown(
        '<p class="card-desc">Genera el informe SSOMA con objetivo, alcance, '
        'desarrollo, recomendaciones y anexos fotográficos.</p>',
        unsafe_allow_html=True,
    )

import streamlit as st
from aprendices import generar_por_aprendiz
from fichas import generar_por_ficha

st.set_page_config(page_title="Generador de Cancelaciones", layout="centered")

if "modulo_seleccionado" not in st.session_state:
    st.session_state.modulo_seleccionado = None

if st.session_state.modulo_seleccionado is None:
    st.title("📄 Generador de Reportes de Cancelación")
    st.markdown("### Selecciona el módulo que deseas usar:")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔍 EVIDENCIAS X APRENDICES"):
            st.session_state.modulo_seleccionado = "aprendices"
            st.rerun()
    with col2:
        if st.button("📂 EVIDENCIA X FICHAS"):
            st.session_state.modulo_seleccionado = "fichas"
            st.rerun()

if st.session_state.modulo_seleccionado == "aprendices":
    generar_por_aprendiz()

elif st.session_state.modulo_seleccionado == "fichas":
    generar_por_ficha()

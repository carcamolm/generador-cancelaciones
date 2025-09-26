import streamlit as st
import os
from generador import generar_documentos

st.set_page_config(page_title="Generador de Cancelaciones", layout="centered")
st.title("📄 Generador de Reportes de Cancelación")

st.markdown("Sube el archivo Excel y selecciona la carpeta con las evidencias.")

excel = st.file_uploader("📥 Archivo Excel (.xlsx)", type=["xlsx"])
carpeta_imagenes = st.text_input("📁 Ruta local de la carpeta con imágenes")

if st.button("Generar documentos"):
    if excel and carpeta_imagenes:
        with open("estudiantes.xlsx", "wb") as f:
            f.write(excel.getbuffer())
        if not os.path.exists(carpeta_imagenes):
            st.error("❌ La carpeta de imágenes no existe.")
        else:
            ruta_pdf = generar_documentos("estudiantes.xlsx", carpeta_imagenes)
            st.success("✅ Documentos generados con éxito.")
            with open(ruta_pdf, "rb") as file:
                st.download_button("📄 Descargar reporte general PDF", file.read(), file_name="reporte_general.pdf")
    else:
        st.warning("⚠️ Por favor sube el Excel y especifica la carpeta de imágenes.")

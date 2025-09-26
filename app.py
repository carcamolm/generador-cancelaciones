import streamlit as st
import pandas as pd
from fpdf import FPDF
from PIL import Image
import io
import zipfile
import tempfile
import os

st.set_page_config(page_title="Generador de Cancelaciones", layout="centered")
st.title("📄 Generador de Reportes de Cancelación")

# 📘 Botón para descargar instructivo
st.markdown("¿Primera vez usando la herramienta? Descarga el instructivo institucional aquí:")
try:
    with open("instructivo_cancelaciones.pdf", "rb") as pdf_file:
        st.download_button(
            label="📘 Descargar instructivo en PDF",
            data=pdf_file.read(),
            file_name="Instructivo_Generador_Cancelaciones.pdf",
            mime="application/pdf"
        )
except FileNotFoundError:
    st.warning("⚠️ El instructivo no se encuentra en el repositorio.")

# 📁 Carga de Excel
st.subheader("📁 Paso 1: Cargar archivo Excel")
excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"])

# 🖼️ Carga de imágenes
st.subheader("🖼️ Paso 2: Cargar evidencias en imagen")
uploaded_images = st.file_uploader("Imágenes (.png, .jpg)", type=["png", "jpg"], accept_multiple_files=True)

# 🔄 Generación de documentos
if st.button("Generar documentos"):
    if not excel_file or not uploaded_images:
        st.error("❗ Debes subir el Excel y al menos una imagen.")
    else:
        df = pd.read_excel(excel_file)
        columnas_requeridas = {"Nombre", "Ficha", "Evidencia"}
        if not columnas_requeridas.issubset(df.columns):
            st.error("❌ El archivo Excel debe tener las columnas: Nombre, Ficha, Evidencia.")
        else:
            imagen_dict = {img.name: img for img in uploaded_images}
            zip_buffer = io.BytesIO()
            resumen_general = ""
            total_aprendices = 0
            temp_files = []

            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                agrupado = df.groupby("Ficha")

                for ficha, grupo in agrupado:
                    ficha_str = str(ficha)
                    resumen_texto = f"Resumen de Ficha: {ficha_str}\nTotal de aprendices: {len(grupo)}\n\n"

                    for _, row in grupo.iterrows():
                        nombre = row["Nombre"].replace(" ", "_")
                        evidencia_nombre = row["Evidencia"]
                        evidencia_file = imagen_dict.get(evidencia_nombre)

                        if not evidencia_file:
                            st.warning(f"❗ No se encontró la imagen: {evidencia_nombre}")
                            continue

                        # Crear PDF individual
                        pdf = FPDF()
                        pdf.add_page()
                        pdf.set_font("Arial", size=12)
                        pdf.cell(0, 10, f"FICHA: {ficha_str}", ln=True)
                        pdf.cell(0, 10, f"APRENDIZ: {row['Nombre']}", ln=True)
                        pdf.cell(0, 10, "EVIDENCIA CORREO", ln=True)

                        image = Image.open(evidencia_file)
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                            image.save(tmp_img.name, format="PNG")
                            pdf.image(tmp_img.name, x=10, y=40, w=100)
                            temp_files.append(tmp_img.name)

                        pdf_bytes = pdf.output(dest='S').encode('latin1')
                        ruta_pdf = f"documentos_pdf/{ficha_str}/{ficha_str}_{nombre}.pdf"
                        zip_file.writestr(ruta_pdf, pdf_bytes)

                        resumen_texto += f"- {row['Nombre']}\n"
                        total_aprendices += 1

                    resumen_path = f"documentos_pdf/{ficha_str}/resumen_ficha_{ficha_str}.txt"
                    zip_file.writestr(resumen_path, resumen_texto)
                    resumen_general += resumen_texto + "\n"

                # Crear reporte general con logo
                pdf_general = FPDF()
                pdf_general.add_page()
                logo_path = "logo_sena.png"
                if os.path.exists(logo_path):
                    pdf_general.image(logo_path, x=10, y=8, w=30)
                pdf_general.set_xy(0, 35)
                pdf_general.set_font("Arial", "B", 14)
                pdf_general.cell(0, 10, "Reporte General de Cancelaciones", ln=True, align="C")
                pdf_general.set_font("Arial", "", 11)
                pdf_general.multi_cell(0, 10, resumen_general)
                pdf_general.cell(0, 10, f"Total de fichas: {len(agrupado)}", ln=True)
                pdf_general.cell(0, 10, f"Total de aprendices: {total_aprendices}", ln=True)

                pdf_general_bytes = pdf_general.output(dest='S').encode('latin1')
                zip_file.writestr("documentos_pdf/reporte_general.pdf", pdf_general_bytes)

            # 🧹 Limpiar archivos temporales
            for path in temp_files:
                try:
                    os.remove(path)
                except Exception:
                    pass

            zip_buffer.seek(0)
            st.success("✅ Documentos generados correctamente.")
            st.download_button(
                label="📥 Descargar carpeta ZIP con todos los documentos",
                data=zip_buffer,
                file_name="cancelaciones.zip",
                mime="application/zip"
            )

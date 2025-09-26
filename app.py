import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile

st.set_page_config(page_title="Generador de Cancelaciones", layout="centered")
st.title("📄 Generador de Reportes de Cancelación")

# Subida de Excel
excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"])
# Subida de imágenes
uploaded_images = st.file_uploader("Sube las evidencias (imágenes)", type=["png", "jpg"], accept_multiple_files=True)

# Botón para generar
if st.button("Generar documentos"):
    if not excel_file or not uploaded_images:
        st.error("⚠️ Debes subir el Excel y al menos una imagen.")
    else:
        df = pd.read_excel(excel_file)
        ficha = str(df.iloc[0]["Ficha"])  # Asegúrate de tener esta columna
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            resumen_texto = ""

            for index, row in df.iterrows():
                nombre = row["Nombre"].replace(" ", "_")
                documento = str(row["Documento"])
                motivo = row["Motivo"]
                archivo_nombre = f"{ficha}_{nombre}"

                # Crear PDF individual
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, txt=f"Cancelación de {nombre}", ln=True)
                pdf.cell(200, 10, txt=f"Documento: {documento}", ln=True)
                pdf.cell(200, 10, txt=f"Motivo: {motivo}", ln=True)
                pdf_output = io.BytesIO()
                pdf.output(pdf_output)
                pdf_output.seek(0)

                # Guardar en carpeta virtual
                ruta_pdf = f"documentos_pdf/{ficha}/{archivo_nombre}.pdf"
                zip_file.writestr(ruta_pdf, pdf_output.read())

                # Agregar al resumen
                resumen_texto += f"{nombre} - {documento} - {motivo}\n"

            # Crear resumen general
            resumen_pdf = FPDF()
            resumen_pdf.add_page()
            resumen_pdf.set_font("Arial", size=12)
            resumen_pdf.multi_cell(0, 10, resumen_texto)
            resumen_output = io.BytesIO()
            resumen_pdf.output(resumen_output)
            resumen_output.seek(0)

            zip_file.writestr(f"documentos_pdf/{ficha}/resumen_ficha_{ficha}.pdf", resumen_output.read())

        zip_buffer.seek(0)

        st.success("✅ Documentos generados correctamente.")
        st.download_button(
            label="📥 Descargar carpeta ZIP con todos los documentos",
            data=zip_buffer,
            file_name=f"cancelaciones_{ficha}.zip",
            mime="application/zip"
        )

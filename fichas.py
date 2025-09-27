import streamlit as st
import pandas as pd
from fpdf import FPDF
from PIL import Image
from datetime import datetime
import os
import io
import zipfile
import tempfile

def generar_por_ficha():
    st.title("📂 Generación por Ficha con Consolidado Institucional")

    st.subheader("📁 Paso 1: Cargar archivo Excel")
    excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"])

    st.subheader("🖼️ Paso 2: Cargar evidencias por ficha")
    uploaded_images = st.file_uploader("Imágenes (.png, .jpg)", type=["png", "jpg"], accept_multiple_files=True)

    if st.button("Generar documentos"):
        if not excel_file or not uploaded_images:
            st.error("❗ Debes subir el Excel y al menos una imagen.")
            return

        df = pd.read_excel(excel_file)
        columnas_requeridas = {"Nombre", "Ficha", "Evidencia"}
        if not columnas_requeridas.issubset(df.columns):
            st.error("❌ El archivo Excel debe tener las columnas: Nombre, Ficha, Evidencia.")
            return

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

                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", size=12)
                    pdf.cell(0, 10, f"FICHA: {ficha_str}", ln=True)
                    pdf.cell(0, 10, f"APRENDIZ: {row['Nombre']}", ln=True)
                    pdf.cell(0, 10, "EVIDENCIA CORREO", ln=True)

                    image = Image.open(evidencia_file)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        image.save(tmp_img.name, format="PNG")


                        y_actual = pdf.get_y() + 5  # 5 mm de espacio después del texto
                        pdf.image(tmp_img.name, x=10, y=y_actual, w=100)

                        temp_files.append(tmp_img.name)

                    nombre_archivo = f"EVIDENCIAS_DESERCIÓN_{nombre}_.pdf"
                    ruta_pdf = f"documentos_pdf/Ficha_{ficha_str}/{nombre_archivo}"
                    pdf_bytes = pdf.output(dest='S').encode('latin1')
                    zip_file.writestr(ruta_pdf, pdf_bytes)

                    resumen_texto += f"- {row['Nombre']}\n"
                    total_aprendices += 1

                resumen_path = f"documentos_pdf/Ficha_{ficha_str}/resumen_ficha_{ficha_str}.txt"
                zip_file.writestr(resumen_path, resumen_texto)
                resumen_general += resumen_texto + "\n"

            # PDF consolidado institucional
            fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
            pdf_general = FPDF()
            pdf_general.add_page()

            logo_path = "logo_sena.png"
            if os.path.exists(logo_path):
                pdf_general.image(logo_path, x=10, y=8, w=30)
            pdf_general.set_xy(0, 35)

            pdf_general.set_font("Arial", "B", 14)
            pdf_general.cell(0, 10, "Reporte de Aprendices enviados a cancelación por Formulario", ln=True, align="C")
            pdf_general.set_font("Arial", "", 11)
            pdf_general.cell(0, 10, f"Fecha de creación: {fecha_actual}", ln=True, align="C")
            pdf_general.ln(10)

            for ficha, grupo in agrupado:
                pdf_general.set_font("Arial", "B", 12)
                pdf_general.cell(0, 10, f"Ficha: {ficha}", ln=True)
                pdf_general.set_font("Arial", "", 11)
                for nombre in grupo["Nombre"]:
                    pdf_general.cell(0, 8, f"- {nombre}", ln=True)
                pdf_general.ln(5)

            pdf_general.set_font("Arial", "B", 12)
            pdf_general.ln(10)
            pdf_general.cell(0, 10, "Resumen Final", ln=True)
            pdf_general.set_font("Arial", "", 11)
            pdf_general.cell(0, 8, f"Total de fichas procesadas: {len(agrupado)}", ln=True)
            pdf_general.cell(0, 8, f"Total de aprendices incluidos: {total_aprendices}", ln=True)

            pdf_general_bytes = pdf_general.output(dest='S').encode('latin1')
            zip_file.writestr("documentos_pdf/reporte_consolidado.pdf", pdf_general_bytes)

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
            file_name="cancelaciones_por_ficha.zip",
            mime="application/zip"
        )

    if st.button("⬅️ Volver al menú principal"):
        st.session_state.modulo_seleccionado = None
        st.rerun()

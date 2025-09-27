import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import os
import io
import zipfile

def generar_por_ficha():
    st.title("📂 Generación por Ficha con Consolidado Institucional")

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

    st.subheader("📁 Paso 1: Cargar archivo Excel")
    excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"])

    st.subheader("🖼️ Paso 2: Cargar evidencias en imagen")
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
        resumen_fichas = {}
        total_general = 0

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            agrupado = df.groupby("Ficha")

            for ficha, grupo in agrupado:
                ficha_str = str(ficha)
                carpeta_ficha = f"documentos_pdf/{ficha_str}/"
                resumen_texto = f"📌 FICHA: {ficha_str}\nAPRENDICES:\n"

                for _, row in grupo.iterrows():
                    nombre = row["Nombre"]
                    nombre_archivo = f"EVIDENCIAS_DESERCIÓN_{nombre.replace(' ', '_')}_.pdf"
                    ruta_pdf = os.path.join(carpeta_ficha, nombre_archivo)
                    evidencia_nombre = row["Evidencia"]
                    evidencia_file = imagen_dict.get(evidencia_nombre)

                    pdf_buffer = io.BytesIO()
                    c = canvas.Canvas(pdf_buffer, pagesize=A4)
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(2*cm, 27*cm, f"FICHA: {ficha_str}")
                    c.drawString(2*cm, 26.2*cm, f"APRENDIZ: {nombre}")
                    c.setFont("Helvetica", 12)
                    c.drawString(2*cm, 25.2*cm, "EVIDENCIA CORREO:")

                    if evidencia_file:
                        try:
                            img = ImageReader(evidencia_file)
                            c.drawImage(img, 2*cm, 12*cm, width=16*cm, preserveAspectRatio=True, mask='auto')
                        except:
                            c.drawString(2*cm, 24.5*cm, "[Error al insertar imagen]")
                    else:
                        c.drawString(2*cm, 24.5*cm, "[Imagen no encontrada]")

                    c.showPage()
                    c.save()
                    pdf_bytes = pdf_buffer.getvalue()
                    zip_file.writestr(ruta_pdf, pdf_bytes)

                    resumen_texto += f"- {nombre}\n"
                    total_general += 1

                resumen_path = os.path.join(carpeta_ficha, f"resumen_ficha_{ficha_str}.txt")
                zip_file.writestr(resumen_path, resumen_texto)
                resumen_fichas[ficha_str] = len(grupo)

            # PDF consolidado institucional
            consolidado_buffer = io.BytesIO()
            c = canvas.Canvas(consolidado_buffer, pagesize=A4)
            c.setFont("Helvetica-Bold", 16)
            c.drawCentredString(10.5*cm, 23.5*cm, "REPORTE CONSOLIDADO DE CANCELACIONES")

            logo_path = "logo_sena.png"
            if os.path.exists(logo_path):
                try:
                    logo = ImageReader(logo_path)
                    c.drawImage(logo, 2*cm, 25.5*cm, width=4*cm, preserveAspectRatio=True, mask='auto')
                except:
                    pass

            c.setFont("Helvetica", 12)
            y = 20.5*cm
            for ficha, grupo in agrupado:
                c.drawString(2*cm, y, f"📌 FICHA: {ficha}")
                y -= 0.6*cm
                c.drawString(2*cm, y, "APRENDICES:")
                y -= 0.5*cm
                for _, row in grupo.iterrows():
                    c.drawString(2.5*cm, y, f"- {row['Nombre']}")
                    y -= 0.4*cm
                    if y < 4*cm:
                        c.showPage()
                        y = 27*cm
                y -= 0.8*cm

            c.showPage()
            c.setFont("Helvetica-Bold", 14)
            c.drawString(2*cm, 27*cm, "📊 RESUMEN FINAL POR FICHA")
            c.setFont("Helvetica", 12)
            y = 26*cm
            for ficha, cantidad in resumen_fichas.items():
                c.drawString(2*cm, y, f"Ficha {ficha}: {cantidad} aprendices")
                y -= 0.5*cm
            c.drawString(2*cm, y - 0.5*cm, f"🧮 TOTAL GENERAL: {total_general} aprendices")
            c.save()

            consolidado_bytes = consolidado_buffer.getvalue()
            zip_file.writestr("documentos_pdf/reporte_consolidado.pdf", consolidado_bytes)

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

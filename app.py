import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from PIL import Image
from datetime import datetime
import io
import zipfile
import tempfile
import os


# Menú principal
st.set_page_config(page_title="Generador de Cancelaciones", layout="centered")
st.title("📄 Generador de Reportes de Cancelación")

opcion = st.radio("Selecciona el módulo que deseas usar:", ["🏷️ EVIDENCIAS X APRENDICES", "📁 EVIDENCIA X FICHAS"])

if opcion == "🏷️ EVIDENCIAS X APRENDICES":
    st.markdown("### 🔍 Generación por aprendiz")
    # Aquí colocas TODO el código que ya tienes funcionando con FPDF
    # Puedes encapsularlo en una función como generar_por_aprendiz()

elif opcion == "📁 EVIDENCIA X FICHAS":
    st.markdown("### 📂 Generación por ficha con consolidado institucional")
    # Aquí colocas el nuevo flujo con reportlab y portada
    # Puedes encapsularlo en una función como generar_por_ficha()





st.set_page_config(page_title="Generador de Cancelaciones", layout="centered")
st.title("📄 Generador de Reportes de Cancelación")

if "carga_id" not in st.session_state:
    st.session_state.carga_id = 0

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

# 🔄 Botón para nueva carga
st.markdown("---")
col1, col2 = st.columns([3, 1])
with col2:
    if st.button("🔄 Nueva carga"):
        st.session_state.carga_id += 1
        st.rerun()

# 📁 Carga de Excel
st.subheader("📁 Paso 1: Cargar archivo Excel")
excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"], key=f"excel_{st.session_state.carga_id}")

# 🖼️ Carga de imágenes
st.subheader("🖼️ Paso 2: Cargar evidencias en imagen")
uploaded_images = st.file_uploader("Imágenes (.png, .jpg)", type=["png", "jpg"], accept_multiple_files=True, key=f"images_{st.session_state.carga_id}")

# 🔄 Generación de documentos
if st.button("Generar documentos", key=f"generar_{st.session_state.carga_id}"):
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
                    carpeta_ficha = f"documentos_pdf/{ficha_str}/"

                    for _, row in grupo.iterrows():
                        nombre = row["Nombre"].replace(" ", "_")
                        evidencia_nombre = row["Evidencia"]
                        evidencia_file = imagen_dict.get(evidencia_nombre)

                        if not evidencia_file:
                            st.warning(f"❗ No se encontró la imagen: {evidencia_nombre}")
                            continue

                        ruta_pdf = f"{carpeta_ficha}{ficha_str}_{nombre}.pdf"
                        pdf_bytes = io.BytesIO()
                        c = canvas.Canvas(pdf_bytes, pagesize=A4)
                        c.setFont("Helvetica-Bold", 14)
                        c.drawString(2*cm, 27*cm, f"FICHA: {ficha_str}")
                        c.drawString(2*cm, 26.2*cm, f"APRENDIZ: {row['Nombre']}")
                        c.setFont("Helvetica", 12)
                        c.drawString(2*cm, 25.2*cm, "EVIDENCIA CORREO:")

                        try:
                            img = Image.open(evidencia_file)
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                                img.save(tmp_img.name, format="PNG")
                                c.drawImage(tmp_img.name, 2*cm, 12*cm, width=16*cm, preserveAspectRatio=True, mask='auto')
                                temp_files.append(tmp_img.name)
                        except:
                            c.drawString(2*cm, 24.5*cm, "[Error al insertar imagen]")

                        c.showPage()
                        c.save()
                        zip_file.writestr(ruta_pdf, pdf_bytes.getvalue())

                        resumen_texto += f"- {row['Nombre']}\n"
                        total_aprendices += 1

                    resumen_path = f"{carpeta_ficha}resumen_ficha_{ficha_str}.txt"
                    zip_file.writestr(resumen_path, resumen_texto)
                    resumen_general += resumen_texto + "\n"

                # PDF consolidado institucional con portada
                pdf_general = io.BytesIO()
                c = canvas.Canvas(pdf_general, pagesize=A4)
                logo_path = "logo_sena.png"

                # Portada
                c.setFont("Helvetica-Bold", 20)
                c.drawCentredString(10.5*cm, 25.5*cm, "SENA")
                if os.path.exists(logo_path):
                    try:
                        logo = ImageReader(logo_path)
                        c.drawImage(logo, 8.5*cm, 19.5*cm, width=4*cm, preserveAspectRatio=True, mask='auto')
                    except:
                        pass
                c.setFont("Helvetica-Bold", 16)
                c.drawCentredString(10.5*cm, 17.5*cm, "REPORTE CONSOLIDADO DE CANCELACIONES")
                fecha_actual = datetime.now().strftime("%d de %B de %Y")
                c.setFont("Helvetica", 12)
                c.drawCentredString(10.5*cm, 16*cm, f"Fecha de generación: {fecha_actual}")
                c.showPage()

                # Contenido por ficha
                y = 27*cm
                resumen_fichas = {}
                for ficha, grupo in agrupado:
                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(2*cm, y, f"📌 FICHA: {ficha}")
                    y -= 0.6*cm
                    c.setFont("Helvetica", 11)
                    c.drawString(2*cm, y, "APRENDICES:")
                    y -= 0.5*cm
                    for _, row in grupo.iterrows():
                        c.drawString(2.5*cm, y, f"- {row['Nombre']}")
                        y -= 0.4*cm
                        if y < 4*cm:
                            c.showPage()
                            y = 27*cm
                    resumen_fichas[ficha] = len(grupo)
                    y -= 0.8*cm

                # Resumen final
                c.showPage()
                c.setFont("Helvetica-Bold", 14)
                c.drawString(2*cm, 27*cm, "📊 RESUMEN FINAL POR FICHA")
                c.setFont("Helvetica", 12)
                y = 26*cm
                for ficha, cantidad in resumen_fichas.items():
                    c.drawString(2*cm, y, f"Ficha {ficha}: {cantidad} aprendices")
                    y -= 0.5*cm
                c.drawString(2*cm, y - 0.5*cm, f"🧮 TOTAL GENERAL: {total_aprendices} aprendices")
                c.save()
                zip_file.writestr("documentos_pdf/reporte_general.pdf", pdf_general.getvalue())

            for path in temp_files:
                try:
                    os.remove(path)
                except:
                    pass

            zip_buffer.seek(0)
            st.success("✅ Documentos generados correctamente.")
            st.download_button(
                label="📥 Descargar carpeta ZIP con todos los documentos",
                data=zip_buffer,
                file_name="cancelaciones.zip",
                mime="application/zip"
            )

# Información de estado actual
st.markdown("---")
st.caption(f"💡 **Estado actual:** Carga #{st.session_state.carga_id + 1}")
if excel_file:
    st.caption(f"📄 Archivo Excel: {excel_file.name}")
if uploaded_images:
    st.caption(f"🖼️ Imágenes: {len(uploaded_images)} archivo(s)")

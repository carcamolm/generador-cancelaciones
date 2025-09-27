import streamlit as st
import pandas as pd
from fpdf import FPDF
from PIL import Image
from datetime import datetime
import io
import zipfile
import tempfile
import os

st.set_page_config(page_title="Generador de Cancelaciones", layout="centered")

# Inicializar navegación
if "modulo_seleccionado" not in st.session_state:
    st.session_state.modulo_seleccionado = None

# Función para validar imagen
def es_imagen_valida(imagen_file):
    """Valida si una imagen es válida"""
    try:
        imagen = Image.open(imagen_file)
        imagen.verify()
        return True
    except:
        return False

# Función para limpiar texto para PDF
def limpiar_texto_pdf(texto):
    """Limpia el texto para evitar problemas de codificación en PDF"""
    # Reemplazar caracteres problemáticos
    replacements = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
        'ñ': 'n', 'Ñ': 'N', 'ü': 'u', 'Ü': 'U'
    }
    
    for old, new in replacements.items():
        texto = texto.replace(old, new)
    
    # Mantener solo caracteres ASCII seguros
    return ''.join(char for char in texto if ord(char) < 128)

# Función para crear PDF individual
def crear_pdf_individual(nombre, ficha, evidencia_file):
    """Crea un PDF individual para cada estudiante"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    
    # Limpiar textos
    nombre_limpio = limpiar_texto_pdf(nombre)
    ficha_limpio = limpiar_texto_pdf(str(ficha))
    
    # Información del estudiante
    pdf.cell(0, 10, f"FICHA: {ficha_limpio}", ln=True)
    pdf.cell(0, 10, f"APRENDIZ: {nombre_limpio}", ln=True)
    pdf.ln(5)
    pdf.cell(0, 10, "EVIDENCIA CORREO", ln=True)
    pdf.ln(10)
    
    # Agregar imagen si es válida
    if evidencia_file and es_imagen_valida(evidencia_file):
        try:
            imagen = Image.open(evidencia_file)
            # Redimensionar imagen si es muy grande
            max_width, max_height = 800, 600
            if imagen.width > max_width or imagen.height > max_height:
                imagen.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
            
            # Guardar imagen temporalmente
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                imagen.save(tmp_img.name, format="PNG")
                pdf.image(tmp_img.name, x=10, y=50, w=120)
                return pdf, tmp_img.name
        except Exception as e:
            pdf.cell(0, 10, f"[Error al cargar imagen: {str(e)}]", ln=True)
    else:
        pdf.cell(0, 10, "[Imagen no válida o no encontrada]", ln=True)
    
    return pdf, None

# Función para crear reporte consolidado
def crear_reporte_consolidado(df, agrupado, logo_disponible=False):
    """Crea el reporte consolidado institucional"""
    pdf = FPDF()
    pdf.add_page()
    
    # Logo si está disponible
    if logo_disponible and os.path.exists("logo_sena.png"):
        try:
            pdf.image("logo_sena.png", x=10, y=8, w=30)
            pdf.ln(35)
        except:
            pdf.ln(10)
    else:
        pdf.ln(10)
    
    # Título principal
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "REPORTE CONSOLIDADO DE CANCELACIONES", ln=True, align="C")
    pdf.ln(10)
    
    # Fecha del reporte
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 8, f"Fecha de generacion: {fecha_actual}", ln=True)
    pdf.ln(5)
    
    # Información por ficha
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "DETALLE POR FICHAS:", ln=True)
    pdf.ln(5)
    
    total_aprendices = 0
    
    for ficha, grupo in agrupado:
        ficha_limpio = limpiar_texto_pdf(str(ficha))
        pdf.set_font("Arial", "B", 11)
        pdf.cell(0, 8, f"FICHA: {ficha_limpio}", ln=True)
        
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 6, f"Cantidad de aprendices: {len(grupo)}", ln=True)
        pdf.cell(0, 6, "Aprendices:", ln=True)
        
        for _, row in grupo.iterrows():
            nombre_limpio = limpiar_texto_pdf(row['Nombre'])
            pdf.cell(0, 5, f"  - {nombre_limpio}", ln=True)
        
        pdf.ln(3)
        total_aprendices += len(grupo)
    
    # Resumen final
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "RESUMEN GENERAL:", ln=True)
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, f"Total de fichas procesadas: {len(agrupado)}", ln=True)
    pdf.cell(0, 8, f"Total de aprendices: {total_aprendices}", ln=True)
    
    return pdf

# Menú principal
if st.session_state.modulo_seleccionado is None:
    st.title("📄 Generador de Reportes de Cancelación")
    st.markdown("### Selecciona el módulo que deseas usar:")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔍 EVIDENCIAS X APRENDICES", use_container_width=True):
            st.session_state.modulo_seleccionado = "aprendices"
            st.rerun()
    with col2:
        if st.button("📂 EVIDENCIA X FICHAS", use_container_width=True):
            st.session_state.modulo_seleccionado = "fichas"
            st.rerun()

# Módulo por aprendiz
elif st.session_state.modulo_seleccionado == "aprendices":
    st.title("🔍 Generación por Aprendiz")

    # Instructivo
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
        st.info("ℹ️ El instructivo no se encuentra disponible.")

    st.subheader("📁 Paso 1: Cargar archivo Excel")
    excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"], key="excel_aprendices")

    st.subheader("🖼️ Paso 2: Cargar evidencias en imagen")
    uploaded_images = st.file_uploader(
        "Imágenes (.png, .jpg)", 
        type=["png", "jpg", "jpeg"], 
        accept_multiple_files=True,
        key="images_aprendices"
    )

    if st.button("Generar documentos", key="generar_aprendices"):
        if not excel_file or not uploaded_images:
            st.error("❗ Debes subir el archivo Excel y al menos una imagen.")
        else:
            try:
                df = pd.read_excel(excel_file)
                columnas_requeridas = {"Nombre", "Ficha", "Evidencia"}
                
                if not columnas_requeridas.issubset(df.columns):
                    st.error("❌ El archivo Excel debe tener las columnas: Nombre, Ficha, Evidencia.")
                else:
                    # Crear diccionario de imágenes
                    imagen_dict = {img.name: img for img in uploaded_images}
                    
                    # Progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    zip_buffer = io.BytesIO()
                    temp_files = []
                    total_procesados = 0
                    total_registros = len(df)
                    
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        agrupado = df.groupby("Ficha")
                        
                        for ficha, grupo in agrupado:
                            ficha_str = str(ficha)
                            resumen_texto = f"RESUMEN FICHA: {ficha_str}\n"
                            resumen_texto += f"Total de aprendices: {len(grupo)}\n"
                            resumen_texto += f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
                            resumen_texto += "APRENDICES:\n"
                            
                            for _, row in grupo.iterrows():
                                nombre = row["Nombre"]
                                # Limpiar nombre para archivo y texto
                                nombre_archivo = limpiar_texto_pdf(nombre).replace(" ", "_").replace("/", "_")
                                evidencia_nombre = row["Evidencia"]
                                evidencia_file = imagen_dict.get(evidencia_nombre)
                                
                                # Actualizar progreso
                                total_procesados += 1
                                progress_bar.progress(total_procesados / total_registros)
                                status_text.text(f"Procesando: {limpiar_texto_pdf(nombre)} (Ficha {ficha_str})")
                                
                                if evidencia_file:
                                    pdf, temp_img_path = crear_pdf_individual(nombre, ficha_str, evidencia_file)
                                    
                                    if temp_img_path:
                                        temp_files.append(temp_img_path)
                                    
                                    # Guardar PDF en ZIP
                                    pdf_bytes = bytes(pdf.output())
                                    ruta_pdf = f"documentos_pdf/Ficha_{ficha_str}/{ficha_str}_{nombre_archivo}.pdf"
                                    zip_file.writestr(ruta_pdf, pdf_bytes)
                                    
                                    resumen_texto += f"- {limpiar_texto_pdf(nombre)} OK\n"
                                else:
                                    st.warning(f"❗ No se encontró la imagen: {evidencia_nombre}")
                                    resumen_texto += f"- {limpiar_texto_pdf(nombre)} (Sin evidencia)\n"
                            
                            # Guardar resumen de ficha
                            resumen_path = f"documentos_pdf/Ficha_{ficha_str}/resumen_ficha_{ficha_str}.txt"
                            zip_file.writestr(resumen_path, resumen_texto)
                        
                        # Crear reporte consolidado
                        status_text.text("Generando reporte consolidado...")
                        pdf_general = crear_reporte_consolidado(df, agrupado, True)
                        pdf_general_bytes = bytes(pdf_general.output())
                        zip_file.writestr("documentos_pdf/REPORTE_CONSOLIDADO.pdf", pdf_general_bytes)
                    
                    # Limpiar archivos temporales
                    for temp_file in temp_files:
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    
                    progress_bar.progress(1.0)
                    status_text.text("✅ Proceso completado exitosamente")
                    
                    zip_buffer.seek(0)
                    st.success("✅ Documentos generados correctamente.")
                    st.download_button(
                        label="📥 Descargar carpeta ZIP con todos los documentos",
                        data=zip_buffer,
                        file_name=f"cancelaciones_aprendices_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                        mime="application/zip"
                    )
                    
            except Exception as e:
                st.error(f"❌ Error al procesar los archivos: {str(e)}")

    if st.button("⬅️ Volver al menú principal", key="volver_aprendices"):
        st.session_state.modulo_seleccionado = None
        st.rerun()

# Módulo por ficha
elif st.session_state.modulo_seleccionado == "fichas":
    st.title("📂 Generación por Ficha con Consolidado Institucional")

    # Instructivo
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
        st.info("ℹ️ El instructivo no se encuentra disponible.")

    st.subheader("📁 Paso 1: Cargar archivo Excel")
    excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"], key="excel_fichas")

    st.subheader("🖼️ Paso 2: Cargar evidencias en imagen")
    uploaded_images = st.file_uploader(
        "Imágenes (.png, .jpg)", 
        type=["png", "jpg", "jpeg"], 
        accept_multiple_files=True,
        key="images_fichas"
    )

    if st.button("Generar documentos", key="generar_fichas"):
        if not excel_file or not uploaded_images:
            st.error("❗ Debes subir el archivo Excel y al menos una imagen.")
        else:
            try:
                df = pd.read_excel(excel_file)
                columnas_requeridas = {"Nombre", "Ficha", "Evidencia"}
                
                if not columnas_requeridas.issubset(df.columns):
                    st.error("❌ El archivo Excel debe tener las columnas: Nombre, Ficha, Evidencia.")
                else:
                    # Crear diccionario de imágenes
                    imagen_dict = {img.name: img for img in uploaded_images}
                    
                    # Progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    zip_buffer = io.BytesIO()
                    temp_files = []
                    total_procesados = 0
                    total_registros = len(df)
                    
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        agrupado = df.groupby("Ficha")
                        
                        # Procesar cada ficha
                        for ficha, grupo in agrupado:
                            ficha_str = str(ficha)
                            
                            for _, row in grupo.iterrows():
                                nombre = row["Nombre"]
                                # Limpiar nombre para archivo y texto
                                nombre_archivo = limpiar_texto_pdf(nombre).replace(" ", "_").replace("/", "_")
                                evidencia_nombre = row["Evidencia"]
                                evidencia_file = imagen_dict.get(evidencia_nombre)
                                
                                # Actualizar progreso
                                total_procesados += 1
                                progress_bar.progress(total_procesados / total_registros)
                                status_text.text(f"Procesando: {limpiar_texto_pdf(nombre)} (Ficha {ficha_str})")
                                
                                if evidencia_file:
                                    pdf, temp_img_path = crear_pdf_individual(nombre, ficha_str, evidencia_file)
                                    
                                    if temp_img_path:
                                        temp_files.append(temp_img_path)
                                    
                                    # Guardar PDF en carpeta de ficha
                                    pdf_bytes = bytes(pdf.output())
                                    ruta_pdf = f"documentos_pdf/Ficha_{ficha_str}/{ficha_str}_{nombre_archivo}.pdf"
                                    zip_file.writestr(ruta_pdf, pdf_bytes)
                                else:
                                    st.warning(f"❗ No se encontró la imagen: {evidencia_nombre} para {limpiar_texto_pdf(nombre)}")
                        
                        # Crear reporte consolidado institucional
                        status_text.text("Generando reporte consolidado institucional...")
                        pdf_consolidado = crear_reporte_consolidado(df, agrupado, True)
                        pdf_consolidado_bytes = bytes(pdf_consolidado.output())
                        zip_file.writestr("REPORTE_CONSOLIDADO_INSTITUCIONAL.pdf", pdf_consolidado_bytes)
                    
                    # Limpiar archivos temporales
                    for temp_file in temp_files:
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    
                    progress_bar.progress(1.0)
                    status_text.text("✅ Proceso completado exitosamente")
                    
                    zip_buffer.seek(0)
                    st.success("✅ Documentos generados correctamente.")
                    st.download_button(
                        label="📥 Descargar carpeta ZIP con todos los documentos",
                        data=zip_buffer,
                        file_name=f"cancelaciones_fichas_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                        mime="application/zip"
                    )
                    
            except Exception as e:
                st.error(f"❌ Error al procesar los archivos: {str(e)}")

    if st.button("⬅️ Volver al menú principal", key="volver_fichas"):
        st.session_state.modulo_seleccionado = None
        st.rerun()
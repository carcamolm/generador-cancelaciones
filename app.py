import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.units import inch
from PIL import Image
from datetime import datetime
import io
import zipfile
import tempfile
import os

# Configurar la página
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

# Función para crear PDF individual usando ReportLab
def crear_pdf_individual(nombre, ficha, evidencia_file):
    """Crea un PDF individual para cada estudiante usando ReportLab"""
    try:
        # Crear buffer para el PDF
        buffer = io.BytesIO()
        
        # Crear canvas
        p = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4
        
        # Información del estudiante
        p.setFont("Helvetica-Bold", 16)
        p.drawString(50, height - 50, f"FICHA: {str(ficha)}")
        
        p.setFont("Helvetica-Bold", 14)
        p.drawString(50, height - 80, f"APRENDIZ: {str(nombre)}")
        
        p.setFont("Helvetica-Bold", 12)
        p.drawString(50, height - 110, "EVIDENCIA CORREO")
        
        # Agregar imagen si es válida
        temp_img_path = None
        if evidencia_file and es_imagen_valida(evidencia_file):
            try:
                imagen = Image.open(evidencia_file)
                
                # Redimensionar imagen si es muy grande
                max_width, max_height = 500, 400
                if imagen.width > max_width or imagen.height > max_height:
                    imagen.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
                
                # Guardar imagen temporalmente
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                    imagen.save(tmp_img.name, format="PNG")
                    temp_img_path = tmp_img.name
                    
                    # Agregar imagen al PDF
                    img_reader = ImageReader(tmp_img.name)
                    img_width, img_height = imagen.size
                    
                    # Calcular posición y tamaño en el PDF
                    scale = min(400/img_width, 300/img_height, 1.0)
                    final_width = img_width * scale
                    final_height = img_height * scale
                    
                    p.drawImage(img_reader, 50, height - 150 - final_height, 
                              width=final_width, height=final_height)
                    
            except Exception as e:
                p.setFont("Helvetica", 10)
                p.drawString(50, height - 140, f"[Error al cargar imagen: {str(e)}]")
        else:
            p.setFont("Helvetica", 10)
            p.drawString(50, height - 140, "[Imagen no válida o no encontrada]")
        
        # Finalizar el PDF
        p.save()
        buffer.seek(0)
        
        return buffer.getvalue(), temp_img_path
        
    except Exception as e:
        st.error(f"Error creando PDF: {str(e)}")
        return None, None

# Función para crear PDF individual por aprendiz (módulo fichas)
def crear_pdf_aprendiz_ficha(nombre, ficha, evidencia_file):
    """Crea un PDF individual para cada estudiante usando el formato específico para fichas"""
    try:
        # Crear buffer para el PDF
        buffer = io.BytesIO()
        
        # Crear canvas
        c = canvas.Canvas(buffer, pagesize=A4)
        
        # Información del estudiante
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2*cm, 27*cm, f"FICHA: {str(ficha)}")
        c.drawString(2*cm, 26.2*cm, f"APRENDIZ: {str(nombre)}")
        
        c.setFont("Helvetica", 12)
        c.drawString(2*cm, 25.2*cm, "EVIDENCIA CORREO:")
        
        # Agregar imagen si es válida
        temp_img_path = None
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
                    temp_img_path = tmp_img.name
                    
                    # Agregar imagen al PDF
                    img = ImageReader(tmp_img.name)
                    c.drawImage(img, 2*cm, 12*cm, width=16*cm, preserveAspectRatio=True, mask='auto')
                    
            except Exception as e:
                c.drawString(2*cm, 24.5*cm, f"[Error al insertar imagen: {str(e)}]")
        else:
            c.drawString(2*cm, 24.5*cm, "[Imagen no encontrada o no válida]")
        
        c.showPage()
        c.save()
        buffer.seek(0)
        
        return buffer.getvalue(), temp_img_path
        
    except Exception as e:
        st.error(f"Error creando PDF: {str(e)}")
        return None, None

# Función para crear reporte consolidado institucional (módulo fichas)
def crear_reporte_consolidado_institucional(df, agrupado, logo_disponible=False):
    """Crea el reporte consolidado institucional sin evidencias individuales"""
    try:
        # Crear buffer para el PDF
        buffer = io.BytesIO()
        
        # Crear canvas
        c = canvas.Canvas(buffer, pagesize=A4)
        
        # Título principal
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(10.5*cm, 23.5*cm, "REPORTE CONSOLIDADO DE CANCELACIONES")
        
        # Logo institucional
        if logo_disponible and os.path.exists("logo_sena.png"):
            try:
                logo = ImageReader("logo_sena.png")
                c.drawImage(logo, 2*cm, 25.5*cm, width=4*cm, preserveAspectRatio=True, mask='auto')
            except:
                pass
        
        c.setFont("Helvetica", 12)
        y = 20.5*cm
        resumen_fichas = {}
        total_general = 0
        
        # Información por ficha
        for ficha, grupo in agrupado:
            # Verificar si necesitamos nueva página
            if y < 6*cm:
                c.showPage()
                y = 27*cm
            
            c.drawString(2*cm, y, f"📌 FICHA: {str(ficha)}")
            y -= 0.6*cm
            c.drawString(2*cm, y, "APRENDICES:")
            y -= 0.5*cm
            
            for _, row in grupo.iterrows():
                if y < 4*cm:
                    c.showPage()
                    y = 27*cm
                c.drawString(2.5*cm, y, f"- {str(row['Nombre'])}")
                y -= 0.4*cm
            
            cantidad = len(grupo)
            resumen_fichas[ficha] = cantidad
            total_general += cantidad
            y -= 0.8*cm
        
        # Resumen final en nueva página
        c.showPage()
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2*cm, 27*cm, "📊 RESUMEN FINAL POR FICHA")
        c.setFont("Helvetica", 12)
        y = 26*cm
        
        for ficha, cantidad in resumen_fichas.items():
            c.drawString(2*cm, y, f"Ficha {str(ficha)}: {cantidad} aprendices")
            y -= 0.5*cm
        
        c.drawString(2*cm, y - 0.5*cm, f"🧮 TOTAL GENERAL: {total_general} aprendices")
        c.save()
        
        buffer.seek(0)
        return buffer.getvalue()
        
    except Exception as e:
        st.error(f"Error creando reporte consolidado institucional: {str(e)}")
        return None
    """Crea el reporte consolidado institucional usando ReportLab"""
    try:
        # Crear buffer para el PDF
        buffer = io.BytesIO()
        
        # Crear canvas
        p = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4
        y_position = height - 50
        
        # Logo si está disponible
        if logo_disponible and os.path.exists("logo_sena.png"):
            try:
                logo_reader = ImageReader("logo_sena.png")
                p.drawImage(logo_reader, 50, y_position - 60, width=80, height=60)
                y_position -= 80
            except:
                y_position -= 20
        else:
            y_position -= 20
        
        # Título principal
        p.setFont("Helvetica-Bold", 18)
        title = "REPORTE CONSOLIDADO DE CANCELACIONES"
        title_width = p.stringWidth(title, "Helvetica-Bold", 18)
        p.drawString((width - title_width) / 2, y_position, title)
        y_position -= 40
        
        # Fecha del reporte
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
        p.setFont("Helvetica", 12)
        p.drawString(50, y_position, f"Fecha de generación: {fecha_actual}")
        y_position -= 30
        
        # Información por ficha
        p.setFont("Helvetica-Bold", 14)
        p.drawString(50, y_position, "DETALLE POR FICHAS:")
        y_position -= 25
        
        total_aprendices = 0
        
        for ficha, grupo in agrupado:
            # Verificar si necesitamos nueva página
            if y_position < 100:
                p.showPage()
                y_position = height - 50
            
            p.setFont("Helvetica-Bold", 12)
            p.drawString(50, y_position, f"FICHA: {str(ficha)}")
            y_position -= 20
            
            p.setFont("Helvetica", 10)
            p.drawString(70, y_position, f"Cantidad de aprendices: {len(grupo)}")
            y_position -= 15
            
            p.drawString(70, y_position, "Aprendices:")
            y_position -= 12
            
            for _, row in grupo.iterrows():
                if y_position < 50:
                    p.showPage()
                    y_position = height - 50
                
                p.setFont("Helvetica", 9)
                p.drawString(90, y_position, f"- {str(row['Nombre'])}")
                y_position -= 12
            
            y_position -= 10
            total_aprendices += len(grupo)
        
        # Resumen final
        if y_position < 100:
            p.showPage()
            y_position = height - 50
            
        p.setFont("Helvetica-Bold", 14)
        p.drawString(50, y_position, "RESUMEN GENERAL:")
        y_position -= 25
        
        p.setFont("Helvetica", 12)
        p.drawString(50, y_position, f"Total de fichas procesadas: {len(agrupado)}")
        y_position -= 20
        p.drawString(50, y_position, f"Total de aprendices: {total_aprendices}")
        
        # Finalizar el PDF
        p.save()
        buffer.seek(0)
        
# Función para crear reporte consolidado usando ReportLab (módulo aprendices)
def crear_reporte_consolidado(df, agrupado, logo_disponible=False):
    """Crea el reporte consolidado institucional usando ReportLab"""
    try:
        # Crear buffer para el PDF
        buffer = io.BytesIO()
        
        # Crear canvas
        p = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4
        y_position = height - 50
        
        # Logo si está disponible
        if logo_disponible and os.path.exists("logo_sena.png"):
            try:
                logo_reader = ImageReader("logo_sena.png")
                p.drawImage(logo_reader, 50, y_position - 60, width=80, height=60)
                y_position -= 80
            except:
                y_position -= 20
        else:
            y_position -= 20
        
        # Título principal
        p.setFont("Helvetica-Bold", 18)
        title = "REPORTE CONSOLIDADO DE CANCELACIONES"
        title_width = p.stringWidth(title, "Helvetica-Bold", 18)
        p.drawString((width - title_width) / 2, y_position, title)
        y_position -= 40
        
        # Fecha del reporte
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
        p.setFont("Helvetica", 12)
        p.drawString(50, y_position, f"Fecha de generación: {fecha_actual}")
        y_position -= 30
        
        # Información por ficha
        p.setFont("Helvetica-Bold", 14)
        p.drawString(50, y_position, "DETALLE POR FICHAS:")
        y_position -= 25
        
        total_aprendices = 0
        
        for ficha, grupo in agrupado:
            # Verificar si necesitamos nueva página
            if y_position < 100:
                p.showPage()
                y_position = height - 50
            
            p.setFont("Helvetica-Bold", 12)
            p.drawString(50, y_position, f"FICHA: {str(ficha)}")
            y_position -= 20
            
            p.setFont("Helvetica", 10)
            p.drawString(70, y_position, f"Cantidad de aprendices: {len(grupo)}")
            y_position -= 15
            
            p.drawString(70, y_position, "Aprendices:")
            y_position -= 12
            
            for _, row in grupo.iterrows():
                if y_position < 50:
                    p.showPage()
                    y_position = height - 50
                
                p.setFont("Helvetica", 9)
                p.drawString(90, y_position, f"- {str(row['Nombre'])}")
                y_position -= 12
            
            y_position -= 10
            total_aprendices += len(grupo)
        
        # Resumen final
        if y_position < 100:
            p.showPage()
            y_position = height - 50
            
        p.setFont("Helvetica-Bold", 14)
        p.drawString(50, y_position, "RESUMEN GENERAL:")
        y_position -= 25
        
        p.setFont("Helvetica", 12)
        p.drawString(50, y_position, f"Total de fichas procesadas: {len(agrupado)}")
        y_position -= 20
        p.drawString(50, y_position, f"Total de aprendices: {total_aprendices}")
        
        # Finalizar el PDF
        p.save()
        buffer.seek(0)
        
        return buffer.getvalue()
        
    except Exception as e:
        st.error(f"Error creando reporte consolidado: {str(e)}")
        return None

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
                                # Crear nombre de archivo seguro
                                nombre_archivo = str(nombre).replace(" ", "_").replace("/", "_").replace("\\", "_")
                                # Remover caracteres problemáticos para nombres de archivo
                                caracteres_problematicos = '<>:"/\\|?*áéíóúñüÁÉÍÓÚÑÜ'
                                for char in caracteres_problematicos:
                                    nombre_archivo = nombre_archivo.replace(char, "_")
                                
                                evidencia_nombre = row["Evidencia"]
                                evidencia_file = imagen_dict.get(evidencia_nombre)
                                
                                # Actualizar progreso
                                total_procesados += 1
                                progress_bar.progress(total_procesados / total_registros)
                                status_text.text(f"Procesando: {str(nombre)[:30]}... (Ficha {ficha_str})")
                                
                                if evidencia_file:
                                    pdf_bytes, temp_img_path = crear_pdf_individual(nombre, ficha_str, evidencia_file)
                                    
                                    if pdf_bytes is not None:
                                        if temp_img_path:
                                            temp_files.append(temp_img_path)
                                        
                                        # Guardar PDF en ZIP
                                        ruta_pdf = f"documentos_pdf/Ficha_{ficha_str}/{ficha_str}_{nombre_archivo}.pdf"
                                        zip_file.writestr(ruta_pdf, pdf_bytes)
                                        
                                        resumen_texto += f"- {str(nombre)} OK\n"
                                    else:
                                        resumen_texto += f"- {str(nombre)} (Error creacion PDF)\n"
                                else:
                                    st.warning(f"❗ No se encontró la imagen: {evidencia_nombre}")
                                    resumen_texto += f"- {str(nombre)} (Sin evidencia)\n"
                            
                            # Guardar resumen de ficha
                            resumen_path = f"documentos_pdf/Ficha_{ficha_str}/resumen_ficha_{ficha_str}.txt"
                            zip_file.writestr(resumen_path, resumen_texto.encode('utf-8'))
                        
                        # Crear reporte consolidado
                        status_text.text("Generando reporte consolidado...")
                        pdf_general_bytes = crear_reporte_consolidado(df, agrupado, True)
                        
                        if pdf_general_bytes is not None:
                            zip_file.writestr("documentos_pdf/REPORTE_CONSOLIDADO.pdf", pdf_general_bytes)
                        else:
                            st.warning("No se pudo crear el reporte consolidado")
                    
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
                                # Crear nombre de archivo seguro
                                nombre_archivo = str(nombre).replace(" ", "_").replace("/", "_").replace("\\", "_")
                                # Remover caracteres problemáticos para nombres de archivo
                                caracteres_problematicos = '<>:"/\\|?*áéíóúñüÁÉÍÓÚÑÜ'
                                for char in caracteres_problematicos:
                                    nombre_archivo = nombre_archivo.replace(char, "_")
                                
                                evidencia_nombre = row["Evidencia"]
                                evidencia_file = imagen_dict.get(evidencia_nombre)
                                
                                # Actualizar progreso
                                total_procesados += 1
                                progress_bar.progress(total_procesados / total_registros)
                                status_text.text(f"Procesando: {str(nombre)[:30]}... (Ficha {ficha_str})")
                                
                                if evidencia_file:
                                    pdf_bytes, temp_img_path = crear_pdf_aprendiz_ficha(nombre, ficha_str, evidencia_file)
                                    
                                    if pdf_bytes is not None:
                                        if temp_img_path:
                                            temp_files.append(temp_img_path)
                                        
                                        # Guardar PDF en carpeta de ficha
                                        ruta_pdf = f"documentos_pdf/{ficha_str}/{ficha_str}_{nombre_archivo}.pdf"
                                        zip_file.writestr(ruta_pdf, pdf_bytes)
                                else:
                                    st.warning(f"❗ No se encontró la imagen: {evidencia_nombre} para {str(nombre)}")
                        
                        # Crear reporte consolidado institucional
                        status_text.text("Generando reporte consolidado institucional...")
                        pdf_consolidado_bytes = crear_reporte_consolidado_institucional(df, agrupado, True)
                        
                        if pdf_consolidado_bytes is not None:
                            zip_file.writestr("REPORTE_CONSOLIDADO_INSTITUCIONAL.pdf", pdf_consolidado_bytes)
                        else:
                            st.warning("No se pudo crear el reporte consolidado institucional")
                    
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
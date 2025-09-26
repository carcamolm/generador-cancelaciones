import streamlit as st
import pandas as pd
from PIL import Image
import io

st.set_page_config(page_title="Generador de Reportes de Cancelación", layout="centered")

st.title("📄 Generador de Reportes de Cancelación")
st.markdown("Sube el archivo Excel y las evidencias en imagen para generar los documentos.")

# 📁 Subir archivo Excel
excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"])

# 🖼️ Subir múltiples imágenes
uploaded_images = st.file_uploader(
    "Sube las evidencias (imágenes)", 
    type=["png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

# 🧠 Procesar al hacer clic
if st.button("Generar documentos"):
    if not excel_file:
        st.error("⚠️ Debes subir un archivo Excel.")
    elif not uploaded_images:
        st.error("⚠️ Debes subir al menos una imagen de evidencia.")
    else:
        try:
            df = pd.read_excel(excel_file)
            st.success(f"✅ Se cargaron {len(df)} registros del Excel.")
            st.success(f"✅ Se cargaron {len(uploaded_images)} imágenes.")

            # Ejemplo de procesamiento: mostrar nombres de archivos
            st.markdown("### 🖼️ Imágenes cargadas:")
            for img in uploaded_images:
                st.write(f"- {img.name}")
                image = Image.open(img)
                st.image(image, caption=img.name, width=150)

            # Aquí puedes agregar la lógica para generar PDFs o documentos
            st.info("✅ Documentos generados correctamente (simulado).")

        except Exception as e:
            st.error(f"❌ Error al procesar: {e}")

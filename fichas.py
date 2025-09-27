from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import pandas as pd
import os

# === CONFIGURACIÓN ===

archivo_excel = "estudiantes.xlsx"
carpeta_pdf = "documentos_pdf"
logo_path = "logo_sena.png"
pdf_consolidado = "reporte_consolidado.pdf"

os.makedirs(carpeta_pdf, exist_ok=True)

# === LECTURA DE DATOS ===

df = pd.read_excel(archivo_excel)
agrupado = df.groupby("Ficha")

# === FUNCIÓN PARA CREAR PDF INDIVIDUAL POR APRENDIZ ===

def crear_pdf_aprendiz(nombre, ficha, imagen_evidencia, ruta_pdf):
    c = canvas.Canvas(ruta_pdf, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(2*cm, 27*cm, f"FICHA: {ficha}")
    c.drawString(2*cm, 26.2*cm, f"APRENDIZ: {nombre}")
    c.setFont("Helvetica", 12)
    c.drawString(2*cm, 25.2*cm, "EVIDENCIA CORREO:")

    if os.path.exists(imagen_evidencia):
        try:
            img = ImageReader(imagen_evidencia)
            c.drawImage(img, 2*cm, 12*cm, width=16*cm, preserveAspectRatio=True, mask='auto')
        except:
            c.drawString(2*cm, 24.5*cm, "[Error al insertar imagen]")
    else:
        c.drawString(2*cm, 24.5*cm, "[Imagen no encontrada]")

    c.showPage()
    c.save()

# === GENERACIÓN DE PDFs INDIVIDUALES AGRUPADOS POR FICHA ===

# === GENERACIÓN DE PDFs INDIVIDUALES AGRUPADOS POR FICHA ===

for index, row in df.iterrows():
    nombre = row["Nombre"]
    ficha = str(row["Ficha"])
    imagen_evidencia = row["Evidencia"]

    carpeta_ficha = os.path.join(carpeta_pdf, ficha)
    os.makedirs(carpeta_ficha, exist_ok=True)

    nombre_archivo = f"EVIDENCIAS_DESERCIÓN_{nombre.replace(' ', '_')}_.pdf"
    ruta_pdf = os.path.join(carpeta_ficha, nombre_archivo)

    crear_pdf_aprendiz(nombre, ficha, imagen_evidencia, ruta_pdf)

# === CREACIÓN DEL PDF CONSOLIDADO INSTITUCIONAL ===

c = canvas.Canvas(pdf_consolidado, pagesize=A4)
c.setFont("Helvetica-Bold", 16)
c.drawCentredString(10.5*cm, 23.5*cm, "REPORTE CONSOLIDADO DE CANCELACIONES")

# Logo institucional
if os.path.exists(logo_path):
    try:
        logo = ImageReader(logo_path)
        c.drawImage(logo, 2*cm, 25.5*cm, width=4*cm, preserveAspectRatio=True, mask='auto')
    except:
        pass

c.setFont("Helvetica", 12)
y = 20.5*cm
resumen_fichas = {}
total_general = 0

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

    cantidad = len(grupo)
    resumen_fichas[ficha] = cantidad
    total_general += cantidad
    y -= 0.8*cm

# === RESUMEN FINAL ===

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

print("✅ PDFs individuales agrupados por ficha y PDF institucional generado sin evidencias.")

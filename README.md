# 🧾 Generador de Reportes de Cancelación

Aplicación institucional desarrollada para automatizar la generación de reportes de cancelación por aprendiz y por ficha, integrando evidencias visuales y consolidado institucional en formato PDF. Diseñada para facilitar el trabajo de instructores, coordinadores y personal administrativo del SENA.

---

## 🚀 ¿Qué hace esta app?

- 📂 Permite cargar un archivo Excel con los datos de aprendices y fichas
- 🖼️ Integra evidencias en imagen (.png, .jpg) por cada aprendiz
- 🧾 Genera documentos PDF individuales por aprendiz con formato institucional
- 📋 Crea reportes agrupados por ficha con portada y consolidado
- 📦 Descarga final en formato ZIP con estructura organizada por carpeta
- 📘 Incluye instructivo institucional descargable

---

## 🧩 Estructura modular

La app está organizada en tres módulos principales:

- `app.py` → Menú principal y navegación
- `aprendices.py` → Generación por aprendiz (FPDF)
- `fichas.py` → Generación por ficha (ReportLab)

---

## 📁 Requisitos técnicos

Antes de ejecutar la app, asegúrate de tener instaladas las siguientes dependencias:

```bash
streamlit
fpdf
reportlab
pandas
Pillow

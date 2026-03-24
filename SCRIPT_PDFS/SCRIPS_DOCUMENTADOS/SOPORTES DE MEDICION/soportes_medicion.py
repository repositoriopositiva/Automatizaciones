import os
import sys
import pandas as pd
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, Image
)
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import styles
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter

# ==========================================================
# 🔍 BUSCAR ARCHIVO EXCEL
# ==========================================================
archivos_excel = [f for f in os.listdir() if f.lower().endswith(".xlsx")]

if not archivos_excel:
    print("❌ No se encontró ningún archivo .xlsx en la carpeta.")
    sys.exit()

nombre_archivo = archivos_excel[0]
print(f"📄 Excel detectado: {nombre_archivo}")

# ==========================================================
# 📖 LEER EXCEL
# ==========================================================
df = pd.read_excel(nombre_archivo)

# ==========================================================
# 🧹 LIMPIAR NOMBRES DE COLUMNAS
# ==========================================================
df.columns = (
    df.columns
    .str.strip()
    .str.replace("\n", " ", regex=False)
    .str.replace("  ", " ", regex=False)
)

# ==========================================================
# 🔎 DETECTAR COLUMNAS AUTOMÁTICAMENTE
# ==========================================================
def buscar_columna(palabra_clave):
    for col in df.columns:
        if palabra_clave.upper() in col.upper():
            return col
    return None

col_cedula   = buscar_columna("CEDULA")
col_nombre   = buscar_columna("NOMBRE")
col_contrato = buscar_columna("CONTRATO")
col_razon    = buscar_columna("RAZON")

# ==========================================================
# 🚨 VALIDAR FORMATO DEL EXCEL
# ==========================================================
if not all([col_cedula, col_nombre, col_contrato, col_razon]):
    print("\n❌ LA INFORMACIÓN ESTÁ EN FORMATO INCORRECTO\n")
    print("📌 El formato del Excel debe ser el siguiente orden:\n")
    print("| Cédula | Nombre | Preguntas variables | Contrato | Razón Social | ... |")
    print("\n⚠ Verifique que existan las columnas:")
    print("- Cédula")
    print("- Nombre")
    print("- Contrato")
    print("- Razón Social\n")
    sys.exit()

# Obtener posiciones reales
idx_cedula = df.columns.get_loc(col_cedula)
idx_nombre = df.columns.get_loc(col_nombre)
idx_contrato = df.columns.get_loc(col_contrato)

# Validar orden lógico
if not (idx_cedula < idx_nombre < idx_contrato):
    print("\n❌ LA INFORMACIÓN ESTÁ EN FORMATO INCORRECTO\n")
    print("📌 El orden correcto debe ser:\n")
    print("| Cédula | Nombre | Preguntas variables | Contrato | Razón Social | ... |")
    print("\n⚠ El campo 'Contrato' debe estar después de las preguntas.\n")
    sys.exit()

# ==========================================================
# 📊 DETECTAR PREGUNTAS VARIABLES
# ==========================================================
columnas_preguntas = df.columns[idx_nombre+1 : idx_contrato].tolist()

if len(columnas_preguntas) == 0:
    print("\n❌ LA INFORMACIÓN ESTÁ EN FORMATO INCORRECTO\n")
    print("📌 No se detectaron preguntas entre 'Nombre' y 'Contrato'\n")
    print("Formato esperado:\n")
    print("| Cédula | Nombre | Preguntas variables | Contrato | Razón Social | ... |\n")
    sys.exit()

# ==========================================================
# 📅 DETECTAR FECHA DESDE EXCEL
# ==========================================================
# La primera columna después de Nombre (columna C) se toma como fecha
col_fecha = columnas_preguntas[0]

# Columnas finales que se exportarán
columnas_respuestas = [col_cedula, col_nombre] + columnas_preguntas

# ==========================================================
# 📂 CARPETA SALIDA
# ==========================================================
carpeta_salida = "SOPORTES_DE_MEDICION_GENERADOS"
os.makedirs(carpeta_salida, exist_ok=True)

# ==========================================================
# 🎨 ESTILOS
# ==========================================================
estilos = styles.getSampleStyleSheet()

estilo_celda = ParagraphStyle(
    'celda',
    parent=estilos['Normal'],
    fontSize=8,
    leading=10
)

# ==========================================================
# 📊 AGRUPAR POR PRESTADOR
# ==========================================================
grupos = df.groupby([col_contrato, col_razon])

for (contrato, razon), grupo in grupos:

    if pd.isna(contrato) or pd.isna(razon):
        continue

    # ======================================================
    # 📅 OBTENER FECHA DESDE EXCEL
    # ======================================================
    fecha_elaboracion = grupo.iloc[0][col_fecha]

    if pd.notna(fecha_elaboracion):

        # Si ya viene como texto (ej: "febrero 2026")
        if isinstance(fecha_elaboracion, str):
            fecha_elaboracion = fecha_elaboracion.strip()

        else:
            try:
                fecha_elaboracion = pd.to_datetime(fecha_elaboracion).strftime("%B %Y")
            except:
                fecha_elaboracion = str(fecha_elaboracion)

    else:
        fecha_elaboracion = ""

    nombre_pdf = f"{contrato} SOPORTES DE MEDICION"
    ruta_pdf = os.path.join(carpeta_salida, f"{nombre_pdf}.pdf")

    doc = SimpleDocTemplate(
        ruta_pdf,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    elementos = []

    # ======================================================
    # ENCABEZADO CON LOGO
    # ======================================================
    encabezado_tabla = []

    titulo = Paragraph(
        "<b>RESULTADOS ENCUESTA DE SATISFACCIÓN</b>",
        estilos["Heading2"]
    )

    if os.path.exists("logo.png"):
        logo = Image("logo.png", width=1.5*inch, height=1*inch)
        encabezado_tabla.append([titulo, logo])
    else:
        encabezado_tabla.append([titulo, ""])

    tabla_encabezado = Table(
        encabezado_tabla,
        colWidths=[doc.width * 0.75, doc.width * 0.25]
    )

    elementos.append(tabla_encabezado)
    elementos.append(Spacer(1, 0.2 * inch))

    elementos.append(Paragraph(f"<b>Contrato:</b> {contrato}", estilos["Normal"]))
    elementos.append(Paragraph(f"<b>Razón Social:</b> {razon}", estilos["Normal"]))
    elementos.append(Paragraph(f"<b>Periodo Evaluado:</b> {fecha_elaboracion}", estilos["Normal"]))
    elementos.append(Spacer(1, 0.3 * inch))

    # ======================================================
    # TABLA DE RESPUESTAS
    # ======================================================
    datos_tabla = []

    encabezados = [Paragraph(f"<b>{col}</b>", estilo_celda) for col in columnas_respuestas]
    datos_tabla.append(encabezados)

    for _, fila in grupo.iterrows():
        fila_formateada = []
        for col in columnas_respuestas:
            valor = fila[col]
            texto = "" if pd.isna(valor) else str(valor)
            fila_formateada.append(Paragraph(texto, estilo_celda))
        datos_tabla.append(fila_formateada)

    # ======================================================
    # ANCHOS DE COLUMNAS
    # ======================================================
    total_width = doc.width
    anchos = []

    for col in columnas_respuestas:
        col_upper = col.upper()

        if "CEDULA" in col_upper:
            anchos.append(total_width * 0.1)

        elif "NOMBRE" in col_upper:
            anchos.append(total_width * 0.1)

        elif "OBSERV" in col_upper:
            anchos.append(total_width * 0.2)

        else:
            anchos.append(total_width * 0.55 / (len(columnas_respuestas) - 3))

    tabla = Table(
        datos_tabla,
        colWidths=anchos,
        repeatRows=1
    )

    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.3, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    elementos.append(tabla)

    doc.build(elementos)

print("✅ PDFs generados correctamente con fecha tomada desde el Excel")
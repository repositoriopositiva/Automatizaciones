import os
import zipfile
import pandas as pd
from fpdf import FPDF
import unicodedata
import shutil
import sys
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# ==========================================================
# 🔍 BUSCAR ARCHIVO EXCEL
# ==========================================================
archivos_excel = [f for f in os.listdir() if f.lower().endswith(".xlsx")]

if not archivos_excel:
    print("❌ No se encontró ningún archivo .xlsx en la carpeta.")
    sys.exit()

if len(archivos_excel) > 1:
    print("⚠️ Se encontraron varios archivos .xlsx. Usando el primero.")

nombre_archivo = archivos_excel[0]
print(f"📄 Excel detectado: {nombre_archivo}")

# ==========================================================
# 🧹 FUNCIÓN LIMPIEZA TEXTO
# ==========================================================
def limpiar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return texto.strip()

# ==========================================================
# 📄 CARGAR EXCEL
# ==========================================================
df = pd.read_excel(nombre_archivo)

df.columns = (
    df.columns.str.replace("\ufeff", "", regex=True)
    .str.strip()
    .str.lower()
)

# ==========================================================
# 🔎 BUSCAR COLUMNAS
# ==========================================================
def buscar_columna(nombre):
    for c in df.columns:
        if nombre.replace(" ", "") == c.replace(" ", ""):
            return c
    raise Exception(f"❌ No se encontró la columna: {nombre}")

col_contrato  = buscar_columna("contrato")
col_nit       = buscar_columna("nit")
col_prestador = buscar_columna("razón social")
col_total     = buscar_columna("total numero de autorización")
col_aprobado  = buscar_columna("aprobado")
col_gestionado= buscar_columna("gestionado")
col_cerrado   = buscar_columna("cerrado")
col_pendiente = buscar_columna("pendiente")
col_anulado   = buscar_columna("anulado")
col_cancelado = buscar_columna("cancelado")

# ==========================================================
# 🧹 LIMPIAR CARPETA DE SALIDA
# ==========================================================
output_folder = "AUTORIZACIONES_GENERADAS"
if os.path.exists(output_folder):
    shutil.rmtree(output_folder)
os.makedirs(output_folder)

print("⏳ Procesando...\n")

# ==========================================================
# 🧾 GENERAR PDFS
# ==========================================================
total_registros = len(df)

for i, (_, row) in enumerate(df.iterrows(), start=1):

    contrato   = limpiar_texto(row[col_contrato])
    nit        = limpiar_texto(row[col_nit])
    prestador  = limpiar_texto(row[col_prestador])
    total      = limpiar_texto(row[col_total])
    aprobado   = limpiar_texto(row[col_aprobado])
    gestionado = limpiar_texto(row[col_gestionado])
    cerrado    = limpiar_texto(row[col_cerrado])
    pendiente  = limpiar_texto(row[col_pendiente])
    anulado    = limpiar_texto(row[col_anulado])
    cancelado  = limpiar_texto(row[col_cancelado])

    print(f"➡ Generando {i}/{total_registros} → contrato: {contrato}")

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    # LOGO
    if os.path.exists("logo.png"):
        pdf.image("logo.png", x=pdf.w - 60, y=8, w=50)

    # TÍTULO
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 15, "AUTORIZACIONES GENERADAS POR ESTADO", ln=True)

    pdf.ln(8)

    # DATOS
    pdf.set_font("Arial", "B", 11)
    pdf.cell(35, 8, "CONTRATO:")
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, contrato, ln=True)

    pdf.set_font("Arial", "B", 11)
    pdf.cell(35, 8, "PRESTADOR:")
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, prestador, ln=True)

    pdf.set_font("Arial", "B", 11)
    pdf.cell(35, 8, "NIT:")
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, nit, ln=True)

    pdf.ln(10)

    # TABLA CENTRADA
    table_width = 100
    start_x = (pdf.w - table_width) / 2

    pdf.set_font("Arial", "B", 12)
    pdf.set_x(start_x)
    pdf.cell(60, 8, "ESTADO", 1, 0, "C")
    pdf.cell(40, 8, "CANTIDAD", 1, 1, "C")

    pdf.set_font("Arial", "", 11)
    for estado, valor in [
        ("APROBADO", aprobado),
        ("GESTIONADO", gestionado),
        ("CERRADO", cerrado),
        ("PENDIENTE", pendiente),
        ("ANULADO", anulado),
        ("CANCELADO", cancelado),
    ]:
        pdf.set_x(start_x)
        pdf.cell(60, 8, estado, 1, 0, "C")
        pdf.cell(40, 8, str(valor), 1, 1, "C")

    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, f"Total autorizaciones: {total}", ln=True)

    nombre_salida = f"{contrato} - AUTORIZACIONES.pdf".replace("/", "-")
    pdf.output(os.path.join(output_folder, nombre_salida))


print("\n✅ PROCESO COMPLETADO")
print(f"📂 PDFs generados en: {output_folder}")

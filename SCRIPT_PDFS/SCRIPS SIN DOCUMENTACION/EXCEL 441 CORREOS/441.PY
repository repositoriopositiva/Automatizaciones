import pandas as pd
import os
import sys
import re

# ==========================================================
# 🔍 BUSCAR ARCHIVO EXCEL
# ==========================================================
archivos_excel = [f for f in os.listdir() if f.lower().endswith(".xlsb") or f.lower().endswith(".xlsx")]  # ✅ Acepta ambos formatos

if not archivos_excel:
    print("❌ No se encontró ningún archivo .xlsx o .xlsb en la carpeta.")
    sys.exit()

if len(archivos_excel) > 1:
    print("⚠️ Se encontraron varios archivos. Usando el primero.")

nombre_archivo = archivos_excel[0]
extension = os.path.splitext(nombre_archivo)[1].lower()  # ✅ Detecta la extensión
print(f"📄 Excel detectado: {nombre_archivo}")

# ==========================================================
# 📂 CONFIGURACIÓN SEGÚN TU EXCEL
# ==========================================================
columna_nit = "NIT"
columna_razon = "RAZON SOCIAL"
columna_fecha = "FECHA"
carpeta_salida = "EXCEL_POR_NIT"

os.makedirs(carpeta_salida, exist_ok=True)

# ==========================================================
# 📖 LEER EXCEL según extensión
# ==========================================================
try:
    if extension == ".xlsb":
        df = pd.read_excel(nombre_archivo, engine="pyxlsb")  # ✅ Motor para .xlsb
    else:
        df = pd.read_excel(nombre_archivo)                   # ✅ Motor por defecto para .xlsx
except ModuleNotFoundError:
    print("❌ Falta instalar la librería: pip install pyxlsb")
    sys.exit()

# ✅ Limpiar espacios en nombres de columna
df.columns = df.columns.str.strip()

# Validar columnas
for col in [columna_nit, columna_razon, columna_fecha]:
    if col not in df.columns:
        raise ValueError(f"❌ La columna '{col}' no existe en el Excel")

# ✅ Convertir todas las columnas a str EXCEPTO la de fecha
cols_a_str = [c for c in df.columns if c != columna_fecha]
df[cols_a_str] = df[cols_a_str].astype(str)

# ==========================================================
# 🧹 FUNCIONES DE LIMPIEZA
# ==========================================================
def limpiar_nit(nit):
    return re.sub(r"[^\d]", "", str(nit))

def limpiar_texto(texto):
    texto = str(texto)
    texto = re.sub(r'[\\/:*?"<>|]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

# ==========================================================
# 📄 CREAR UN EXCEL POR CADA NIT
# ==========================================================
for nit, df_nit in df.groupby(columna_nit):
    nit_limpio = limpiar_nit(nit)

    nombre_salida = f"{nit_limpio}_INDICADORES DE GESTION.xlsx"
    ruta_salida = os.path.join(carpeta_salida, nombre_salida)

    df_nit.to_excel(ruta_salida, index=False)

    print(f"✔ Archivo creado: {nombre_salida}")

print("✅ Proceso finalizado correctamente")
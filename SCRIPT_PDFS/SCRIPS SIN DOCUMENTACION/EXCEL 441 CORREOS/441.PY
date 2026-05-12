import pandas as pd
import os
import sys
import re

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
# 📂 CONFIGURACIÓN SEGÚN TU EXCEL
# ==========================================================
columna_nit = "NIT"
columna_razon = "RAZON SOCIAL"
carpeta_salida = "EXCEL_POR_NIT"

os.makedirs(carpeta_salida, exist_ok=True)

# ==========================================================
# 📖 LEER EXCEL
# ==========================================================
df = pd.read_excel(nombre_archivo, dtype=str)

# Validar columnas
for col in [columna_nit, columna_razon]:
    if col not in df.columns:
        raise ValueError(f"❌ La columna '{col}' no existe en el Excel")

# ==========================================================
# 🧹 FUNCIONES DE LIMPIEZA
# ==========================================================
def limpiar_nit(nit):
    return re.sub(r"[^\d]", "", str(nit))  # quita comas y cualquier cosa que no sea número

def limpiar_texto(texto):
    texto = str(texto)
    texto = re.sub(r'[\\/:*?"<>|]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

# ==========================================================
# 📄 CREAR UN EXCEL POR CADA NIT
# ==========================================================
for nit, df_nit in df.groupby(columna_nit):
    razon_social = df_nit[columna_razon].iloc[0]

    nit_limpio = limpiar_nit(nit)
    razon_limpia = limpiar_texto(razon_social)
    INDICADORES_DE_GESTION = "INDICADORES DE GESTION"


    nombre_salida = f"{nit_limpio}_{INDICADORES_DE_GESTION}.xlsx"
    ruta_salida = os.path.join(carpeta_salida, nombre_salida)

    df_nit.to_excel(ruta_salida, index=False)

    print(f"✔ Archivo creado: {nombre_salida}")

print("✅ Proceso finalizado correctamente")

import os
import re
import unicodedata
import pandas as pd

# ==================================================
# CONFIGURACIÓN
# ==================================================
CARPETA_EXCEL = r"C:\Users\Trabajo\Desktop\SCRIPT PDFS\ALGORITMO DE BUSQUEDA\PASO 5 estandarizacion\raiz\EVALUACION_DE_INFORME"
SALIDA_REPORTE = "reporte.xlsx"

TEXTO_FINAL = "EVALUACION DE OBLIGACIONES CONTRACTUALES"

# ==================================================
# FUNCIONES
# ==================================================
def quitar_tildes(texto):
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

def extraer_contrato(nombre):
    match = re.search(r"\d{4}-\d{4}", nombre)
    return match.group(0) if match else None

# ==================================================
# PROCESO
# ==================================================
registros = []

for archivo in os.listdir(CARPETA_EXCEL):

    if not archivo.lower().endswith((".xlsx", ".xls", ".csv")):
        continue

    ruta_original = os.path.join(CARPETA_EXCEL, archivo)

    nombre_limpio = quitar_tildes(archivo).upper()

    contrato = extraer_contrato(nombre_limpio)

    if not contrato:
        registros.append({
            "Archivo Original": archivo,
            "Archivo Nuevo": "",
            "Estado": "❌ NO SE IDENTIFICO CONTRATO"
        })
        continue

    nuevo_nombre = f"{contrato} {TEXTO_FINAL}.xlsx"
    ruta_nueva = os.path.join(CARPETA_EXCEL, nuevo_nombre)

    # Evitar sobrescribir
    contador = 1
    while os.path.exists(ruta_nueva):
        nuevo_nombre = f"{contrato} {TEXTO_FINAL} ({contador}).xlsx"
        ruta_nueva = os.path.join(CARPETA_EXCEL, nuevo_nombre)
        contador += 1

    os.rename(ruta_original, ruta_nueva)

    registros.append({
        "Archivo Original": archivo,
        "Archivo Nuevo": nuevo_nombre,
        "Estado": "✅ RENOMBRADO"
    })

# ==================================================
# REPORTE
# ==================================================
df = pd.DataFrame(registros)
df.to_excel(SALIDA_REPORTE, index=False)

print("🎉 Proceso finalizado")
print(f"📄 Reporte generado: {SALIDA_REPORTE}")
print(f"📊 Total procesados: {len(df)}")

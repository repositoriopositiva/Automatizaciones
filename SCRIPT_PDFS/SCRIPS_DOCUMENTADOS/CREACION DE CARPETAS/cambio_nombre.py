import os
import re

# ------------------------------------------------------
# ⚙️ CONFIGURACIÓN - Cambia estos valores si necesitas
# ------------------------------------------------------
SUFIJO = "FEBRERO - 2026"

CARPETAS = [
    "TECNICO",
    "OBLIGACIONES"
]

# ------------------------------------------------------
# 🔄 RENOMBRAR ARCHIVOS
# ------------------------------------------------------
carpeta_base = os.getcwd()

for nombre_carpeta in CARPETAS:
    ruta_carpeta = os.path.join(carpeta_base, nombre_carpeta)

    if not os.path.exists(ruta_carpeta):
        print(f"⚠️ Carpeta no encontrada: {nombre_carpeta}")
        continue

    print(f"\n📂 Procesando: {nombre_carpeta}")
    archivos = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith(".xlsx")]

    if not archivos:
        print("   ⚠️ No se encontraron archivos .xlsx")
        continue

    for archivo in archivos:
        nombre_sin_ext = os.path.splitext(archivo)[0]
        nuevo_nombre   = f"{nombre_sin_ext} {SUFIJO}.xlsx"

        ruta_original = os.path.join(ruta_carpeta, archivo)
        ruta_nueva    = os.path.join(ruta_carpeta, nuevo_nombre)

        # Evitar renombrar si ya tiene el sufijo
        if SUFIJO in nombre_sin_ext:
            print(f"   ⏭️ Ya tiene sufijo, omitido: {archivo}")
            continue

        os.rename(ruta_original, ruta_nueva)
        print(f"   ✅ {archivo}")
        print(f"      → {nuevo_nombre}")

print("\n🎉 Renombrado finalizado correctamente")
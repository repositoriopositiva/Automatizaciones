import os
import pandas as pd

# ------------------------------------------------------
# 📍 CONFIGURACIÓN
# ------------------------------------------------------
RUTA_RAIZ = r"C:\SCRIPT PDFS\ALGORITMO DE BUSQUEDA\PASO 1 eliminacion de pdf\raiz"
EXT_EXCEL = (".xlsx", ".xls", ".xlsm", ".csv")
EXT_PDF = ".pdf",".word"



reporte = []

# ------------------------------------------------------
# 🔍 RECORRIDO DE CARPETAS
# ------------------------------------------------------
for carpeta_padre in os.listdir(RUTA_RAIZ):
    ruta_padre = os.path.join(RUTA_RAIZ, carpeta_padre)

    if not os.path.isdir(ruta_padre):
        continue

    # Recorre carpetas hijo (INFORMES)
    for carpeta_hijo in os.listdir(ruta_padre):
        ruta_hijo = os.path.join(ruta_padre, carpeta_hijo)

        if not os.path.isdir(ruta_hijo):
            continue

        excels_encontrados = []

        for archivo in os.listdir(ruta_hijo):
            ruta_archivo = os.path.join(ruta_hijo, archivo)

            extension = os.path.splitext(archivo)[1].lower()

            # ✅ Registrar Excels
            if extension in EXT_EXCEL:
                excels_encontrados.append(archivo)

            # ❌ Eliminar cualquier otro archivo
            else:
                try:
                    os.remove(ruta_archivo)
                except Exception as e:
                    print(f"⚠️ No se pudo eliminar {ruta_archivo}: {e}")

        # Registrar en reporte
        if excels_encontrados:
            for excel in excels_encontrados:
                reporte.append({
                    "Carpeta Padre": carpeta_padre,
                    "Carpeta Informe": carpeta_hijo,
                    "Archivo Excel": excel
                })
        else:
            # Por si una carpeta quedó sin Excel
            reporte.append({
                "Carpeta Padre": carpeta_padre,
                "Carpeta Informe": carpeta_hijo,
                "Archivo Excel": "NO TIENE EXCEL"
            })

# ------------------------------------------------------
# 📊 GENERAR REPORTE
# ------------------------------------------------------
df = pd.DataFrame(reporte)
df.to_excel("reporte_excels_por_carpeta.xlsx", index=False)

print("\n✅ Proceso finalizado")
print("📄 Reporte generado: reporte_excels_por_carpeta.xlsx")

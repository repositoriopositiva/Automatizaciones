import os
import shutil
import unicodedata
import pandas as pd

# ======================================================
# 📂 CONFIGURACIÓN
# ======================================================
RUTA_RAIZ = r"C:\SCRIPT PDFS\ALGORITMO DE BUSQUEDA\PASO 2 organizacion de archivos\raiz"



CARPETA_TECNICOS = os.path.join(RUTA_RAIZ, "INFORMES_TECNICOS")
CARPETA_EVALUACION = os.path.join(RUTA_RAIZ, "EVALUACION_DE_INFORME")

EXT_EXCEL = (".xlsx", ".xls", ".csv")

reporte = []

# ======================================================
# 🏗 CREAR CARPETAS DESTINO (FORZADO)
# ======================================================
os.makedirs(CARPETA_TECNICOS, exist_ok=True)
os.makedirs(CARPETA_EVALUACION, exist_ok=True)

print("📁 Carpetas destino verificadas/creadas")

# ======================================================
# 🔧 FUNCIONES
# ======================================================
def quitar_tildes(texto):
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")

def normalizar(texto):
    return quitar_tildes(texto).lower()

# ======================================================
# 🔍 RECORRER CARPETAS PADRE E HIJO
# ======================================================
for carpeta_padre in os.listdir(RUTA_RAIZ):
    ruta_padre = os.path.join(RUTA_RAIZ, carpeta_padre)

    # ❌ No entrar a carpetas destino
    if carpeta_padre in ["INFORMES_TECNICOS", "EVALUACION_DE_INFORME"]:
        continue

    if not os.path.isdir(ruta_padre):
        continue

    for carpeta_hijo in os.listdir(ruta_padre):
        ruta_hijo = os.path.join(ruta_padre, carpeta_hijo)

        if not os.path.isdir(ruta_hijo):
            continue

        for archivo in os.listdir(ruta_hijo):
            if not archivo.lower().endswith(EXT_EXCEL):
                continue

            ruta_archivo = os.path.join(ruta_hijo, archivo)
            nombre_norm = normalizar(archivo)
            print("🔎 Analizando:", ruta_archivo)
            nombre_norm = normalizar(archivo)

            # 🟩 EVALUACION DE INFORME (PRIMERO)
            if (
                "evaluacion" in nombre_norm
                or "evaluaccion" in nombre_norm
                or "evaluiacion" in nombre_norm
                or "obligaciones contractuales" in nombre_norm
                or "obligaciones" in nombre_norm
            ):
                shutil.move(
                    ruta_archivo,
                    os.path.join(CARPETA_EVALUACION, archivo)
                )

            # 🟦 INFORMES TECNICOS (SEGUNDO)
            elif (
                "informe tecnico" in nombre_norm
                or "informe tecnco" in nombre_norm
                or "informe de interventoria" in nombre_norm
                or "informe interventoria" in nombre_norm
                or nombre_norm.startswith("informe")
            ):
                shutil.move(
                    ruta_archivo,
                    os.path.join(CARPETA_TECNICOS, archivo)
                )

            # ❓ NO CLASIFICADO
            else:
                reporte.append({
                    "Carpeta Padre": carpeta_padre,
                    "Carpeta Informe": carpeta_hijo,
                    "Archivo": archivo,
                    "Motivo": "No coincide con ningún patrón definido"
                })


# ======================================================
# 📊 GENERAR REPORTE EN EXCEL
# ======================================================
df = pd.DataFrame(reporte)
df.to_excel("reporte_excels_no_clasificados.xlsx", index=False)

print("✅ Proceso finalizado correctamente")
print("📄 Reporte generado: reporte_excels_no_clasificados.xlsx")

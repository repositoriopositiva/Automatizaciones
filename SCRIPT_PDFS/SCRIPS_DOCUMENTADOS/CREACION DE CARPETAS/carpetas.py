import os
import sys
import pandas as pd
import shutil
import re

# ------------------------------------------------------
# 🔍 DETECTAR AUTOMÁTICAMENTE EL ARCHIVO EXCEL O CSV
# ------------------------------------------------------
archivos_excel = [f for f in os.listdir() if f.lower().endswith((".xlsx", ".csv"))]

if not archivos_excel:
    print("❌ No se encontró ningún archivo .xlsx o .csv en la carpeta.")
    sys.exit()

if len(archivos_excel) > 1:
    print("⚠️ Se encontraron varios archivos:")
    for i, f in enumerate(archivos_excel, 1):
        print(f"   {i}. {f}")
    print("➡ Usando el primero automáticamente.")

nombre_archivo = archivos_excel[0]
print(f"📄 Archivo detectado: {nombre_archivo}")

# ------------------------------------------------------
# 📊 LEER ARCHIVO
# ------------------------------------------------------
if nombre_archivo.lower().endswith(".csv"):
    df = pd.read_csv(nombre_archivo)
else:
    df = pd.read_excel(nombre_archivo)

df = df.dropna(how="all")

# ------------------------------------------------------
# 📆 MAPA DE MESES
# ------------------------------------------------------
meses = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# ------------------------------------------------------
# 🧠 FUNCIÓN NOMBRE CARPETA
# ------------------------------------------------------
def crear_nombre_carpeta(row):
    try:
        responsable = str(row.iloc[0]).strip().upper()
        contrato = str(row.iloc[1]).strip()
        mes_actual = int(row.iloc[2])
        mes_creacion = int(row.iloc[3])
        anio = int(row.iloc[4])

        numero_informe = (mes_actual - mes_creacion) + 1
        mes_texto = meses.get(mes_actual, "MES_INVALIDO")

        return f"INFORME {numero_informe:02d} {mes_texto} {anio} {contrato}"

    except:
        return "ERROR_EN_FILA"

df["NOMBRE_CARPETA"] = df.apply(crear_nombre_carpeta, axis=1)

# ------------------------------------------------------
# 📦 INDEXAR EXCELS DE LOS REPOSITORIOS
# ------------------------------------------------------
carpeta_base = os.getcwd()

repo_david = os.path.join(carpeta_base, "TECNICOS")
repo_tecnico = os.path.join(carpeta_base, "OBLIGACIONES")

repositorio_excels = {}

def indexar_excels(carpeta):
    if not os.path.exists(carpeta):
        return
    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith(".xlsx"):
            match = re.search(r"\d{4}-\d{4}", archivo)
            if match:
                clave = match.group()
                repositorio_excels.setdefault(clave, []).append(
                    os.path.join(carpeta, archivo)
                )

indexar_excels(repo_david)
indexar_excels(repo_tecnico)

# ------------------------------------------------------
# 📁 CREAR CARPETA RAÍZ ANALISTAS
# ------------------------------------------------------
ruta_analistas = os.path.join(carpeta_base, "ANALISTAS")
os.makedirs(ruta_analistas, exist_ok=True)

# ------------------------------------------------------
# 📁 CREAR CARPETAS Y COPIAR EXCELS
# ------------------------------------------------------
for index, row in df.iterrows():
    responsable = str(row.iloc[0]).strip().upper()
    nombre_carpeta = row["NOMBRE_CARPETA"]

    if nombre_carpeta == "ERROR_EN_FILA":
        print(f"⚠️ Fila {index + 2} omitida por error")
        continue

    # ANALISTAS / RESPONSABLE
    ruta_responsable = os.path.join(ruta_analistas, responsable)
    os.makedirs(ruta_responsable, exist_ok=True)

    # ANALISTAS / RESPONSABLE / INFORME
    ruta_informe = os.path.join(ruta_responsable, nombre_carpeta)
    os.makedirs(ruta_informe, exist_ok=True)

    print(f"✅ Carpeta creada: {ruta_informe}")

    # 🔗 Enlace por ZZZZ-YYYY
    match = re.search(r"\d{4}-\d{4}", nombre_carpeta)
    if match:
        clave = match.group()
        if clave in repositorio_excels:
            for excel in repositorio_excels[clave]:
                shutil.copy(excel, ruta_informe)
                print(f"   📎 Excel copiado: {os.path.basename(excel)}")
        else:
            print(f"   ⚠️ No se encontraron Excel para {clave}")

# ------------------------------------------------------
# 💾 GUARDAR RESULTADO
# ------------------------------------------------------
df.to_excel("resultado_prueba_carpetas.xlsx", index=False)

print("\n🎉 Proceso finalizado correctamente")
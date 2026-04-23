import os
import pandas as pd

# 📁 Carpeta a leer (usa "." para la carpeta actual)
ruta_carpeta = "."

# 📄 Obtener lista de archivos (solo archivos, no carpetas)
archivos = [
    f for f in os.listdir(ruta_carpeta)
    if os.path.isfile(os.path.join(ruta_carpeta, f))
]

# 📊 Crear DataFrame
df = pd.DataFrame(archivos, columns=["Nombre del archivo"])

# 📁 Guardar en Excel
nombre_excel = "lista_archivos.xlsx"
df.to_excel(nombre_excel, index=False)

print(f"✅ Archivo '{nombre_excel}' creado correctamente con {len(archivos)} archivos.")

import os

carpeta = r"C:\Usuariios\SCRIPT PDFS\SCRIPS SIN DOCUMENTACION\CAMBIO DE NOMBRE\archivos"

archivos = [f for f in os.listdir(carpeta) if f.endswith(".pdf")]

if not archivos:
    print("No se encontraron archivos PDF en la carpeta.")
else:
    total = len(archivos)
    print("Se encontraron " + str(total) + " archivos. Renombrando...\n")
    for archivo in archivos:
        nombre_sin_ext = os.path.splitext(archivo)[0]
        nuevo_nombre = nombre_sin_ext + "-REPS.pdf"
        ruta_original = os.path.join(carpeta, archivo)
        ruta_nueva = os.path.join(carpeta, nuevo_nombre)
        os.rename(ruta_original, ruta_nueva)
        print("  " + archivo + "  ->  " + nuevo_nombre)

    print("\n" + str(total) + " archivos renombrados correctamente.")
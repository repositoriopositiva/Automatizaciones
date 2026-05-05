"""
excel_a_pdf.py
Convierte archivos Excel (.xlsx / .xls) a PDF manteniendo el nombre original.
Requiere: LibreOffice instalado en el sistema.

Uso:
    python excel_a_pdf.py                        # convierte todos los .xlsx/.xls de la carpeta actual
    python excel_a_pdf.py ruta/carpeta           # convierte todos los Excel de esa carpeta
    python excel_a_pdf.py archivo1.xlsx archivo2.xlsx   # convierte archivos específicos
"""

import sys
import subprocess
import shutil
from pathlib import Path


def libreoffice_disponible() -> bool:
    """Verifica que LibreOffice esté instalado."""
    return shutil.which("libreoffice") is not None or shutil.which("soffice") is not None


def obtener_comando_libreoffice() -> str:
    """Retorna el comando disponible para LibreOffice."""
    if shutil.which("libreoffice"):
        return "libreoffice"
    if shutil.which("soffice"):
        return "soffice"
    raise EnvironmentError(
        "LibreOffice no está instalado o no está en el PATH.\n"
        "Instálalo con:\n"
        "  macOS:   brew install --cask libreoffice\n"
        "  Ubuntu:  sudo apt install libreoffice\n"
        "  Windows: https://www.libreoffice.org/download/download-libreoffice/"
    )


def convertir_excel_a_pdf(ruta_excel: Path, carpeta_salida: Path) -> Path:
    """
    Convierte un archivo Excel a PDF usando LibreOffice.
    El PDF se guarda en carpeta_salida con el mismo nombre base.
    Retorna la ruta del PDF generado.
    """
    comando = obtener_comando_libreoffice()

    result = subprocess.run(
        [
            comando,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(carpeta_salida),
            str(ruta_excel),
        ],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"Error al convertir '{ruta_excel.name}':\n{result.stderr.strip()}"
        )

    pdf_generado = carpeta_salida / (ruta_excel.stem + ".pdf")
    if not pdf_generado.exists():
        raise FileNotFoundError(
            f"LibreOffice no generó el PDF esperado en: {pdf_generado}"
        )

    return pdf_generado


def recopilar_archivos_excel(argumentos: list[str]) -> list[Path]:
    """
    Determina qué archivos Excel procesar según los argumentos recibidos:
    - Sin argumentos      → todos los Excel de la carpeta actual
    - Una carpeta         → todos los Excel de esa carpeta
    - Archivos específicos → esos archivos
    """
    extensiones = {".xlsx", ".xls", ".xlsm", ".xlsb"}

    if not argumentos:
        carpeta = Path(".")
        archivos = sorted(p for p in carpeta.iterdir() if p.suffix.lower() in extensiones)
        if not archivos:
            print("No se encontraron archivos Excel en la carpeta actual.")
        return archivos

    if len(argumentos) == 1:
        ruta = Path(argumentos[0])
        if ruta.is_dir():
            archivos = sorted(p for p in ruta.iterdir() if p.suffix.lower() in extensiones)
            if not archivos:
                print(f"No se encontraron archivos Excel en: {ruta}")
            return archivos

    # Uno o más archivos explícitos
    archivos = []
    for arg in argumentos:
        ruta = Path(arg)
        if not ruta.exists():
            print(f"  [ADVERTENCIA] No se encontró el archivo: {ruta}")
        elif ruta.suffix.lower() not in extensiones:
            print(f"  [ADVERTENCIA] No es un archivo Excel: {ruta}")
        else:
            archivos.append(ruta)
    return archivos


def main():
    if not libreoffice_disponible():
        print(
            "ERROR: LibreOffice no está instalado.\n"
            "Instálalo con:\n"
            "  macOS:   brew install --cask libreoffice\n"
            "  Ubuntu:  sudo apt install libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/download-libreoffice/"
        )
        sys.exit(1)

    archivos = recopilar_archivos_excel(sys.argv[1:])

    if not archivos:
        sys.exit(0)

    print(f"\n{'='*50}")
    print(f"  Archivos a convertir: {len(archivos)}")
    print(f"{'='*50}\n")

    exitosos = []
    fallidos = []

    for ruta_excel in archivos:
        carpeta_salida = ruta_excel.parent  # PDF queda junto al Excel
        print(f"  Convirtiendo: {ruta_excel.name} ...", end=" ", flush=True)
        try:
            pdf = convertir_excel_a_pdf(ruta_excel, carpeta_salida)
            print(f"OK  →  {pdf.name}")
            exitosos.append(pdf)
        except Exception as e:
            print(f"ERROR\n    {e}")
            fallidos.append(ruta_excel.name)

    print(f"\n{'='*50}")
    print(f"  Resultado: {len(exitosos)} convertidos, {len(fallidos)} fallidos")
    if fallidos:
        print("  Fallidos:")
        for nombre in fallidos:
            print(f"    - {nombre}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
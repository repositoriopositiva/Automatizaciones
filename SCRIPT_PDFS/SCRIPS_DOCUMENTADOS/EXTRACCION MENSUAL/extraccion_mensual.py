"""
SCRIPT 2 - extraccion_mensual.py
Procesa la carpeta ANALISTAS (estructura: ANALISTAS > RESPONSABLE > INFORME XX MES AÑO CONTRATO):
  1. Recorre todas las subcarpetas de informe dentro de ANALISTAS
  2. Mueve Excel de informes técnicos → INFORMES TECNICOS ENERO
  3. Mueve Excel de obligaciones      → OBLIGACIONES CONTRACTUALES ENERO
"""

import shutil
from pathlib import Path
import re
import unicodedata

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
BASE_DIR             = Path(__file__).parent
CARPETA_ANALISTAS    = BASE_DIR / "ANALISTAS"
CARPETA_INF_TEC      = BASE_DIR / "INFORMES TECNICOS ENERO"
CARPETA_OBLIGACIONES = BASE_DIR / "EVALUACION OBLIGACIONES CONTRACTUALES ENERO"

# CAMBIAR MANUALMENTE CADA MES
MES_EVALUADO = "MARZO 2026"


# ─────────────────────────────────────────────
# FUNCIONES AUXILIARES
# ─────────────────────────────────────────────

def normalizar(texto: str) -> str:
    """
    Quita tildes, espacios extra y pasa a minúsculas
    """
    texto = str(texto).strip().lower()

    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ascii", "ignore").decode("ascii")

    return texto


def obtener_contrato(nombre_carpeta: str):
    """
    Extrae número de contrato tipo 0106-2023
    """
    patron = r"\d{3,5}-\d{4}"
    match = re.search(patron, nombre_carpeta)
    if match:
        return match.group()
    return "CONTRATO_DESCONOCIDO"


def mover_renombrando(origen: Path, carpeta_destino: Path, nuevo_nombre: str):
    """
    Mueve archivo renombrándolo
    """
    destino = carpeta_destino / nuevo_nombre

    if destino.exists():
        destino = carpeta_destino / f"{destino.stem}_dup{destino.suffix}"

    shutil.move(str(origen), str(destino))
    return destino


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  SCRIPT 2 - Extracción desde ANALISTAS")
    print("=" * 60)

    if not CARPETA_ANALISTAS.exists():
        print(f"\n❌ No se encontró la carpeta: {CARPETA_ANALISTAS}")
        return

    CARPETA_INF_TEC.mkdir(exist_ok=True)
    CARPETA_OBLIGACIONES.mkdir(exist_ok=True)

    print("\nClasificando Excel desde carpetas de ANALISTAS...")

    inf_tec_movidos = []
    oblig_movidos   = []
    sin_clasificar  = []

    # Nivel 1: responsables
    for responsable_dir in sorted(CARPETA_ANALISTAS.iterdir()):
        if not responsable_dir.is_dir():
            continue

        # Nivel 2: carpetas de informe
        for informe_dir in sorted(responsable_dir.iterdir()):
            if not informe_dir.is_dir():
                continue

            contrato = obtener_contrato(informe_dir.name)

            excels = list(informe_dir.glob("*.xlsx")) + list(informe_dir.glob("*.xls"))

            for archivo in excels:
                if not archivo.is_file():
                    continue

                nombre_norm = normalizar(archivo.name)

                # ── INFORME TECNICO ─────────────────────
                if "informe" in nombre_norm and "tecnico" in nombre_norm:

                    nuevo_nombre = f"{contrato} - INFORME TECNICO - {MES_EVALUADO}{archivo.suffix}"

                    mover_renombrando(
                        archivo,
                        CARPETA_INF_TEC,
                        nuevo_nombre
                    )

                    inf_tec_movidos.append(nuevo_nombre)
                    print(f"  ✔ Téc  [{responsable_dir.name}] {nuevo_nombre}")

                # ── OBLIGACIONES CONTRACTUALES ───────────
                elif ("obligacion" in nombre_norm or
                      "obligaciones" in nombre_norm):

                    nuevo_nombre = f"{contrato} - EVALUACION DE OBLIGACIONES CONTRACTUALES - {MES_EVALUADO}{archivo.suffix}"

                    mover_renombrando(
                        archivo,
                        CARPETA_OBLIGACIONES,
                        nuevo_nombre
                    )

                    oblig_movidos.append(nuevo_nombre)
                    print(f"  ✔ Obl  [{responsable_dir.name}] {nuevo_nombre}")

                else:
                    sin_clasificar.append(archivo.name)
                    print(f"  ⚠ Sin clasificar [{responsable_dir.name}]: {archivo.name}")

    # ── RESUMEN ─────────────────────────────
    print("\n" + "=" * 60)
    print("  RESUMEN FINAL")
    print("=" * 60)

    print(f"  📁 INFORMES TECNICOS ENERO           : {len(inf_tec_movidos)} archivos")
    print(f"  📁 OBLIGACIONES CONTRACTUALES ENERO  : {len(oblig_movidos)} archivos")

    if sin_clasificar:
        print(f"\n  ⚠ Sin clasificar ({len(sin_clasificar)}):")
        for f in sin_clasificar:
            print(f"     - {f}")

    print("\n  ✅ Script 2 finalizado.\n")


if __name__ == "__main__":
    main()
    
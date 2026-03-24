"""
SCRIPT 2 - extraccion_mensual.py
Procesa la carpeta ANALISTAS (estructura: ANALISTAS > RESPONSABLE > INFORME XX MES AÑO CONTRATO):
  1. Recorre todas las subcarpetas de informe dentro de ANALISTAS
  2. Mueve Excel de informes técnicos → INFORMES TECNICOS ENERO
  3. Mueve Excel de obligaciones      → OBLIGACIONES CONTRACTUALES ENERO

Ubicación esperada (todo en la misma carpeta):
  extraccion_mensual.py
  ANALISTAS/  (carpeta con subcarpetas de analistas y contratos)
"""

import shutil
from pathlib import Path

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
BASE_DIR             = Path(__file__).parent
CARPETA_ANALISTAS    = BASE_DIR / "ANALISTAS"
CARPETA_INF_TEC      = BASE_DIR / "INFORMES TECNICOS ENERO"
CARPETA_OBLIGACIONES = BASE_DIR / "OBLIGACIONES CONTRACTUALES ENERO"

# ─────────────────────────────────────────────
# FUNCIÓN AUXILIAR
# ─────────────────────────────────────────────

def normalizar(texto: str) -> str:
    return str(texto).strip().lower()


def mover_sin_colision(origen: Path, carpeta_destino: Path) -> Path:
    """Mueve archivo; si ya existe en destino agrega sufijo _dup."""
    destino = carpeta_destino / origen.name
    if destino.exists():
        destino = carpeta_destino / f"{origen.stem}_dup{origen.suffix}"
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

    # Crear carpetas destino si no existen
    CARPETA_INF_TEC.mkdir(exist_ok=True)
    CARPETA_OBLIGACIONES.mkdir(exist_ok=True)

    print("\nClasificando Excel desde carpetas de ANALISTAS...")

    inf_tec_movidos = []
    oblig_movidos   = []
    sin_clasificar  = []

    # Nivel 1: responsables (ANDRES, MAURICIO, etc.)
    for responsable_dir in sorted(CARPETA_ANALISTAS.iterdir()):
        if not responsable_dir.is_dir():
            continue

        # Nivel 2: carpetas de informe (INFORME 02 FEBRERO 2026 0106-2023)
        for informe_dir in sorted(responsable_dir.iterdir()):
            if not informe_dir.is_dir():
                continue

            # Nivel 3: archivos Excel dentro de cada carpeta de informe
            excels = list(informe_dir.glob("*.xlsx")) + list(informe_dir.glob("*.xls"))

            for archivo in excels:
                if not archivo.is_file():
                    continue
                nombre_norm = normalizar(archivo.name)

                if "informe tecnico" in nombre_norm or "informe técnico" in nombre_norm:
                    mover_sin_colision(archivo, CARPETA_INF_TEC)
                    inf_tec_movidos.append(archivo.name)
                    print(f"  ✔ Téc  [{responsable_dir.name}] {archivo.name}")

                elif ("obligaciones contractuales" in nombre_norm or
                      "ogligaciones contractuales" in nombre_norm):
                    mover_sin_colision(archivo, CARPETA_OBLIGACIONES)
                    oblig_movidos.append(archivo.name)
                    print(f"  ✔ Obl  [{responsable_dir.name}] {archivo.name}")

                else:
                    sin_clasificar.append(archivo.name)
                    print(f"  ⚠ Sin clasificar [{responsable_dir.name}]: {archivo.name}")

    # ── Resumen final ──────────────────────────────────────
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
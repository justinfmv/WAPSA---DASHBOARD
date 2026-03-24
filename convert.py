"""
convert.py
----------
Convierte el archivo exportado de Azure DevOps ("WAPSA Team - Epics *.xlsx")
al formato que consume el dashboard (data.xlsx, hoja "Hoja2").

Uso:
    python convert.py                          # usa el archivo fuente por defecto
    python convert.py "WAPSA Team - Epics (6).xlsx"   # especifica otro archivo
"""

import sys
import re
from pathlib import Path
import pandas as pd

# ── Configuración ──────────────────────────────────────────────────────────────
DEFAULT_SOURCE = "WAPSA Team - Epics (5).xlsx"
OUTPUT_FILE    = "data.xlsx"
OUTPUT_SHEET   = "Hoja2"

# Columnas que debe tener Hoja2 (en este orden)
TARGET_COLS = [
    "ID", "Work Item Type", "Title", "Iteration Path", "Estados", "Tags",
    "Start Date", "Finish Date", "Finish Date Chinalco", "Aplicación",
    "Ticket Chinalco", "Ticket SyC", "Assigned To", "Description",
    "Tipo Caso", "Tipo Evento", "Tipo Estabilización", "Parent", "Epic",
]
# ──────────────────────────────────────────────────────────────────────────────


def detect_source() -> Path:
    """Si se pasa un argumento lo usa; si no, busca el xlsx más reciente con 'Epics' en el nombre."""
    if len(sys.argv) > 1:
        p = Path(sys.argv[1])
        if not p.exists():
            sys.exit(f"❌ Archivo no encontrado: {p}")
        return p

    # Buscar automáticamente el archivo más reciente
    candidates = sorted(
        Path(".").glob("WAPSA Team - Epics*.xlsx"),
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    if candidates:
        return candidates[0]

    # Fallback al nombre por defecto
    p = Path(DEFAULT_SOURCE)
    if not p.exists():
        sys.exit(
            f"❌ No se encontró el archivo fuente.\n"
            f"   Asegúrate de que exista '{DEFAULT_SOURCE}' en esta carpeta\n"
            f"   o pásalo como argumento:  python convert.py <archivo.xlsx>"
        )
    return p


def normalize_col_name(name: str) -> str:
    """Quita espacios extra y normaliza acentos comunes para tolerar variaciones de exportación."""
    replacements = {
        "Aplicacion": "Aplicación",
        "Tipo Estabilizacion": "Tipo Estabilización",
    }
    name = name.strip()
    return replacements.get(name, name)


def main():
    source = detect_source()
    print(f"[OK] Leyendo fuente: {source}")

    # Leer todas las hojas del xlsx fuente (suele tener una sola)
    xf = pd.ExcelFile(source)
    df_raw = pd.read_excel(xf, sheet_name=xf.sheet_names[0])

    # Normalizar nombres de columnas
    df_raw.columns = [normalize_col_name(c) for c in df_raw.columns]

    print(f"   Filas totales: {len(df_raw)}")
    print(f"   Columnas: {list(df_raw.columns)}")

    # ── Construir mapa Epic ID → Título ───────────────────────────────────────
    epics = df_raw[df_raw["Work Item Type"] == "Epic"][["ID", "Title"]].copy()
    epic_map: dict = dict(zip(epics["ID"], epics["Title"]))
    print(f"   Épicas encontradas: {len(epic_map)}")

    # ── Filtrar solo Tickets ──────────────────────────────────────────────────
    tickets = df_raw[df_raw["Work Item Type"] == "Ticket"].copy()
    print(f"   Tickets: {len(tickets)}")

    # ── Agregar columna Epic ──────────────────────────────────────────────────
    def resolve_epic(parent_id):
        if pd.isna(parent_id):
            return ""
        return epic_map.get(int(parent_id), "")

    tickets["Epic"] = tickets["Parent"].apply(resolve_epic)

    # ── Seleccionar y ordenar columnas ────────────────────────────────────────
    # Agregar columnas faltantes con valor vacío (por si la exportación varía)
    for col in TARGET_COLS:
        if col not in tickets.columns:
            print(f"   [WARN] Columna no encontrada en fuente, se agrega vacia: '{col}'")
            tickets[col] = ""

    tickets = tickets[TARGET_COLS].reset_index(drop=True)

    # ── Escribir data.xlsx ────────────────────────────────────────────────────
    output = Path(OUTPUT_FILE)
    if output.exists():
        # Preservar otras hojas que pueda tener data.xlsx
        with pd.ExcelWriter(output, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            tickets.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
    else:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            tickets.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)

    print(f"\n[DONE] {OUTPUT_FILE} actualizado -- {len(tickets)} tickets en hoja '{OUTPUT_SHEET}'")
    print(f"   Épicas únicas en los datos: {tickets['Epic'].nunique()}")


if __name__ == "__main__":
    main()

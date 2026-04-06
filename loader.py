# loader.py — Lectura de archivos SAP, detección de formato, validación de columnas

import re
import sys
from pathlib import Path

import chardet
import pandas as pd

from config import MB51_KEY_COLS, PO_KEY_COLS


# ── Detección de encoding ─────────────────────────────────────────────────────

def detect_encoding(filepath: str) -> str:
    """Detecta el encoding del archivo usando chardet. Fallback: cp1252."""
    with open(filepath, "rb") as f:
        raw = f.read(50_000)
    result = chardet.detect(raw)
    encoding = result.get("encoding") or "cp1252"
    confidence = result.get("confidence") or 0
    if confidence < 0.7:
        encoding = "cp1252"
    return encoding


# ── Detección de separador ────────────────────────────────────────────────────

def detect_separator(filepath: str, encoding: str) -> str:
    """Detecta el separador de columnas contando ';', ',', '\\t' en las primeras líneas."""
    try:
        with open(filepath, "r", encoding=encoding, errors="replace") as f:
            sample = "".join(f.readline() for _ in range(5))
    except Exception:
        return ";"
    counts = {
        ";": sample.count(";"),
        ",": sample.count(","),
        "\t": sample.count("\t"),
    }
    return max(counts, key=counts.get)


# ── Lectura del archivo SAP ───────────────────────────────────────────────────

def read_sap_file(filepath: str, file_label: str = "archivo") -> pd.DataFrame:
    """
    Lee un archivo exportado de SAP (CSV/TXT) de forma robusta:
    - Detecta encoding y separador automáticamente.
    - Lee todo como string para evitar conversiones automáticas.
    - Reintenta con skiprows 0–4 para saltar filas de encabezado SAP.
    Devuelve el DataFrame con columnas tal como aparecen en el archivo.
    """
    filepath = str(filepath)
    if not Path(filepath).exists():
        raise FileNotFoundError(f"No se encontró el archivo: {filepath}")

    encoding = detect_encoding(filepath)
    sep = detect_separator(filepath, encoding)

    last_error = None
    for skip in range(5):
        try:
            df = pd.read_csv(
                filepath,
                sep=sep,
                encoding=encoding,
                dtype=str,
                skiprows=skip,
                on_bad_lines="warn",
            )
            # Limpiar espacios en nombres de columnas
            df.columns = [c.strip() for c in df.columns]
            # Verificar que hay al menos 3 columnas con nombre real
            named_cols = [c for c in df.columns if c and not c.startswith("Unnamed")]
            if len(named_cols) >= 3:
                print(f"  [{file_label}] Leído con encoding={encoding}, sep={repr(sep)}, "
                      f"skiprows={skip}, {len(df)} filas, {len(df.columns)} columnas.")
                return df
        except Exception as e:
            last_error = e
            continue

    raise ValueError(
        f"No se pudo leer '{filepath}' como CSV.\n"
        f"Último error: {last_error}\n"
        "Verifica que el archivo sea un CSV/TXT exportado de SAP."
    )


# ── Resolución de columnas ────────────────────────────────────────────────────

def resolve_columns(df: pd.DataFrame, col_aliases: dict) -> dict:
    """
    Intenta encontrar cada columna lógica buscando los alias en las columnas reales del DataFrame.
    Retorna {nombre_logico: nombre_real_en_df} para las encontradas.
    Las no encontradas se incluyen con valor None.
    """
    actual_cols_lower = {c.strip().lower(): c for c in df.columns}
    mapping = {}
    for logical_name, aliases in col_aliases.items():
        found = None
        for alias in aliases:
            if alias.strip().lower() in actual_cols_lower:
                found = actual_cols_lower[alias.strip().lower()]
                break
        mapping[logical_name] = found
    return mapping


# ── Confirmación interactiva de columnas ──────────────────────────────────────

def confirm_columns(col_map: dict, file_label: str) -> bool:
    """
    Muestra al usuario las columnas detectadas y pide confirmación.
    Retorna True si el usuario confirma, False si cancela.
    """
    print(f"\n{'='*60}")
    print(f"  Columnas detectadas en: {file_label}")
    print(f"{'='*60}")
    print(f"  {'Campo lógico':<30} {'Columna en archivo'}")
    print(f"  {'-'*28} {'-'*28}")
    missing = []
    for logical, actual in col_map.items():
        status = actual if actual else "--- NO ENCONTRADA ---"
        print(f"  {logical:<30} {status}")
        if actual is None:
            missing.append(logical)
    print(f"{'='*60}")

    if missing:
        print(f"\n  ADVERTENCIA: Las siguientes columnas no fueron encontradas:")
        for m in missing:
            print(f"    - {m}")
        print("  El análisis continuará, pero las filas sin estos datos serán marcadas.\n")

    while True:
        resp = input("  ¿Las columnas son correctas? (S/n): ").strip().lower()
        if resp in ("", "s", "si", "sí", "y", "yes"):
            return True
        if resp in ("n", "no"):
            print("  Operación cancelada por el usuario. Verifica el archivo e inténtalo de nuevo.")
            return False
        print("  Por favor responde S (sí) o N (no).")


# ── Marcado de filas con datos clave faltantes ────────────────────────────────

def flag_missing_key_data(df: pd.DataFrame, col_map: dict, key_col_names: list) -> pd.DataFrame:
    """
    Agrega una columna '_advertencia' con texto descriptivo para filas con datos clave ausentes.
    key_col_names: lista de nombres lógicos (claves de col_map) que son obligatorios.
    """
    df = df.copy()
    df["_advertencia"] = ""

    for logical_name in key_col_names:
        actual_col = col_map.get(logical_name)
        if actual_col is None:
            # La columna no existe en el archivo — marcar todas las filas
            mask = pd.Series([True] * len(df), index=df.index)
        else:
            mask = df[actual_col].isna() | (df[actual_col].astype(str).str.strip() == "")

        if mask.any():
            desc = f"Sin {logical_name}"
            df.loc[mask, "_advertencia"] = df.loc[mask, "_advertencia"].apply(
                lambda x: f"{x}; {desc}".lstrip("; ")
            )

    return df


# ── Función auxiliar para leer y validar un archivo completo ─────────────────

def load_and_validate(filepath: str, col_aliases: dict, key_cols: list,
                      file_label: str, interactive: bool = True):
    """
    Wrapper completo: lee el archivo, resuelve columnas, pide confirmación (si interactive=True),
    y marca filas con datos faltantes.
    Retorna (df, col_map) o lanza SystemExit si el usuario cancela.
    """
    df = read_sap_file(filepath, file_label=file_label)
    col_map = resolve_columns(df, col_aliases)

    if interactive:
        ok = confirm_columns(col_map, file_label)
        if not ok:
            sys.exit(0)

    df = flag_missing_key_data(df, col_map, key_cols)
    return df, col_map

# transformer.py — Normalización de datos, cruce MB51 ↔ Pedidos, agregación

import re
import warnings
from typing import Optional

import numpy as np
import pandas as pd

from config import MOV_ENTRY, MOV_CANCEL


# ── Parsing de fechas SAP ─────────────────────────────────────────────────────

DATE_FORMATS = ["%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y"]


def parse_sap_date(series: pd.Series) -> pd.Series:
    """
    Convierte una Serie de strings de fecha SAP a datetime.
    Prueba múltiples formatos; retorna NaT para valores no parseables.
    """
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    for fmt in DATE_FORMATS:
        mask = result.isna() & series.notna() & (series.astype(str).str.strip() != "")
        if not mask.any():
            break
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            parsed = pd.to_datetime(
                series[mask].astype(str).str.strip(),
                format=fmt,
                errors="coerce",
            )
        result[mask] = parsed

    return result


# ── Parsing de decimales españoles ────────────────────────────────────────────

def parse_spanish_decimal(series: pd.Series) -> pd.Series:
    """
    Convierte valores con formato español (1.234,56) a float.
    Elimina puntos de miles primero, luego convierte coma a punto decimal.
    """
    s = series.astype(str).str.strip()
    # Eliminar puntos de miles (solo si hay coma decimal presente)
    s = s.str.replace(r"\.(?=\d{3}([,\s]|$))", "", regex=True)
    # Convertir coma decimal a punto
    s = s.str.replace(",", ".", regex=False)
    # Eliminar caracteres no numéricos salvo punto y signo
    s = s.str.replace(r"[^\d.\-]", "", regex=True)
    s = s.replace("", np.nan)
    return pd.to_numeric(s, errors="coerce")


# ── Normalización de tipo de movimiento ───────────────────────────────────────

def normalize_movement_type(series: pd.Series) -> pd.Series:
    """
    Convierte el tipo de movimiento a entero.
    Maneja tanto '101' como 'WE' (entradas de mercancías en SAP).
    """
    WE_MAP = {
        "WE": 101, "WA": 261, "ST": 301,
        "UM": 309, "WL": 601,
    }
    s = series.astype(str).str.strip()
    # Intentar conversión numérica directa
    numeric = pd.to_numeric(s, errors="coerce")
    # Para los que no son numéricos, buscar en el mapa
    mask_non_numeric = numeric.isna() & (s != "") & (s != "nan")
    if mask_non_numeric.any():
        numeric[mask_non_numeric] = s[mask_non_numeric].map(WE_MAP)
    return numeric


# ── Construcción de clave de cruce ────────────────────────────────────────────

def build_join_key(order_series: pd.Series, pos_series: pd.Series) -> pd.Series:
    """
    Construye la clave de cruce como 'PEDIDO_ZFILL10-POS_ZFILL5'.
    Normaliza ceros a la izquierda para unificar variaciones de exportación SAP.
    """
    order = order_series.astype(str).str.strip().str.zfill(10)
    pos   = pos_series.astype(str).str.strip().str.zfill(5)
    return order + "-" + pos


# ── Agregación de movimientos MB51 ───────────────────────────────────────────

def aggregate_movements(mb51_df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    Agrupa el MB51 por clave de cruce (pedido+posición) y tipo de movimiento.
    Para mov. 101: suma cantidades e importes, toma la última fecha contable, junta usuarios.
    Para mov. 102: suma cantidades.
    Retorna un DataFrame con una fila por (join_key, origen lógico).
    """
    df = mb51_df.copy()

    # Mapear a nombres estándar internos
    def col(logical):
        return col_map.get(logical)

    req = ["order", "position", "movement", "quantity"]
    for r in req:
        if col(r) is None:
            raise ValueError(f"Columna obligatoria '{r}' no encontrada en MB51.")

    # Construir clave de cruce
    df["_join_key"] = build_join_key(df[col("order")], df[col("position")])

    # Normalizar tipos
    df["_mov"] = normalize_movement_type(df[col("movement")])
    df["_qty"] = parse_spanish_decimal(df[col("quantity")])

    if col("accounting_date"):
        df["_acc_date"] = parse_sap_date(df[col("accounting_date")])
    else:
        df["_acc_date"] = pd.NaT

    if col("amount"):
        df["_amount"] = parse_spanish_decimal(df[col("amount")])
    else:
        df["_amount"] = np.nan

    if col("user"):
        df["_user"] = df[col("user")].astype(str).str.strip()
    else:
        df["_user"] = ""

    if col("warehouse"):
        df["_warehouse"] = df[col("warehouse")].astype(str).str.strip()
    else:
        df["_warehouse"] = ""

    if col("doc_date"):
        df["_doc_date_mb51"] = parse_sap_date(df[col("doc_date")])
    else:
        df["_doc_date_mb51"] = pd.NaT

    # Separar 101 y 102
    mov101 = df[df["_mov"] == 101].copy()
    mov102 = df[df["_mov"] == 102].copy()

    # Agregar 101
    agg101 = (
        mov101.groupby("_join_key", as_index=False)
        .agg(
            qty_101=("_qty", "sum"),
            amount_101=("_amount", "sum"),
            last_acc_date=("_acc_date", "max"),
            users_101=("_user", lambda x: ", ".join(sorted(set(x.dropna()) - {"", "nan"}))),
            warehouse=("_warehouse", "first"),
            doc_date_mb51=("_doc_date_mb51", "first"),
        )
    )

    # Agregar 102
    agg102 = (
        mov102.groupby("_join_key", as_index=False)
        .agg(qty_102=("_qty", "sum"))
    )

    # Combinar
    result = agg101.merge(agg102, on="_join_key", how="outer")

    # Filas que solo tienen 102 (cancelaciones sin entrada previa)
    result["qty_101"] = result["qty_101"].fillna(0)
    result["amount_101"] = result["amount_101"].fillna(0)
    result["qty_102"] = result["qty_102"].fillna(0)
    result["users_101"] = result["users_101"].fillna("")
    result["warehouse"] = result["warehouse"].fillna("")

    return result


# ── Cruce MB51 con archivo de pedidos ─────────────────────────────────────────

def merge_mb51_with_po(
    agg_mb51: pd.DataFrame,
    po_df: pd.DataFrame,
    po_col_map: dict,
) -> pd.DataFrame:
    """
    Realiza el join RIGHT entre los movimientos agregados de MB51 y las líneas de pedido.
    Conserva todas las líneas de pedido, incluso sin movimientos.
    """

    def po_col(logical):
        return po_col_map.get(logical)

    # Validar columnas clave del archivo de pedidos
    for req in ["order", "position"]:
        if po_col(req) is None:
            raise ValueError(f"Columna obligatoria '{req}' no encontrada en archivo de pedidos.")

    # Construir clave en el archivo de pedidos
    po_df = po_df.copy()
    po_df["_join_key"] = build_join_key(po_df[po_col("order")], po_df[po_col("position")])

    # Normalizar campos del pedido
    if po_col("quantity"):
        po_df["_qty_ordered"] = parse_spanish_decimal(po_df[po_col("quantity")])
    else:
        po_df["_qty_ordered"] = np.nan

    if po_col("doc_date"):
        po_df["_po_doc_date"] = parse_sap_date(po_df[po_col("doc_date")])
    else:
        po_df["_po_doc_date"] = pd.NaT

    if po_col("description"):
        po_df["_description"] = po_df[po_col("description")].astype(str).str.strip()
    elif po_col("material"):
        po_df["_description"] = po_df[po_col("material")].astype(str).str.strip()
    else:
        po_df["_description"] = ""

    if po_col("supplier"):
        po_df["_supplier"] = po_df[po_col("supplier")].astype(str).str.strip()
    else:
        po_df["_supplier"] = ""

    if po_col("material"):
        po_df["_material_po"] = po_df[po_col("material")].astype(str).str.strip()
    else:
        po_df["_material_po"] = ""

    if po_col("pending"):
        po_df["_qty_pending_sap"] = parse_spanish_decimal(po_df[po_col("pending")])
    else:
        po_df["_qty_pending_sap"] = np.nan

    # JOIN: right join (todas las líneas de pedido)
    merged = po_df.merge(agg_mb51, on="_join_key", how="left")

    # Rellenar nulls de agregación para líneas sin movimientos
    merged["qty_101"]      = merged["qty_101"].fillna(0)
    merged["amount_101"]   = merged["amount_101"].fillna(0)
    merged["qty_102"]      = merged["qty_102"].fillna(0)
    merged["users_101"]    = merged["users_101"].fillna("")
    merged["warehouse"]    = merged["warehouse"].fillna("")
    merged["last_acc_date"] = merged.get("last_acc_date", pd.NaT)

    # Reconstruir Pedido y Posición limpios desde la clave
    merged["_order"]    = merged["_join_key"].str.split("-").str[0].str.lstrip("0")
    merged["_position"] = merged["_join_key"].str.split("-").str[1].str.lstrip("0")

    return merged

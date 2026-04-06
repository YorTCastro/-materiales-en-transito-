# analyzer.py — Métricas de negocio, KPIs, alertas de color, análisis adicionales

import numpy as np
import pandas as pd

from config import (
    ALERT_GREEN, ALERT_YELLOW, ALERT_RED, ALERT_NO_ENTRY, ALERT_NO_DATE,
    GREEN_MAX, YELLOW_MAX,
    ORIGIN_MAP_WAREHOUSE, ORIGIN_KEYWORDS, ORIGIN_UNKNOWN,
    TOP_MATERIALS_N, TOP_USERS_N, TOP_PENDING_N, MAX_RECOMMENDATIONS,
)


# ── Días transcurridos ────────────────────────────────────────────────────────

def compute_days_elapsed(merged_df: pd.DataFrame) -> pd.Series:
    """
    Calcula los días entre la fecha del pedido (_po_doc_date) y la última
    fecha contable de ingreso (last_acc_date).
    Retorna NaN para filas sin entrada o con fechas inválidas.
    """
    has_entry = merged_df["qty_101"] > 0
    has_dates = merged_df["_po_doc_date"].notna() & merged_df["last_acc_date"].notna()

    days = pd.Series(np.nan, index=merged_df.index)
    mask = has_entry & has_dates
    if mask.any():
        delta = (merged_df.loc[mask, "last_acc_date"] - merged_df.loc[mask, "_po_doc_date"])
        days[mask] = delta.dt.days

    return days


# ── Alerta de color ───────────────────────────────────────────────────────────

def assign_color_alert(days_series: pd.Series, qty_101_series: pd.Series,
                       po_date_series: pd.Series, acc_date_series: pd.Series,
                       policy_days: int = None) -> pd.Series:
    """
    Asigna la alerta de color basada en días transcurridos.
    policy_days sobrescribe YELLOW_MAX si se proporciona.
    """
    if policy_days is None:
        policy_days = YELLOW_MAX

    green_max   = GREEN_MAX
    yellow_max  = policy_days

    alert = pd.Series(ALERT_NO_ENTRY, index=days_series.index)

    # Sin fecha válida pero con entrada
    no_date_mask = qty_101_series > 0 & (po_date_series.isna() | acc_date_series.isna())
    alert[no_date_mask] = ALERT_NO_DATE

    # Con días calculados
    has_days = days_series.notna()
    alert[has_days & (days_series <= green_max)]                              = ALERT_GREEN
    alert[has_days & (days_series > green_max) & (days_series <= yellow_max)] = ALERT_YELLOW
    alert[has_days & (days_series > yellow_max)]                              = ALERT_RED

    return alert


# ── Determinación de origen ───────────────────────────────────────────────────

def compute_origin(merged_df: pd.DataFrame) -> pd.Series:
    """
    Determina el origen (La Estrella / Taller / Desconocido) por:
    1. Código de almacén en ORIGIN_MAP_WAREHOUSE
    2. Búsqueda de keywords en el campo _supplier
    """
    origin = pd.Series(ORIGIN_UNKNOWN, index=merged_df.index)

    # Por almacén
    if "warehouse" in merged_df.columns and ORIGIN_MAP_WAREHOUSE:
        mapped = merged_df["warehouse"].map(ORIGIN_MAP_WAREHOUSE)
        origin[mapped.notna()] = mapped[mapped.notna()]

    # Por supplier (keyword)
    if "_supplier" in merged_df.columns:
        supplier_lower = merged_df["_supplier"].astype(str).str.lower()
        for keyword, label in ORIGIN_KEYWORDS.items():
            mask = supplier_lower.str.contains(keyword, na=False) & (origin == ORIGIN_UNKNOWN)
            origin[mask] = label

    return origin


# ── Cantidades pendientes ─────────────────────────────────────────────────────

def compute_pending_qty(qty_ordered: pd.Series, qty_101: pd.Series,
                        qty_102: pd.Series) -> pd.Series:
    """
    pendiente = qty_ordenada − (qty_101 − qty_102), mínimo 0.
    """
    net = qty_101 - qty_102
    pending = qty_ordered - net
    return pending.clip(lower=0)


# ── Importe pendiente estimado ────────────────────────────────────────────────

def estimate_pending_amount(amount_101: pd.Series, qty_101: pd.Series,
                             pending_qty: pd.Series) -> pd.Series:
    """
    Estima el importe pendiente usando el precio unitario implícito.
    precio_unitario = amount_101 / qty_101
    importe_pendiente = precio_unitario * pending_qty
    """
    unit_price = amount_101.where(qty_101 > 0, other=np.nan) / qty_101.replace(0, np.nan)
    return (unit_price * pending_qty).fillna(0)


# ── Construcción del DataFrame de detalle completo ───────────────────────────

def build_detail_df(merged_df: pd.DataFrame, policy_days: int = None) -> pd.DataFrame:
    """
    Combina todos los cálculos y retorna el DataFrame de detalle listo para reportar.
    """
    df = merged_df.copy()

    df["dias_transcurridos"] = compute_days_elapsed(df)
    df["alerta"] = assign_color_alert(
        df["dias_transcurridos"],
        df["qty_101"],
        df["_po_doc_date"],
        df.get("last_acc_date", pd.Series(pd.NaT, index=df.index)),
        policy_days=policy_days,
    )
    df["origen"] = compute_origin(df)
    df["qty_neta"] = (df["qty_101"] - df["qty_102"]).clip(lower=0)
    df["qty_pendiente"] = compute_pending_qty(df["_qty_ordered"], df["qty_101"], df["qty_102"])
    df["importe_pendiente"] = estimate_pending_amount(df["amount_101"], df["qty_101"], df["qty_pendiente"])
    df["tiene_anulacion"] = (df["qty_102"] > 0).map({True: "Sí", False: "No"})

    return df


# ── KPIs ──────────────────────────────────────────────────────────────────────

def compute_kpis(detail_df: pd.DataFrame) -> dict:
    """
    Calcula los 7 KPIs del resumen ejecutivo.
    """
    df = detail_df
    total = len(df)

    # Solo líneas con al menos una entrada (mov. 101)
    with_entry = df[df["qty_101"] > 0]
    n_with_entry = len(with_entry)

    n_green   = (df["alerta"] == ALERT_GREEN).sum()
    n_yellow  = (df["alerta"] == ALERT_YELLOW).sum()
    n_red     = (df["alerta"] == ALERT_RED).sum()
    n_no_entry = (df["alerta"] == ALERT_NO_ENTRY).sum()
    n_no_date  = (df["alerta"] == ALERT_NO_DATE).sum()

    pct_on_time = (n_green + n_yellow) / n_with_entry * 100 if n_with_entry > 0 else 0
    pct_overdue = n_red / n_with_entry * 100 if n_with_entry > 0 else 0

    n_cancellations = (df["qty_102"] > 0).sum()
    pct_cancellations = n_cancellations / total * 100 if total > 0 else 0

    n_partial = ((df["qty_pendiente"] > 0) & (df["qty_101"] > 0)).sum()
    pct_partial = n_partial / n_with_entry * 100 if n_with_entry > 0 else 0

    total_pending_amount = df["importe_pendiente"].sum()

    # Por origen
    origin_kpis = {}
    for origin_val in df["origen"].unique():
        sub = df[(df["origen"] == origin_val) & (df["qty_101"] > 0)]
        if len(sub) == 0:
            origin_kpis[origin_val] = None
            continue
        n_ot = ((sub["alerta"] == ALERT_GREEN) | (sub["alerta"] == ALERT_YELLOW)).sum()
        origin_kpis[origin_val] = n_ot / len(sub) * 100

    return {
        "total_lineas":          total,
        "n_con_entrada":         n_with_entry,
        "n_sin_entrada":         int(n_no_entry),
        "n_verde":               int(n_green),
        "n_amarillo":            int(n_yellow),
        "n_rojo":                int(n_red),
        "pct_oportuno":          round(pct_on_time, 1),
        "pct_vencido":           round(pct_overdue, 1),
        "n_anulaciones":         int(n_cancellations),
        "pct_anulaciones":       round(pct_cancellations, 1),
        "n_parciales":           int(n_partial),
        "pct_parciales":         round(pct_partial, 1),
        "importe_pendiente_total": round(total_pending_amount, 2),
        "por_origen":            origin_kpis,
    }


# ── Análisis A: Top materiales por tiempo promedio ───────────────────────────

def top_materials_by_avg_time(detail_df: pd.DataFrame, n: int = TOP_MATERIALS_N) -> pd.DataFrame:
    """Top N materiales con mayor tiempo promedio de ingreso."""
    df = detail_df[detail_df["dias_transcurridos"].notna()].copy()
    if df.empty:
        return pd.DataFrame(columns=["Material", "Descripción", "Días Promedio", "N Ingresos"])

    result = (
        df.groupby(["_material_po", "_description"], as_index=False)
        .agg(
            dias_promedio=("dias_transcurridos", "mean"),
            n_ingresos=("dias_transcurridos", "count"),
        )
        .sort_values("dias_promedio", ascending=False)
        .head(n)
    )
    result.columns = ["Material", "Descripción", "Días Promedio", "N Ingresos"]
    result["Días Promedio"] = result["Días Promedio"].round(1)
    return result


# ── Análisis B: Usuarios con más ingresos vencidos ────────────────────────────

def top_users_overdue(detail_df: pd.DataFrame, n: int = TOP_USERS_N) -> pd.DataFrame:
    """Top N usuarios con más registros de ingreso ROJO."""
    df = detail_df[detail_df["alerta"] == ALERT_RED].copy()
    if df.empty:
        return pd.DataFrame(columns=["Usuario", "Ingresos Vencidos", "% del Total Vencido"])

    # Expandir usuarios (pueden ser varios separados por coma)
    rows = []
    for _, row in df.iterrows():
        users = [u.strip() for u in str(row["users_101"]).split(",") if u.strip() and u.strip() != "nan"]
        for u in users:
            rows.append(u)

    if not rows:
        return pd.DataFrame(columns=["Usuario", "Ingresos Vencidos", "% del Total Vencido"])

    from collections import Counter
    counts = Counter(rows)
    total_red = sum(counts.values())
    result = pd.DataFrame(
        [(u, c, round(c / total_red * 100, 1)) for u, c in counts.most_common(n)],
        columns=["Usuario", "Ingresos Vencidos", "% del Total Vencido"],
    )
    return result


# ── Análisis C: Tendencia semanal/mensual ─────────────────────────────────────

def compute_trend(detail_df: pd.DataFrame):
    """
    Retorna (weekly_df, monthly_df) con % de ingresos oportunos por período.
    """
    df = detail_df[detail_df["_po_doc_date"].notna() & (detail_df["qty_101"] > 0)].copy()
    df = df.set_index("_po_doc_date")

    def trend_resample(freq, label):
        if df.empty:
            return pd.DataFrame(columns=["Período", "Total", "Oportunos", "% Oportuno"])
        grp = df.resample(freq)
        total    = grp.size().rename("Total")
        oportuno = grp["alerta"].apply(
            lambda x: ((x == ALERT_GREEN) | (x == ALERT_YELLOW)).sum()
        ).rename("Oportunos")
        result = pd.concat([total, oportuno], axis=1).reset_index()
        result.columns = ["Período", "Total", "Oportunos"]
        result["% Oportuno"] = (result["Oportunos"] / result["Total"] * 100).round(1)
        return result

    weekly  = trend_resample("W-MON", "Semana")
    monthly = trend_resample("MS", "Mes")
    return weekly, monthly


# ── Análisis D: Materiales con mayor importe pendiente ───────────────────────

def top_pending_amount(detail_df: pd.DataFrame, n: int = TOP_PENDING_N) -> pd.DataFrame:
    """Top N materiales por importe pendiente estimado."""
    df = detail_df[detail_df["importe_pendiente"] > 0].copy()
    if df.empty:
        return pd.DataFrame(columns=["Material", "Descripción", "Importe Pendiente", "Qty Pendiente"])

    result = (
        df.groupby(["_material_po", "_description"], as_index=False)
        .agg(
            importe_pendiente=("importe_pendiente", "sum"),
            qty_pendiente=("qty_pendiente", "sum"),
        )
        .sort_values("importe_pendiente", ascending=False)
        .head(n)
    )
    result.columns = ["Material", "Descripción", "Importe Pendiente", "Qty Pendiente"]
    return result


# ── Análisis E: Tasa de anulaciones por origen y usuario ─────────────────────

def cancellation_rate(detail_df: pd.DataFrame) -> pd.DataFrame:
    """Tasa de anulaciones (filas con qty_102 > 0) por origen y usuario."""
    rows = []
    for _, row in detail_df.iterrows():
        users = [u.strip() for u in str(row["users_101"]).split(",") if u.strip() and u.strip() != "nan"]
        if not users:
            users = ["(sin usuario)"]
        for u in users:
            rows.append({
                "origen":       row["origen"],
                "usuario":      u,
                "tiene_anulacion": 1 if row["tiene_anulacion"] == "Sí" else 0,
                "total":        1,
            })

    if not rows:
        return pd.DataFrame(columns=["Origen", "Usuario", "Total", "Con Anulación", "% Anulación"])

    df_rows = pd.DataFrame(rows)
    result = (
        df_rows.groupby(["origen", "usuario"], as_index=False)
        .agg(total=("total", "sum"), con_anulacion=("tiene_anulacion", "sum"))
        .assign(pct_anulacion=lambda x: (x["con_anulacion"] / x["total"] * 100).round(1))
        .sort_values("pct_anulacion", ascending=False)
    )
    result.columns = ["Origen", "Usuario", "Total", "Con Anulación", "% Anulación"]
    return result


# ── Recomendaciones automáticas ──────────────────────────────────────────────

def generate_recommendations(kpis: dict, detail_df: pd.DataFrame,
                              top_mat: pd.DataFrame, top_users: pd.DataFrame,
                              policy_days: int) -> list:
    """
    Genera hasta MAX_RECOMMENDATIONS recomendaciones basadas en los hallazgos.
    """
    recs = []

    # 1. Alto % de vencidos
    if kpis["pct_vencido"] > 30:
        recs.append(
            f"PRIORIDAD ALTA — El {kpis['pct_vencido']}% de los ingresos superaron el plazo de "
            f"{policy_days} días. Revisar con urgencia los procesos de confirmación de recepción "
            f"para reducir esta tasa."
        )
    elif kpis["pct_vencido"] > 15:
        recs.append(
            f"El {kpis['pct_vencido']}% de los ingresos están fuera de la política ({policy_days} días). "
            f"Implementar alertas tempranas al día {policy_days - 1} para anticipar vencimientos."
        )

    # 2. Materiales más lentos
    if not top_mat.empty:
        top3 = top_mat.head(3)
        mat_list = ", ".join(
            f"{row['Material']} ({row['Días Promedio']} días)"
            for _, row in top3.iterrows()
        )
        recs.append(
            f"Los materiales con mayor demora promedio son: {mat_list}. "
            f"Revisar si existen problemas de tránsito o de registro en SAP para estos ítems."
        )

    # 3. Usuarios con más vencidos
    if not top_users.empty:
        top_user = top_users.iloc[0]
        recs.append(
            f"El usuario '{top_user['Usuario']}' concentra {top_user['Ingresos Vencidos']} ingresos "
            f"vencidos ({top_user['% del Total Vencido']}% del total ROJO). "
            f"Considerar capacitación o redistribución de carga de trabajo."
        )

    # 4. Alto % de ingresos parciales
    if kpis["pct_parciales"] > 25:
        recs.append(
            f"El {kpis['pct_parciales']}% de los ingresos son parciales (quedan cantidades pendientes). "
            f"Verificar si los despachos del proveedor son incompletos o si hay problemas de confirmación."
        )

    # 5. Comparativo por origen
    origen_kpis = kpis.get("por_origen", {})
    if len(origen_kpis) >= 2:
        sorted_origins = sorted(
            [(k, v) for k, v in origen_kpis.items() if v is not None],
            key=lambda x: x[1]
        )
        if sorted_origins:
            worst = sorted_origins[0]
            best  = sorted_origins[-1]
            if worst[1] < best[1] - 15:
                recs.append(
                    f"Existe una brecha de {round(best[1] - worst[1], 1)}% en cumplimiento entre "
                    f"'{best[0]}' ({best[1]}% oportuno) y '{worst[0]}' ({worst[1]}% oportuno). "
                    f"Investigar las causas del rezago en '{worst[0]}'."
                )

    # Limitar al máximo configurado
    return recs[:MAX_RECOMMENDATIONS]

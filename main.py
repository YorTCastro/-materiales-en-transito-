# main.py — Orquestador CLI de la herramienta de Materiales en Tránsito

import sys
import traceback
from datetime import date
from pathlib import Path

from openpyxl import Workbook

import config
from loader import load_and_validate
from transformer import aggregate_movements, merge_mb51_with_po
from analyzer import (
    build_detail_df,
    cancellation_rate,
    compute_kpis,
    compute_trend,
    generate_recommendations,
    top_materials_by_avg_time,
    top_pending_amount,
    top_users_overdue,
)
from reporter import (
    save_workbook,
    write_analysis_sheet,
    write_detail_sheet,
    write_kpis_sheet,
    write_recommendations_sheet,
    write_trend_sheet,
)


# ── Banner ────────────────────────────────────────────────────────────────────

BANNER = r"""
╔══════════════════════════════════════════════════════════════╗
║        MATERIALES EN TRÁNSITO — Análisis de Ingresos        ║
║              Cruce MB51 × Archivo de Pedidos                 ║
╚══════════════════════════════════════════════════════════════╝
"""


# ── Entrada de rutas ──────────────────────────────────────────────────────────

def ask_filepath(prompt: str, must_exist: bool = True) -> str:
    while True:
        raw = input(prompt).strip().strip('"').strip("'")
        if not raw:
            print("  Por favor ingresa una ruta.")
            continue
        p = Path(raw)
        if must_exist and not p.exists():
            print(f"  Archivo no encontrado: {raw}")
            continue
        return str(p)


def ask_policy_days() -> int:
    default = config.POLICY_DAYS
    raw = input(f"\n  Umbral de política en días (Enter para usar {default}): ").strip()
    if not raw:
        return default
    try:
        val = int(raw)
        if val < 1:
            raise ValueError
        return val
    except ValueError:
        print(f"  Valor inválido. Se usará el default: {default} días.")
        return default


def ask_output_path() -> str:
    default_name = f"Informe_Transito_{date.today().strftime('%Y%m%d')}.xlsx"
    raw = input(f"\n  Ruta/nombre del archivo Excel de salida\n"
                f"  (Enter para '{default_name}' en el directorio actual): ").strip().strip('"').strip("'")
    if not raw:
        return default_name
    p = Path(raw)
    if p.suffix.lower() != ".xlsx":
        p = p.with_suffix(".xlsx")
    return str(p)


# ── Pipeline principal ────────────────────────────────────────────────────────

def run():
    print(BANNER)

    # 1. Solicitar archivos
    print("  PASO 1 — Proporciona las rutas de los archivos SAP\n")
    mb51_path = ask_filepath("  Ruta del archivo MB51 (movimientos): ")
    po_path   = ask_filepath("  Ruta del archivo de Pedidos (Doc. compra): ")

    # 2. Umbral de política
    policy_days = ask_policy_days()
    print(f"  Umbral de política: {policy_days} días\n")

    # 3. Ruta de salida
    output_path = ask_output_path()

    print("\n" + "─" * 62)
    print("  PASO 2 — Cargando y validando archivos...")
    print("─" * 62)

    # 4. Cargar MB51
    mb51_df, mb51_col_map = load_and_validate(
        mb51_path,
        config.MB51_COLS,
        config.MB51_KEY_COLS,
        file_label="MB51 (movimientos)",
        interactive=True,
    )

    # 5. Cargar archivo de pedidos
    po_df, po_col_map = load_and_validate(
        po_path,
        config.PO_COLS,
        config.PO_KEY_COLS,
        file_label="Archivo de Pedidos",
        interactive=True,
    )

    print("\n" + "─" * 62)
    print("  PASO 3 — Procesando datos...")
    print("─" * 62)

    # 6. Agregar movimientos MB51
    print("  Agregando movimientos por pedido+posición...")
    agg_mb51 = aggregate_movements(mb51_df, mb51_col_map)
    print(f"  → {len(agg_mb51)} líneas únicas en MB51 (pos. con mov. 101/102)")

    # 7. Cruce con pedidos
    print("  Cruzando con archivo de pedidos...")
    merged_df = merge_mb51_with_po(agg_mb51, po_df, po_col_map)
    print(f"  → {len(merged_df)} líneas de pedido en el resultado")

    # 8. Calcular métricas de detalle
    print("  Calculando métricas de negocio...")
    detail_df = build_detail_df(merged_df, policy_days=policy_days)

    # 9. KPIs
    kpis = compute_kpis(detail_df)

    # Imprimir resumen en consola
    print("\n" + "─" * 62)
    print("  RESUMEN EJECUTIVO (consola)")
    print("─" * 62)
    print(f"  Total líneas analizadas:      {kpis['total_lineas']}")
    print(f"  Con entrada:                  {kpis['n_con_entrada']}")
    print(f"  Sin entrada:                  {kpis['n_sin_entrada']}")
    print(f"  % Oportuno (VERDE+AMARILLO):  {kpis['pct_oportuno']}%")
    print(f"  % Vencido (ROJO):             {kpis['pct_vencido']}%")
    print(f"  % Con anulación (102):        {kpis['pct_anulaciones']}%")
    print(f"  % Ingresos parciales:         {kpis['pct_parciales']}%")
    print(f"  Importe total pendiente:      {kpis['importe_pendiente_total']:,.2f}")
    if kpis["por_origen"]:
        print("  Por origen:")
        for origen, pct in kpis["por_origen"].items():
            pct_str = f"{round(pct, 1)}%" if pct is not None else "Sin datos"
            print(f"    {origen:<25} {pct_str}")

    # 10. Análisis adicionales
    print("\n  Calculando análisis adicionales...")
    top_mat   = top_materials_by_avg_time(detail_df)
    top_users = top_users_overdue(detail_df)
    weekly, monthly = compute_trend(detail_df)
    top_pend  = top_pending_amount(detail_df)
    cancel_df = cancellation_rate(detail_df)
    recs      = generate_recommendations(kpis, detail_df, top_mat, top_users, policy_days)

    # 11. Generar Excel
    print("\n" + "─" * 62)
    print("  PASO 4 — Generando Excel...")
    print("─" * 62)

    wb = Workbook()
    # Eliminar hoja por defecto
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    write_kpis_sheet(wb, kpis, policy_days)
    write_detail_sheet(wb, detail_df)
    write_analysis_sheet(wb, top_mat,   "top_materials", "Top 10 Materiales por Tiempo Promedio de Ingreso")
    write_analysis_sheet(wb, top_users, "top_users",     "Top 5 Usuarios con Más Ingresos Vencidos")
    write_trend_sheet(wb, weekly, monthly)
    write_analysis_sheet(wb, top_pend,  "pending",       "Materiales con Mayor Importe Pendiente")
    write_analysis_sheet(wb, cancel_df, "cancellations", "Tasa de Anulaciones por Origen y Usuario")
    write_recommendations_sheet(wb, recs)

    save_workbook(wb, output_path)

    print("\n  Recomendaciones generadas:")
    for i, r in enumerate(recs, 1):
        print(f"  {i}. {r[:100]}{'...' if len(r) > 100 else ''}")

    print("\n  ¡Análisis completado exitosamente!")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        # Soporte para argumento "Ajustar política: X días"
        if len(sys.argv) > 1:
            import re
            arg = " ".join(sys.argv[1:])
            m = re.search(r"ajustar\s+pol[ií]tica[:\s]+(\d+)\s*d[ií]as?", arg, re.IGNORECASE)
            if m:
                config.POLICY_DAYS = int(m.group(1))
                print(f"  Política ajustada a {config.POLICY_DAYS} días desde argumento de línea de comandos.")

        run()

    except KeyboardInterrupt:
        print("\n\n  Operación cancelada por el usuario.")
        sys.exit(0)

    except FileNotFoundError as e:
        print(f"\n  ERROR — Archivo no encontrado: {e}")
        sys.exit(1)

    except ValueError as e:
        print(f"\n  ERROR — Problema con los datos: {e}")
        print("  Verifica que los archivos sean exportaciones válidas de SAP.")
        sys.exit(1)

    except PermissionError as e:
        print(f"\n  ERROR — Sin permisos para escribir el archivo de salida: {e}")
        print("  Cierra el Excel si está abierto e intenta de nuevo.")
        sys.exit(1)

    except Exception as e:
        print(f"\n  ERROR inesperado: {e}")
        print("  Detalle técnico:")
        traceback.print_exc()
        print("\n  Si el problema persiste, reporta el error anterior al equipo de soporte.")
        sys.exit(1)

    finally:
        input("\n  Presiona Enter para salir...")

# app.py — Interfaz web Streamlit con login para Materiales en Tránsito

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth

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
    write_analysis_sheet,
    write_detail_sheet,
    write_kpis_sheet,
    write_recommendations_sheet,
    write_trend_sheet,
    save_workbook,
)
from config import MB51_COLS, PO_COLS, MB51_KEY_COLS, PO_KEY_COLS, POLICY_DAYS
from openpyxl import Workbook

# ── Configuración de página ───────────────────────────────────────────────────

st.set_page_config(
    page_title="Materiales en Tránsito",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Estilos CSS ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .metric-card {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 12px 16px;
        border-left: 4px solid #2F5597;
        margin-bottom: 8px;
    }
    .alert-verde    { background-color: #e8f5e9; color: #1b5e20; padding: 3px 8px; border-radius: 4px; font-weight: bold; }
    .alert-amarillo { background-color: #fff8e1; color: #f57f17; padding: 3px 8px; border-radius: 4px; font-weight: bold; }
    .alert-rojo     { background-color: #ffebee; color: #b71c1c; padding: 3px 8px; border-radius: 4px; font-weight: bold; }
    .login-title { text-align: center; font-size: 2rem; font-weight: bold; color: #2F5597; margin-bottom: 0.5rem; }
    .login-sub   { text-align: center; color: #666; margin-bottom: 2rem; }
</style>
""", unsafe_allow_html=True)


# ── Autenticación ─────────────────────────────────────────────────────────────

# Inicializar claves de sesión requeridas por streamlit-authenticator
for _key in ["authentication_status", "name", "username", "logout"]:
    if _key not in st.session_state:
        st.session_state[_key] = None

with open("credentials.yaml") as f:
    _config = yaml.load(f, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    _config["credentials"],
    _config["cookie"]["name"],
    _config["cookie"]["key"],
    _config["cookie"]["expiry_days"],
)

# Mostrar login centrado cuando no está autenticado
if st.session_state.get("authentication_status") is not True:
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown('<p class="login-title">📦 Materiales en Tránsito</p>', unsafe_allow_html=True)
        st.markdown('<p class="login-sub">Sistema de análisis de ingresos SAP</p>', unsafe_allow_html=True)
        authenticator.login(location="main")

        if st.session_state.get("authentication_status") is False:
            st.error("Usuario o contraseña incorrectos.")
        elif st.session_state.get("authentication_status") is None:
            st.info("Ingresa tus credenciales para continuar.")
    st.stop()


# ── App principal (solo si está autenticado) ──────────────────────────────────

# Sidebar
with st.sidebar:
    st.markdown(f"### 👤 {st.session_state.get('name', 'Usuario')}")
    st.caption(f"Usuario: `{st.session_state.get('username', '')}`")
    st.divider()
    authenticator.logout("Cerrar sesión", location="sidebar")
    st.divider()

    st.markdown("### ⚙️ Configuración")
    policy_days = st.number_input(
        "Umbral de política (días)",
        min_value=1, max_value=30,
        value=POLICY_DAYS,
        help="Días máximos permitidos entre fecha del pedido y fecha de ingreso."
    )
    st.caption(f"🟢 Verde: 1–3 días  \n🟡 Amarillo: 4–{policy_days} días  \n🔴 Rojo: >{policy_days} días")
    st.divider()

    st.markdown("### 📄 Plantillas SAP")
    st.caption("Descarga los formatos de ejemplo para preparar tus archivos.")

    @st.cache_data(show_spinner=False)
    def generar_plantilla_mb51():
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        AZUL = "FF2F5597"; BLANCO = "FFFFFFFF"
        AMARILLO = "FFFFF2CC"; VERDE = "FFE2EFDA"

        def hc(cell, texto, bg=AZUL):
            cell.value = texto
            cell.fill  = PatternFill("solid", fgColor=bg)
            cell.font  = Font(bold=True, color=BLANCO if bg==AZUL else "FF000000", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            t = Side(style="thin", color="FF000000")
            cell.border = Border(left=t, right=t, top=t, bottom=t)

        def dc(cell, valor, bg=None, italic=False, size=10):
            cell.value = valor
            if bg: cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(size=size, italic=italic, color="FF595959" if italic else "FF000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            t = Side(style="thin", color="FFD9D9D9")
            cell.border = Border(left=t, right=t, top=t, bottom=t)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MB51 - Movimientos"

        ws.merge_cells("A1:R1")
        ws["A1"].value = "PLANTILLA MB51 — Movimientos de Materiales (Exportación SAP)"
        ws["A1"].fill  = PatternFill("solid", fgColor=AZUL)
        ws["A1"].font  = Font(bold=True, color=BLANCO, size=12)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        ws.merge_cells("A2:R2")
        ws["A2"].value = "FORMATO: Exportar desde SAP transacción MB51 como .TXT o .CSV separado por punto y coma (;). Movimiento 101 = Ingreso | 102 = Anulación"
        ws["A2"].fill  = PatternFill("solid", fgColor=AMARILLO)
        ws["A2"].font  = Font(color="FF7F6000", size=10, italic=True)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 30

        cols = [
            ("Material","Código material SAP\nEj: 100001234"),
            ("Texto breve de material","Descripción del material"),
            ("CMv","Tipo movimiento\n101=Ingreso\n102=Anulación"),
            ("Alm.","Código almacén\nEj: HB01"),
            ("Pedido","N° pedido compra\nEj: 4500001234"),
            ("Pos.","Posición pedido\nEj: 10"),
            ("Doc.mat.","N° documento material"),
            ("Pos","Posición documento\nEj: 1"),
            ("Cantidad","Cantidad\nComa decimal\nEj: 10,000"),
            ("UMB","Unidad medida\nEj: UN, KG, M"),
            ("Fecha doc.","Fecha documento\nDD.MM.AAAA"),
            ("Reserva","N° reserva\nPuede ir vacío"),
            ("Fe.contab.","Fecha contabilización\nDD.MM.AAAA"),
            ("Hora","Hora movimiento\nEj: 14:30:00"),
            ("Importe ML","Importe moneda local\nEj: 1.500,00"),
            ("Texto cab.documento","Texto cabecera\nPuede ir vacío"),
            ("Referencia","Referencia externa\nPuede ir vacío"),
            ("Usuario","Usuario SAP\nEj: JPEREZ"),
        ]
        for i, (n, _) in enumerate(cols, 1):
            hc(ws.cell(3, i), n)
            ws.column_dimensions[get_column_letter(i)].width = 16
        ws.row_dimensions[3].height = 20
        for i, (_, t) in enumerate(cols, 1):
            dc(ws.cell(4, i), t, bg="FFF2F2F2", italic=True, size=8)
        ws.row_dimensions[4].height = 60

        ejemplos = [
            ["100001234","VALVULA ESFERICA 1","101","HB01","4500001234","10","5000001001","1","10,000","UN","01.03.2024","","05.03.2024","10:30:00","1.500,00","","","JPEREZ"],
            ["100001234","VALVULA ESFERICA 1","101","HB01","4500001234","20","5000001002","1","5,000","UN","01.03.2024","","04.03.2024","09:15:00","750,00","","","MGARCIA"],
            ["100002567","TORNILLO HEX M12x50","101","HB02","4500001235","10","5000001003","1","100,000","UN","01.03.2024","","15.03.2024","14:00:00","200,00","","","JPEREZ"],
            ["100002567","TORNILLO HEX M12x50","102","HB02","4500001235","10","5000001004","1","20,000","UN","01.03.2024","","16.03.2024","11:00:00","40,00","ANULACION","","MGARCIA"],
            ["100003891","TUBO ACERO 2 SCH40","101","HB01","4500001236","10","5000001005","1","50,000","M","05.03.2024","","07.03.2024","08:45:00","5.000,00","","","CLOPEZ"],
        ]
        for r, row in enumerate(ejemplos, 5):
            bg = VERDE if row[2] == "101" else "FFFFCCCC"
            for c, v in enumerate(row, 1):
                dc(ws.cell(r, c), v, bg=bg)
            ws.row_dimensions[r].height = 18

        ws.merge_cells("A11:R11")
        ws["A11"].value = "VERDE = Movimiento 101 (Ingreso)     ROJO = Movimiento 102 (Anulación)"
        ws["A11"].fill  = PatternFill("solid", fgColor="FFF2F2F2")
        ws["A11"].font  = Font(italic=True, size=9)
        ws["A11"].alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A5"

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    @st.cache_data(show_spinner=False)
    def generar_plantilla_pedidos():
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        AZUL = "FF2F5597"; BLANCO = "FFFFFFFF"
        AMARILLO = "FFFFF2CC"; VERDE = "FFE2EFDA"

        def hc(cell, texto, bg=AZUL):
            cell.value = texto
            cell.fill  = PatternFill("solid", fgColor=bg)
            cell.font  = Font(bold=True, color=BLANCO if bg==AZUL else "FF000000", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            t = Side(style="thin", color="FF000000")
            cell.border = Border(left=t, right=t, top=t, bottom=t)

        def dc(cell, valor, bg=None, italic=False, size=10):
            cell.value = valor
            if bg: cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(size=size, italic=italic, color="FF595959" if italic else "FF000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            t = Side(style="thin", color="FFD9D9D9")
            cell.border = Border(left=t, right=t, top=t, bottom=t)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Pedidos - Doc. Compra"

        ws.merge_cells("A1:J1")
        ws["A1"].value = "PLANTILLA PEDIDOS — Documentos de Compra (Exportación SAP)"
        ws["A1"].fill  = PatternFill("solid", fgColor=AZUL)
        ws["A1"].font  = Font(bold=True, color=BLANCO, size=12)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        ws.merge_cells("A2:J2")
        ws["A2"].value = "FORMATO: Exportar desde SAP transacciones ME2M / ME2L como .TXT o .CSV separado por punto y coma (;)."
        ws["A2"].fill  = PatternFill("solid", fgColor=AMARILLO)
        ws["A2"].font  = Font(color="FF7F6000", size=10, italic=True)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 25

        cols = [
            ("Material","Código material SAP\nEj: 100001234"),
            ("Texto breve","Descripción del material"),
            ("Cantidad","Cantidad total pedido\nComa decimal\nEj: 15,000"),
            ("Por entrg.","Cantidad pendiente\npor entregar\nEj: 5,000"),
            ("UMP","Unidad de medida\nEj: UN, KG, M"),
            ("Fecha doc.","Fecha del pedido\nDD.MM.AAAA\nEj: 01.03.2024"),
            ("Doc.compr.","N° documento compra\nEj: 4500001234"),
            ("Pos.","Posición pedido\nEj: 10"),
            ("Proveedor/Centro suministrador","Nombre proveedor o\ncentro suministrador\nEj: La Estrella / Taller"),
            ("Mon.","Moneda\nEj: CLP, USD"),
        ]
        for i, (n, _) in enumerate(cols, 1):
            hc(ws.cell(3, i), n)
            ws.column_dimensions[get_column_letter(i)].width = 20
        ws.row_dimensions[3].height = 20
        for i, (_, t) in enumerate(cols, 1):
            dc(ws.cell(4, i), t, bg="FFF2F2F2", italic=True, size=8)
        ws.row_dimensions[4].height = 60

        ejemplos = [
            ["100001234","VALVULA ESFERICA 1","15,000","0,000","UN","01.03.2024","4500001234","10","Almacen La Estrella","CLP"],
            ["100001234","VALVULA ESFERICA 1","10,000","0,000","UN","01.03.2024","4500001234","20","Almacen La Estrella","CLP"],
            ["100002567","TORNILLO HEX M12x50","100,000","20,000","UN","01.03.2024","4500001235","10","Taller Central","CLP"],
            ["100003891","TUBO ACERO 2 SCH40","50,000","0,000","M","05.03.2024","4500001236","10","Almacen La Estrella","CLP"],
            ["100004512","BRIDA SLIP-ON 2","30,000","30,000","UN","10.03.2024","4500001237","10","Taller Central","CLP"],
        ]
        for r, row in enumerate(ejemplos, 5):
            for c, v in enumerate(row, 1):
                dc(ws.cell(r, c), v, bg=VERDE)
            ws.row_dimensions[r].height = 18

        ws.merge_cells("A11:J11")
        ws["A11"].value = "El campo Proveedor/Centro suministrador debe contener 'La Estrella' o 'Taller' para identificar el origen automaticamente."
        ws["A11"].fill  = PatternFill("solid", fgColor=AMARILLO)
        ws["A11"].font  = Font(italic=True, size=9, color="FF7F6000")
        ws["A11"].alignment = Alignment(horizontal="center", wrap_text=True)
        ws.freeze_panes = "A5"

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    st.download_button(
        label="⬇️ Plantilla MB51",
        data=generar_plantilla_mb51(),
        file_name="Plantilla_MB51.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.download_button(
        label="⬇️ Plantilla Pedidos",
        data=generar_plantilla_pedidos(),
        file_name="Plantilla_Pedidos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.divider()
    st.caption("v1.0 · Materiales en Tránsito")

# Título principal
st.title("📦 Análisis de Materiales en Tránsito")
st.caption("Carga los dos archivos SAP para generar el análisis completo.")

# ── Carga de archivos ─────────────────────────────────────────────────────────

col_a, col_b = st.columns(2)
with col_a:
    mb51_file = st.file_uploader(
        "📂 Archivo MB51 (Movimientos de materiales)",
        type=["csv", "txt", "xlsx"],
        help="Exportación de la transacción MB51 de SAP."
    )
with col_b:
    po_file = st.file_uploader(
        "📂 Archivo de Pedidos (Documentos de compra)",
        type=["csv", "txt", "xlsx"],
        help="Exportación del listado de pedidos/documentos de compra de SAP."
    )

if not mb51_file or not po_file:
    st.info("⬆️ Sube los dos archivos para comenzar el análisis.")
    st.stop()

# ── Procesamiento ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def process_files(mb51_bytes, mb51_name, po_bytes, po_name, policy_days):
    """Procesa los archivos y retorna todos los DataFrames y KPIs."""

    def save_temp(content, name):
        suffix = Path(name).suffix or ".csv"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(content)
        tmp.flush()
        return tmp.name

    mb51_path = save_temp(mb51_bytes, mb51_name)
    po_path   = save_temp(po_bytes, po_name)

    mb51_df, mb51_col_map = load_and_validate(
        mb51_path, MB51_COLS, MB51_KEY_COLS, "MB51", interactive=False
    )
    po_df, po_col_map = load_and_validate(
        po_path, PO_COLS, PO_KEY_COLS, "Pedidos", interactive=False
    )

    agg_mb51  = aggregate_movements(mb51_df, mb51_col_map)
    merged    = merge_mb51_with_po(agg_mb51, po_df, po_col_map)
    detail    = build_detail_df(merged, policy_days=policy_days)
    kpis      = compute_kpis(detail)
    top_mat   = top_materials_by_avg_time(detail)
    top_users = top_users_overdue(detail)
    weekly, monthly = compute_trend(detail)
    top_pend  = top_pending_amount(detail)
    cancel_df = cancellation_rate(detail)
    recs      = generate_recommendations(kpis, detail, top_mat, top_users, policy_days)

    return detail, kpis, top_mat, top_users, weekly, monthly, top_pend, cancel_df, recs

with st.spinner("Procesando archivos SAP..."):
    try:
        detail, kpis, top_mat, top_users, weekly, monthly, top_pend, cancel_df, recs = process_files(
            mb51_file.read(), mb51_file.name,
            po_file.read(), po_file.name,
            policy_days,
        )
    except Exception as e:
        st.error(f"Error al procesar los archivos: {e}")
        st.caption("Verifica que los archivos sean exportaciones válidas de SAP con las columnas correctas.")
        st.stop()

# ── KPIs ──────────────────────────────────────────────────────────────────────

st.divider()
st.subheader("📊 Resumen Ejecutivo")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total líneas",         kpis["total_lineas"])
k2.metric("Con entrada",          kpis["n_con_entrada"])
k3.metric("Sin entrada",          kpis["n_sin_entrada"])
k4.metric("Importe pendiente",    f"{kpis['importe_pendiente_total']:,.0f}")

st.markdown("")
c1, c2, c3, c4 = st.columns(4)
c1.metric("🟢 Oportunos",  f"{kpis['pct_oportuno']}%",  f"{kpis['n_verde']+kpis['n_amarillo']} registros")
c2.metric("🔴 Vencidos",   f"{kpis['pct_vencido']}%",   f"{kpis['n_rojo']} registros",    delta_color="inverse")
c3.metric("⚠️ Anulaciones", f"{kpis['pct_anulaciones']}%", f"{kpis['n_anulaciones']} registros", delta_color="inverse")
c4.metric("📦 Parciales",  f"{kpis['pct_parciales']}%",  f"{kpis['n_parciales']} registros")

# Comparativo por origen
if kpis.get("por_origen"):
    st.markdown("")
    st.markdown("**Cumplimiento por origen:**")
    orig_cols = st.columns(len(kpis["por_origen"]))
    for i, (origen, pct) in enumerate(sorted(kpis["por_origen"].items(),
                                              key=lambda x: (x[1] or 0), reverse=True)):
        pct_str = f"{round(pct, 1)}%" if pct is not None else "Sin datos"
        orig_cols[i].metric(origen, pct_str)

# ── Tabs de análisis ──────────────────────────────────────────────────────────

st.divider()
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Detalle por línea",
    "📈 Análisis adicionales",
    "📅 Tendencia",
    "💡 Recomendaciones",
    "⬇️ Descargar Excel",
])

# -- Tab 1: Detalle --
with tab1:
    st.markdown("Ordenado por días transcurridos (mayor a menor). Las filas en naranja tienen datos incompletos.")

    display_cols = {
        "Material":           "_material_po",
        "Descripción":        "_description",
        "Origen":             "origen",
        "Pedido":             "_order",
        "Pos.":               "_position",
        "Fecha Pedido":       "_po_doc_date",
        "Fecha Ingreso":      "last_acc_date",
        "Días":               "dias_transcurridos",
        "Alerta":             "alerta",
        "Qty Ordenada":       "_qty_ordered",
        "Qty Entrada":        "qty_101",
        "Qty Anulada":        "qty_102",
        "Qty Neta":           "qty_neta",
        "Qty Pendiente":      "qty_pendiente",
        "Importe ML":         "amount_101",
        "Usuarios":           "users_101",
        "¿Anulación?":        "tiene_anulacion",
    }

    df_show = detail[[c for c in display_cols.values() if c in detail.columns]].copy()
    df_show.columns = [k for k, v in display_cols.items() if v in detail.columns]
    df_show = df_show.sort_values("Días", ascending=False, na_position="last")

    def color_alert_row(row):
        alerta = row.get("Alerta", "")
        colors = {
            "VERDE":       "background-color: #e8f5e9",
            "AMARILLO":    "background-color: #fff8e1",
            "ROJO":        "background-color: #ffebee",
            "SIN ENTRADA": "background-color: #f5f5f5",
        }
        c = colors.get(alerta, "")
        return [c] * len(row)

    styled = df_show.style.apply(color_alert_row, axis=1)
    st.dataframe(styled, use_container_width=True, height=500)

# -- Tab 2: Análisis adicionales --
with tab2:
    sub1, sub2, sub3, sub4 = st.columns([1,1,1,1])

    with sub1:
        st.markdown("**🏆 Top materiales por tiempo promedio**")
        if not top_mat.empty:
            st.dataframe(top_mat, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin datos.")

    with sub2:
        st.markdown("**👤 Usuarios con más ingresos vencidos**")
        if not top_users.empty:
            st.dataframe(top_users, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin datos.")

    with sub3:
        st.markdown("**💰 Mayor importe pendiente**")
        if not top_pend.empty:
            st.dataframe(top_pend, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin datos.")

    with sub4:
        st.markdown("**🔄 Tasa de anulaciones**")
        if not cancel_df.empty:
            st.dataframe(cancel_df, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin datos.")

# -- Tab 3: Tendencia --
with tab3:
    col_w, col_m = st.columns(2)

    with col_w:
        st.markdown("**Tendencia semanal**")
        if not weekly.empty:
            st.line_chart(
                weekly.set_index("Período")["% Oportuno"],
                use_container_width=True,
            )
            st.dataframe(weekly, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin suficientes datos para tendencia semanal.")

    with col_m:
        st.markdown("**Tendencia mensual**")
        if not monthly.empty:
            st.line_chart(
                monthly.set_index("Período")["% Oportuno"],
                use_container_width=True,
            )
            st.dataframe(monthly, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin suficientes datos para tendencia mensual.")

# -- Tab 4: Recomendaciones --
with tab4:
    if recs:
        for i, rec in enumerate(recs, 1):
            st.markdown(f"**{i}.** {rec}")
            st.markdown("")
    else:
        st.success("No se detectaron problemas críticos que recomienden acción inmediata.")

# -- Tab 5: Descargar Excel --
with tab5:
    st.markdown("Genera el informe completo en Excel con todas las hojas y formatos.")

    @st.cache_data(show_spinner=False)
    def generate_excel(detail_bytes, kpis_str, policy_days):
        """Genera el Excel en memoria y retorna los bytes."""
        import pickle, hashlib
        # Reconstruir objetos desde caché
        detail_local = pd.read_parquet(io.BytesIO(detail_bytes))

        kpis_local      = compute_kpis(detail_local)
        top_mat_local   = top_materials_by_avg_time(detail_local)
        top_users_local = top_users_overdue(detail_local)
        weekly_l, monthly_l = compute_trend(detail_local)
        top_pend_local  = top_pending_amount(detail_local)
        cancel_local    = cancellation_rate(detail_local)
        recs_local      = generate_recommendations(kpis_local, detail_local,
                                                    top_mat_local, top_users_local, policy_days)

        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        write_kpis_sheet(wb, kpis_local, policy_days)
        write_detail_sheet(wb, detail_local)
        write_analysis_sheet(wb, top_mat_local,   "top_materials", "Top 10 Materiales")
        write_analysis_sheet(wb, top_users_local, "top_users",     "Top 5 Usuarios Demora")
        write_trend_sheet(wb, weekly_l, monthly_l)
        write_analysis_sheet(wb, top_pend_local,  "pending",       "Mayor Importe Pendiente")
        write_analysis_sheet(wb, cancel_local,    "cancellations", "Tasa de Anulaciones")
        write_recommendations_sheet(wb, recs_local)

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # Serializar detail para caché
    detail_buf = io.BytesIO()
    detail.to_parquet(detail_buf)

    if st.button("📥 Generar Excel", type="primary"):
        with st.spinner("Generando Excel..."):
            try:
                excel_bytes = generate_excel(detail_buf.getvalue(), str(kpis), policy_days)
                from datetime import date
                filename = f"Informe_Transito_{date.today().strftime('%Y%m%d')}.xlsx"
                st.download_button(
                    label="⬇️ Descargar informe Excel",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.success(f"Excel listo: {filename}")
            except Exception as e:
                st.error(f"Error generando Excel: {e}")

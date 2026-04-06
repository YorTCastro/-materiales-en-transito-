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

    st.divider()
    st.caption("v1.0 · Materiales en Tránsito")

# Título principal
st.title("📦 Análisis de Materiales en Tránsito")

# ── Carga de archivos + Plantillas ────────────────────────────────────────────

col_a, col_b = st.columns(2)
with col_a:
    mb51_file = st.file_uploader(
        "📂 Archivo MB51 (Movimientos de materiales)",
        type=["csv", "txt", "xlsx"],
        help="Exportación de la transacción MB51 de SAP."
    )
    st.download_button(
        label="📄 Descargar plantilla MB51",
        data=generar_plantilla_mb51(),
        file_name="Plantilla_MB51.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_b:
    po_file = st.file_uploader(
        "📂 Archivo de Pedidos (Documentos de compra)",
        type=["csv", "txt", "xlsx"],
        help="Exportación del listado de pedidos/documentos de compra de SAP."
    )
    st.download_button(
        label="📄 Descargar plantilla Pedidos",
        data=generar_plantilla_pedidos(),
        file_name="Plantilla_Pedidos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
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

import plotly.graph_objects as go
import plotly.express as px

st.divider()
st.subheader("📊 Resumen Ejecutivo")

# Fila 1 — tarjetas de indicadores clave
def kpi_card(label, value, sublabel="", color="#2F5597"):
    return f"""
    <div style="background:{color}10; border-left:4px solid {color};
                border-radius:8px; padding:14px 18px; height:90px;">
        <div style="font-size:12px; color:#555; font-weight:600;
                    text-transform:uppercase; letter-spacing:0.5px;">{label}</div>
        <div style="font-size:28px; font-weight:700; color:{color}; line-height:1.2;">{value}</div>
        <div style="font-size:11px; color:#888; margin-top:2px;">{sublabel}</div>
    </div>"""

col1, col2, col3, col4, col5 = st.columns(5)
col1.markdown(kpi_card("Total líneas", kpis["total_lineas"], "pedidos analizados"), unsafe_allow_html=True)
col2.markdown(kpi_card("Con entrada", kpis["n_con_entrada"], f"{kpis['n_sin_entrada']} sin entrada", "#1565C0"), unsafe_allow_html=True)
col3.markdown(kpi_card("% Oportuno", f"{kpis['pct_oportuno']}%", f"{kpis['n_verde']+kpis['n_amarillo']} registros", "#2E7D32"), unsafe_allow_html=True)
col4.markdown(kpi_card("% Vencido", f"{kpis['pct_vencido']}%", f"{kpis['n_rojo']} registros ROJO", "#C62828"), unsafe_allow_html=True)
col5.markdown(kpi_card("Imp. Pendiente", f"{kpis['importe_pendiente_total']:,.0f}", "moneda local", "#E65100"), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Fila 2 — gráficas principales
gc1, gc2, gc3 = st.columns([1.1, 1.1, 0.8])

# Gráfica 1: Donut distribución de alertas
with gc1:
    n_verde   = kpis["n_verde"]
    n_amarillo= kpis["n_amarillo"]
    n_rojo    = kpis["n_rojo"]
    n_sin     = kpis["n_sin_entrada"]
    labels    = ["Verde (1-3 días)", f"Amarillo (4-{policy_days} días)", f"Rojo (>{policy_days} días)", "Sin entrada"]
    values    = [n_verde, n_amarillo, n_rojo, n_sin]
    colors_d  = ["#00B050", "#FFC000", "#FF0000", "#BDBDBD"]

    fig_donut = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.55,
        marker=dict(colors=colors_d, line=dict(color="#ffffff", width=2)),
        textinfo="percent+value",
        textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>Cantidad: %{value}<br>%{percent}<extra></extra>",
    ))
    fig_donut.add_annotation(
        text=f"<b>{kpis['total_lineas']}</b><br>líneas",
        x=0.5, y=0.5, font_size=14, showarrow=False
    )
    fig_donut.update_layout(
        title=dict(text="Distribución por Estado de Ingreso", font_size=14, x=0.5),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5, font_size=11),
        margin=dict(t=45, b=10, l=10, r=10),
        height=280,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_donut, use_container_width=True)

# Gráfica 2: Barras por origen
with gc2:
    origen_kpis = kpis.get("por_origen", {})
    if origen_kpis:
        orig_names = list(origen_kpis.keys())
        orig_vals  = [round(v, 1) if v is not None else 0 for v in origen_kpis.values()]
        orig_colors= ["#2196F3" if v >= 80 else "#FF9800" if v >= 60 else "#F44336" for v in orig_vals]

        fig_bar = go.Figure(go.Bar(
            x=orig_names, y=orig_vals,
            marker=dict(color=orig_colors, line=dict(color="#ffffff", width=1)),
            text=[f"{v}%" for v in orig_vals],
            textposition="outside",
            textfont=dict(size=13, color="#333"),
            hovertemplate="<b>%{x}</b><br>% Oportuno: %{y}%<extra></extra>",
        ))
        fig_bar.add_hline(y=80, line_dash="dash", line_color="#2E7D32",
                          annotation_text="Meta 80%", annotation_position="top right",
                          annotation_font_size=10)
        fig_bar.update_layout(
            title=dict(text="Cumplimiento por Origen", font_size=14, x=0.5),
            yaxis=dict(title="% Oportuno", range=[0, 110], ticksuffix="%", gridcolor="#eee"),
            xaxis=dict(title=""),
            margin=dict(t=45, b=10, l=10, r=10),
            height=280,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Sin datos de origen.")

# Gráfica 3: Indicadores secundarios (tarjetas mini)
with gc3:
    st.markdown("""
    <div style="display:flex; flex-direction:column; gap:10px; padding-top:8px;">
    """, unsafe_allow_html=True)

    def mini_card(icon, label, val, color):
        return f"""<div style="background:#fff; border:1px solid #e0e0e0; border-radius:8px;
                    padding:10px 14px; display:flex; align-items:center; gap:12px;">
            <div style="font-size:22px;">{icon}</div>
            <div>
                <div style="font-size:11px; color:#666; text-transform:uppercase;">{label}</div>
                <div style="font-size:20px; font-weight:700; color:{color};">{val}</div>
            </div></div>"""

    st.markdown(
        mini_card("⚠️", "Con anulaciones", f"{kpis['pct_anulaciones']}%", "#E65100") +
        mini_card("📦", "Ingresos parciales", f"{kpis['pct_parciales']}%", "#1565C0") +
        mini_card("❌", "Sin entrada", kpis["n_sin_entrada"], "#C62828") +
        mini_card("✅", "Verde (oportuno)", kpis["n_verde"], "#2E7D32"),
        unsafe_allow_html=True
    )
    st.markdown("</div>", unsafe_allow_html=True)

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
    st.caption("Ordenado por días transcurridos (mayor a menor).")

    display_cols = {
        "Material":      "_material_po",
        "Descripción":   "_description",
        "Origen":        "origen",
        "Pedido":        "_order",
        "Pos.":          "_position",
        "Fecha Pedido":  "_po_doc_date",
        "Fecha Ingreso": "last_acc_date",
        "Días":          "dias_transcurridos",
        "Alerta":        "alerta",
        "Qty Ordenada":  "_qty_ordered",
        "Qty Entrada":   "qty_101",
        "Qty Anulada":   "qty_102",
        "Qty Neta":      "qty_neta",
        "Qty Pendiente": "qty_pendiente",
        "Importe ML":    "amount_101",
        "Usuarios":      "users_101",
        "¿Anulación?":   "tiene_anulacion",
    }

    df_show = detail[[c for c in display_cols.values() if c in detail.columns]].copy()
    df_show.columns = [k for k, v in display_cols.items() if v in detail.columns]
    df_show = df_show.sort_values("Días", ascending=False, na_position="last")

    def color_alert_row(row):
        alerta = row.get("Alerta", "")
        colors = {
            "VERDE":       "background-color: #c6efce; color: #000000",
            "AMARILLO":    "background-color: #ffeb9c; color: #000000",
            "ROJO":        "background-color: #ffc7ce; color: #000000",
            "SIN ENTRADA": "background-color: #eeeeee; color: #000000",
        }
        c = colors.get(alerta, "color: #000000")
        return [c] * len(row)

    styled = df_show.style.apply(color_alert_row, axis=1)
    st.dataframe(styled, use_container_width=True, height=500)

# -- Tab 2: Análisis adicionales --
with tab2:
    r1c1, r1c2 = st.columns(2)

    # Gráfica: Top materiales por días promedio
    with r1c1:
        st.markdown("**🏆 Top materiales por tiempo promedio de ingreso**")
        if not top_mat.empty:
            fig_mat = px.bar(
                top_mat.sort_values("Días Promedio"),
                x="Días Promedio", y="Material",
                orientation="h",
                text="Días Promedio",
                color="Días Promedio",
                color_continuous_scale=["#00B050","#FFC000","#FF0000"],
                labels={"Días Promedio": "Días promedio", "Material": ""},
            )
            fig_mat.update_traces(texttemplate="%{text:.1f} días", textposition="outside")
            fig_mat.update_layout(
                height=320, margin=dict(t=10, b=10, l=10, r=30),
                coloraxis_showscale=False,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(tickfont=dict(size=11)),
            )
            st.plotly_chart(fig_mat, use_container_width=True)
        else:
            st.caption("Sin datos.")

    # Gráfica: Usuarios con más vencidos
    with r1c2:
        st.markdown("**👤 Usuarios con más ingresos vencidos**")
        if not top_users.empty:
            fig_usr = px.bar(
                top_users.sort_values("Ingresos Vencidos"),
                x="Ingresos Vencidos", y="Usuario",
                orientation="h",
                text="Ingresos Vencidos",
                color="Ingresos Vencidos",
                color_continuous_scale=["#FF9800","#F44336"],
                labels={"Ingresos Vencidos": "Ingresos vencidos", "Usuario": ""},
            )
            fig_usr.update_traces(textposition="outside")
            fig_usr.update_layout(
                height=320, margin=dict(t=10, b=10, l=10, r=30),
                coloraxis_showscale=False,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_usr, use_container_width=True)
        else:
            st.caption("Sin datos.")

    r2c1, r2c2 = st.columns(2)

    # Gráfica: Importe pendiente por material
    with r2c1:
        st.markdown("**💰 Materiales con mayor importe pendiente**")
        if not top_pend.empty:
            fig_pend = px.bar(
                top_pend.sort_values("Importe Pendiente"),
                x="Importe Pendiente", y="Material",
                orientation="h",
                text="Importe Pendiente",
                color_discrete_sequence=["#1565C0"],
                labels={"Importe Pendiente": "Importe pendiente", "Material": ""},
            )
            fig_pend.update_traces(
                texttemplate="%{text:,.0f}",
                textposition="outside",
            )
            fig_pend.update_layout(
                height=320, margin=dict(t=10, b=10, l=10, r=30),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_pend, use_container_width=True)
        else:
            st.caption("Sin datos.")

    # Tabla: anulaciones
    with r2c2:
        st.markdown("**🔄 Tasa de anulaciones por origen y usuario**")
        if not cancel_df.empty:
            def color_cancel(val):
                if isinstance(val, (int, float)):
                    if val >= 50: return "background-color:#FF9800; color:#000"
                    if val >= 20: return "background-color:#FFE0B2; color:#000"
                return ""
            st.dataframe(
                cancel_df.style.map(color_cancel, subset=["% Anulación"]),
                use_container_width=True, hide_index=True, height=320,
            )
        else:
            st.caption("Sin datos.")

# -- Tab 3: Tendencia --
with tab3:
    col_w, col_m = st.columns(2)

    def trend_chart(df_trend, title):
        if df_trend.empty:
            st.caption("Sin suficientes datos.")
            return
        df_trend = df_trend.copy()
        df_trend["Período"] = df_trend["Período"].astype(str).str[:10]
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_trend["Período"], y=df_trend["% Oportuno"],
            mode="lines+markers+text",
            name="% Oportuno",
            line=dict(color="#2F5597", width=2.5),
            marker=dict(size=8, color="#2F5597"),
            text=df_trend["% Oportuno"].apply(lambda v: f"{v:.0f}%"),
            textposition="top center",
            textfont=dict(size=11),
            fill="tozeroy",
            fillcolor="rgba(47,85,151,0.08)",
        ))
        fig.add_hline(y=80, line_dash="dash", line_color="#2E7D32",
                      annotation_text="Meta 80%", annotation_font_size=10)
        fig.update_layout(
            title=dict(text=title, font_size=14, x=0.5),
            yaxis=dict(range=[0, 110], ticksuffix="%", gridcolor="#eee", title="% Oportuno"),
            xaxis=dict(title="", tickangle=-30),
            height=320, margin=dict(t=45, b=40, l=10, r=10),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(df_trend, use_container_width=True, hide_index=True)

    with col_w:
        trend_chart(weekly, "Tendencia Semanal — % Ingresos Oportunos")
    with col_m:
        trend_chart(monthly, "Tendencia Mensual — % Ingresos Oportunos")

# -- Tab 4: Recomendaciones --
with tab4:
    if recs:
        for i, rec in enumerate(recs, 1):
            color = "#C62828" if "PRIORIDAD ALTA" in rec else "#E65100" if i == 1 else "#1565C0"
            icon  = "🚨" if "PRIORIDAD ALTA" in rec else "⚠️" if i <= 2 else "💡"
            st.markdown(f"""
            <div style="background:{color}08; border-left:4px solid {color};
                        border-radius:6px; padding:14px 18px; margin-bottom:10px;">
                <span style="font-size:16px;">{icon}</span>
                <span style="font-weight:600; color:{color}; margin-left:6px;">Recomendación {i}</span>
                <p style="margin:6px 0 0 0; color:#333; line-height:1.5;">{rec}</p>
            </div>""", unsafe_allow_html=True)
    else:
        st.success("✅ No se detectaron problemas críticos que recomienden acción inmediata.")

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

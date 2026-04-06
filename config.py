# config.py — Constantes y configuración central

# ── Política de ingreso ───────────────────────────────────────────────────────
POLICY_DAYS = 5        # Días máximos permitidos entre fecha pedido y Fe.contab.
GREEN_MAX   = 3        # 1–3 días: VERDE
YELLOW_MAX  = 5        # 4–5 días: AMARILLO  (> YELLOW_MAX = ROJO)

# ── Tipos de movimiento SAP ───────────────────────────────────────────────────
MOV_ENTRY  = 101
MOV_CANCEL = 102

# ── Mapeo de código de almacén → origen legible ───────────────────────────────
# Ajusta los códigos según tu sistema SAP.
# También se usa búsqueda por substring en el campo Proveedor/Centro suministrador.
ORIGIN_MAP_WAREHOUSE = {
    # Ejemplos; reemplaza con los códigos reales:
    # "HB01": "La Estrella",
    # "HB02": "Taller",
}
ORIGIN_KEYWORDS = {
    "estrella": "La Estrella",
    "taller":   "Taller",
}
ORIGIN_UNKNOWN = "Desconocido"

# ── Alias de columnas MB51 ────────────────────────────────────────────────────
# Cada clave es el nombre lógico interno; los valores son posibles nombres en el CSV.
MB51_COLS = {
    "material":         ["Material"],
    "description":      ["Texto breve de material", "Texto breve mat.", "Texto breve"],
    "movement":         ["CMv", "Cl.mov.", "TipoMov", "Tipo mov.", "Clase de movimiento"],
    "warehouse":        ["Alm.", "Almacén", "Almacen"],
    "order":            ["Pedido"],
    "position":         ["Pos.", "Pos"],
    "doc_mat":          ["Doc.mat.", "Documento material"],
    "quantity":         ["Cantidad"],
    "unit":             ["UMB"],
    "doc_date":         ["Fecha doc.", "Fecha documento"],
    "accounting_date":  ["Fe.contab.", "Fecha contab.", "Fecha contabilización"],
    "hour":             ["Hora"],
    "amount":           ["Importe ML", "Importe en moneda local"],
    "doc_header_text":  ["Texto cab.documento", "Texto cab. documento"],
    "reference":        ["Referencia"],
    "user":             ["Usuario"],
    "reservation":      ["Reserva"],
}

# Columnas MB51 sin las cuales no se puede procesar la fila (se marcan con advertencia)
MB51_KEY_COLS = ["order", "position", "movement"]
MB51_DATE_COLS = ["accounting_date", "doc_date"]

# ── Alias de columnas Archivo de Pedidos ─────────────────────────────────────
PO_COLS = {
    "material":    ["Material"],
    "description": ["Texto breve", "Texto breve de material", "Texto breve mat."],
    "quantity":    ["Cantidad"],
    "pending":     ["Por entrg.", "Por entregar", "Pte. entrg."],
    "unit":        ["UMP", "UMB", "UM"],
    "doc_date":    ["Fecha doc.", "Fecha documento"],
    "order":       ["Doc.compr.", "Pedido", "Documento de compras"],
    "position":    ["Pos.", "Pos"],
    "supplier":    ["Proveedor/Centro suministrador", "Proveedor", "Centro suministrador"],
    "currency":    ["Mon.", "Moneda"],
}

# Columnas de pedido obligatorias
PO_KEY_COLS = ["order", "position", "quantity", "doc_date"]

# ── Nombres de hojas del Excel de salida ─────────────────────────────────────
OUTPUT_SHEETS = {
    "kpis":            "Resumen KPIs",
    "detail":          "Detalle por Línea",
    "top_materials":   "Top Materiales",
    "top_users":       "Usuarios Demora",
    "trend":           "Tendencia",
    "pending":         "Pendientes",
    "cancellations":   "Cancelaciones",
    "recommendations": "Recomendaciones",
}

# ── Colores para el Excel (hexadecimal ARGB sin #) ───────────────────────────
COLOR_GREEN        = "FF00B050"   # Verde oportuno
COLOR_YELLOW       = "FFFFC000"   # Amarillo límite
COLOR_RED          = "FFFF0000"   # Rojo vencido
COLOR_NO_ENTRY     = "FFD9D9D9"   # Gris sin entrada
COLOR_WARNING_ROW  = "FFFFD966"   # Amarillo claro para filas con advertencia
COLOR_HEADER       = "FF2F5597"   # Azul oscuro para cabeceras
COLOR_HEADER_FONT  = "FFFFFFFF"   # Texto blanco
COLOR_HEADER_ORANGE= "FFE36209"   # Naranja para cabeceras de recomendaciones
COLOR_SUBHEADER    = "FFD6E4BC"   # Verde claro para subcabeceras
COLOR_ALT_ROW      = "FFF2F2F2"   # Gris muy claro para filas alternas

# ── Alertas ───────────────────────────────────────────────────────────────────
ALERT_GREEN    = "VERDE"
ALERT_YELLOW   = "AMARILLO"
ALERT_RED      = "ROJO"
ALERT_NO_ENTRY = "SIN ENTRADA"
ALERT_NO_DATE  = "SIN FECHA"

ALERT_COLOR_MAP = {
    ALERT_GREEN:    COLOR_GREEN,
    ALERT_YELLOW:   COLOR_YELLOW,
    ALERT_RED:      COLOR_RED,
    ALERT_NO_ENTRY: COLOR_NO_ENTRY,
    ALERT_NO_DATE:  COLOR_WARNING_ROW,
}

# ── Parámetros de análisis ────────────────────────────────────────────────────
TOP_MATERIALS_N = 10
TOP_USERS_N     = 5
TOP_PENDING_N   = 10
MAX_RECOMMENDATIONS = 5

# ── Máx ancho de columna en Excel (caracteres) ────────────────────────────────
MAX_COL_WIDTH = 55

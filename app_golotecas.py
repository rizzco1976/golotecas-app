import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import io
import re
from pathlib import Path

# ── Configuración ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Carteras · Análisis de Ventas",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Estilos ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.main { background-color: #0f1117; }

.metric-card {
    background: linear-gradient(135deg, #1a1d2e 0%, #16192b 100%);
    border: 1px solid #2a2d3e;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 0.5rem;
}
.metric-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #6b7280;
    margin-bottom: 0.3rem;
}
.metric-value {
    font-size: 2rem;
    font-weight: 700;
    color: #f9fafb;
    line-height: 1;
}
.metric-delta {
    font-size: 0.78rem;
    margin-top: 0.25rem;
}
.delta-pos { color: #34d399; }
.delta-neg { color: #f87171; }
.delta-neu { color: #9ca3af; }

.section-title {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #6366f1;
    padding: 0.5rem 0 0.8rem 0;
    border-bottom: 1px solid #1f2937;
    margin-bottom: 1rem;
}

.producto-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.55rem 0.8rem;
    border-radius: 8px;
    margin-bottom: 0.3rem;
    background: #1a1d2e;
    border-left: 3px solid #6366f1;
}
.producto-nombre {
    font-size: 0.78rem;
    color: #d1d5db;
    flex: 1;
}
.producto-valor {
    font-size: 0.85rem;
    font-weight: 600;
    color: #a5b4fc;
    margin-left: 1rem;
    white-space: nowrap;
}

.proyeccion-card {
    background: linear-gradient(135deg, #1e1b4b 0%, #1a1d2e 100%);
    border: 1px solid #4338ca;
    border-radius: 12px;
    padding: 1.4rem;
    margin-bottom: 1rem;
}
.proyeccion-title {
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #818cf8;
    margin-bottom: 0.8rem;
}
.proyeccion-valor {
    font-size: 2.4rem;
    font-weight: 700;
    color: #a5b4fc;
}
.proyeccion-rango {
    font-size: 0.78rem;
    color: #6b7280;
    margin-top: 0.2rem;
}

.tag-verde {
    display: inline-block;
    background: rgba(52,211,153,0.15);
    color: #34d399;
    font-size: 0.68rem;
    font-weight: 600;
    padding: 0.15rem 0.55rem;
    border-radius: 20px;
    letter-spacing: 0.05em;
}
.tag-rojo {
    display: inline-block;
    background: rgba(248,113,113,0.15);
    color: #f87171;
    font-size: 0.68rem;
    font-weight: 600;
    padding: 0.15rem 0.55rem;
    border-radius: 20px;
    letter-spacing: 0.05em;
}
.tag-amarillo {
    display: inline-block;
    background: rgba(251,191,36,0.15);
    color: #fbbf24;
    font-size: 0.68rem;
    font-weight: 600;
    padding: 0.15rem 0.55rem;
    border-radius: 20px;
    letter-spacing: 0.05em;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #0d1018;
    border-right: 1px solid #1f2937;
}

/* Ocultar elementos default de Streamlit */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ── Constantes ─────────────────────────────────────────────────────────────────
MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
         'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

TRIMESTRE_DE_MES = {
    'Enero': 1, 'Febrero': 1, 'Marzo': 1,
    'Abril': 2, 'Mayo': 2, 'Junio': 2,
    'Julio': 3, 'Agosto': 3, 'Septiembre': 3,
    'Octubre': 4, 'Noviembre': 4, 'Diciembre': 4,
}

# ── Funciones de carga y parseo ────────────────────────────────────────────────
# Formato único (2025+2026 en una sola hoja, una hoja por cartera) usado por
# Info_Franco.xlsx: tabla dinámica de Excel con fila de años, fila de meses,
# y debajo cuenta (10 dígitos) -> negocio (2 dígitos) -> categoría (4 dígitos) -> producto.
NEGOCIO_PATTERN = re.compile(r'^\d{2} - ')
CATEGORIA_PATTERN = re.compile(r'^\d{4} - ')
CUENTA_PATTERN = re.compile(r'^\d{10}$')
PRODUCTO_PATTERN = re.compile(r'^(.*)\s+\((\d+)\)$')
NOTA_PARCIAL_PATTERN = re.compile(r'(\d{4}):\s*Info hasta\s*(\d{1,2})/(\d{1,2})')


def leer_filas(ws, max_vacias=300):
    """
    Lee las filas de una hoja cortando la 'cola' de filas vacías que Excel arrastra
    en las tablas dinámicas: este archivo declara ~261.000 filas cuando los datos
    reales terminan cerca de la 4.500. Sin este corte cada carga tarda ~40 s.
    """
    filas = []
    vacias = 0
    for row in ws.iter_rows(values_only=True):
        if all(c is None for c in row):
            vacias += 1
            if vacias >= max_vacias and filas:
                break
            filas.append(row)
            continue
        vacias = 0
        filas.append(row)
    return filas


def encontrar_columnas_mes(rows):
    """
    Ubica automáticamente la fila de encabezado (la que arranca con 'Etiquetas
    de fila') sin importar el offset vertical -que varía entre hojas del mismo
    archivo (ej. Medina trae una fila extra 'Cliente')-, lee la fila de año
    2 filas arriba y arma el mapa columna -> (año, mes).
    """
    header_idx = None
    for i, row in enumerate(rows):
        if row and row[0] == 'Etiquetas de fila':
            header_idx = i
            break
    if header_idx is None:
        return None, {}

    fila_meses = rows[header_idx]
    fila_anio = rows[header_idx - 2] if header_idx >= 2 else []

    # Forward-fill del año a lo largo de las columnas (Excel sólo lo escribe 1 vez por bloque)
    anio_por_col = {}
    anio_actual = None
    for i in range(len(fila_meses)):
        val = fila_anio[i] if i < len(fila_anio) else None
        if val is not None and re.match(r'^(19|20)\d{2}$', str(val).strip()):
            anio_actual = int(str(val).strip())
        anio_por_col[i] = anio_actual

    col_mes = {}
    for i, val in enumerate(fila_meses):
        if val in MESES:
            anio = anio_por_col.get(i)
            if anio:
                col_mes[i] = (anio, val)
    return header_idx, col_mes


def parsear_cartera_sheet(rows, cartera_nombre, col_mes, header_idx):
    """Parsea una hoja (una cartera) ya localizado su encabezado de meses."""
    registros = []
    cuenta_actual = None
    negocio_actual = None
    categoria_actual = None

    for row in rows[header_idx + 1:]:
        col_a = row[0]
        if col_a is None:
            continue
        col_a_str = str(col_a).strip()

        # Detectar cuenta: código puramente numérico de 10 dígitos
        if CUENTA_PATTERN.match(col_a_str):
            cuenta_actual = col_a_str.lstrip('0') or '0'
            negocio_actual = None
            categoria_actual = None
            continue

        if not cuenta_actual:
            continue

        # Detectar negocio (2 dígitos) o categoría (4 dígitos)
        if NEGOCIO_PATTERN.match(col_a_str):
            negocio_actual = col_a_str
            categoria_actual = None
            continue
        if CATEGORIA_PATTERN.match(col_a_str):
            categoria_actual = col_a_str
            continue

        # Detectar producto: termina en "(código)"
        m = PRODUCTO_PATTERN.match(col_a_str)
        if m:
            producto_limpio = m.group(1).strip()
            codigo = int(m.group(2))
            for col_idx, (anio, mes) in col_mes.items():
                if col_idx >= len(row):
                    continue
                val = row[col_idx]
                cantidad = float(val) if val else 0.0
                if cantidad == 0.0:
                    continue
                registros.append({
                    'cartera': cartera_nombre,
                    'cuenta': f'cuenta {cuenta_actual}',
                    'codigo': codigo,
                    'producto': col_a_str,
                    'producto_limpio': producto_limpio,
                    'negocio': negocio_actual,
                    'categoria': categoria_actual,
                    'anio': anio,
                    'mes': mes,
                    'cantidad': cantidad
                })
    return pd.DataFrame(registros)


def detectar_mes_parcial(rows):
    """
    Busca en una hoja resumen ('Por negocio') una nota tipo '2026: Info hasta
    13/09' y devuelve (año, nombre_mes) del mes incompleto, o None si no hay.
    """
    for row in rows:
        for cell in row:
            if isinstance(cell, str):
                m = NOTA_PARCIAL_PATTERN.search(cell)
                if m:
                    anio = int(m.group(1))
                    mes_idx = int(m.group(3)) - 1  # formato dd/mm -> el mes es el 2do número
                    if 0 <= mes_idx < 12:
                        return anio, MESES[mes_idx]
    return None


@st.cache_data(show_spinner=False)
def cargar_info_franco(archivo_bytes):
    """
    Parsea el Excel consolidado multi-cartera (una hoja por cartera, cada una
    seguida opcionalmente de su hoja resumen 'Por negocio' / 'Por negocio (n)').
    Devuelve (df_todas_las_carteras, notas_parciales), donde notas_parciales es
    {cartera: mes_num_incompleto} para excluir ese mes de las proyecciones.
    """
    wb = openpyxl.load_workbook(io.BytesIO(archivo_bytes), read_only=True, data_only=True)
    nombres = wb.sheetnames

    dfs = []
    notas_parciales = {}

    for i, nombre in enumerate(nombres):
        if nombre.strip().lower().startswith('por negocio'):
            continue  # se procesa como complemento de la cartera anterior

        ws = wb[nombre]
        rows = leer_filas(ws)
        header_idx, col_mes = encontrar_columnas_mes(rows)
        if header_idx is None or not col_mes:
            continue

        cartera_nombre = nombre.strip()
        df_cartera = parsear_cartera_sheet(rows, cartera_nombre, col_mes, header_idx)
        if not df_cartera.empty:
            dfs.append(df_cartera)

        # La hoja resumen de esta cartera (si existe) trae la nota de mes parcial
        if i + 1 < len(nombres) and nombres[i + 1].strip().lower().startswith('por negocio'):
            rows_neg = leer_filas(wb[nombres[i + 1]])
            nota = detectar_mes_parcial(rows_neg)
            if nota:
                anio_p, mes_p = nota
                notas_parciales[cartera_nombre] = anio_p * 12 + MESES.index(mes_p)

    if not dfs:
        return pd.DataFrame(), {}

    df = pd.concat(dfs, ignore_index=True)
    # mes_num continuo y comparable entre años: 2025*12+idx < 2026*12+idx, cronológico siempre
    df['mes_num'] = df['anio'] * 12 + df['mes'].apply(lambda m: MESES.index(m))
    return df, notas_parciales


def limpiar_nombre(nombre):
    """Limpia el nombre del producto removiendo el código al final."""
    import re
    return re.sub(r'\s*\(\d+\)\s*$', '', nombre).strip()


FORMATO_PATTERNS = {
    'Display': re.compile(r'DISPLAY'),
    'Peso (Kg/Grs)': re.compile(r'\bKG\b|X\s*[\d.,]+\s*GRS?\.?\b|X\s*[\d.,]+\s*G\.?\s*$|\d+\s*GS\b|\d+GR\b'),
    'Unidad suelta': re.compile(r'X\s*\d+\s*U(?:NID)?(?:ADES)?\.?\b|\d+\s*U\.?\s*$'),
    'Volumen (Cc/Lt)': re.compile(r'\bCC\b|\bML\b|\bLTS?\b'),
}


def clasificar_formato(nombre):
    """
    Heurística basada en el texto del nombre del producto para inferir su formato
    de venta (display, peso, unidad suelta, volumen). No es un dato garantizado por
    Arcor — es una inferencia de texto, y ~40% de los nombres vienen abreviados por
    el sistema origen y quedan 'Sin clasificar'.
    """
    n = nombre.upper()
    for formato, patron in FORMATO_PATTERNS.items():
        if patron.search(n):
            return formato
    return 'Sin clasificar'


FORMATO_COLOR = {
    'Display': '#818cf8',
    'Peso (Kg/Grs)': '#34d399',
    'Unidad suelta': '#fbbf24',
    'Volumen (Cc/Lt)': '#60a5fa',
    'Sin clasificar': '#6b7280',
}

# Excepciones confirmadas manualmente: cuántas unidades físicas trae cada bulto.
# A diferencia de FORMATO (heurística de texto), esto es un dato exacto confirmado,
# clave para convertir BU -> unidades reales sin sesgo.
PACK_OVERRIDE = {
    15001: 21,   # ALF CHOCO CREAMY 21X
    13757: 21,   # ALF MINITORTA DARK 2 (21X, nombre truncado)
    13357: 21,   # ALF. MINITORTA BCO 2 (21X, nombre truncado)
    13359: 21,   # ALF. MINITORTA BRO 2 (21X, nombre truncado)
    13358: 21,   # ALF. MINITORTA CLA 2 (21X, nombre truncado)
    13360: 21,   # ALF. MINITORTA COC 2 (21X, nombre truncado)
    15489: 21,   # ALF.GOAT NEGRO 21 X 75GR
    6596: 21,    # ALFAJOR COFLER BLOCK 21 x 60 G
    14884: 21,   # BOB TRIPLE 21X60G
    15274: 21,   # BOB TRIPLE BLANCO X60 GRS (21X, confirmado)
}


def badge_pack(codigo, cantidad_bultos):
    """Si el código tiene un pack confirmado, devuelve HTML con la conversión a unidades reales."""
    unidades_x_bulto = PACK_OVERRIDE.get(codigo)
    if not unidades_x_bulto:
        return ''
    total_unidades = cantidad_bultos * unidades_x_bulto
    return f' <span style="color:#f472b6; font-size:0.65rem; font-weight:600;">· ×{unidades_x_bulto}/bulto = {total_unidades:,.0f} uds.</span>'


def etiqueta_mes(mes_num):
    """Convierte mes_num continuo (año*12+idx_mes) a etiqueta corta tipo 'Ene25'."""
    anio = mes_num // 12
    mes_idx = mes_num % 12
    return f"{MESES[mes_idx][:3]}{anio % 100:02d}"


def etiqueta_mes_marcada(mes_num, mes_num_parcial):
    """Igual que etiqueta_mes, pero agrega '*' si es el mes en curso (parcial)."""
    base = etiqueta_mes(mes_num)
    return f"{base}*" if mes_num_parcial is not None and mes_num == mes_num_parcial else base


def proyectar_proximo_pedido(serie_mensual):
    """
    Proyección simple pero robusta:
    - Promedio móvil de los últimos 3 meses con datos
    - Con intervalo ±1 desvío estándar
    """
    valores = [v for v in serie_mensual if v > 0]
    if len(valores) == 0:
        return 0, 0, 0
    recientes = valores[-3:]
    prom = np.mean(recientes)
    std = np.std(recientes) if len(recientes) > 1 else prom * 0.2
    return round(prom, 2), round(max(0, prom - std), 2), round(prom + std, 2)


def calcular_tendencia(serie):
    """Retorna pendiente normalizada de tendencia (-1 a 1)."""
    valores = [v for v in serie if v > 0]
    if len(valores) < 2:
        return 0
    x = np.arange(len(valores))
    coef = np.polyfit(x, valores, 1)
    media = np.mean(valores) if np.mean(valores) != 0 else 1
    return coef[0] / media


def identificar_oportunidades(df_cuenta, df_cadena):
    """
    Identifica productos con potencial de crecimiento en un local:
    - El local compra menos que el promedio del RESTO de la cadena (excluyéndose a sí mismo)
    """
    cuenta_actual = df_cuenta['cuenta'].iloc[0] if not df_cuenta.empty else None

    # Total por producto en el local
    local_total = df_cuenta.groupby(['codigo','producto'])['cantidad'].sum().reset_index()
    local_total.columns = ['codigo','producto','total_local']

    # Cadena SIN el local actual, para que el promedio no se calcule incluyéndose a sí mismo
    df_resto = df_cadena[df_cadena['cuenta'] != cuenta_actual] if cuenta_actual else df_cadena
    n_cuentas_resto = df_resto['cuenta'].nunique()
    if n_cuentas_resto == 0:
        return pd.DataFrame(columns=['codigo','producto','total_local','promedio_cadena','ratio','gap'])

    cadena_total = df_resto.groupby(['codigo','producto'])['cantidad'].sum().reset_index()
    cadena_total['promedio_cadena'] = cadena_total['cantidad'] / n_cuentas_resto
    cadena_total = cadena_total[['codigo','producto','promedio_cadena']]

    merged = local_total.merge(cadena_total, on=['codigo','producto'])
    merged = merged[merged['promedio_cadena'] > 0]
    merged['ratio'] = merged['total_local'] / merged['promedio_cadena']
    merged['gap'] = merged['promedio_cadena'] - merged['total_local']

    # Solo productos con gap positivo (local compra menos que el resto de la cadena)
    oportunidades = merged[merged['gap'] > 0.5].sort_values('gap', ascending=False)
    return oportunidades


# ── Sidebar ────────────────────────────────────────────────────────────────────
APP_VERSION = "v12 · multi-cartera (Golotecas / Open 25 / Medina) · 2026-09-23"

# Cuentas de ruido a excluir, por cartera (volumen insignificante, no son locales reales)
RUIDO_POR_CARTERA = {
    'Golotecas': ['cuenta 3080021'],
}

with st.sidebar:
    st.markdown("## 🗂️ Carteras · Ventas")
    st.caption(f"🔖 {APP_VERSION}")
    st.markdown("---")

    archivo = st.file_uploader(
        "Excel consolidado de carteras",
        type=["xlsx"],
        help="Archivo con una hoja por cartera (ej. Golotecas, Open 25, Medina), formato tabla dinámica mensual con 2025 y 2026"
    )

    cartera_sel = None
    vista = None
    df_todas = pd.DataFrame()
    notas_parciales = {}

    if archivo:
        with st.spinner("Procesando datos..."):
            df_todas, notas_parciales = cargar_info_franco(archivo.read())

        if df_todas.empty:
            st.error("No se pudieron leer datos del archivo. Verificá el formato.")
        else:
            st.success("✓ Archivo cargado")
            CARTERAS = sorted(df_todas['cartera'].unique())

            st.markdown("---")
            cartera_sel = st.selectbox("Cartera", CARTERAS)
            _df_sel = df_todas[df_todas['cartera'] == cartera_sel]
            # Mismo criterio que el resto de la app: sin las cuentas de ruido
            _df_sel = _df_sel[~_df_sel['cuenta'].isin(RUIDO_POR_CARTERA.get(cartera_sel, []))]
            _anios = " + ".join(str(a) for a in sorted(_df_sel['anio'].unique()))
            st.caption(f"📅 Datos {_anios} · {_df_sel['cuenta'].nunique()} cuentas en esta cartera")

            st.markdown("---")
            st.markdown("##### Navegación")

            opciones_vista = ["📊 Resumen cadena", "🏪 Análisis por local", "🎯 Próximo pedido", "📈 Comparativa locales", "🗂️ Por negocio", "🗓️ Análisis trimestral"]

            vista = st.radio(
                "Vista",
                opciones_vista,
                label_visibility="collapsed"
            )

# ── Main content ───────────────────────────────────────────────────────────────
if not archivo:
    st.markdown("# 🗂️ Carteras · Análisis de Ventas")
    st.markdown("---")
    col1, col2 = st.columns([2,1])
    with col1:
        st.markdown("""
        ### Cargá el archivo Excel para comenzar

        Esta herramienta analiza las ventas de tus carteras de clientes (cada
        hoja del Excel es una cartera, ej. Golotecas, Open 25, Medina) y te da,
        para la cartera que elijas en el panel izquierdo:

        - **Proyección del próximo pedido** por local, con intervalo de confianza
        - **Productos a empujar** en cada local según su comportamiento vs. el resto de la cartera
        - **Comparativa entre locales** para detectar oportunidades
        - **Tendencias mensuales y trimestrales** por categoría y producto

        Usá el panel izquierdo para subir el Excel consolidado.
        """)
    st.stop()

if df_todas.empty:
    st.stop()  # el error ya se mostró en el sidebar

# Filtrar a la cartera seleccionada. A partir de acá, df_raw y CUENTAS quedan
# con los mismos nombres/estructura que antes, así que el resto de las vistas
# funciona sin cambios sobre esta porción ya filtrada.
df_raw = df_todas[df_todas['cartera'] == cartera_sel].copy()

cuentas_ruido = RUIDO_POR_CARTERA.get(cartera_sel, [])
if cuentas_ruido:
    df_raw = df_raw[~df_raw['cuenta'].isin(cuentas_ruido)]

df_raw['trimestre'] = df_raw['mes'].map(TRIMESTRE_DE_MES)
df_raw['trimestre_label'] = df_raw['anio'].astype(str) + ' T' + df_raw['trimestre'].astype(str)

df_raw['formato'] = df_raw['producto_limpio'].apply(clasificar_formato)

# Mes en curso (incompleto) para esta cartera, si el archivo trae la nota -
# se excluye de las proyecciones de próximo pedido y se marca con "*" en los gráficos.
MES_NUM_PARCIAL = notas_parciales.get(cartera_sel)

# El trimestre que contiene ese mes queda incompleto: no es comparable contra el
# mismo trimestre del año anterior sin igualar los meses.
if MES_NUM_PARCIAL is not None:
    ANIO_PARCIAL = MES_NUM_PARCIAL // 12
    MES_PARCIAL = MESES[MES_NUM_PARCIAL % 12]
    TRIMESTRE_PARCIAL = TRIMESTRE_DE_MES[MES_PARCIAL]
    TRIMESTRE_PARCIAL_LABEL = f"{ANIO_PARCIAL} T{TRIMESTRE_PARCIAL}"
else:
    ANIO_PARCIAL = MES_PARCIAL = TRIMESTRE_PARCIAL = TRIMESTRE_PARCIAL_LABEL = None

CUENTAS = sorted(df_raw['cuenta'].unique())

# ══════════════════════════════════════════════════════════════════════════════
# VISTA 1: RESUMEN CADENA
# ══════════════════════════════════════════════════════════════════════════════
if vista == "📊 Resumen cadena":
    st.markdown(f"# Resumen · {cartera_sel}")
    st.markdown("---")

    # Métricas globales
    total_cadena = df_raw['cantidad'].sum()
    total_productos = df_raw['codigo'].nunique()
    total_locales = df_raw['cuenta'].nunique()

    # Mes con más ventas
    por_mes = df_raw.groupby('mes_num')['cantidad'].sum()
    mes_pico_num = por_mes.idxmax()
    mes_pico = etiqueta_mes(mes_pico_num)
    valor_pico = por_mes[mes_pico_num]

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total vendido</div>
            <div class="metric-value">{total_cadena:,.0f}</div>
            <div class="metric-delta delta-neu">BU · {'-'.join(str(a) for a in sorted(df_raw['anio'].unique()))}</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Locales activos</div>
            <div class="metric-value">{total_locales}</div>
            <div class="metric-delta delta-neu">cuentas en {cartera_sel}</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Productos distintos</div>
            <div class="metric-value">{total_productos:,}</div>
            <div class="metric-delta delta-neu">SKUs únicos</div>
        </div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Mes pico</div>
            <div class="metric-value">{mes_pico}</div>
            <div class="metric-delta delta-pos">{valor_pico:,.0f} BU</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.markdown('<div class="section-title">Evolución mensual · Cadena completa</div>', unsafe_allow_html=True)
        evol = df_raw.groupby('mes_num')['cantidad'].sum().reset_index()
        evol['mes'] = evol['mes_num'].apply(lambda m: etiqueta_mes_marcada(m, MES_NUM_PARCIAL))
        # Orden cronológico forzado (Streamlit/Altair ordena texto alfabéticamente por defecto)
        evol['mes'] = pd.Categorical(evol['mes'], categories=evol['mes'].tolist(), ordered=True)
        st.bar_chart(evol.set_index('mes')['cantidad'], color="#6366f1", height=240)
        if MES_NUM_PARCIAL is not None and MES_NUM_PARCIAL in evol['mes_num'].values:
            st.caption("* mes en curso, con datos parciales")

    with col_right:
        st.markdown('<div class="section-title">Volumen por local</div>', unsafe_allow_html=True)
        por_local = df_raw.groupby('cuenta')['cantidad'].sum().sort_values(ascending=False)
        por_local.index = [c.replace('cuenta ', '') for c in por_local.index]
        st.bar_chart(por_local, color="#818cf8", height=240)

    # Top productos cadena
    st.markdown('<div class="section-title">Top 15 productos · Cadena</div>', unsafe_allow_html=True)
    top_prods = (df_raw.groupby(['codigo','producto_limpio'])['cantidad']
                 .sum().reset_index()
                 .sort_values('cantidad', ascending=False)
                 .head(15))

    for _, row in top_prods.iterrows():
        pct = row['cantidad'] / total_cadena * 100
        fmt = clasificar_formato(row['producto_limpio'])
        color_fmt = FORMATO_COLOR[fmt]
        pack_html = badge_pack(row['codigo'], row['cantidad'])
        st.markdown(f"""
        <div class="producto-row">
            <span class="producto-nombre">{row['producto_limpio'][:50]} <span style="color:{color_fmt}; font-size:0.65rem; font-weight:600;">· {fmt}</span>{pack_html}</span>
            <span class="producto-valor">{row['cantidad']:,.0f} BU &nbsp;·&nbsp; {pct:.1f}%</span>
        </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# VISTA 2: ANÁLISIS POR LOCAL
# ══════════════════════════════════════════════════════════════════════════════
elif vista == "🏪 Análisis por local":
    st.markdown("# Análisis · Por local")
    st.markdown("---")

    cuenta_sel = st.selectbox(
        "Seleccioná el local",
        CUENTAS,
        format_func=lambda x: x.replace('cuenta ', 'Local ')
    )

    df_local = df_raw[df_raw['cuenta'] == cuenta_sel]

    # Métricas del local
    total_local = df_local['cantidad'].sum()
    total_cadena = df_raw['cantidad'].sum()
    n_locales = df_raw['cuenta'].nunique()
    promedio_cadena = total_cadena / n_locales
    participacion = total_local / total_cadena * 100
    vs_promedio = (total_local - promedio_cadena) / promedio_cadena * 100

    col1, col2, col3 = st.columns(3)
    with col1:
        signo = "+" if vs_promedio >= 0 else ""
        clase = "delta-pos" if vs_promedio >= 0 else "delta-neg"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total local</div>
            <div class="metric-value">{total_local:,.0f}</div>
            <div class="metric-delta {clase}">{signo}{vs_promedio:.1f}% vs promedio cadena</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Participación en cadena</div>
            <div class="metric-value">{participacion:.1f}%</div>
            <div class="metric-delta delta-neu">del total cadena</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        prods_activos = (df_local.groupby('codigo')['cantidad'].sum() > 0).sum()
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Productos activos</div>
            <div class="metric-value">{prods_activos}</div>
            <div class="metric-delta delta-neu">con al menos 1 pedido</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.markdown(f'<div class="section-title">Evolución mensual · {cuenta_sel.replace("cuenta ", "Local ")}</div>', unsafe_allow_html=True)
        evol_local = df_local.groupby('mes_num')['cantidad'].sum().reset_index()
        evol_local['mes'] = evol_local['mes_num'].apply(lambda m: etiqueta_mes_marcada(m, MES_NUM_PARCIAL))
        evol_local['mes'] = pd.Categorical(evol_local['mes'], categories=evol_local['mes'].tolist(), ordered=True)
        st.bar_chart(evol_local.set_index('mes')['cantidad'], color="#6366f1", height=220)

    with col_right:
        st.markdown('<div class="section-title">Top 5 categorías</div>', unsafe_allow_html=True)
        # Extraer categoría del nombre de producto usando el codigo (aproximado)
        top5 = (df_local.groupby('producto_limpio')['cantidad']
                .sum().sort_values(ascending=False).head(5))
        for prod, val in top5.items():
            pct = val / total_local * 100
            st.markdown(f"""
            <div class="producto-row">
                <span class="producto-nombre">{prod[:40]}</span>
                <span class="producto-valor">{val:.0f} · {pct:.0f}%</span>
            </div>""", unsafe_allow_html=True)

    # Oportunidades
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">Productos con oportunidad · Este local compra menos que el promedio de la cadena</div>', unsafe_allow_html=True)

    oport = identificar_oportunidades(df_local, df_raw)
    if not oport.empty:
        oport['producto_limpio'] = oport['producto'].apply(limpiar_nombre)
        for _, row in oport.head(10).iterrows():
            gap = row['gap']
            ratio = row['ratio']
            if ratio < 0.3:
                tag = '<span class="tag-rojo">Muy por debajo</span>'
            elif ratio < 0.7:
                tag = '<span class="tag-amarillo">Por debajo</span>'
            else:
                tag = '<span class="tag-verde">Leve oportunidad</span>'

            st.markdown(f"""
            <div class="producto-row">
                <span class="producto-nombre">{row['producto_limpio'][:50]} &nbsp; {tag}</span>
                <span class="producto-valor">gap {gap:.1f} BU</span>
            </div>""", unsafe_allow_html=True)
    else:
        st.info("Este local está por encima del promedio en todos los productos activos.")


# ══════════════════════════════════════════════════════════════════════════════
# VISTA 3: PRÓXIMO PEDIDO
# ══════════════════════════════════════════════════════════════════════════════
elif vista == "🎯 Próximo pedido":
    st.markdown("# Proyección · Próximo pedido")
    st.markdown("---")

    cuenta_sel = st.selectbox(
        "Seleccioná el local",
        CUENTAS,
        format_func=lambda x: x.replace('cuenta ', 'Local ')
    )

    df_local = df_raw[df_raw['cuenta'] == cuenta_sel]

    # Proyección total del local
    # Se excluye el mes en curso (parcial) para no subestimar la proyección con un mes incompleto
    meses_disp = [m for m in sorted(df_raw['mes_num'].unique()) if m != MES_NUM_PARCIAL]
    serie_total = [df_local[df_local['mes_num'] == m]['cantidad'].sum() for m in meses_disp]
    proy_total, low, high = proyectar_proximo_pedido(serie_total)

    st.markdown(f"""
    <div class="proyeccion-card">
        <div class="proyeccion-title">Proyección próximo pedido · {cuenta_sel.replace("cuenta ", "Local ")}</div>
        <div class="proyeccion-valor">{proy_total:,.0f} BU</div>
        <div class="proyeccion-rango">Rango estimado: {low:,.0f} – {high:,.0f} BU &nbsp;·&nbsp; Basado en promedio móvil 3 meses</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Proyección por producto (los más relevantes)
    st.markdown('<div class="section-title">Proyección por producto · Top 20 por volumen</div>', unsafe_allow_html=True)

    top_prods = (df_local.groupby(['codigo','producto_limpio'])['cantidad']
                 .sum().reset_index()
                 .sort_values('cantidad', ascending=False)
                 .head(20))

    col1, col2 = st.columns(2)
    for idx, (_, prod_row) in enumerate(top_prods.iterrows()):
        df_prod = df_local[df_local['codigo'] == prod_row['codigo']]
        serie = [df_prod[df_prod['mes_num'] == m]['cantidad'].sum() for m in meses_disp]
        proy, lo, hi = proyectar_proximo_pedido(serie)

        # Calcular tendencia
        tend = calcular_tendencia(serie)
        if tend > 0.1:
            tend_txt = '<span class="tag-verde">↑ sube</span>'
        elif tend < -0.1:
            tend_txt = '<span class="tag-rojo">↓ baja</span>'
        else:
            tend_txt = '<span class="tag-amarillo">→ estable</span>'

        contenido = f"""
        <div class="producto-row" style="border-left-color: {'#34d399' if tend > 0.1 else '#f87171' if tend < -0.1 else '#fbbf24'}">
            <span class="producto-nombre">{prod_row['producto_limpio'][:45]} &nbsp; {tend_txt}</span>
            <span class="producto-valor">{proy:.1f} BU</span>
        </div>"""

        if idx % 2 == 0:
            with col1:
                st.markdown(contenido, unsafe_allow_html=True)
        else:
            with col2:
                st.markdown(contenido, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:0.72rem; color:#4b5563; padding: 0.8rem; background:#111827; border-radius:8px;">
    📌 La proyección usa promedio móvil de los últimos 3 meses con datos. El rango refleja la variabilidad histórica reciente. 
    Para pedidos irregulares, tomá la proyección como referencia base, no como número exacto.
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# VISTA 4: COMPARATIVA LOCALES
# ══════════════════════════════════════════════════════════════════════════════
elif vista == "📈 Comparativa locales":
    st.markdown("# Comparativa · Todos los locales")
    st.markdown("---")

    # Tabla resumen por local
    resumen = []
    for cuenta in CUENTAS:
        df_c = df_raw[df_raw['cuenta'] == cuenta]
        total = df_c['cantidad'].sum()
        prods_activos = (df_c.groupby('codigo')['cantidad'].sum() > 0).sum()
        meses_disp_c = [m for m in sorted(df_raw['mes_num'].unique()) if m != MES_NUM_PARCIAL]
        serie = [df_c[df_c['mes_num'] == m]['cantidad'].sum() for m in meses_disp_c]
        proy, _, _ = proyectar_proximo_pedido(serie)
        tend = calcular_tendencia(serie)
        top_prod = (df_c.groupby('producto_limpio')['cantidad'].sum()
                    .sort_values(ascending=False).index[0] if not df_c.empty else '-')
        resumen.append({
            'Local': cuenta.replace('cuenta ', 'Local '),
            'Total': total,
            'Productos activos': prods_activos,
            'Próx. pedido (est.)': proy,
            'Tendencia': tend,
            'Producto #1': top_prod[:40]
        })

    df_resumen = pd.DataFrame(resumen).sort_values('Total', ascending=False)

    st.markdown('<div class="section-title">Ranking de locales por volumen</div>', unsafe_allow_html=True)
    for _, row in df_resumen.iterrows():
        tend = row['Tendencia']
        if tend > 0.1:
            tend_txt = '<span class="tag-verde">↑ tendencia positiva</span>'
        elif tend < -0.1:
            tend_txt = '<span class="tag-rojo">↓ tendencia negativa</span>'
        else:
            tend_txt = '<span class="tag-amarillo">→ estable</span>'

        st.markdown(f"""
        <div class="metric-card" style="margin-bottom:0.6rem;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div>
                    <div style="font-size:1rem; font-weight:700; color:#f9fafb;">{row['Local']}</div>
                    <div style="font-size:0.72rem; color:#6b7280; margin-top:0.2rem;">
                        Top producto: {row['Producto #1']} &nbsp; {tend_txt}
                    </div>
                </div>
                <div style="text-align:right;">
                    <div style="font-size:1.4rem; font-weight:700; color:#a5b4fc;">{row['Total']:,.0f} BU</div>
                    <div style="font-size:0.72rem; color:#6b7280;">
                        {row['Productos activos']} SKUs · próx. pedido ~{row['Próx. pedido (est.)']:,.0f}
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Evolución mensual superpuesta
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">Evolución mensual · Todos los locales</div>', unsafe_allow_html=True)

    evol_pivot = df_raw.groupby(['cuenta','mes_num'])['cantidad'].sum().reset_index()
    evol_pivot['cuenta'] = evol_pivot['cuenta'].str.replace('cuenta ', 'L')
    pivot = evol_pivot.pivot(index='mes_num', columns='cuenta', values='cantidad').fillna(0)
    etiquetas_orden = [etiqueta_mes_marcada(i, MES_NUM_PARCIAL) for i in pivot.index]
    pivot.index = pd.CategoricalIndex(etiquetas_orden, categories=etiquetas_orden, ordered=True)
    st.line_chart(pivot, height=300)

    # Producto estrella por local
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">Producto estrella por local</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    for idx, cuenta in enumerate(CUENTAS):
        df_c = df_raw[df_raw['cuenta'] == cuenta]
        top3 = (df_c.groupby('producto_limpio')['cantidad']
                .sum().sort_values(ascending=False).head(3))
        contenido = f"""
        <div class="metric-card" style="margin-bottom:0.4rem;">
            <div class="metric-label">{cuenta.replace('cuenta ', 'Local ')}</div>
            {''.join([f'<div style="font-size:0.78rem; color:#d1d5db; padding:0.1rem 0;">{p[:45]} <span style="color:#6366f1;">{v:.0f} BU</span></div>' for p, v in top3.items()])}
        </div>"""
        if idx % 2 == 0:
            with col1:
                st.markdown(contenido, unsafe_allow_html=True)
        else:
            with col2:
                st.markdown(contenido, unsafe_allow_html=True)



# ══════════════════════════════════════════════════════════════════════════════
# VISTA 5: POR NEGOCIO
# ══════════════════════════════════════════════════════════════════════════════
elif vista == "🗂️ Por negocio":
    st.markdown("# Movimientos · Por negocio")
    st.markdown("---")

    if 'negocio' not in df_raw.columns or df_raw['negocio'].isna().all():
        st.warning("El archivo cargado no tiene información de negocio/categoría detectada.")
        st.stop()

    # Filtro opcional por local y por formato
    col_local, col_formato = st.columns([1, 1])
    with col_local:
        filtro_local = st.selectbox(
            "Local",
            ["Cadena completa"] + CUENTAS,
            format_func=lambda x: x if x == "Cadena completa" else x.replace('cuenta ', 'Local ')
        )
    with col_formato:
        filtro_formato = st.selectbox(
            "Formato de venta",
            ["Todos"] + list(FORMATO_COLOR.keys()),
            help="Inferido del nombre del producto (heurística de texto, no un dato garantizado por Arcor)"
        )

    df_neg = df_raw if filtro_local == "Cadena completa" else df_raw[df_raw['cuenta'] == filtro_local]
    df_neg = df_neg[df_neg['negocio'].notna()]
    if filtro_formato != "Todos":
        df_neg = df_neg[df_neg['formato'] == filtro_formato]

    # Resumen por negocio
    resumen_negocio = (df_neg.groupby('negocio')['cantidad']
                       .sum().sort_values(ascending=False).reset_index())
    total_general = resumen_negocio['cantidad'].sum()

    st.markdown('<div class="section-title">Resumen · Todos los negocios</div>', unsafe_allow_html=True)
    for _, row_n in resumen_negocio.iterrows():
        pct = row_n['cantidad'] / total_general * 100 if total_general else 0
        st.markdown(f"""
        <div class="metric-card" style="margin-bottom:0.5rem;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div style="font-size:0.95rem; font-weight:700; color:#f9fafb;">{row_n['negocio']}</div>
                <div style="text-align:right;">
                    <div style="font-size:1.3rem; font-weight:700; color:#a5b4fc;">{row_n['cantidad']:,.0f} BU</div>
                    <div style="font-size:0.72rem; color:#6b7280;">{pct:.1f}% del total</div>
                </div>
            </div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Drill-down: elegir negocio y ver sus categorías + productos
    st.markdown('<div class="section-title">Detalle · Categorías y productos dentro de un negocio</div>', unsafe_allow_html=True)

    negocios_disp = sorted(df_neg['negocio'].unique())
    negocio_sel = st.selectbox("Elegí un negocio para ver el detalle", negocios_disp)

    df_negocio_sel = df_neg[df_neg['negocio'] == negocio_sel]

    # Categorías dentro del negocio elegido
    cat_resumen = (df_negocio_sel[df_negocio_sel['categoria'].notna()]
                   .groupby('categoria')['cantidad'].sum()
                   .sort_values(ascending=False).reset_index())

    total_negocio = df_negocio_sel['cantidad'].sum()

    col_cat, col_prod = st.columns(2)

    with col_cat:
        st.markdown(f'<div style="font-size:0.72rem; font-weight:700; letter-spacing:0.08em; color:#818cf8; text-transform:uppercase; margin-bottom:0.6rem;">Categorías en {negocio_sel}</div>', unsafe_allow_html=True)
        for _, row_c in cat_resumen.iterrows():
            pct_c = row_c['cantidad'] / total_negocio * 100 if total_negocio else 0
            st.markdown(f"""
            <div class="producto-row" style="border-left-color:#818cf8;">
                <span class="producto-nombre">{row_c['categoria']}</span>
                <span class="producto-valor">{row_c['cantidad']:,.0f} BU · {pct_c:.0f}%</span>
            </div>""", unsafe_allow_html=True)

    with col_prod:
        st.markdown(f'<div style="font-size:0.72rem; font-weight:700; letter-spacing:0.08em; color:#34d399; text-transform:uppercase; margin-bottom:0.6rem;">Top 15 productos en {negocio_sel}</div>', unsafe_allow_html=True)
        top_prod_negocio = (df_negocio_sel.groupby(['codigo', 'producto_limpio'])['cantidad']
                            .sum().reset_index().sort_values('cantidad', ascending=False).head(15))
        for _, r in top_prod_negocio.iterrows():
            prod, val, cod = r['producto_limpio'], r['cantidad'], r['codigo']
            pct_p = val / total_negocio * 100 if total_negocio else 0
            fmt = clasificar_formato(prod)
            color_fmt = FORMATO_COLOR[fmt]
            pack_html = badge_pack(cod, val)
            st.markdown(f"""
            <div class="producto-row" style="border-left-color:#34d399;">
                <span class="producto-nombre">{prod[:32]} <span style="color:{color_fmt}; font-size:0.62rem; font-weight:600;">· {fmt}</span>{pack_html}</span>
                <span class="producto-valor">{val:,.0f} · {pct_p:.0f}%</span>
            </div>""", unsafe_allow_html=True)

    # Filtro adicional: elegir una categoría específica y ver solo sus productos
    if not cat_resumen.empty:
        st.markdown("<br>", unsafe_allow_html=True)
        categoria_sel = st.selectbox(
            "O elegí una categoría puntual para ver todos sus productos",
            ["(ninguna)"] + list(cat_resumen['categoria'])
        )
        if categoria_sel != "(ninguna)":
            df_cat_sel = df_negocio_sel[df_negocio_sel['categoria'] == categoria_sel]
            prods_cat = (df_cat_sel.groupby(['codigo', 'producto_limpio'])['cantidad']
                        .sum().reset_index().sort_values('cantidad', ascending=False))
            st.markdown(f'<div class="section-title">Todos los productos en {categoria_sel}</div>', unsafe_allow_html=True)
            for _, r in prods_cat.iterrows():
                prod, val, cod = r['producto_limpio'], r['cantidad'], r['codigo']
                pack_html = badge_pack(cod, val)
                st.markdown(f"""
                <div class="producto-row">
                    <span class="producto-nombre">{prod[:55]}{pack_html}</span>
                    <span class="producto-valor">{val:,.0f} BU</span>
                </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# VISTA 6: ANÁLISIS TRIMESTRAL
# ══════════════════════════════════════════════════════════════════════════════
elif vista == "🗓️ Análisis trimestral":
    st.markdown("# Análisis · Trimestral y estacionalidad")
    st.markdown("---")

    col_local, _ = st.columns([1, 2])
    with col_local:
        filtro_local_tri = st.selectbox(
            "Local",
            ["Cadena completa"] + CUENTAS,
            format_func=lambda x: x if x == "Cadena completa" else x.replace('cuenta ', 'Local '),
            key="sel_trimestral"
        )

    df_tri = df_raw if filtro_local_tri == "Cadena completa" else df_raw[df_raw['cuenta'] == filtro_local_tri]

    # ── Evolución por trimestre (todos los trimestres disponibles en orden) ──
    st.markdown('<div class="section-title">Volumen por trimestre</div>', unsafe_allow_html=True)

    orden_trimestres = (df_tri[['anio', 'trimestre', 'trimestre_label']]
                        .drop_duplicates()
                        .sort_values(['anio', 'trimestre']))
    por_trimestre = df_tri.groupby('trimestre_label')['cantidad'].sum()
    por_trimestre = por_trimestre.reindex(orden_trimestres['trimestre_label'])

    # Marcar con "*" el trimestre que está incompleto por el mes en curso
    if TRIMESTRE_PARCIAL_LABEL and TRIMESTRE_PARCIAL_LABEL in por_trimestre.index:
        etiquetas_tri = [f"{t}*" if t == TRIMESTRE_PARCIAL_LABEL else t for t in por_trimestre.index]
        por_trimestre.index = pd.CategoricalIndex(etiquetas_tri, categories=etiquetas_tri, ordered=True)

    st.bar_chart(por_trimestre, color="#6366f1", height=260)

    if TRIMESTRE_PARCIAL_LABEL and TRIMESTRE_PARCIAL_LABEL in orden_trimestres['trimestre_label'].values:
        st.caption(f"* {TRIMESTRE_PARCIAL_LABEL} está incompleto: {MES_PARCIAL} viene cortado por la fecha del informe.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Comparativa año contra año, mismo trimestre ──
    st.markdown('<div class="section-title">Comparativa · Mismo trimestre, año contra año</div>', unsafe_allow_html=True)

    trimestres_presentes = sorted(df_tri['trimestre'].unique())
    anios_presentes = sorted(df_tri['anio'].unique())

    if len(anios_presentes) >= 2:
        for t in trimestres_presentes:
            df_t = df_tri[df_tri['trimestre'] == t]
            por_anio_t = df_t.groupby('anio')['cantidad'].sum()

            # Solo comparar si el trimestre tiene datos en más de un año
            if len(por_anio_t) < 2:
                continue

            anio_base = min(por_anio_t.index)
            anio_comp = max(por_anio_t.index)

            # Si el trimestre del año más reciente está cortado por el mes en curso,
            # se comparan sólo los meses completos que existen en los dos años.
            # Sin esto, un T3 con 13 días de septiembre aparenta una caída que no existe.
            nota_equiparada = ""
            if TRIMESTRE_PARCIAL is not None and t == TRIMESTRE_PARCIAL and anio_comp == ANIO_PARCIAL:
                meses_completos = [m for m in df_t[df_t['anio'] == anio_comp]['mes'].unique() if m != MES_PARCIAL]
                meses_comunes = [m for m in meses_completos if m in df_t[df_t['anio'] == anio_base]['mes'].unique()]
                if not meses_comunes:
                    continue
                df_eq = df_t[df_t['mes'].isin(meses_comunes)]
                por_anio_t = df_eq.groupby('anio')['cantidad'].sum()
                if len(por_anio_t) < 2:
                    continue
                abrev = ", ".join(m[:3] for m in sorted(meses_comunes, key=MESES.index))
                nota_equiparada = f" · comparando sólo {abrev} (se excluye {MES_PARCIAL}, mes en curso)"

            val_base = por_anio_t[anio_base]
            val_comp = por_anio_t[anio_comp]
            var_pct = (val_comp - val_base) / val_base * 100 if val_base else 0

            clase = "delta-pos" if var_pct >= 0 else "delta-neg"
            signo = "+" if var_pct >= 0 else ""
            color_barra = "#34d399" if var_pct >= 0 else "#f87171"

            st.markdown(f"""
            <div class="metric-card" style="margin-bottom:0.5rem;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div>
                        <div style="font-size:0.95rem; font-weight:700; color:#f9fafb;">Trimestre {t}</div>
                        <div style="font-size:0.72rem; color:#6b7280; margin-top:0.15rem;">
                            {anio_base} T{t}: {val_base:,.0f} BU &nbsp;→&nbsp; {anio_comp} T{t}: {val_comp:,.0f} BU{nota_equiparada}
                        </div>
                    </div>
                    <div style="text-align:right;">
                        <div style="font-size:1.4rem; font-weight:700; color:{color_barra};">{signo}{var_pct:.1f}%</div>
                    </div>
                </div>
            </div>""", unsafe_allow_html=True)
    else:
        st.info("Con un solo año de datos no hay comparativa año contra año todavía. Se habilita cuando el archivo trae dos años o más para esta cartera.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Estacionalidad por producto: elegir un trimestre y ver qué domina ──
    st.markdown('<div class="section-title">Productos que dominan cada trimestre</div>', unsafe_allow_html=True)

    trimestre_sel = st.selectbox(
        "Elegí un trimestre para ver su top de productos",
        sorted(orden_trimestres['trimestre_label'].unique(), key=lambda x: (x.split(' T')[0], x.split(' T')[1]))
    )

    df_trim_sel = df_tri[df_tri['trimestre_label'] == trimestre_sel]

    if TRIMESTRE_PARCIAL_LABEL and trimestre_sel == TRIMESTRE_PARCIAL_LABEL:
        st.caption(f"Ojo: este trimestre está incompleto ({MES_PARCIAL} viene cortado), así que los volúmenes son menores a lo que va a cerrar.")

    top_prod_trim = (df_trim_sel.groupby('producto_limpio')['cantidad']
                     .sum().sort_values(ascending=False).head(15))
    total_trim_sel = df_trim_sel['cantidad'].sum()

    for prod, val in top_prod_trim.items():
        pct = val / total_trim_sel * 100 if total_trim_sel else 0
        st.markdown(f"""
        <div class="producto-row">
            <span class="producto-nombre">{prod[:45]}</span>
            <span class="producto-valor">{val:,.0f} BU · {pct:.0f}%</span>
        </div>""", unsafe_allow_html=True)

    # ── Comparar el mix de productos de este trimestre vs mismo trimestre año anterior ──
    trimestre_num_sel = int(trimestre_sel.split(' T')[1])
    anio_num_sel = int(trimestre_sel.split(' T')[0])
    anio_anterior = anio_num_sel - 1

    df_trim_anterior = df_tri[(df_tri['trimestre'] == trimestre_num_sel) & (df_tri['anio'] == anio_anterior)]

    if not df_trim_anterior.empty:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">Qué cambió · T{trimestre_num_sel} {anio_anterior} → T{trimestre_num_sel} {anio_num_sel}</div>', unsafe_allow_html=True)

        # Mismo criterio que arriba: si el trimestre elegido está cortado por el mes
        # en curso, se comparan sólo los meses completos presentes en ambos años,
        # para que ningún producto aparezca "cayendo" por un mes a medio terminar.
        df_actual_cmp = df_trim_sel
        df_anterior_cmp = df_trim_anterior
        if TRIMESTRE_PARCIAL is not None and trimestre_num_sel == TRIMESTRE_PARCIAL and anio_num_sel == ANIO_PARCIAL:
            meses_completos_sel = [m for m in df_trim_sel['mes'].unique() if m != MES_PARCIAL]
            meses_comunes_sel = [m for m in meses_completos_sel if m in df_trim_anterior['mes'].unique()]
            if meses_comunes_sel:
                df_actual_cmp = df_trim_sel[df_trim_sel['mes'].isin(meses_comunes_sel)]
                df_anterior_cmp = df_trim_anterior[df_trim_anterior['mes'].isin(meses_comunes_sel)]
                abrev_sel = ", ".join(m[:3] for m in sorted(meses_comunes_sel, key=MESES.index))
                st.caption(f"Comparando sólo {abrev_sel}: {MES_PARCIAL} está incompleto y se excluye de los dos años.")

        actual_por_prod = df_actual_cmp.groupby('producto_limpio')['cantidad'].sum()
        anterior_por_prod = df_anterior_cmp.groupby('producto_limpio')['cantidad'].sum()

        comparativa = pd.DataFrame({'actual': actual_por_prod, 'anterior': anterior_por_prod}).fillna(0)
        comparativa['var_abs'] = comparativa['actual'] - comparativa['anterior']
        comparativa = comparativa.sort_values('var_abs', ascending=False)

        col_sube_t, col_baja_t = st.columns(2)
        with col_sube_t:
            st.markdown('<div style="font-size:0.72rem; font-weight:700; letter-spacing:0.08em; color:#34d399; text-transform:uppercase; margin-bottom:0.6rem;">↑ Crecieron más</div>', unsafe_allow_html=True)
            for prod, row_c in comparativa[comparativa['var_abs'] > 0].head(10).iterrows():
                es_nuevo = row_c['anterior'] == 0
                tag_nuevo = ' <span class="tag-amarillo">nuevo</span>' if es_nuevo else ''
                st.markdown(f"""
                <div class="producto-row" style="border-left-color:#34d399;">
                    <span class="producto-nombre">{prod[:34]}{tag_nuevo}</span>
                    <span class="producto-valor" style="color:#34d399;">+{row_c['var_abs']:.0f} BU</span>
                </div>""", unsafe_allow_html=True)
        with col_baja_t:
            st.markdown('<div style="font-size:0.72rem; font-weight:700; letter-spacing:0.08em; color:#f87171; text-transform:uppercase; margin-bottom:0.6rem;">↓ Bajaron más</div>', unsafe_allow_html=True)
            for prod, row_c in comparativa[comparativa['var_abs'] < 0].sort_values('var_abs').head(10).iterrows():
                es_discontinuado = row_c['actual'] == 0
                tag_disc = ' <span class="tag-rojo">discontinuado</span>' if es_discontinuado else ''
                st.markdown(f"""
                <div class="producto-row" style="border-left-color:#f87171;">
                    <span class="producto-nombre">{prod[:34]}{tag_disc}</span>
                    <span class="producto-valor" style="color:#f87171;">{row_c['var_abs']:.0f} BU</span>
                </div>""", unsafe_allow_html=True)
        st.caption("💡 \"nuevo\" = no existía en el trimestre anterior (no es crecimiento orgánico). \"discontinuado\" = no tuvo ventas en el trimestre actual.")
    else:
        st.caption(f"No hay datos de T{trimestre_num_sel} {anio_anterior} para comparar contra este trimestre.")

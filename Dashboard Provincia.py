import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import calendar
import numpy as np

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GTM SAC - REGIÓN PROVINCIA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CONSTANTES ────────────────────────────────────────────────────────────
META_DIARIA = 5

# ═══════════════════════════════════════════════════════════════════════════
# ZONAS EXCLUIDAS, MAYORISTAS Y REGIONES
# ═══════════════════════════════════════════════════════════════════════════
ZONAS_EXCLUIR = ["OFICINA", "MODERNO", "BROKER", "MARCAS PROPIAS"]
ZONAS_MAYORISTAS = ["MAYORISTAS"]

# ✅ REGIÓN FIJA: PROVINCIA
REGION_FIJA = "REGION PROVINCIA"
REGION_ZONAS = {
    "REGION PROVINCIA": ["CENTRO 1", "CENTRO 2", "ORIENTE 1", "SUR 1", "SUR 2", "SUR CHICO 1", "SUR 3"],
}

TODAS_ZONAS = REGION_ZONAS["REGION PROVINCIA"]

# ─── CSS ────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header { color: #DC2626; font-size: 28px; font-weight: 800; margin-bottom: 4px; }
    .sub-header { color: #64748B; font-size: 14px; margin-bottom: 20px; }
    .alert-warning { background: #FEF3C7; border-left: 4px solid #F59E0B; padding: 12px 16px; border-radius: 8px; margin: 8px 0; }
    .alert-danger { background: #FEE2E2; border-left: 4px solid #DC2626; padding: 12px 16px; border-radius: 8px; margin: 8px 0; }
    .alert-success { background: #DCFCE7; border-left: 4px solid #16A34A; padding: 12px 16px; border-radius: 8px; margin: 8px 0; }
    .alert-info { background: #EFF6FF; border-left: 4px solid #3B82F6; padding: 12px 16px; border-radius: 8px; margin: 8px 0; }
    
    .welcome-title {
        text-align: center;
        font-size: 36px;
        font-weight: 800;
        color: #DC2626;
        margin: 4px 0 0 0;
    }
    .welcome-subtitle {
        text-align: center;
        font-size: 16px;
        color: #64748B;
        font-weight: 400;
        margin: -4px 0 20px 0;
    }
    .welcome-box {
        text-align: center;
        background: #F1F5F9;
        padding: 16px 20px;
        border-radius: 12px;
        margin-bottom: 24px;
    }
    .welcome-box h2 {
        color: #1E293B;
        margin: 0;
        font-size: 28px;
    }
    .welcome-box p {
        color: #64748B;
        margin: 4px 0 0 0;
        font-size: 15px;
    }
    .welcome-instructions {
        background: #EFF6FF;
        border-left: 4px solid #3B82F6;
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 16px;
    }
    .welcome-instructions p {
        margin: 0;
        color: #1E293B;
        font-weight: 600;
    }
    .file-card {
        background: #F8FAFC;
        padding: 16px 20px;
        border-radius: 10px;
        margin-bottom: 12px;
        border: 2px solid #E2E8F0;
        transition: all 0.2s ease;
    }
    .file-card:hover {
        border-color: #3B82F6;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .file-card .icon { font-size: 22px; margin-right: 8px; }
    .file-card .name { font-weight: 700; font-size: 15px; color: #1E293B; }
    .file-card .ext { font-size: 13px; color: #94A3B8; }
    .file-card .status-valid { color: #16A34A; font-weight: 600; font-size: 14px; }
    .file-card .status-error { color: #DC2626; font-weight: 600; font-size: 14px; }
    .file-card .status-pending { color: #94A3B8; font-weight: 500; font-size: 14px; }
    .file-card-valid { border-color: #16A34A; background: #F0FDF4; }
    .file-card-error { border-color: #DC2626; background: #FEF2F2; }
    .welcome-footer {
        text-align: center;
        color: #94A3B8;
        font-size: 13px;
        margin-top: 20px;
        padding-top: 16px;
        border-top: 1px solid #E2E8F0;
    }
    
    .executive-table {
        width: 100%;
        border-collapse: collapse;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        font-family: 'Segoe UI', sans-serif;
        font-size: 14px;
    }
    .executive-table thead tr {
        background: #1E293B;
        color: #FFFFFF;
        font-weight: 600;
        text-align: left;
    }
    .executive-table thead th {
        padding: 12px 16px;
    }
    .executive-table tbody tr {
        border-bottom: 1px solid #E2E8F0;
    }
    .executive-table tbody tr:nth-child(even) {
        background: #F8FAFC;
    }
    .executive-table tbody tr:nth-child(odd) {
        background: #FFFFFF;
    }
    .executive-table tbody td {
        padding: 10px 16px;
        color: #1E293B;
    }
    .executive-table tfoot {
        background: #F1F5F9;
        font-size: 13px;
        color: #64748B;
    }
    .executive-table tfoot td {
        padding: 8px 16px;
        border-top: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

# ─── FUNCIÓN PARA DETECTAR HEADER ──────────────────────────────────────
def detectar_header(df_raw, palabras_clave):
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(val).upper().strip() for val in row.values if pd.notna(val)])
        coincidencias = sum(1 for palabra in palabras_clave if palabra.upper() in row_str)
        if coincidencias >= len(palabras_clave) - 1:
            return idx
    return 0

# ═══════════════════════════════════════════════════════════════════════════
# FUNCIONES DE CÁLCULO
# ═══════════════════════════════════════════════════════════════════════════

def calcular_dias_habiles(mes_sel, df_log):
    if mes_sel == 'Todos' or mes_sel is None:
        fecha_actual = datetime.now()
        year = fecha_actual.year
        month = fecha_actual.month
    else:
        try:
            year = int(mes_sel.split('-')[0])
            month = int(mes_sel.split('-')[1])
        except:
            fecha_actual = datetime.now()
            year = fecha_actual.year
            month = fecha_actual.month
    
    # ─── FERIADOS PERÚ ──────────────────────────────────────────────
    feriados = [
        datetime(year, 1, 1),
        datetime(year, 4, 9),
        datetime(year, 4, 10),
        datetime(year, 5, 1),
        datetime(year, 6, 29),
        datetime(year, 7, 28),
        datetime(year, 7, 29),
        datetime(year, 8, 30),
        datetime(year, 10, 8),
        datetime(year, 11, 1),
        datetime(year, 12, 8),
        datetime(year, 12, 25),
    ]
    
    primer_dia = datetime(year, month, 1)
    if month == 12:
        ultimo_dia = datetime(year, month, 31)
    else:
        ultimo_dia = datetime(year, month + 1, 1) - timedelta(days=1)
    
    dias_habiles = 0
    fecha_actual_calc = primer_dia
    while fecha_actual_calc <= ultimo_dia:
        if fecha_actual_calc.weekday() < 5 and fecha_actual_calc not in feriados:
            dias_habiles += 1
        fecha_actual_calc += timedelta(days=1)
    return dias_habiles

def calcular_dias_habiles_transcurridos(mes_sel):
    """
    Calcula días hábiles transcurridos desde inicio del mes hasta hoy.
    No cuenta sábados, domingos ni feriados.
    """
    if mes_sel == 'Todos' or mes_sel is None:
        fecha_actual = datetime.now()
        year = fecha_actual.year
        month = fecha_actual.month
    else:
        try:
            year = int(mes_sel.split('-')[0])
            month = int(mes_sel.split('-')[1])
        except:
            fecha_actual = datetime.now()
            year = fecha_actual.year
            month = fecha_actual.month
    
    # ─── FERIADOS PERÚ ──────────────────────────────────────────────
    feriados = [
        datetime(year, 1, 1),
        datetime(year, 4, 9),
        datetime(year, 4, 10),
        datetime(year, 5, 1),
        datetime(year, 6, 29),
        datetime(year, 7, 28),
        datetime(year, 7, 29),
        datetime(year, 8, 30),
        datetime(year, 10, 8),
        datetime(year, 11, 1),
        datetime(year, 12, 8),
        datetime(year, 12, 25),
    ]
    
    fecha_inicio = datetime(year, month, 1)
    fecha_hoy = datetime.now()
    
    # Último día del mes
    if month == 12:
        ultimo_dia = datetime(year, month, 31)
    else:
        ultimo_dia = datetime(year, month + 1, 1) - timedelta(days=1)
    
    # Si la fecha actual es del mes seleccionado, usar hoy
    if fecha_hoy.month == month and fecha_hoy.year == year:
        fecha_limite = fecha_hoy
    else:
        fecha_limite = ultimo_dia
    
    dias_habiles = 0
    fecha_actual = fecha_inicio
    
    while fecha_actual <= fecha_limite:
        # Lunes a viernes y NO feriado
        if fecha_actual.weekday() < 5 and fecha_actual not in feriados:
            dias_habiles += 1
        fecha_actual += timedelta(days=1)
    
    return dias_habiles

def calcular_meta_visitas(mes_sel, df_log):
    dias_habiles = calcular_dias_habiles(mes_sel, df_log)
    return META_DIARIA * dias_habiles

def obtener_meta_por_cliente(zonas_filtro):
    if zonas_filtro and 'Todas' not in zonas_filtro:
        for zona in zonas_filtro:
            if zona in ZONAS_MAYORISTAS:
                return 1
    return 2

def calcular_objetivo_clientes_nuevos(df_cartera, zonas_filtro=None):
    if df_cartera is None or df_cartera.empty:
        return 0
    cartera = df_cartera.copy()
    if zonas_filtro and 'Todas' not in zonas_filtro:
        cartera = cartera[cartera['zona'].isin(zonas_filtro)]
    zonas_count = cartera['zona'].value_counts()
    objetivo_total = 0
    for zona, count in zonas_count.items():
        if zona in ZONAS_MAYORISTAS:
            objetivo_total += count * 1
        else:
            objetivo_total += count * 2
    return objetivo_total

def calcular_cartera_vigente(clientes_activos, clientes_nuevos):
    return clientes_activos + clientes_nuevos

def calcular_tasa_conversion(clientes_nuevos, total_prospectos):
    if total_prospectos == 0:
        return 0.0
    return (clientes_nuevos / total_prospectos * 100)

def obtener_mes_anterior(mes_sel, df_log):
    if mes_sel == 'Todos':
        meses_disponibles = sorted(df_log['mes_ano'].unique())
        if not meses_disponibles:
            return None
        ultimo_mes = meses_disponibles[-1]
    else:
        ultimo_mes = mes_sel
    
    year = int(ultimo_mes.split('-')[0])
    month = int(ultimo_mes.split('-')[1])
    if month == 1:
        return f"{year-1}-12"
    else:
        return f"{year}-{month-1:02d}"

def calcular_avance_lineal(mes_sel):
    """Calcula el avance lineal esperado según los días hábiles transcurridos"""
    if mes_sel == 'Todos' or mes_sel is None:
        fecha_actual = datetime.now()
        year = fecha_actual.year
        month = fecha_actual.month
    else:
        try:
            year = int(mes_sel.split('-')[0])
            month = int(mes_sel.split('-')[1])
        except:
            fecha_actual = datetime.now()
            year = fecha_actual.year
            month = fecha_actual.month
    
    # Obtener días hábiles totales del mes
    total_dias_habiles = calcular_dias_habiles(mes_sel, None)
    
    # Usar la nueva función para días transcurridos
    dias_transcurridos = calcular_dias_habiles_transcurridos(mes_sel)
    
    # Calcular avance lineal
    if total_dias_habiles > 0:
        avance_lineal = (dias_transcurridos / total_dias_habiles) * 100
    else:
        avance_lineal = 0
    
    return {
        'avance_lineal': avance_lineal,
        'dias_transcurridos': dias_transcurridos,
        'dias_totales': total_dias_habiles,
        'dias_restantes': total_dias_habiles - dias_transcurridos
    }

# ═══════════════════════════════════════════════════════════════════════════
# FUNCIONES DE CARGA
# ═══════════════════════════════════════════════════════════════════════════

@st.cache_data
def cargar_log_visitas(file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Log", header=0)
    except Exception as e:
        st.error(f"❌ Error al leer la hoja Log: {e}")
        return pd.DataFrame()
    
    if df.empty:
        st.error("❌ La hoja Log está vacía")
        return pd.DataFrame()
    
    df.columns = [str(col).strip() for col in df.columns]
    
    col_fecha = None
    col_zona = None
    col_cliente = None
    col_tipo = None
    col_task = None
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if 'date' in col_lower or 'fecha' in col_lower:
            col_fecha = col
        elif 'zona' in col_lower or 'territorio' in col_lower:
            col_zona = col
        elif 'cliente o prospecto' in col_lower:
            col_cliente = col
        elif col_lower == 'tipo':
            col_tipo = col
        elif 'task' in col_lower or 'tarea' in col_lower:
            col_task = col
    
    if col_cliente is None:
        st.error("❌ No se encontró columna 'Cliente o Prospecto' en la hoja Log")
        st.info(f"📋 Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()
    
    if col_fecha is None:
        st.error("❌ No se encontró columna de fecha (Date / Fecha) en la hoja Log")
        st.info(f"📋 Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()
    
    df_resultado = pd.DataFrame()
    df_resultado['cliente'] = df[col_cliente].astype(str).str.strip()
    df_resultado['fecha'] = pd.to_datetime(df[col_fecha], errors='coerce')
    df_resultado = df_resultado.dropna(subset=['cliente', 'fecha'])
    
    if col_zona:
        df_resultado['zona'] = df[col_zona].astype(str).str.strip().str.upper()
    else:
        df_resultado['zona'] = 'SIN ZONA'
    
    df_resultado = df_resultado[~df_resultado['zona'].isin(ZONAS_EXCLUIR)]
    
    if col_tipo:
        df_resultado['tipo'] = df[col_tipo].astype(str).str.strip().str.upper()
    else:
        df_resultado['tipo'] = 'MANTENIMIENTO'
    
    if col_task:
        df_resultado['task'] = df[col_task].astype(str).str.strip()
    else:
        df_resultado['task'] = ''
    
    df_resultado['tipo_visita'] = 'FÍSICA'
    df_resultado = df_resultado[df_resultado['cliente'] != '']
    df_resultado = df_resultado[df_resultado['cliente'].str.upper() != 'GO TO MARKET']
    df_resultado = df_resultado[df_resultado['cliente'].str.upper() != 'NAN']
    
    if not df_resultado.empty:
        df_resultado['semana_inicio'] = df_resultado['fecha'] - pd.to_timedelta(df_resultado['fecha'].dt.weekday, unit='D')
        df_resultado['semana'] = df_resultado['semana_inicio'].dt.strftime('Semana %W')
        df_resultado['mes'] = df_resultado['fecha'].dt.strftime('%B %Y')
        df_resultado['dia_semana'] = df_resultado['fecha'].dt.day_name()
        df_resultado['mes_ano'] = df_resultado['fecha'].dt.strftime('%Y-%m')
    
    return df_resultado

@st.cache_data
def cargar_lead(file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Lead", header=0)
    except Exception as e:
        st.error(f"❌ Error al leer la hoja Lead: {e}")
        return pd.DataFrame()
    
    if df.empty:
        st.error("❌ La hoja Lead está vacía")
        return pd.DataFrame()
    
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    col_zona = None
    col_cliente = None
    col_fecha = None
    
    for col in df.columns:
        col_upper = col.upper()
        if 'ZONA' in col_upper or 'TERRITORIO' in col_upper:
            col_zona = col
        elif 'LEAD' in col_upper or 'NOMBRE' in col_upper or 'CLIENTE' in col_upper:
            col_cliente = col
        elif 'FECHA' in col_upper or 'CREACION' in col_upper or 'REGISTRO' in col_upper:
            col_fecha = col
    
    if col_cliente is None and len(df.columns) >= 2:
        col_cliente = df.columns[1]
    if col_zona is None and len(df.columns) >= 1:
        col_zona = df.columns[0]
    if col_fecha is None and len(df.columns) >= 3:
        col_fecha = df.columns[2]
    
    if col_cliente is None:
        st.error("❌ No se encontró columna de cliente en la hoja Lead")
        st.info(f"📋 Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()
    
    df_resultado = pd.DataFrame()
    df_resultado['cliente'] = df[col_cliente].astype(str).str.strip().str.upper()
    
    if col_zona:
        df_resultado['zona'] = df[col_zona].astype(str).str.strip().str.upper()
    else:
        df_resultado['zona'] = 'SIN ZONA'
    
    if col_fecha:
        df_resultado['fecha_registro'] = pd.to_datetime(df[col_fecha], errors='coerce')
    else:
        df_resultado['fecha_registro'] = pd.NaT
    
    df_resultado = df_resultado[~df_resultado['zona'].isin(ZONAS_EXCLUIR)]
    df_resultado = df_resultado[df_resultado['cliente'] != '']
    df_resultado = df_resultado[df_resultado['cliente'] != 'NAN']
    df_resultado = df_resultado[df_resultado['cliente'] != 'NONE']
    
    df_resultado = df_resultado.sort_values('fecha_registro', ascending=True)
    
    return df_resultado

@st.cache_data
def cargar_cartera_activa(file_bytes):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    except Exception as e:
        st.error(f"❌ Error al leer el archivo de cartera activa: {e}")
        return pd.DataFrame()
    
    if df_raw.empty:
        st.error("❌ El archivo de cartera activa está vacío")
        return pd.DataFrame()
    
    header_row = detectar_header(df_raw, ['ZONA', 'CLIENTE'])
    
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
    except:
        df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    
    if df.empty:
        st.error("❌ El archivo de cartera activa está vacío")
        return pd.DataFrame()
    
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    col_zona = None
    col_cliente = None
    col_clase = None
    
    for col in df.columns:
        if 'ZONA' in col or 'ZONAS' in col:
            col_zona = col
        elif 'CLIENTE' in col:
            col_cliente = col
        elif 'CLASE' in col:
            col_clase = col
    
    if col_cliente is None and len(df.columns) >= 2:
        col_cliente = df.columns[1]
    if col_zona is None and len(df.columns) >= 3:
        col_zona = df.columns[2]
    if col_clase is None and len(df.columns) >= 4:
        col_clase = df.columns[3]
    
    if col_cliente is None:
        st.error("❌ No se encontró la columna 'CLIENTE'")
        st.info(f"📋 Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()
    
    df_resultado = pd.DataFrame()
    df_resultado['zona'] = df[col_zona].astype(str).str.strip().str.upper() if col_zona in df.columns else ''
    df_resultado['cliente'] = df[col_cliente].astype(str).str.strip().str.upper()
    
    if col_clase and col_clase in df.columns:
        df_resultado['clase'] = df[col_clase].astype(str).str.strip().str.upper()
    else:
        df_resultado['clase'] = 'ESTÁNDAR'
    
    df_resultado = df_resultado[~df_resultado['zona'].isin(ZONAS_EXCLUIR)]
    df_resultado = df_resultado[df_resultado['cliente'] != '']
    df_resultado = df_resultado[df_resultado['cliente'] != 'NAN']
    df_resultado = df_resultado[df_resultado['cliente'] != 'NONE']
    
    return df_resultado

@st.cache_data
def cargar_clientes_riesgo(file_bytes):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    except Exception as e:
        st.error(f"❌ Error al leer el archivo de clientes en riesgo: {e}")
        return pd.DataFrame()
    
    if df_raw.empty:
        st.error("❌ El archivo de clientes en riesgo está vacío")
        return pd.DataFrame()
    
    header_row = detectar_header(df_raw, ['ZONA', 'CLIENTE'])
    
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
    except:
        df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    
    if df.empty:
        st.error("❌ El archivo de clientes en riesgo está vacío")
        return pd.DataFrame()
    
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    col_zona = None
    col_cliente = None
    col_dias = None
    
    for col in df.columns:
        if 'ZONA' in col or 'ZONAS' in col:
            col_zona = col
        elif 'CLIENTE' in col:
            col_cliente = col
        elif 'DIAS' in col or 'SIN COMPRA' in col:
            col_dias = col
    
    if col_cliente is None and len(df.columns) >= 2:
        col_cliente = df.columns[1]
    if col_zona is None and len(df.columns) >= 3:
        col_zona = df.columns[2]
    if col_dias is None and len(df.columns) >= 6:
        col_dias = df.columns[5]
    
    if col_cliente is None:
        st.error("❌ No se encontró la columna 'CLIENTE'")
        st.info(f"📋 Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()
    
    df_resultado = pd.DataFrame()
    df_resultado['zona'] = df[col_zona].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True) if col_zona in df.columns else ''
    df_resultado['cliente'] = df[col_cliente].astype(str).str.strip().str.upper()
    
    if col_dias and col_dias in df.columns:
        df_resultado['dias_sin_compra'] = pd.to_numeric(df[col_dias], errors='coerce').fillna(0)
    else:
        df_resultado['dias_sin_compra'] = 0
    
    df_resultado = df_resultado[df_resultado['zona'] != '']
    df_resultado = df_resultado[df_resultado['zona'] != 'N°']
    df_resultado = df_resultado[df_resultado['zona'] != 'NAN']
    df_resultado = df_resultado[df_resultado['cliente'] != '']
    df_resultado = df_resultado[df_resultado['cliente'].str.upper() != 'CLIENTE']
    df_resultado = df_resultado[df_resultado['cliente'].str.upper() != 'NAN']
    df_resultado = df_resultado[~df_resultado['cliente'].str.contains('CLIENTES CON', na=False)]
    df_resultado = df_resultado[~df_resultado['zona'].isin(ZONAS_EXCLUIR)]
    
    return df_resultado

@st.cache_data
def cargar_ventas_categoria(file_bytes):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    except Exception as e:
        st.error(f"❌ Error al leer el archivo de ventas: {e}")
        return pd.DataFrame()
    
    if df_raw.empty:
        st.error("❌ El archivo de ventas está vacío")
        return pd.DataFrame()
    
    header_row = None
    for idx, row in df_raw.iterrows():
        row_str = ' '.join([str(val).upper().strip() for val in row.values if pd.notna(val)])
        if 'OBJ' in row_str and 'REAL' in row_str:
            header_row = idx
            break
    
    if header_row is None:
        for idx, row in df_raw.iterrows():
            row_str = ' '.join([str(val).upper().strip() for val in row.values if pd.notna(val)])
            if 'ETIQUETA' in row_str or 'JEFATURA' in row_str:
                header_row = idx
                break
    
    if header_row is None:
        header_row = 0
        st.warning("⚠️ No se detectó el header automáticamente, usando fila 1 (índice 0)")
    
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=header_row)
    except:
        df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    
    if df.empty:
        st.error("❌ El archivo de ventas está vacío")
        return pd.DataFrame()
    
    df.columns = [str(col).strip() for col in df.columns]
    
    col_jefatura = None
    col_zona = None
    col_categoria = None
    col_objetivo = None
    col_avance = None
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        col_sin_tilde = col_lower.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
        
        if 'jefatura' in col_sin_tilde or 'jefaturas' in col_sin_tilde or 'region' in col_sin_tilde:
            col_jefatura = col
        elif 'zona' in col_sin_tilde or 'territorio' in col_sin_tilde:
            col_zona = col
        elif 'categoria' in col_sin_tilde:
            col_categoria = col
        elif 'objetivo' in col_sin_tilde or 'obj' in col_sin_tilde:
            col_objetivo = col
        elif 'avance' in col_sin_tilde or 'real' in col_sin_tilde:
            col_avance = col
    
    if col_jefatura is None and len(df.columns) >= 1:
        col_jefatura = df.columns[0]
    if col_zona is None and len(df.columns) >= 2:
        col_zona = df.columns[1]
    if col_categoria is None and len(df.columns) >= 3:
        col_categoria = df.columns[2]
    if col_objetivo is None and len(df.columns) >= 4:
        col_objetivo = df.columns[3]
    if col_avance is None and len(df.columns) >= 5:
        col_avance = df.columns[4]
    
    if col_objetivo is None or col_avance is None:
        st.error("❌ No se encontraron columnas de objetivo y avance")
        st.info(f"📋 Columnas disponibles: {list(df.columns)}")
        return pd.DataFrame()
    
    df_resultado = pd.DataFrame()
    df_resultado['jefatura'] = df[col_jefatura].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
    df_resultado['zona'] = df[col_zona].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
    df_resultado['categoria'] = df[col_categoria].astype(str).str.strip()
    
    df_resultado['objetivo'] = pd.to_numeric(df[col_objetivo], errors='coerce').fillna(0.0)
    df_resultado['avance'] = pd.to_numeric(df[col_avance], errors='coerce').fillna(0.0)
    
    df_resultado = df_resultado[~df_resultado['zona'].isin(ZONAS_EXCLUIR)]
    df_resultado = generar_totales_ventas(df_resultado)
    
    return df_resultado

def generar_totales_ventas(df):
    if df.empty:
        return df
    
    df_resultado = df.copy()
    
    totales_zona = df.groupby(['jefatura', 'zona'], as_index=False).agg({
        'objetivo': 'sum',
        'avance': 'sum'
    })
    totales_zona['categoria'] = 'TOTAL ZONA'
    
    totales_region = df.groupby('jefatura', as_index=False).agg({
        'objetivo': 'sum',
        'avance': 'sum'
    })
    totales_region['zona'] = 'TOTAL REGION'
    totales_region['categoria'] = 'TOTAL REGION'
    
    df_final = pd.concat([
        df_resultado,
        totales_zona,
        totales_region
    ], ignore_index=True)
    
    return df_final

# ═══════════════════════════════════════════════════════════════════════════
# FUNCIONES DE PROCESAMIENTO
# ═══════════════════════════════════════════════════════════════════════════

def obtener_rango_mes(mes_sel):
    if mes_sel == 'Todos' or mes_sel is None:
        fecha_actual = datetime.now()
        year = fecha_actual.year
        month = fecha_actual.month
    else:
        try:
            year = int(mes_sel.split('-')[0])
            month = int(mes_sel.split('-')[1])
        except:
            fecha_actual = datetime.now()
            year = fecha_actual.year
            month = fecha_actual.month
    
    fecha_inicio = datetime(year, month, 1)
    if month == 12:
        fecha_fin = datetime(year, month, 31)
    else:
        fecha_fin = datetime(year, month + 1, 1) - timedelta(days=1)
    
    return fecha_inicio, fecha_fin

def procesar_visitas(df_log, zonas_filtro=None, mes_sel=None):
    if df_log is None or df_log.empty:
        return {
            'total_visitas': 0,
            'visitas_prospeccion': 0,
            'visitas_mantenimiento': 0,
            'clientes_unicos': 0,
            'meta_total': 0,
            'pct_cumplimiento': 0,
            'num_semanas': 0
        }
    
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df_log = df_log[df_log['zona'].isin(zonas_filtro)]
    
    df_log_unicas = df_log.drop_duplicates(subset=['cliente', 'fecha'])
    
    total_visitas = len(df_log_unicas)
    visitas_prospeccion = len(df_log_unicas[df_log_unicas['tipo'] == 'PROSPECCIÓN']) if 'tipo' in df_log_unicas.columns else 0
    visitas_mantenimiento = len(df_log_unicas[df_log_unicas['tipo'] == 'MANTENIMIENTO']) if 'tipo' in df_log_unicas.columns else 0
    clientes_unicos = df_log_unicas['cliente'].nunique()
    num_semanas = df_log_unicas['semana'].nunique() if 'semana' in df_log_unicas.columns else 1
    
    meta_total = calcular_meta_visitas(mes_sel, df_log)
    pct_cumplimiento = (total_visitas / meta_total * 100) if meta_total > 0 else 0
    
    return {
        'total_visitas': total_visitas,
        'visitas_prospeccion': visitas_prospeccion,
        'visitas_mantenimiento': visitas_mantenimiento,
        'clientes_unicos': clientes_unicos,
        'meta_total': meta_total,
        'pct_cumplimiento': pct_cumplimiento,
        'num_semanas': num_semanas
    }

def procesar_calidad_visita(df_log, df_cartera, zonas_filtro=None):
    if df_log is None or df_log.empty or df_cartera is None or df_cartera.empty:
        return {
            'total_visitas': 0,
            'visitas_a_activos': 0,
            'visitas_a_prospectos': 0,
            'visitas_a_otros_no_activos': 0,
            'pct_calidad': 0,
            'total_cartera': 0,
            'cartera_visitada': 0,
            'pct_cobertura': 0,
            'clientes_activos_detalle': [],
            'clientes_prospectos_detalle': [],
            'clientes_otros_detalle': []
        }
    
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df_log = df_log[df_log['zona'].isin(zonas_filtro)]
        df_cartera = df_cartera[df_cartera['zona'].isin(zonas_filtro)]
    
    df_log_unicas = df_log.drop_duplicates(subset=['cliente', 'fecha'])
    
    clientes_activos = set(df_cartera['cliente'].str.upper().str.strip())
    df_log_unicas['es_activo'] = df_log_unicas['cliente'].str.upper().str.strip().isin(clientes_activos)
    
    total_visitas = len(df_log_unicas)
    
    df_activos = df_log_unicas[df_log_unicas['es_activo']]
    visitas_a_activos = len(df_activos)
    
    df_no_activos = df_log_unicas[~df_log_unicas['es_activo']]
    
    df_prospectos = df_no_activos[df_no_activos['tipo'] == 'PROSPECCIÓN']
    visitas_a_prospectos = len(df_prospectos)
    
    df_otros = df_no_activos[df_no_activos['tipo'] != 'PROSPECCIÓN']
    visitas_a_otros_no_activos = len(df_otros)
    
    clientes_visitados = set(df_log_unicas['cliente'].str.upper().str.strip())
    cartera_visitada = clientes_activos.intersection(clientes_visitados)
    
    pct_cobertura = (len(cartera_visitada) / len(clientes_activos) * 100) if len(clientes_activos) > 0 else 0
    pct_calidad = (visitas_a_activos / total_visitas * 100) if total_visitas > 0 else 0
    
    def contar_visitas_por_cliente(df):
        if df.empty:
            return []
        conteo = df['cliente'].str.upper().str.strip().value_counts().to_dict()
        return sorted([{'cliente': k, 'visitas': v} for k, v in conteo.items()], key=lambda x: x['visitas'], reverse=True)
    
    clientes_activos_detalle = contar_visitas_por_cliente(df_activos)
    clientes_prospectos_detalle = contar_visitas_por_cliente(df_prospectos)
    clientes_otros_detalle = contar_visitas_por_cliente(df_otros)
    
    return {
        'total_visitas': total_visitas,
        'visitas_a_activos': int(visitas_a_activos),
        'visitas_a_prospectos': int(visitas_a_prospectos),
        'visitas_a_otros_no_activos': int(visitas_a_otros_no_activos),
        'pct_calidad': pct_calidad,
        'total_cartera': len(clientes_activos),
        'cartera_visitada': len(cartera_visitada),
        'pct_cobertura': pct_cobertura,
        'clientes_activos_detalle': clientes_activos_detalle,
        'clientes_prospectos_detalle': clientes_prospectos_detalle,
        'clientes_otros_detalle': clientes_otros_detalle
    }

def procesar_prospectos_iniciales_lead(df_lead, df_log, zonas_filtro=None, mes_sel=None):
    if df_lead is None or df_lead.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    df_lead_filtrado = df_lead.copy()
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df_lead_filtrado = df_lead_filtrado[df_lead_filtrado['zona'].isin(zonas_filtro)]
    
    if df_lead_filtrado.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    fecha_inicio, fecha_fin = obtener_rango_mes(mes_sel)
    
    df_iniciales = df_lead_filtrado[
        (df_lead_filtrado['fecha_registro'] < fecha_inicio) |
        (df_lead_filtrado['fecha_registro'].isna())
    ]
    
    if df_iniciales.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    if df_log is not None and not df_log.empty:
        df_log_filtrado = df_log.copy()
        if zonas_filtro and 'Todas' not in zonas_filtro:
            df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
        
        df_cierres_anteriores = df_log_filtrado[
            (df_log_filtrado['tipo'] == 'PROSPECCIÓN') & 
            (df_log_filtrado['task'].astype(str).str.upper().str.contains('CIERRE', na=False)) &
            (~df_log_filtrado['task'].astype(str).str.upper().str.contains('NO CIERRE', na=False)) &
            (df_log_filtrado['fecha'] < fecha_inicio)
        ]
        
        clientes_con_cierre_anterior = set(df_cierres_anteriores['cliente'].str.upper().str.strip())
        df_iniciales = df_iniciales[~df_iniciales['cliente'].isin(clientes_con_cierre_anterior)]
    
    if df_iniciales.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    clientes_detalle = []
    for _, row in df_iniciales.iterrows():
        cliente = row['cliente']
        fecha_reg = row['fecha_registro'].strftime('%d/%m/%Y') if pd.notna(row['fecha_registro']) else 'Sin fecha'
        zona = row['zona']
        clientes_detalle.append({
            'nombre': cliente,
            'fecha': fecha_reg,
            'zona': zona
        })
    
    clientes_detalle.sort(key=lambda x: x['fecha'])
    nombres = [c['nombre'] for c in clientes_detalle[:10]]
    
    return {
        'cantidad': len(clientes_detalle),
        'clientes': clientes_detalle,
        'nombres': nombres
    }

def procesar_sqls_lead(df_lead, zonas_filtro=None, mes_sel=None):
    if df_lead is None or df_lead.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    df_lead_filtrado = df_lead.copy()
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df_lead_filtrado = df_lead_filtrado[df_lead_filtrado['zona'].isin(zonas_filtro)]
    
    if df_lead_filtrado.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    fecha_inicio, fecha_fin = obtener_rango_mes(mes_sel)
    
    df_sqls = df_lead_filtrado[
        (df_lead_filtrado['fecha_registro'] >= fecha_inicio) &
        (df_lead_filtrado['fecha_registro'] <= fecha_fin)
    ]
    
    if df_sqls.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    clientes_detalle = []
    for _, row in df_sqls.iterrows():
        cliente = row['cliente']
        fecha_reg = row['fecha_registro'].strftime('%d/%m/%Y')
        zona = row['zona']
        clientes_detalle.append({
            'nombre': cliente,
            'fecha': fecha_reg,
            'zona': zona
        })
    
    clientes_detalle.sort(key=lambda x: x['fecha'])
    nombres = [c['nombre'] for c in clientes_detalle[:10]]
    
    return {
        'cantidad': len(clientes_detalle),
        'clientes': clientes_detalle,
        'nombres': nombres
    }

def procesar_embudo_completo(df_log, df_lead, zonas_filtro=None, mes_sel=None):
    if df_lead is None or df_lead.empty:
        return []
    
    fecha_inicio, fecha_fin = obtener_rango_mes(mes_sel)
    
    df_lead_filtrado = df_lead.copy()
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df_lead_filtrado = df_lead_filtrado[df_lead_filtrado['zona'].isin(zonas_filtro)]
    
    df_prospectos_mes = df_lead_filtrado[
        (df_lead_filtrado['fecha_registro'] < fecha_fin) |
        (df_lead_filtrado['fecha_registro'].isna())
    ]
    
    if df_prospectos_mes.empty:
        etapas = ['PROSPECCIÓN', 'CALIFICACIÓN', 'VISITA', 'PROPUESTA', 'NEGOCIACIÓN', 'CIERRE']
        return [{'etapa': e, 'cantidad': 0, 'clientes': [], 'dias_prom': 0, 'tooltip_text': 'Sin prospectos'} for e in etapas]
    
    if df_log is not None and not df_log.empty:
        df_log_filtrado = df_log.copy()
        if zonas_filtro and 'Todas' not in zonas_filtro:
            df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
        
        df_cierres_anteriores = df_log_filtrado[
            (df_log_filtrado['tipo'] == 'PROSPECCIÓN') & 
            (df_log_filtrado['task'].astype(str).str.upper().str.contains('CIERRE', na=False)) &
            (~df_log_filtrado['task'].astype(str).str.upper().str.contains('NO CIERRE', na=False)) &
            (df_log_filtrado['fecha'] < fecha_inicio)
        ]
        clientes_con_cierre_anterior = set(df_cierres_anteriores['cliente'].str.upper().str.strip())
        df_prospectos_mes = df_prospectos_mes[~df_prospectos_mes['cliente'].isin(clientes_con_cierre_anterior)]
        
        df_cierres_despues = df_log_filtrado[
            (df_log_filtrado['tipo'] == 'PROSPECCIÓN') & 
            (df_log_filtrado['task'].astype(str).str.upper().str.contains('CIERRE', na=False)) &
            (~df_log_filtrado['task'].astype(str).str.upper().str.contains('NO CIERRE', na=False)) &
            (df_log_filtrado['fecha'] > fecha_fin)
        ]
        clientes_cierres_despues = set(df_cierres_despues['cliente'].str.upper().str.strip())
        df_prospectos_mes = df_prospectos_mes[~df_prospectos_mes['cliente'].isin(clientes_cierres_despues)]
    
    if df_prospectos_mes.empty:
        etapas = ['PROSPECCIÓN', 'CALIFICACIÓN', 'VISITA', 'PROPUESTA', 'NEGOCIACIÓN', 'CIERRE']
        return [{'etapa': e, 'cantidad': 0, 'clientes': [], 'dias_prom': 0, 'tooltip_text': 'Sin prospectos'} for e in etapas]
    
    cliente_etapa = {}
    cliente_fechas = {}
    cliente_visitas = {}
    
    orden_etapas = {
        'PROSPECCIÓN': 1,
        'CALIFICACIÓN': 2,
        'VISITA': 3,
        'PROPUESTA': 4,
        'NEGOCIACIÓN': 5,
        'CIERRE': 6
    }
    
    def asignar_etapa(task):
        if pd.isna(task):
            return 'PROSPECCIÓN'
        task_upper = str(task).upper()
        if 'CIERRE' in task_upper and 'NO CIERRE' not in task_upper:
            return 'CIERRE'
        elif 'NEGOCIACIÓN' in task_upper or 'NEGOCIACION' in task_upper:
            return 'NEGOCIACIÓN'
        elif 'PROPUESTA' in task_upper:
            return 'PROPUESTA'
        elif 'VISITA' in task_upper:
            return 'VISITA'
        elif 'CALIFICACIÓN' in task_upper or 'CALIFICACION' in task_upper:
            return 'CALIFICACIÓN'
        else:
            return 'PROSPECCIÓN'
    
    for _, row in df_prospectos_mes.iterrows():
        cliente = row['cliente']
        cliente_etapa[cliente] = 'PROSPECCIÓN'
        cliente_fechas[cliente] = row['fecha_registro'] if pd.notna(row['fecha_registro']) else datetime.now()
        cliente_visitas[cliente] = 0
    
    if df_log is not None and not df_log.empty:
        df_log_filtrado = df_log.copy()
        if zonas_filtro and 'Todas' not in zonas_filtro:
            df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
        
        df_log_prospectos = df_log_filtrado[df_log_filtrado['tipo'] == 'PROSPECCIÓN'].copy()
        
        if not df_log_prospectos.empty:
            for _, row in df_log_prospectos.iterrows():
                cliente = row['cliente']
                if cliente in cliente_etapa:
                    etapa = asignar_etapa(row['task'])
                    fecha = row['fecha']
                    
                    if orden_etapas.get(etapa, 0) > orden_etapas.get(cliente_etapa[cliente], 0):
                        cliente_etapa[cliente] = etapa
                        cliente_fechas[cliente] = fecha
                    
                    cliente_visitas[cliente] = cliente_visitas.get(cliente, 0) + 1
    
    etapas = ['PROSPECCIÓN', 'CALIFICACIÓN', 'VISITA', 'PROPUESTA', 'NEGOCIACIÓN', 'CIERRE']
    resultado = []
    
    fecha_actual = datetime.now()
    
    for etapa in etapas:
        clientes_en_etapa = []
        
        for cliente, etapa_cliente in cliente_etapa.items():
            if etapa_cliente == etapa:
                fecha_reg = cliente_fechas.get(cliente)
                visitas = cliente_visitas.get(cliente, 0)
                dias_en_etapa = (fecha_actual - fecha_reg).days if fecha_reg else 0
                
                clientes_en_etapa.append({
                    'nombre': cliente,
                    'dias': dias_en_etapa,
                    'visitas': visitas
                })
        
        if clientes_en_etapa:
            dias_prom = round(sum(c['dias'] for c in clientes_en_etapa) / len(clientes_en_etapa))
            clientes_en_etapa.sort(key=lambda x: x['dias'], reverse=True)
        else:
            dias_prom = 0
        
        nombres = [c['nombre'] for c in clientes_en_etapa[:10]]
        if len(clientes_en_etapa) > 10:
            tooltip_text = ", ".join(nombres) + f" (+{len(clientes_en_etapa)-10} más)"
        else:
            tooltip_text = ", ".join(nombres) if nombres else "Sin prospectos"
        
        resultado.append({
            'etapa': etapa,
            'cantidad': len(clientes_en_etapa),
            'clientes': clientes_en_etapa,
            'dias_prom': dias_prom,
            'tooltip_text': tooltip_text,
            'nombres': nombres
        })
    
    return resultado

def procesar_clientes_nuevos(df_log, zonas_filtro=None):
    if df_log is None or df_log.empty:
        return {
            'cantidad': 0,
            'clientes': [],
            'nombres': []
        }
    
    df_log_filtrado = df_log.copy()
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
    
    df_cierres = df_log_filtrado[
        (df_log_filtrado['tipo'] == 'PROSPECCIÓN') & 
        (df_log_filtrado['task'].astype(str).str.upper().str.contains('CIERRE', na=False)) &
        (~df_log_filtrado['task'].astype(str).str.upper().str.contains('NO CIERRE', na=False))
    ]
    
    if df_cierres.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    clientes_cierres = {}
    for _, row in df_cierres.iterrows():
        cliente = row['cliente'].upper().strip()
        if cliente not in clientes_cierres:
            clientes_cierres[cliente] = {
                'nombre': cliente,
                'fecha': row['fecha'].strftime('%d/%m/%Y'),
                'zona': row['zona']
            }
    
    clientes_lista = list(clientes_cierres.values())
    clientes_lista.sort(key=lambda x: x['fecha'])
    nombres = [c['nombre'] for c in clientes_lista[:10]]
    
    return {
        'cantidad': len(clientes_lista),
        'clientes': clientes_lista,
        'nombres': nombres
    }

def procesar_leads_desechados(df_log, zonas_filtro=None):
    if df_log is None or df_log.empty or 'task' not in df_log.columns:
        return {
            'cantidad': 0,
            'clientes': [],
            'nombres': []
        }
    
    df_log_filtrado = df_log.copy()
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
    
    df_no_cierres = df_log_filtrado[
        (df_log_filtrado['tipo'] == 'PROSPECCIÓN') & 
        (df_log_filtrado['task'].astype(str).str.upper().str.contains('NO CIERRE', na=False))
    ]
    
    if df_no_cierres.empty:
        return {'cantidad': 0, 'clientes': [], 'nombres': []}
    
    clientes_no_cierres = {}
    for _, row in df_no_cierres.iterrows():
        cliente = row['cliente'].upper().strip()
        if cliente not in clientes_no_cierres:
            clientes_no_cierres[cliente] = {
                'nombre': cliente,
                'fecha': row['fecha'].strftime('%d/%m/%Y'),
                'zona': row['zona']
            }
    
    clientes_lista = list(clientes_no_cierres.values())
    clientes_lista.sort(key=lambda x: x['fecha'])
    nombres = [c['nombre'] for c in clientes_lista[:10]]
    
    return {
        'cantidad': len(clientes_lista),
        'clientes': clientes_lista,
        'nombres': nombres
    }

def procesar_ventas(df_ventas, zonas_filtro=None, region_filtro=None):
    if df_ventas is None or df_ventas.empty:
        return {
            'totales_zona': pd.DataFrame(),
            'totales_categoria': pd.DataFrame(),
            'total_objetivo': 0.0,
            'total_avance': 0.0,
            'pct_general': 0.0
        }
    
    df = df_ventas.copy()
    
    df['objetivo'] = df['objetivo'].astype(float)
    df['avance'] = df['avance'].astype(float)
    
    # Filtro por región (ahora siempre fijo)
    df = df[df['jefatura'] == REGION_FIJA]
    
    if zonas_filtro and 'Todas' not in zonas_filtro:
        df = df[df['zona'].isin(zonas_filtro)]
    
    if df.empty:
        return {
            'totales_zona': pd.DataFrame(),
            'totales_categoria': pd.DataFrame(),
            'total_objetivo': 0.0,
            'total_avance': 0.0,
            'pct_general': 0.0
        }
    
    df_categorias = df[~df['categoria'].str.contains('TOTAL', na=False)]
    
    if not df_categorias.empty:
        totales_zona = df_categorias.groupby('zona', as_index=False).agg({
            'objetivo': 'sum',
            'avance': 'sum'
        })
        totales_zona['pct'] = (totales_zona['avance'] / totales_zona['objetivo'] * 100).fillna(0.0)
        
        totales_categoria = df_categorias.groupby('categoria', as_index=False).agg({
            'objetivo': 'sum',
            'avance': 'sum'
        })
        totales_categoria['pct'] = (totales_categoria['avance'] / totales_categoria['objetivo'] * 100).fillna(0.0)
    else:
        totales_zona = pd.DataFrame()
        totales_categoria = pd.DataFrame()
    
    total_objetivo = float(df_categorias['objetivo'].sum() if not df_categorias.empty else 0)
    total_avance = float(df_categorias['avance'].sum() if not df_categorias.empty else 0)
    pct_general = (total_avance / total_objetivo * 100) if total_objetivo > 0 else 0.0
    
    return {
        'totales_zona': totales_zona,
        'totales_categoria': totales_categoria,
        'total_objetivo': total_objetivo,
        'total_avance': total_avance,
        'pct_general': pct_general
    }

# ═══════════════════════════════════════════════════════════════════════════
# FUNCIONES AUXILIARES DE RENDER
# ═══════════════════════════════════════════════════════════════════════════

def render_tabla_html(df, titulo, columnas=None, total_texto=None):
    if df is None or df.empty:
        st.info(f"No hay datos para mostrar en {titulo}")
        return
    
    if columnas:
        df = df.rename(columns=columnas)
    
    html = f"""
    <table class="executive-table">
        <thead>
            <tr>
    """
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += """
            </tr>
        </thead>
        <tbody>
    """
    for _, row in df.iterrows():
        html += "<tr>"
        for col in df.columns:
            html += f"<td>{row[col]}</td>"
        html += "</tr>"
    html += """
        </tbody>
        <tfoot>
            <tr>
                <td colspan="{0}">{1}</td>
            </tr>
        </tfoot>
    </table>
    """.format(len(df.columns), total_texto if total_texto else f"{len(df)} registros")
    
    st.markdown(html, unsafe_allow_html=True)

def render_embudo(embudo, titulo="📊 Embudo de Ventas"):
    st.markdown(f"#### {titulo}")
    if embudo and sum(e['cantidad'] for e in embudo) > 0:
        colores_embudo = {
            "PROSPECCIÓN": "#DC2626",
            "CALIFICACIÓN": "#F59E0B",
            "VISITA": "#3B82F6",
            "PROPUESTA": "#10B981",
            "NEGOCIACIÓN": "#8B5CF6",
            "CIERRE": "#16A34A"
        }
        
        total_inicial = embudo[0]['cantidad'] if embudo else 1
        anchos = [95, 90, 85, 80, 75, 70]
        
        for i, d in enumerate(embudo):
            etapa = d['etapa']
            cantidad = d['cantidad']
            porcentaje = (cantidad / total_inicial * 100) if total_inicial > 0 else 0
            color = colores_embudo.get(etapa, "#94A3B8")
            ancho = anchos[i] if i < len(anchos) else 80
            tooltip = d['tooltip_text']
            dias_prom = d['dias_prom']
            
            col_a, col_b, col_c = st.columns([0.5, 4, 0.5])
            with col_b:
                st.markdown(f"""
                <div style="width: {ancho}%; margin: 0 auto 2px auto; 
                            background: white; 
                            border-left: 5px solid {color}; 
                            border-radius: 8px; 
                            padding: 8px 14px; 
                            box-shadow: 0 1px 4px rgba(0,0,0,0.06);
                            cursor: help;"
                     title="{tooltip}">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-size: 12px; font-weight: 700; color: {color};">
                            {etapa}
                        </span>
                        <div style="display: flex; align-items: center; gap: 6px;">
                            <span style="font-size: 16px; font-weight: 800; color: #1E293B;">
                                {cantidad}
                            </span>
                            <span style="font-size: 9px; color: #64748B; background: #F1F5F9; padding: 1px 8px; border-radius: 10px;">
                                {porcentaje:.0f}%
                            </span>
                        </div>
                    </div>
                    <div style="display: flex; justify-content: space-between; font-size: 9px; color: #94A3B8; margin-top: 1px;">
                        <span>{cantidad} prospectos</span>
                        <span>⏱️ {dias_prom} días</span>
                    </div>
                    <div style="font-size: 8px; color: #94A3B8; margin-top: 1px; font-style: italic; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
                        {tooltip}
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                if i < len(embudo) - 1:
                    st.markdown(
                        '<div style="text-align: center; font-size: 12px; color: #CBD5E1; line-height: 0.8; margin: 0;">▼</div>',
                        unsafe_allow_html=True
                    )
        
        with st.expander("📋 Ver detalle completo de todos los prospectos por etapa"):
            for d in embudo:
                if d['cantidad'] > 0 and d['clientes']:
                    st.markdown(f"### {d['etapa']} - {len(d['clientes'])} prospectos")
                    df_detalle = pd.DataFrame([
                        {
                            "Prospecto": c['nombre'],
                            "Días en etapa": c['dias'],
                            "Total visitas": c['visitas']
                        }
                        for c in d['clientes']
                    ])
                    st.dataframe(df_detalle, use_container_width=True, hide_index=True)
                    st.markdown("---")
    else:
        st.info('No hay datos de prospectos para el embudo')

def render_distribucion_cartera(cartera):
    st.markdown("#### 📊 Distribución de Cartera por Clase")
    if not cartera.empty and 'clase' in cartera.columns:
        clases_counts = cartera['clase'].value_counts()
        total = len(cartera)
        colores = {'ORO': '🟡', 'PLATA': '⚪', 'BRONCE': '🟤', 'ESTÁNDAR': '🔵'}
        for clase in ['ORO', 'PLATA', 'BRONCE', 'ESTÁNDAR']:
            count = clases_counts.get(clase, 0)
            pct = (count / total * 100) if total > 0 else 0
            color = colores.get(clase, '⬜')
            bar_width = min(pct, 100)
            bar_fill = '█' * int(bar_width / 2)
            bar_empty = '░' * (50 - len(bar_fill))
            st.markdown(
                f"""
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">
                    <span style="font-size: 14px; font-weight: 600; min-width: 80px;">{color} {clase.title()}</span>
                    <span style="font-family: monospace; font-size: 14px; letter-spacing: 1px; background: #F1F5F9; padding: 2px 8px; border-radius: 4px; flex: 1;">
                        {bar_fill}{bar_empty}
                    </span>
                    <span style="font-size: 14px; font-weight: 600; min-width: 60px; text-align: right;">{pct:.0f}%</span>
                    <span style="font-size: 12px; color: #64748B; min-width: 50px;">({count})</span>
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.info('No hay datos de cartera para mostrar')

def render_tabla_clientes_riesgo(df_riesgo_filtrado):
    st.markdown("#### ⚠️ Clientes en Riesgo")
    if not df_riesgo_filtrado.empty:
        df_riesgo_mostrar = df_riesgo_filtrado[['cliente', 'dias_sin_compra']].copy()
        df_riesgo_mostrar = df_riesgo_mostrar[df_riesgo_mostrar['dias_sin_compra'] > 0]
        df_riesgo_mostrar.columns = ['Cliente en Riesgo', 'Días sin Compra']
        render_tabla_html(df_riesgo_mostrar, "Clientes en Riesgo", total_texto=f"{len(df_riesgo_mostrar)} clientes en riesgo")
    else:
        st.info('✅ No hay clientes en riesgo para los filtros seleccionados')

def render_tabla_clientes_nuevos(clientes_nuevos, clientes_nuevos_reales):
    st.markdown("#### ✅ Clientes Nuevos (Cierres)")
    if clientes_nuevos_reales > 0:
        df_nuevos = pd.DataFrame(clientes_nuevos['clientes'])[['nombre', 'fecha']]
        df_nuevos.columns = ['Cliente Nuevo (Cierre)', 'Fecha Cierre']
        render_tabla_html(df_nuevos, "Clientes Nuevos", total_texto=f"{len(df_nuevos)} clientes nuevos")
    else:
        st.info("✅ No hay clientes nuevos (cierres) en el período")

def render_tabla_sqls(sqls):
    st.markdown("#### 🆕 Nuevos Prospectos (SQLs)")
    if sqls['cantidad'] > 0:
        df_sqls = pd.DataFrame(sqls['clientes'])[['nombre', 'fecha']]
        df_sqls.columns = ['SQL (Prospecto Calificado)', '1era Visita']
        render_tabla_html(df_sqls, "SQLs", total_texto=f"{len(df_sqls)} SQLs encontrados")
    else:
        st.info("🆕 No hay SQLs en el período")

def render_tabla_leads_desechados(leads_desechados):
    st.markdown("#### ❌ Leads Desechados")
    if leads_desechados['cantidad'] > 0:
        df_desechados = pd.DataFrame(leads_desechados['clientes'])[['nombre', 'fecha']]
        df_desechados.columns = ['Lead Desechado', 'Fecha']
        render_tabla_html(df_desechados, "Leads Desechados", total_texto=f"{len(df_desechados)} leads desechados")
    else:
        st.info("✅ No hay leads desechados (NO CIERRE) en el período")

def render_botones_ejes(eje_actual):
    st.markdown("#### 📊 Filtrar por Eje de Gestión")
    
    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
    
    with col_btn1:
        if st.button("📊 CARTERA ACTIVA", key=f"btn_cartera_{eje_actual}", use_container_width=True):
            st.session_state[f"eje_seleccionado_{eje_actual}"] = "cartera"
            st.rerun()
        if st.session_state.get(f"eje_seleccionado_{eje_actual}") == "cartera":
            st.markdown('<div style="background:#16A34A20; border-radius:8px; padding:2px; margin-top:-8px; text-align:center; font-size:11px; color:#16A34A;">✅ Seleccionado</div>', unsafe_allow_html=True)
    
    with col_btn2:
        if st.button("📈 PRODUCTIVIDAD", key=f"btn_productividad_{eje_actual}", use_container_width=True):
            st.session_state[f"eje_seleccionado_{eje_actual}"] = "productividad"
            st.rerun()
        if st.session_state.get(f"eje_seleccionado_{eje_actual}") == "productividad":
            st.markdown('<div style="background:#3B82F620; border-radius:8px; padding:2px; margin-top:-8px; text-align:center; font-size:11px; color:#3B82F6;">✅ Seleccionado</div>', unsafe_allow_html=True)
    
    with col_btn3:
        if st.button("🎯 PROSPECCIÓN", key=f"btn_prospeccion_{eje_actual}", use_container_width=True):
            st.session_state[f"eje_seleccionado_{eje_actual}"] = "prospeccion"
            st.rerun()
        if st.session_state.get(f"eje_seleccionado_{eje_actual}") == "prospeccion":
            st.markdown('<div style="background:#F59E0B20; border-radius:8px; padding:2px; margin-top:-8px; text-align:center; font-size:11px; color:#F59E0B;">✅ Seleccionado</div>', unsafe_allow_html=True)
    
    with col_btn4:
        if st.button("💰 VENTAS", key=f"btn_ventas_{eje_actual}", use_container_width=True):
            st.session_state[f"eje_seleccionado_{eje_actual}"] = "ventas"
            st.rerun()
        if st.session_state.get(f"eje_seleccionado_{eje_actual}") == "ventas":
            st.markdown('<div style="background:#DC262620; border-radius:8px; padding:2px; margin-top:-8px; text-align:center; font-size:11px; color:#DC2626;">✅ Seleccionado</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    return st.session_state.get(f"eje_seleccionado_{eje_actual}", "todos")

def mostrar_pantalla_carga():
    col1, col2, col3 = st.columns([1, 2.2, 1])
    with col2:
        try:
            st.image("goidmarket.png", width=140)
        except:
            pass
        
        st.markdown(
            f"""
            <div class="welcome-title">Go To Market</div>
            <div class="welcome-subtitle">Gestión de los 4 ejes Comerciales - {REGION_FIJA}</div>
            """,
            unsafe_allow_html=True
        )
        
        st.markdown(
            f"""
            <div class="welcome-box">
                <h2>📊 121 Semanal - {REGION_FIJA}</h2>
                <p>Seguimiento de visitas, prospección y ventas</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        st.markdown(
            """
            <div class="welcome-instructions">
                <p>📋 Carga los 4 archivos para comenzar:</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        with st.container():
            col_file1, col_file2 = st.columns([3, 1])
            with col_file1:
                st.markdown('<span class="icon">📊</span> <span class="name">Log de Visitas</span> <span class="ext">.xlsx / .xls</span>', unsafe_allow_html=True)
            with col_file2:
                file_log = st.file_uploader("", type=["xlsx", "xls"], key="upload_log", label_visibility="collapsed")
        
        if file_log is not None:
            try:
                df_test_log = cargar_log_visitas(file_log.read())
                file_log.seek(0)
                df_test_lead = cargar_lead(file_log.read())
                file_log.seek(0)
                
                if not df_test_log.empty and not df_test_lead.empty:
                    st.markdown('<div class="file-card-valid" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-valid">✅ Archivo válido: Hojas Log y Lead encontradas</span></div>', unsafe_allow_html=True)
                elif not df_test_log.empty:
                    st.markdown('<div class="file-card-warning" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-warning">⚠️ Hoja Log OK, pero no se encontró la hoja Lead</span></div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: Archivo vacío o formato incorrecto</span></div>', unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f'<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: {str(e)[:80]}...</span></div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        with st.container():
            col_file1, col_file2 = st.columns([3, 1])
            with col_file1:
                st.markdown('<span class="icon">📋</span> <span class="name">Cartera Activa</span> <span class="ext">.xlsx / .xls</span>', unsafe_allow_html=True)
            with col_file2:
                file_cartera = st.file_uploader("", type=["xlsx", "xls"], key="upload_cartera", label_visibility="collapsed")
        
        if file_cartera is not None:
            try:
                df_test = cargar_cartera_activa(file_cartera.read())
                file_cartera.seek(0)
                if not df_test.empty:
                    st.markdown('<div class="file-card-valid" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-valid">✅ Archivo válido: {} columnas</span></div>'.format(len(df_test.columns)), unsafe_allow_html=True)
                else:
                    st.markdown('<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: Archivo vacío o formato incorrecto</span></div>', unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f'<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: {str(e)[:80]}...</span></div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        with st.container():
            col_file1, col_file2 = st.columns([3, 1])
            with col_file1:
                st.markdown('<span class="icon">⚠️</span> <span class="name">Clientes en Riesgo</span> <span class="ext">.xlsx / .xls</span>', unsafe_allow_html=True)
            with col_file2:
                file_riesgo = st.file_uploader("", type=["xlsx", "xls"], key="upload_riesgo", label_visibility="collapsed")
        
        if file_riesgo is not None:
            try:
                df_test = cargar_clientes_riesgo(file_riesgo.read())
                file_riesgo.seek(0)
                if not df_test.empty:
                    st.markdown('<div class="file-card-valid" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-valid">✅ Archivo válido: {} columnas</span></div>'.format(len(df_test.columns)), unsafe_allow_html=True)
                else:
                    st.markdown('<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: Archivo vacío o formato incorrecto</span></div>', unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f'<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: {str(e)[:80]}...</span></div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        with st.container():
            col_file1, col_file2 = st.columns([3, 1])
            with col_file1:
                st.markdown('<span class="icon">💰</span> <span class="name">Ventas por Categoría</span> <span class="ext">.xlsx / .xls</span>', unsafe_allow_html=True)
            with col_file2:
                file_ventas = st.file_uploader("", type=["xlsx", "xls"], key="upload_ventas", label_visibility="collapsed")
        
        if file_ventas is not None:
            try:
                df_test = cargar_ventas_categoria(file_ventas.read())
                file_ventas.seek(0)
                if not df_test.empty:
                    st.markdown('<div class="file-card-valid" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-valid">✅ Archivo válido: {} columnas</span></div>'.format(len(df_test.columns)), unsafe_allow_html=True)
                else:
                    st.markdown('<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: Archivo vacío o formato incorrecto</span></div>', unsafe_allow_html=True)
            except Exception as e:
                st.markdown(f'<div class="file-card-error" style="padding: 4px 12px; border-radius: 6px; margin-bottom: 8px;"><span class="status-error">❌ Error: {str(e)[:80]}...</span></div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        st.markdown(
            f"""
            <div class="welcome-footer">
                📅 {datetime.now().strftime('%d/%m/%Y %H:%M')} · Go To Market - {REGION_FIJA}
            </div>
            """,
            unsafe_allow_html=True
        )
    
    return file_log, file_cartera, file_riesgo, file_ventas

# ═══════════════════════════════════════════════════════════════════════════
# FUNCIONES DE RENDER DE EJES
# ═══════════════════════════════════════════════════════════════════════════

def render_eje_cartera(cartera_activa, meta_por_cliente, cartera_vigente, cartera, df_riesgo, clientes_nuevos, clientes_nuevos_reales):
    st.markdown("### 📋 CARTERA ACTIVA")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("**Cartera Activa Inicial**", cartera_activa)
    with col2:
        st.metric(
            "**Objetivo Clientes Nuevos**",
            clientes_nuevos_reales,
            delta=f"Meta: {meta_por_cliente}"
        )
    with col3:
        st.metric("**Cartera Activa Vigente**", cartera_vigente)
    
    st.divider()
    
    col_tab1, col_tab2 = st.columns(2)
    with col_tab1:
        render_distribucion_cartera(cartera)
    with col_tab2:
        render_tabla_clientes_riesgo(df_riesgo)
    
    st.markdown("---")
    render_tabla_clientes_nuevos(clientes_nuevos, clientes_nuevos_reales)

def render_eje_productividad(visitas, calidad, mes_sel=None):
    st.markdown("### 📈 PRODUCTIVIDAD")
    
    # ─── CALCULAR AVANCE LINEAL ────────────────────────────────────────
    avance_info = calcular_avance_lineal(mes_sel)
    avance_lineal = avance_info['avance_lineal']
    dias_trans = avance_info['dias_transcurridos']
    dias_totales = avance_info['dias_totales']
    dias_restantes = avance_info['dias_restantes']
    
    meta = visitas['meta_total']
    realizadas = visitas['total_visitas']
    pct = visitas['pct_cumplimiento']
    
    # ─── KPIs ──────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)
    
    with col1:
        pct_cobertura = calidad['pct_cobertura'] if calidad else 0
        cartera_visitada = calidad['cartera_visitada'] if calidad else 0
        total_cartera = calidad['total_cartera'] if calidad else 0
        st.metric(
            "**Objetivo Clientes Visitados**",
            f"{pct_cobertura:.0f}%",
            delta=f"{cartera_visitada} / {total_cartera} clientes"
        )
    
    with col2:
        st.metric(
            "**Objetivo de Visitas**",
            f"{realizadas:.0f} / {meta:.0f}",
            delta=f"{pct:.0f}% cumplimiento"
        )
    
    st.divider()
    
    # ─── BLOQUE INFORMATIVO (con días hábiles y alineación) ───────────
    
    # Calcular visitas esperadas (lineal)
    visitas_esperadas = (meta / dias_totales) * dias_trans if dias_totales > 0 else 0
    
    # Visitas que faltan para estar alineado
    faltan_alinear = visitas_esperadas - realizadas
    
    # Visitas que faltan para cumplir meta
    faltan_meta = meta - realizadas
    
    # Ritmo necesario
    ritmo = faltan_meta / dias_restantes if dias_restantes > 0 else 0
    
    # Estado
    if realizadas >= visitas_esperadas:
        estado = "🟢 Adelantado"
        color_estado = "#16A34A"
    elif faltan_alinear <= visitas_esperadas * 0.1:  # Menos del 10% de retraso
        estado = "⚠️ Al ritmo"
        color_estado = "#F59E0B"
    else:
        estado = "🔴 Retrasado"
        color_estado = "#DC2626"
    
    # ─── MOSTRAR BLOQUE ────────────────────────────────────────────────
    st.markdown(f"""
    <div style="
        background: #F8FAFC;
        border-radius: 10px;
        padding: 16px 20px;
        border-left: 5px solid {color_estado};
        font-size: 14px;
        line-height: 1.8;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    ">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <span>📅 <strong>Días hábiles transcurridos:</strong> {dias_trans} de {dias_totales}</span>
            <span style="color: {color_estado}; font-weight: 700;">{estado}</span>
        </div>
        <div style="display: flex; justify-content: space-between; margin-top: 2px;">
            <span>📈 <strong>Avance lineal esperado:</strong> {avance_lineal:.1f}% ({visitas_esperadas:.0f}/{meta:.0f})</span>
            <span style="color: {'#16A34A' if faltan_alinear <= 0 else '#DC2626'}; font-weight: 600;">
                📌 Te faltan {max(0, int(faltan_alinear))} visitas para estar alineado
            </span>
        </div>
        <div style="display: flex; justify-content: space-between; margin-top: 2px; color: #64748B;">
            <span>⏱️ <strong>{dias_restantes}</strong> días hábiles restantes</span>
            <span>🎯 <strong>Ritmo necesario:</strong> {ritmo:.1f} visitas/día</span>
        </div>
        <div style="margin-top: 4px; font-size: 13px; color: #94A3B8;">
            ⚠️ Necesitas {ritmo:.1f} visitas por día para cumplir la meta
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    # ─── BARRAS DE PROGRESO ──────────────────────────────────────────
    col_det1, col_det2 = st.columns(2)
    with col_det1:
        st.markdown("#### 📊 Visitas Cobertura")
        pct_cobertura = calidad['pct_cobertura'] if calidad else 0
        st.progress(min(pct_cobertura / 100, 1.0))
        st.caption(f"Cobertura de cartera: {pct_cobertura:.0f}%")
    
    with col_det2:
        st.markdown("#### 📊 Detalle de Visitas vs Meta")
        st.progress(min(pct / 100, 1.0))
        st.caption(f"Cumplimiento de visitas: {pct:.0f}%")
        
        st.markdown("#### 📈 Avance Lineal")
        st.progress(min(avance_lineal / 100, 1.0))
        st.caption(f"Avance lineal esperado: {avance_lineal:.0f}%")
    
    st.divider()
    
    # ─── DESGLOSE DE VISITAS ──────────────────────────────────────────
    if calidad and calidad['total_visitas'] > 0:
        st.markdown("#### 📋 Desglose de visitas")
        
        col_res1, col_res2, col_res3, col_res4 = st.columns(4)
        with col_res1:
            st.metric("**Total visitas**", calidad['total_visitas'])
        with col_res2:
            st.metric("**Clientes activos**", calidad['visitas_a_activos'])
        with col_res3:
            st.metric("**Prospectos**", calidad['visitas_a_prospectos'])
        with col_res4:
            st.metric("**Otros no activos**", calidad['visitas_a_otros_no_activos'])
        
        with st.expander("📌 Ver detalle de clientes visitados"):
            if calidad['clientes_activos_detalle']:
                total_clientes = len(calidad['clientes_activos_detalle'])
                total_visitas_cat = calidad['visitas_a_activos']
                st.markdown(f"**✅ Clientes activos visitados ({total_clientes} clientes únicos, {total_visitas_cat} visitas totales)**")
                df_activos = pd.DataFrame(calidad['clientes_activos_detalle'])
                df_activos.columns = ['Cliente', 'Visitas']
                st.dataframe(df_activos, use_container_width=True, hide_index=True)
            else:
                st.info("No hay clientes activos visitados")
            
            st.markdown("---")
            
            if calidad['clientes_prospectos_detalle']:
                total_clientes = len(calidad['clientes_prospectos_detalle'])
                total_visitas_cat = calidad['visitas_a_prospectos']
                st.markdown(f"**🆕 Prospectos visitados ({total_clientes} prospectos únicos, {total_visitas_cat} visitas totales)**")
                df_prospectos = pd.DataFrame(calidad['clientes_prospectos_detalle'])
                df_prospectos.columns = ['Prospecto', 'Visitas']
                st.dataframe(df_prospectos, use_container_width=True, hide_index=True)
            else:
                st.info("No hay prospectos visitados")
            
            st.markdown("---")
            
            if calidad['clientes_otros_detalle']:
                total_clientes = len(calidad['clientes_otros_detalle'])
                total_visitas_cat = calidad['visitas_a_otros_no_activos']
                st.markdown(f"**📌 Otros no activos visitados ({total_clientes} clientes únicos, {total_visitas_cat} visitas totales)**")
                df_otros = pd.DataFrame(calidad['clientes_otros_detalle'])
                df_otros.columns = ['Cliente (no activo)', 'Visitas']
                st.dataframe(df_otros, use_container_width=True, hide_index=True)
            else:
                st.info("No hay otros no activos visitados")
    else:
        st.info("No hay visitas registradas para los filtros seleccionados")

def render_eje_prospeccion(prospectos_iniciales, sqls, leads_desechados, tasa_conversion, clientes_nuevos_reales, total_prospectos_embudo, embudo, sqls_lead, leads_desechados_lead, meta_por_cliente):
    st.markdown("### 🎯 PROSPECCIÓN")
    
    pasan = total_prospectos_embudo - clientes_nuevos_reales - leads_desechados['cantidad']
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        st.metric("**Leads Embudo Iniciales**", prospectos_iniciales['cantidad'])
    with col2:
        st.metric("**Nuevos Leads (SQLs)**", sqls['cantidad'])
    with col3:
        st.metric("**Leads Desechados**", leads_desechados['cantidad'])
    with col4:
        st.metric(
            "**Objetivo Clientes Nuevos**",
            clientes_nuevos_reales,
            delta=f"Meta: {meta_por_cliente}"
        )
    with col5:
        st.metric(
            "**Tasa de Conversión**",
            f"{tasa_conversion:.1f}%",
            delta=f"{clientes_nuevos_reales}/{total_prospectos_embudo} cierres",
            delta_color="normal" if tasa_conversion >= 30 else "inverse"
        )
    with col6:
        st.metric("**Leads Embudo Finales**", pasan)
    
    st.divider()
    
    col_graf, col_tab = st.columns([3, 2])
    with col_graf:
        render_embudo(embudo)
    with col_tab:
        render_tabla_sqls(sqls_lead)
        st.markdown("---")
        render_tabla_leads_desechados(leads_desechados_lead)

def render_eje_ventas(ventas):
    st.markdown("### 💰 VENTAS")
    
    st.markdown("#### 📋 Objetivo de Ventas por Zona")
    
    if not ventas['totales_zona'].empty:
        df_zona = ventas['totales_zona'].copy()
        df_zona['objetivo'] = df_zona['objetivo'].apply(lambda x: f"S/ {x:,.0f}")
        df_zona['avance'] = df_zona['avance'].apply(lambda x: f"S/ {x:,.0f}")
        df_zona['pct'] = df_zona['pct'].apply(lambda x: f"{x:.1f}%")
        df_zona.columns = ['Zona', 'Objetivo', 'Real', '% Cumplimiento']
        
        total_objetivo = ventas['total_objetivo']
        total_avance = ventas['total_avance']
        total_pct = (total_avance / total_objetivo * 100) if total_objetivo > 0 else 0
        
        df_total = pd.DataFrame([{
            'Zona': '**TOTAL**',
            'Objetivo': f"S/ {total_objetivo:,.0f}",
            'Real': f"S/ {total_avance:,.0f}",
            '% Cumplimiento': f"{total_pct:.1f}%"
        }])
        
        df_zona = pd.concat([df_zona, df_total], ignore_index=True)
        render_tabla_html(df_zona, "Objetivo de Ventas por Zona")
    else:
        st.info("No hay datos de ventas por zona")
    
    st.markdown("---")
    
    st.markdown("#### 📊 Venta por Categoría (Agrupada)")
    
    if not ventas['totales_categoria'].empty:
        categorias_agrupadas = {
            "Perfumería + Skin Care": ["PERFUMERIA", "SKIN CARE"],
            "Cuidado Personal + Kids": ["CUIDADO PERSONAL", "KIDS Y BEBES"],
            "Toallitas Húmedas + Pañales": ["TOALLITAS HUMEDAS", "PAÑALES"]
        }
        
        df_cat = ventas['totales_categoria'].copy()
        df_cat['categoria'] = df_cat['categoria'].str.upper().str.strip()
        
        df_agrupado = pd.DataFrame()
        for grupo, categorias in categorias_agrupadas.items():
            df_grupo = df_cat[df_cat['categoria'].isin(categorias)]
            objetivo = df_grupo['objetivo'].sum()
            real = df_grupo['avance'].sum()
            pct = (real / objetivo * 100) if objetivo > 0 else 0
            
            df_agrupado = pd.concat([df_agrupado, pd.DataFrame({
                'Categoría': [grupo],
                'Objetivo': [f"S/ {objetivo:,.0f}"],
                'Real': [f"S/ {real:,.0f}"],
                '% Cumplimiento': [f"{pct:.1f}%"]
            })])
        
        render_tabla_html(df_agrupado, "Venta por Categoría (Agrupada)")
    else:
        st.info("No hay datos de ventas por categoría")

# ═══════════════════════════════════════════════════════════════════════════
# FUNCIONES RENDER (S1, S2, S3, S4)
# ═══════════════════════════════════════════════════════════════════════════

def render_s1_planificar(df_log, df_log_completo, df_lead, df_cartera, df_riesgo, df_ventas, zona_sel, region_sel, mes_sel):
    st.session_state.mes_sel = mes_sel
    
    st.markdown('<div class="main-header">🎯 S1 - Planificar</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Definición de objetivos y prioridades para la semana</div>', unsafe_allow_html=True)
    
    mes_cierre = obtener_mes_anterior(mes_sel, df_log)
    
    if mes_cierre:
        st.info(f"📅 **Datos del cierre:** {mes_cierre} (mes anterior al seleccionado)")
        df_log_filtrado = df_log[df_log['mes_ano'] == mes_cierre].copy() if df_log is not None else pd.DataFrame()
    else:
        st.warning("⚠️ No hay datos disponibles para el mes anterior")
        df_log_filtrado = pd.DataFrame()
    
    zonas_filtro = None if zona_sel == 'Todas' else [zona_sel]
    
    if not df_log_filtrado.empty and zonas_filtro:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
    
    cartera = df_cartera.copy() if df_cartera is not None else pd.DataFrame()
    if not cartera.empty and zonas_filtro:
        cartera = cartera[cartera['zona'].isin(zonas_filtro)]
    
    df_riesgo_filtrado = df_riesgo.copy() if df_riesgo is not None else pd.DataFrame()
    if not df_riesgo_filtrado.empty:
        if zonas_filtro and 'Todas' not in zonas_filtro:
            df_riesgo_filtrado = df_riesgo_filtrado[df_riesgo_filtrado['zona'].isin(zonas_filtro)]
    
    cartera_activa = len(cartera)
    meta_por_cliente = obtener_meta_por_cliente(zonas_filtro)
    
    clientes_nuevos = procesar_clientes_nuevos(df_log_filtrado, zonas_filtro)
    clientes_nuevos_reales = clientes_nuevos['cantidad']
    cartera_vigente = calcular_cartera_vigente(cartera_activa, clientes_nuevos_reales)
    visitas = procesar_visitas(df_log_filtrado, zonas_filtro, mes_cierre)
    calidad = procesar_calidad_visita(df_log_filtrado, cartera, zonas_filtro)
    
    prospectos_iniciales_lead = procesar_prospectos_iniciales_lead(df_lead, df_log, zonas_filtro, mes_cierre)
    sqls_lead = procesar_sqls_lead(df_lead, zonas_filtro, mes_cierre)
    embudo = procesar_embudo_completo(df_log, df_lead, zonas_filtro, mes_cierre)
    leads_desechados = procesar_leads_desechados(df_log_filtrado, zonas_filtro)
    
    total_prospectos_embudo = prospectos_iniciales_lead['cantidad'] + sqls_lead['cantidad']
    tasa_conversion = calcular_tasa_conversion(clientes_nuevos_reales, total_prospectos_embudo)
    ventas = procesar_ventas(df_ventas, zonas_filtro, region_sel)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("📋 Cartera Activa Inicial", cartera_activa)
    with col2:
        st.metric(
            "🎯 Objetivo Clientes Nuevos",
            clientes_nuevos_reales,
            delta=f"Meta: {meta_por_cliente}",
            delta_color="normal" if clientes_nuevos_reales >= meta_por_cliente else "inverse"
        )
    with col3:
        st.metric(
            "📊 Cartera Activa Vigente",
            cartera_vigente,
            delta=f"Activos: {cartera_activa} + Nuevos: {clientes_nuevos_reales}",
            delta_color="normal"
        )
    with col4:
        meta = visitas['meta_total']
        realizadas = visitas['total_visitas']
        pct = visitas['pct_cumplimiento']
        st.metric(
            "📊 Objetivo de Visitas",
            f"{realizadas:.0f}/{meta:.0f}",
            delta=f"{pct:.0f}% cumplimiento (5×días hábiles)",
            delta_color="normal" if pct >= 80 else "inverse"
        )
    with col5:
        pct_cobertura = calidad['pct_cobertura']
        cartera_visitada = calidad['cartera_visitada']
        total_cartera = calidad['total_cartera']
        st.metric(
            "📋 Objetivo Clientes Visitados",
            f"{pct_cobertura:.0f}%",
            delta=f"{cartera_visitada}/{total_cartera} clientes",
            delta_color="normal" if pct_cobertura >= 70 else "inverse"
        )
    
    st.divider()
    
    st.markdown("#### 📊 Resumen de Prospección")
    resumen = procesar_resumen_prospeccion(df_lead, df_log, df_log_filtrado, zonas_filtro, mes_cierre)
    col_r1, col_r2, col_r3, col_r4, col_r5, col_r6 = st.columns(6)
    with col_r1:
        st.metric("📌 Iniciales", resumen['iniciales'])
    with col_r2:
        st.metric("🆕 Nuevos", resumen['sqls'])
    with col_r3:
        st.metric("📊 Total Activos", resumen['total_activos'])
    with col_r4:
        st.metric("✅ Cierres", resumen['cierres'])
    with col_r5:
        st.metric("❌ Desechados", resumen['desechados'])
    with col_r6:
        st.metric("🔄 Pasan", resumen['pasan'])
    
    st.divider()
    
    eje_actual = "s1"
    eje_seleccionado = render_botones_ejes(eje_actual)
    
    if eje_seleccionado == "cartera":
        render_eje_cartera(cartera_activa, meta_por_cliente, cartera_vigente, cartera, df_riesgo_filtrado, clientes_nuevos, clientes_nuevos_reales)
    elif eje_seleccionado == "productividad":
        render_eje_productividad(visitas, calidad, mes_cierre)
    elif eje_seleccionado == "prospeccion":
        render_eje_prospeccion(prospectos_iniciales_lead, sqls_lead, leads_desechados, tasa_conversion, clientes_nuevos_reales, total_prospectos_embudo, embudo, sqls_lead, leads_desechados, meta_por_cliente)
    elif eje_seleccionado == "ventas":
        render_eje_ventas(ventas)
    else:
        st.info("👈 Selecciona un eje de gestión para visualizar los indicadores detallados")

def render_s2_ejecutar(df_log, df_log_completo, df_lead, df_cartera, df_riesgo, df_ventas, zona_sel, region_sel, mes_sel):
    st.session_state.mes_sel = mes_sel
    
    st.markdown('<div class="main-header">🚀 S2 - Ejecutar</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Seguimiento de visitas y atención a clientes</div>', unsafe_allow_html=True)
    
    df_log_filtrado = df_log.copy() if df_log is not None else pd.DataFrame()
    if mes_sel != 'Todos' and not df_log_filtrado.empty and 'mes_ano' in df_log_filtrado.columns:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['mes_ano'] == mes_sel]
    
    zonas_filtro = None if zona_sel == 'Todas' else [zona_sel]
    
    if not df_log_filtrado.empty and zonas_filtro:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
    
    cartera = df_cartera.copy() if df_cartera is not None else pd.DataFrame()
    if not cartera.empty and zonas_filtro:
        cartera = cartera[cartera['zona'].isin(zonas_filtro)]
    
    df_riesgo_filtrado = df_riesgo.copy() if df_riesgo is not None else pd.DataFrame()
    if not df_riesgo_filtrado.empty:
        if zonas_filtro and 'Todas' not in zonas_filtro:
            df_riesgo_filtrado = df_riesgo_filtrado[df_riesgo_filtrado['zona'].isin(zonas_filtro)]
    
    cartera_activa = len(cartera)
    meta_por_cliente = obtener_meta_por_cliente(zonas_filtro)
    
    clientes_nuevos = procesar_clientes_nuevos(df_log_filtrado, zonas_filtro)
    clientes_nuevos_reales = clientes_nuevos['cantidad']
    cartera_vigente = calcular_cartera_vigente(cartera_activa, clientes_nuevos_reales)
    visitas = procesar_visitas(df_log_filtrado, zonas_filtro, mes_sel)
    calidad = procesar_calidad_visita(df_log_filtrado, cartera, zonas_filtro)
    
    prospectos_iniciales_lead = procesar_prospectos_iniciales_lead(df_lead, df_log, zonas_filtro, mes_sel)
    sqls_lead = procesar_sqls_lead(df_lead, zonas_filtro, mes_sel)
    embudo = procesar_embudo_completo(df_log, df_lead, zonas_filtro, mes_sel)
    leads_desechados = procesar_leads_desechados(df_log_filtrado, zonas_filtro)
    
    total_prospectos_embudo = prospectos_iniciales_lead['cantidad'] + sqls_lead['cantidad']
    tasa_conversion = calcular_tasa_conversion(clientes_nuevos_reales, total_prospectos_embudo)
    ventas = procesar_ventas(df_ventas, zonas_filtro, region_sel)
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.metric(
            "📊 Objetivo de Visitas",
            f"{visitas['total_visitas']}/{visitas['meta_total']:.0f}",
            delta=f"{visitas['pct_cumplimiento']:.0f}% cumplimiento",
            delta_color="normal" if visitas['pct_cumplimiento'] >= 80 else "inverse"
        )
    
    with col2:
        st.metric("📈 % Cumplimiento Visitas", f"{visitas['pct_cumplimiento']:.0f}%")
    
    with col3:
        pct_cobertura = calidad['pct_cobertura']
        cartera_visitada = calidad['cartera_visitada']
        total_cartera = calidad['total_cartera']
        st.metric(
            "📋 Objetivo Clientes Visitados",
            f"{pct_cobertura:.0f}%",
            delta=f"{cartera_visitada}/{total_cartera} clientes",
            delta_color="normal" if pct_cobertura >= 70 else "inverse"
        )
    
    with col4:
        st.metric("⚠️ Clientes en Riesgo", len(df_riesgo_filtrado), delta="requieren visita" if len(df_riesgo_filtrado) > 0 else None, delta_color="inverse" if len(df_riesgo_filtrado) > 0 else "normal")
    
    with col5:
        st.metric(
            "💰 Objetivo de Ventas",
            f"S/ {ventas['total_avance']:,.0f}",
            delta=f"{ventas['pct_general']:.0f}% vs presupuesto",
            delta_color="normal" if ventas['pct_general'] >= 80 else "inverse"
        )
    
    with col6:
        st.metric("📋 Cartera Activa Inicial", cartera_activa)
    
    st.divider()
    
    eje_actual = "s2"
    eje_seleccionado = render_botones_ejes(eje_actual)
    
    if eje_seleccionado == "cartera":
        render_eje_cartera(cartera_activa, meta_por_cliente, cartera_vigente, cartera, df_riesgo_filtrado, clientes_nuevos, clientes_nuevos_reales)
    elif eje_seleccionado == "productividad":
        render_eje_productividad(visitas, calidad, mes_sel)
    elif eje_seleccionado == "prospeccion":
        render_eje_prospeccion(prospectos_iniciales_lead, sqls_lead, leads_desechados, tasa_conversion, clientes_nuevos_reales, total_prospectos_embudo, embudo, sqls_lead, leads_desechados, meta_por_cliente)
    elif eje_seleccionado == "ventas":
        render_eje_ventas(ventas)
    else:
        st.info("👈 Selecciona un eje de gestión para visualizar los indicadores detallados")

def render_s3_convertir(df_log, df_log_completo, df_lead, df_ventas, zona_sel, region_sel, mes_sel):
    st.session_state.mes_sel = mes_sel
    
    st.markdown('<div class="main-header">🔄 S3 - Convertir</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Conversión de visitas en ventas y cierres</div>', unsafe_allow_html=True)
    
    df_log_filtrado = df_log.copy() if df_log is not None else pd.DataFrame()
    if mes_sel != 'Todos' and not df_log_filtrado.empty and 'mes_ano' in df_log_filtrado.columns:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['mes_ano'] == mes_sel]
    
    zonas_filtro = None if zona_sel == 'Todas' else [zona_sel]
    
    if not df_log_filtrado.empty and zonas_filtro:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
    
    cartera = st.session_state.df_cartera if 'df_cartera' in st.session_state else pd.DataFrame()
    if not cartera.empty and zonas_filtro:
        cartera = cartera[cartera['zona'].isin(zonas_filtro)]
    
    cartera_activa = len(cartera)
    meta_por_cliente = obtener_meta_por_cliente(zonas_filtro)
    
    clientes_nuevos = procesar_clientes_nuevos(df_log_filtrado, zonas_filtro)
    clientes_nuevos_reales = clientes_nuevos['cantidad']
    cartera_vigente = calcular_cartera_vigente(cartera_activa, clientes_nuevos_reales)
    visitas = procesar_visitas(df_log_filtrado, zonas_filtro, mes_sel)
    calidad = procesar_calidad_visita(df_log_filtrado, cartera, zonas_filtro)
    ventas = procesar_ventas(df_ventas, zonas_filtro, region_sel)
    
    prospectos_iniciales_lead = procesar_prospectos_iniciales_lead(df_lead, df_log, zonas_filtro, mes_sel)
    sqls_lead = procesar_sqls_lead(df_lead, zonas_filtro, mes_sel)
    embudo = procesar_embudo_completo(df_log, df_lead, zonas_filtro, mes_sel)
    leads_desechados = procesar_leads_desechados(df_log_filtrado, zonas_filtro)
    
    total_prospectos_embudo = prospectos_iniciales_lead['cantidad'] + sqls_lead['cantidad']
    tasa_conversion = calcular_tasa_conversion(clientes_nuevos_reales, total_prospectos_embudo)
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            "📊 Objetivo de Visitas",
            f"{visitas['total_visitas']}/{visitas['meta_total']:.0f}",
            delta=f"{visitas['pct_cumplimiento']:.0f}% cumplimiento",
            delta_color="normal" if visitas['pct_cumplimiento'] >= 80 else "inverse"
        )
    
    with col2:
        st.metric("🆕 Nuevos Prospectos (SQLs)", sqls_lead['cantidad'])
    
    with col3:
        st.metric(
            "💰 Objetivo de Ventas",
            f"S/ {ventas['total_avance']:,.0f}",
            delta=f"{ventas['pct_general']:.0f}% vs presupuesto",
            delta_color="normal" if ventas['pct_general'] >= 80 else "inverse"
        )
    
    with col4:
        st.metric("📊 % Cumplimiento Ventas", f"{ventas['pct_general']:.0f}%")
    
    with col5:
        st.metric("✅ Clientes Nuevos", clientes_nuevos_reales)
    
    st.divider()
    
    eje_actual = "s3"
    eje_seleccionado = render_botones_ejes(eje_actual)
    
    if eje_seleccionado == "cartera":
        render_eje_cartera(cartera_activa, meta_por_cliente, cartera_vigente, cartera, pd.DataFrame(), clientes_nuevos, clientes_nuevos_reales)
    elif eje_seleccionado == "productividad":
        render_eje_productividad(visitas, calidad, mes_sel)
    elif eje_seleccionado == "prospeccion":
        render_eje_prospeccion(prospectos_iniciales_lead, sqls_lead, leads_desechados, tasa_conversion, clientes_nuevos_reales, total_prospectos_embudo, embudo, sqls_lead, leads_desechados, meta_por_cliente)
    elif eje_seleccionado == "ventas":
        render_eje_ventas(ventas)
    else:
        st.info("👈 Selecciona un eje de gestión para visualizar los indicadores detallados")

def render_s4_cerrar(df_log, df_log_completo, df_lead, df_cartera, df_riesgo, df_ventas, zona_sel, region_sel, mes_sel):
    st.session_state.mes_sel = mes_sel
    
    st.markdown('<div class="main-header">🏁 S4 - Cerrar</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Resumen ejecutivo y próximos pasos</div>', unsafe_allow_html=True)
    
    df_log_filtrado = df_log.copy() if df_log is not None else pd.DataFrame()
    if mes_sel != 'Todos' and not df_log_filtrado.empty and 'mes_ano' in df_log_filtrado.columns:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['mes_ano'] == mes_sel]
    
    zonas_filtro = None if zona_sel == 'Todas' else [zona_sel]
    
    if not df_log_filtrado.empty and zonas_filtro:
        df_log_filtrado = df_log_filtrado[df_log_filtrado['zona'].isin(zonas_filtro)]
    
    cartera = df_cartera.copy() if df_cartera is not None else pd.DataFrame()
    if not cartera.empty and zonas_filtro:
        cartera = cartera[cartera['zona'].isin(zonas_filtro)]
    
    df_riesgo_filtrado = df_riesgo.copy() if df_riesgo is not None else pd.DataFrame()
    if not df_riesgo_filtrado.empty:
        if zonas_filtro and 'Todas' not in zonas_filtro:
            df_riesgo_filtrado = df_riesgo_filtrado[df_riesgo_filtrado['zona'].isin(zonas_filtro)]
    
    cartera_activa = len(cartera)
    meta_por_cliente = obtener_meta_por_cliente(zonas_filtro)
    
    clientes_nuevos = procesar_clientes_nuevos(df_log_filtrado, zonas_filtro)
    clientes_nuevos_reales = clientes_nuevos['cantidad']
    cartera_vigente = calcular_cartera_vigente(cartera_activa, clientes_nuevos_reales)
    visitas = procesar_visitas(df_log_filtrado, zonas_filtro, mes_sel)
    calidad = procesar_calidad_visita(df_log_filtrado, cartera, zonas_filtro)
    ventas = procesar_ventas(df_ventas, zonas_filtro, region_sel)
    
    prospectos_iniciales_lead = procesar_prospectos_iniciales_lead(df_lead, df_log, zonas_filtro, mes_sel)
    sqls_lead = procesar_sqls_lead(df_lead, zonas_filtro, mes_sel)
    embudo = procesar_embudo_completo(df_log, df_lead, zonas_filtro, mes_sel)
    leads_desechados = procesar_leads_desechados(df_log_filtrado, zonas_filtro)
    
    total_prospectos_embudo = prospectos_iniciales_lead['cantidad'] + sqls_lead['cantidad']
    tasa_conversion = calcular_tasa_conversion(clientes_nuevos_reales, total_prospectos_embudo)
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.metric(
            "📊 Objetivo de Visitas",
            f"{visitas['total_visitas']}/{visitas['meta_total']:.0f}",
            delta=f"{visitas['pct_cumplimiento']:.0f}% cumplimiento",
            delta_color="normal" if visitas['pct_cumplimiento'] >= 80 else "inverse"
        )
    
    with col2:
        st.metric(
            "💰 Objetivo de Ventas",
            f"S/ {ventas['total_avance']:,.0f}",
            delta=f"{ventas['pct_general']:.0f}% vs presupuesto",
            delta_color="normal" if ventas['pct_general'] >= 80 else "inverse"
        )
    
    with col3:
        pct_cobertura = calidad['pct_cobertura']
        cartera_visitada = calidad['cartera_visitada']
        total_cartera = calidad['total_cartera']
        st.metric(
            "📋 Objetivo Clientes Visitados",
            f"{pct_cobertura:.0f}%",
            delta=f"{cartera_visitada}/{total_cartera} clientes",
            delta_color="normal" if pct_cobertura >= 70 else "inverse"
        )
    
    with col4:
        st.metric("⚠️ Clientes en Riesgo", len(df_riesgo_filtrado), delta="pendientes" if len(df_riesgo_filtrado) > 0 else "✅ recuperados", delta_color="inverse" if len(df_riesgo_filtrado) > 0 else "normal")
    
    with col5:
        st.metric("🆕 Nuevos Prospectos (SQLs)", sqls_lead['cantidad'])
    
    with col6:
        st.metric("✅ Clientes Nuevos", clientes_nuevos_reales)
    
    st.divider()
    
    eje_actual = "s4"
    eje_seleccionado = render_botones_ejes(eje_actual)
    
    if eje_seleccionado == "cartera":
        render_eje_cartera(cartera_activa, meta_por_cliente, cartera_vigente, cartera, df_riesgo_filtrado, clientes_nuevos, clientes_nuevos_reales)
    elif eje_seleccionado == "productividad":
        render_eje_productividad(visitas, calidad, mes_sel)
    elif eje_seleccionado == "prospeccion":
        render_eje_prospeccion(prospectos_iniciales_lead, sqls_lead, leads_desechados, tasa_conversion, clientes_nuevos_reales, total_prospectos_embudo, embudo, sqls_lead, leads_desechados, meta_por_cliente)
    elif eje_seleccionado == "ventas":
        render_eje_ventas(ventas)
    else:
        st.info("👈 Selecciona un eje de gestión para visualizar los indicadores detallados")

def procesar_resumen_prospeccion(df_lead, df_log, df_log_filtrado, zonas_filtro=None, mes_sel=None):
    iniciales = procesar_prospectos_iniciales_lead(df_lead, df_log, zonas_filtro, mes_sel)
    sqls = procesar_sqls_lead(df_lead, zonas_filtro, mes_sel)
    total_activos = iniciales['cantidad'] + sqls['cantidad']
    
    cierres = 0
    desechados = 0
    
    if df_log_filtrado is not None and not df_log_filtrado.empty:
        df_cierres = df_log_filtrado[
            (df_log_filtrado['tipo'] == 'PROSPECCIÓN') & 
            (df_log_filtrado['task'].astype(str).str.upper().str.contains('CIERRE', na=False)) &
            (~df_log_filtrado['task'].astype(str).str.upper().str.contains('NO CIERRE', na=False))
        ]
        cierres = df_cierres['cliente'].nunique()
        
        df_desechados = df_log_filtrado[
            (df_log_filtrado['tipo'] == 'PROSPECCIÓN') & 
            (df_log_filtrado['task'].astype(str).str.upper().str.contains('NO CIERRE', na=False))
        ]
        desechados = df_desechados['cliente'].nunique()
    
    pasan = total_activos - cierres - desechados
    
    return {
        'iniciales': iniciales['cantidad'],
        'iniciales_clientes': iniciales['clientes'],
        'sqls': sqls['cantidad'],
        'sqls_clientes': sqls['clientes'],
        'total_activos': total_activos,
        'cierres': cierres,
        'desechados': desechados,
        'pasan': pasan
    }

# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

def main():
    if 'eje_seleccionado_s1' not in st.session_state:
        st.session_state.eje_seleccionado_s1 = "todos"
    if 'eje_seleccionado_s2' not in st.session_state:
        st.session_state.eje_seleccionado_s2 = "todos"
    if 'eje_seleccionado_s3' not in st.session_state:
        st.session_state.eje_seleccionado_s3 = "todos"
    if 'eje_seleccionado_s4' not in st.session_state:
        st.session_state.eje_seleccionado_s4 = "todos"
    
    if 'archivos_cargados' not in st.session_state:
        st.session_state.archivos_cargados = False
        st.session_state.file_log = None
        st.session_state.file_cartera = None
        st.session_state.file_riesgo = None
        st.session_state.file_ventas = None
        st.session_state.df_log = None
        st.session_state.df_lead = None
        st.session_state.df_cartera = None
        st.session_state.df_riesgo = None
        st.session_state.df_ventas = None
        st.session_state.df_log_completo = None
    
    st.sidebar.markdown("### ⚙️ Panel de Control")
    
    if not st.session_state.archivos_cargados:
        file_log, file_cartera, file_riesgo, file_ventas = mostrar_pantalla_carga()
        
        archivos_validos = True
        if not all([file_log, file_cartera, file_riesgo, file_ventas]):
            archivos_validos = False
        else:
            try:
                df_log = cargar_log_visitas(file_log.read())
                file_log.seek(0)
                if df_log.empty:
                    archivos_validos = False
            except:
                archivos_validos = False
            
            try:
                df_lead = cargar_lead(file_log.read())
                file_log.seek(0)
                if df_lead.empty:
                    st.warning("⚠️ No se encontró la hoja 'Lead' en el archivo Log. La prospección se verá limitada.")
            except:
                df_lead = pd.DataFrame()
                st.warning("⚠️ Error al leer la hoja 'Lead'. La prospección se verá limitada.")
            
            try:
                df_cartera = cargar_cartera_activa(file_cartera.read())
                file_cartera.seek(0)
                if df_cartera.empty:
                    archivos_validos = False
            except:
                archivos_validos = False
            
            try:
                df_riesgo = cargar_clientes_riesgo(file_riesgo.read())
                file_riesgo.seek(0)
                if df_riesgo.empty:
                    archivos_validos = False
            except:
                archivos_validos = False
            
            try:
                df_ventas = cargar_ventas_categoria(file_ventas.read())
                file_ventas.seek(0)
                if df_ventas.empty:
                    archivos_validos = False
            except:
                archivos_validos = False
        
        if not archivos_validos:
            st.sidebar.warning("⚠️ Carga los 4 archivos correctamente para comenzar")
            return
        
        st.session_state.archivos_cargados = True
        st.session_state.file_log = file_log
        st.session_state.file_cartera = file_cartera
        st.session_state.file_riesgo = file_riesgo
        st.session_state.file_ventas = file_ventas
        st.session_state.df_log = df_log
        st.session_state.df_lead = df_lead if 'df_lead' in locals() else pd.DataFrame()
        st.session_state.df_cartera = df_cartera
        st.session_state.df_riesgo = df_riesgo
        st.session_state.df_ventas = df_ventas
        st.session_state.df_log_completo = df_log.copy()
        
        st.rerun()
        return
    
    df_log = st.session_state.df_log
    df_lead = st.session_state.df_lead
    df_cartera = st.session_state.df_cartera
    df_riesgo = st.session_state.df_riesgo
    df_ventas = st.session_state.df_ventas
    df_log_completo = st.session_state.df_log_completo
    
    st.sidebar.success("✅ Archivos cargados correctamente")
    st.sidebar.divider()
    
    st.sidebar.markdown("#### 🎯 Filtros")
    
    meses_disponibles = ['Todos']
    if not df_log.empty and 'mes_ano' in df_log.columns:
        meses = sorted(df_log['mes_ano'].unique())
        meses_disponibles.extend(meses)
    
    mes_sel = st.sidebar.selectbox("📅 Mes (Log de Visitas)", meses_disponibles)
    
    # ✅ REGIÓN FIJA: No mostramos selector de región
    st.sidebar.info(f"🌎 Región: {REGION_FIJA}")
    
    # ✅ Zonas disponibles solo de esta región
    zonas_opciones = ["Todas"] + sorted(TODAS_ZONAS)
    zona_sel = st.sidebar.selectbox("📍 Zona", zonas_opciones)
    
    st.sidebar.divider()
    st.sidebar.caption(f"📊 Datos actualizados: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    if st.sidebar.button("🔄 Recargar archivos"):
        st.session_state.archivos_cargados = False
        st.rerun()
    
    tab1, tab2, tab3, tab4 = st.tabs(["🎯 S1 - Planificar", "🚀 S2 - Ejecutar", "🔄 S3 - Convertir", "🏁 S4 - Cerrar"])
    
    with tab1:
        render_s1_planificar(df_log, df_log_completo, df_lead, df_cartera, df_riesgo, df_ventas, zona_sel, REGION_FIJA, mes_sel)
    with tab2:
        render_s2_ejecutar(df_log, df_log_completo, df_lead, df_cartera, df_riesgo, df_ventas, zona_sel, REGION_FIJA, mes_sel)
    with tab3:
        render_s3_convertir(df_log, df_log_completo, df_lead, df_ventas, zona_sel, REGION_FIJA, mes_sel)
    with tab4:
        render_s4_cerrar(df_log, df_log_completo, df_lead, df_cartera, df_riesgo, df_ventas, zona_sel, REGION_FIJA, mes_sel)

if __name__ == "__main__":
    main()
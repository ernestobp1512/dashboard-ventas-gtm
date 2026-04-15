import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Reporte de Visitas Comerciales",
    page_icon="📊",
    layout="wide",
)

# ─── ESTILOS ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Contenedores generales */
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 20px;
        margin-top: 10px;
        margin-bottom: 25px;
    }
    .kpi-card {
        flex: 1;
        background-color: var(--secondary-background-color) !important;
        border-radius: 8px;
        padding: 24px 20px;
        border-left: 4px solid #1c64f2; 
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        color: var(--text-color) !important;
        position: relative;
    }
    .kpi-card-icon {
        position: absolute;
        top: 20px;
        right: 20px;
        width: 36px;
        height: 36px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
    }
    
    .kpi-card:nth-child(1) { border-left-color: #3b82f6; }
    .kpi-card:nth-child(1) .kpi-card-icon { background-color: #dbeafe; color: #1d4ed8; }
    .kpi-card:nth-child(2) { border-left-color: #10b981; }
    .kpi-card:nth-child(2) .kpi-card-icon { background-color: #d1fae5; color: #047857; }
    .kpi-card:nth-child(3) { border-left-color: #f59e0b; }
    .kpi-card:nth-child(3) .kpi-card-icon { background-color: #fef3c7; color: #b45309; }
    .kpi-card:nth-child(4) { border-left-color: #8b5cf6; }
    .kpi-card:nth-child(4) .kpi-card-icon { background-color: #ede9fe; color: #6d28d9; }

    .kpi-value {
        font-size: 32px;
        font-weight: 800;
        margin-top: 10px;
        margin-bottom: 5px;
        color: #1e3a8a;
    }
    .kpi-label {
        font-size: 14px;
        opacity: 0.7;
    }
    
    /* Bloques de gráficos de región */
    .dashboard-grid {
        display: grid;
        grid-template-columns: 2fr 1fr;
        gap: 20px;
        margin-bottom: 25px;
    }
    .dashboard-panel {
        background-color: var(--secondary-background-color) !important;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        color: var(--text-color) !important;
        border: 1px solid rgba(128,128,128,0.2);
    }
    .panel-title {
        font-size: 18px;
        font-weight: 700;
        color: #1e3a8a;
        margin-bottom: 4px;
    }
    .panel-subtitle {
        font-size: 13px;
        opacity: 0.6;
        margin-bottom: 20px;
    }
    
    /* Rutas */
    .rutas-container {
        display: flex;
        flex-wrap: wrap;
        gap: 15px;
    }
    .ruta-card {
        flex: 1 1 calc(25% - 15px);
        min-width: 140px;
        border: 1px solid rgba(128,128,128,0.2);
        border-radius: 10px;
        padding: 15px;
    }
    .ruta-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 12px;
    }
    .ruta-name {
        font-size: 14px;
        opacity: 0.8;
    }
    .ruta-value {
        font-weight: 700;
        font-size: 16px;
        color: #1e3a8a;
    }
    .ruta-progress-container {
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .ruta-progress-bar {
        flex-grow: 1;
        background-color: rgba(128,128,128,0.2);
        height: 6px;
        border-radius: 3px;
        overflow: hidden;
    }
    .ruta-progress-fill {
        height: 100%;
        border-radius: 3px;
    }
    .ruta-pct {
        font-size: 12px;
        opacity: 0.6;
        min-width: 45px;
        text-align: right;
    }
    
    /* Conversiones */
    .conv-list {
        display: flex;
        flex-direction: column;
        gap: 16px;
    }
    .conv-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .conv-item-left {
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .conv-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
    }
    .conv-name {
        font-size: 14px;
    }
    .conv-value {
        font-weight: 700;
        font-size: 14px;
    }

    /* Tabla de Estados */
    .styled-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }
    .styled-table thead tr {
        border-bottom: 2px solid rgba(128,128,128,0.2);
        color: #1e3a8a;
        text-align: left;
    }
    .styled-table th, .styled-table td {
        padding: 12px 10px;
    }
    .styled-table tbody tr {
        border-bottom: 1px solid rgba(128,128,128,0.1);
    }
    .styled-table tbody tr:last-of-type {
        border-bottom: none;
    }
    
    [data-baseweb="tab"][aria-selected="true"] {
        border-bottom: 2px solid var(--primary-color) !important;
        color: var(--primary-color) !important;
    }
    
    .header-title-container-main {
        text-align: center;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    .header-title {
        color: #1e3a8a;
        font-size: 38px;
        font-weight: 800;
        margin-bottom: 5px;
    }
    
</style>
""", unsafe_allow_html=True)

_theme   = st.get_option("theme.base") or "dark"
_is_dark = (_theme == "dark")

if _is_dark:
    st.markdown("""
    <style>
        .header-title, .panel-title, .kpi-value, .ruta-value { color: #60a5fa !important; }
        .styled-table thead tr { color: #60a5fa !important; }
    </style>
    """, unsafe_allow_html=True)


# ─── CONSTANTES ────────────────────────────────────────────────────────────────
EXCEL_FILE = "visitas_ventas.xlsx"

COL_FECHA    = "Date"
COL_TIPO     = "Tipo" 
COL_TIPO_CLI = "Giro"
COL_CLIENTE  = "Cliente o Prospecto"
COL_DISTRITO = "Distrito"
COL_MOTIVO   = "Task"
COL_RESULTADO= "Obs"
COL_ZONA     = "Zona"
COL_REGION   = "Región"
COL_DEPARTA  = "Departamento"
COL_PROVINCIA= "Provincia"
COL_TIPO_VIS = "Tipo Visita" 
COL_VENDEDOR = "Vendedor"

ETAPAS_EMBUDO = [
    "PROSPECCIÓN",
    "CALIFICACIÓN DE LEADS",
    "VISITA",
    "PROPUESTA",
    "NEGOCIACIÓN",
    "CIERRE",
    "NO CIERRE",
]

PALETA_RUTAS = ["#3b82f6", "#10b981", "#f59e0b", "#8b5cf6", "#ef4444", "#0ea5e9", "#14b8a6", "#f43f5e"]

# ─── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_datos_completos(ruta):
    df_log = pd.read_excel(ruta, sheet_name="Log", engine="openpyxl")
    df_users = pd.read_excel(ruta, sheet_name="Users", engine="openpyxl")
    df_zona = pd.read_excel(ruta, sheet_name="Zona", engine="openpyxl")
    
    df_log.columns = df_log.columns.str.strip()
    df_users.columns = df_users.columns.str.strip()
    df_zona.columns = df_zona.columns.str.strip()

    if 'User' in df_log.columns and 'Email' in df_users.columns:
        df_log = df_log.merge(df_users[['Email', 'Name']], left_on='User', right_on='Email', how='left')
        df_log[COL_VENDEDOR] = df_log['Name'].fillna("Desconocido")
    else:
        df_log[COL_VENDEDOR] = "Desconocido"

    if 'Zona' in df_log.columns and 'Zona' in df_zona.columns:
        df_log = df_log.merge(df_zona[['Zona', 'Tipo Zona']], on='Zona', how='left')
        df_log[COL_REGION] = df_log['Tipo Zona'].fillna("Desconocido")
    else:
        df_log[COL_REGION] = "Desconocido"

    df_log[COL_FECHA] = pd.to_datetime(df_log[COL_FECHA], dayfirst=True, errors="coerce")
    df_log = df_log.dropna(subset=[COL_FECHA])
    
    df_log["_semana_nro"] = df_log[COL_FECHA].dt.isocalendar().week.astype(str).str.zfill(2)
    df_log["_anio_semana"] = df_log[COL_FECHA].dt.isocalendar().year.astype(str)
    df_log["_sem_lbl"] = df_log["_anio_semana"] + "-S" + df_log["_semana_nro"]
    df_log["_mes_lbl"] = df_log[COL_FECHA].dt.strftime("%Y-%m")

    for col in [COL_VENDEDOR, COL_TIPO, COL_TIPO_CLI, COL_CLIENTE, COL_DISTRITO, COL_MOTIVO, COL_TIPO_VIS, COL_REGION, COL_ZONA]:
        if col in df_log.columns:
            df_log[col] = df_log[col].astype(str).str.strip()

    # Carga de hojas de Estado
    try:
        df_estado_semana = pd.read_excel(ruta, sheet_name="Estado_Semana", engine="openpyxl")
    except Exception:
        df_estado_semana = pd.DataFrame(columns=["Estado", "Cantidad"])

    try:
        df_estado_mes = pd.read_excel(ruta, sheet_name="Estado_Mes", engine="openpyxl")
    except Exception:
        df_estado_mes = pd.DataFrame(columns=["Estado", "Cantidad"])

    return df_log, df_estado_semana, df_estado_mes

# ── Selector de archivo ───────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    archivo = st.file_uploader(
        "Cargar Excel de visitas",
        type=["xlsx", "xls", "xlsm"],
    )

    ruta_usar = Path(EXCEL_FILE) if archivo is None else archivo
    if archivo is None:
        st.warning("⚠️ Sube un archivo Excel para comenzar.")
        st.stop()

    try:
        df_raw, df_estado_sem, df_estado_mes = cargar_datos_completos(ruta_usar)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    if df_raw.empty:
        st.error("El archivo está vacío o no tiene filas válidas.")
        st.stop()

    st.divider()
    # ── Filtros ───────────────────────────────────────────────────────────────
    st.markdown("### 🔽 Filtros")

    vendedores_list = ["Todos"] + sorted(df_raw[COL_VENDEDOR].dropna().unique().tolist())
    sel_vendedor = st.selectbox("Vendedor", vendedores_list)

    st.markdown("#### Periodo de Filtro")
    modo_fecha = st.radio("Agrupar por:", ["Mes", "Semana"], horizontal=True)
    df_raw = df_raw.sort_values(COL_FECHA)
    
    if modo_fecha == "Mes":
        lista_opciones = sorted(df_raw["_mes_lbl"].dropna().unique().tolist())
        sel_rango = st.select_slider("Rango de Meses", options=lista_opciones, value=(lista_opciones[0], lista_opciones[-1]) if len(lista_opciones)>1 else lista_opciones[0]) if lista_opciones else []
    else:
        lista_opciones = sorted(df_raw["_sem_lbl"].dropna().unique().tolist())
        sel_rango = st.select_slider("Rango de Semanas", options=lista_opciones, value=(lista_opciones[0], lista_opciones[-1]) if len(lista_opciones)>1 else lista_opciones[0]) if lista_opciones else []

    st.divider()
    if st.button("↻ Limpiar caché e ir a inicio"):
        st.cache_data.clear()
        st.rerun()

# ── Aplicar filtros y factor de periodos ──────────────────────────────────────
dff = df_raw.copy()
if sel_vendedor != "Todos":
    dff = dff[dff[COL_VENDEDOR] == sel_vendedor]

rango_label = ""
num_periodos = 1

if sel_rango:
    col_filtro = "_mes_lbl" if modo_fecha == "Mes" else "_sem_lbl"
    if isinstance(sel_rango, tuple) and len(sel_rango) == 2:
        idx1 = lista_opciones.index(sel_rango[0])
        idx2 = lista_opciones.index(sel_rango[1])
        num_periodos = abs(idx2 - idx1) + 1
        
        dff = dff[(dff[col_filtro] >= sel_rango[0]) & (dff[col_filtro] <= sel_rango[1])]
        rango_label = f"entre {sel_rango[0]} y {sel_rango[1]}"
    else:
        num_periodos = 1
        dff = dff[dff[col_filtro] == sel_rango]
        rango_label = f"en {sel_rango}"

df_estado_usar = df_estado_mes if modo_fecha == "Mes" else df_estado_sem

with st.sidebar:
    st.caption(f"**{len(dff):,}** registros filtrados en **{num_periodos}** {'meses' if modo_fecha=='Mes' else 'semana(s)'}.")

# ══════════════════════════════════════════════════════════════════════════════

# ── HEADER PRINCIPAL ──────────────────────────────────────────────────────────
st.markdown("""
<div class="header-title-container-main">
    <div class="header-title">🗺️ Reporte de Visitas Comerciales</div>
</div>
""", unsafe_allow_html=True)
st.markdown(f"<div style='text-align: center; margin-bottom: 2rem;'><span class='header-subtitle'>Análisis de rutas, conversión y cobertura comercial {rango_label.lower()}</span></div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS DE CALCULO
# ══════════════════════════════════════════════════════════════════════════════
def calc_kpis(df_filtro):
    if COL_TIPO not in df_filtro.columns:
        return 0, 0, 0, 0.0
    
    df_vis = df_filtro[(df_filtro[COL_TIPO].str.upper().isin(["PROSPECCIÓN", "PROSPECCION", "MANTENIMIENTO"])) & (df_filtro.get(COL_TIPO_VIS, "").str.upper().isin(["FÍSICA", "FISICA"]))]
    tot_vis = len(df_vis)
    
    df_pros = df_filtro[df_filtro[COL_TIPO].str.upper().isin(["PROSPECCIÓN", "PROSPECCION"])]
    tot_pros = df_pros[COL_CLIENTE].nunique() if not df_pros.empty else 0
    
    n_cierres = 0
    if not df_pros.empty and COL_MOTIVO in df_pros.columns:
        orden_etapa = {e: i for i, e in enumerate(ETAPAS_EMBUDO)}
        valid = df_pros[df_pros[COL_MOTIVO].str.upper().isin([e.upper() for e in ETAPAS_EMBUDO])].copy()
        if not valid.empty:
            valid["_orden"] = valid[COL_MOTIVO].str.upper().map(orden_etapa)
            ultima_etapa = valid.sort_values("_orden").groupby(COL_CLIENTE).last()[[COL_MOTIVO]].reset_index()
            n_cierres = ultima_etapa[ultima_etapa[COL_MOTIVO].str.upper() == "CIERRE"].shape[0]

    t_conv = round((n_cierres / tot_pros * 100), 2) if tot_pros > 0 else 0.0
    return tot_vis, tot_pros, n_cierres, t_conv

def obtener_estado(visitas, meta_df, num_p):
    """Calcula el string del Estado para una zona con base en la tabla umbral x periodos"""
    if meta_df.empty or "Cantidad" not in meta_df.columns or "Estado" not in meta_df.columns:
        return "Sin Datos"
    # Convertir a número por seguridad
    meta_df["Cantidad"] = pd.to_numeric(meta_df["Cantidad"], errors='coerce').fillna(0)
    mdf = meta_df.sort_values(by="Cantidad", ascending=False)
    for _, row in mdf.iterrows():
        umbral = row["Cantidad"] * num_p
        if visitas >= umbral:
            return str(row["Estado"]).strip()
    return str(mdf.iloc[-1]["Estado"]).strip()


def render_region_dashboard(df_region, is_todos=False):
    tot_vis, tot_pros, n_cierres, t_conv = calc_kpis(df_region)

    st.markdown(f"""<div class="kpi-container">
<div class="kpi-card">
<div class="kpi-label">Visitas Totales</div>
<div class="kpi-value">{tot_vis}</div>
<div class="kpi-card-icon">👥</div>
</div>
<div class="kpi-card">
<div class="kpi-label">Prospectos</div>
<div class="kpi-value">{tot_pros}</div>
<div class="kpi-card-icon">👤</div>
</div>
<div class="kpi-card">
<div class="kpi-label">Cierres</div>
<div class="kpi-value">{n_cierres}</div>
<div class="kpi-card-icon">✔️</div>
</div>
<div class="kpi-card">
<div class="kpi-label">Conversión</div>
<div class="kpi-value">{t_conv}%</div>
<div class="kpi-card-icon">%</div>
</div>
</div>""", unsafe_allow_html=True)
    
    if is_todos: return

    if COL_ZONA not in df_region.columns:
        st.warning("No se cuenta con la columna 'Zona' para mostrar métricas.")
        return
        
    df_act_zona = df_region[(df_region[COL_TIPO].str.upper().isin(["PROSPECCIÓN", "PROSPECCION", "MANTENIMIENTO"])) & (df_region.get(COL_TIPO_VIS, "").str.upper().isin(["FÍSICA", "FISICA"]))]
    act_grupos = df_act_zona.groupby(COL_ZONA).size().reset_index(name="Visitas")
    act_grupos = act_grupos.sort_values("Visitas", ascending=False)
    total_visitas_zonas = act_grupos["Visitas"].sum()
    
    html_rutas = ""
    for i, row in act_grupos.iterrows():
        zona_nombre = row[COL_ZONA]
        visor = row["Visitas"]
        pct = round((visor / total_visitas_zonas * 100), 2) if total_visitas_zonas > 0 else 0
        color = PALETA_RUTAS[i % len(PALETA_RUTAS)]
        html_rutas += f"""<div class="ruta-card">
<div class="ruta-header">
<span class="ruta-name">{zona_nombre}</span>
<span class="ruta-value">{visor}</span>
</div>
<div class="ruta-progress-container">
<div class="ruta-progress-bar">
<div class="ruta-progress-fill" style="width: {pct}%; background-color: {color};"></div>
</div>
<span class="ruta-pct">{pct}%</span>
</div>
</div>"""
        
    zonas_unicas = df_region[COL_ZONA].dropna().unique()
    convs_data = []
    for i, z in enumerate(zonas_unicas):
        df_z = df_region[df_region[COL_ZONA] == z]
        v_tz, p_tz, c_tz, conv_tz = calc_kpis(df_z)
        if p_tz > 0:
            convs_data.append({"zona": z, "tasa": conv_tz, "color": PALETA_RUTAS[i % len(PALETA_RUTAS)]})
            
    convs_data = sorted(convs_data, key=lambda x: x["tasa"], reverse=True)
    
    html_convs = ""
    if len(convs_data) > 0:
        for c in convs_data:
            html_convs += f"""<div class="conv-item">
<div class="conv-item-left">
<div class="conv-dot" style="background-color: {c['color']};"></div>
<span class="conv-name">{c['zona']}</span>
</div>
<span class="conv-value">{c['tasa']}%</span>
</div>"""
    else:
        html_convs = "<span style='opacity: 0.6; font-size: 13px;'>Sin datos de prospección suficientes.</span>"
    
    # ── PANEL 1: Actividad y Conversiones ──
    st.markdown(f"""
<div class="dashboard-grid">
<div class="dashboard-panel">
<div class="panel-title">Distribución General</div>
<div class="panel-subtitle">Volumen total de visitas físicas</div>
<div class="rutas-container">
{html_rutas if html_rutas else "<p style='opacity: 0.6; font-size: 13px; margin: 0;'>No hay visitas registradas.</p>"}
</div>
</div>
<div class="dashboard-panel">
<div class="panel-title">Conversiones</div>
<div class="panel-subtitle">% de cierres / prospectos</div>
<div class="conv-list">
{html_convs}
</div>
</div>
</div>
""", unsafe_allow_html=True)

    # ── PANEL 2: NUEVA SECCIÓN Visitas por Ruta/Zona (Bar Chart y Tabla de Estado) ──
    # Componente visual azul de separación
    bg_azul = "#1d4ed8" if not _is_dark else "#1e3a8a"
    st.markdown(f"""
<div style="background-color: {bg_azul}; border-radius: 8px; padding: 15px 20px; margin-top: 10px; display: flex; justify-content: space-between; align-items: center; color: white;">
<div style="text-align: left;">
<div style="font-size: 20px; font-weight: 800; display: flex; align-items: center; gap: 10px;">
<span>🗺️</span> Visitas por Ruta/Zona
</div>
<div style="font-size: 13px; opacity: 0.8; margin-top: 2px;">Distribución de visitas comerciales por frente y evaluación tabular de Estados</div>
</div>
<div style="font-size: 14px; font-weight: 600;">
Total: {total_visitas_zonas} visitas físicas
</div>
</div>
""", unsafe_allow_html=True)
    
    col_chart, col_table = st.columns([2, 1])
    
    # Pre-calculando métricas y tabla
    html_rutas_table = """
<table class="styled-table">
<thead>
<tr>
<th>Ruta</th>
<th>Visitas</th>
<th>%</th>
<th>Estado</th>
</tr>
</thead>
<tbody>
"""
    rutas_label = f"{len(act_grupos)} rutas evaluadas - {rango_label.replace('en ', '').replace('entre ', '').capitalize()}" if not act_grupos.empty else "Sin rutas"
    
    if not act_grupos.empty:
        # Gráfico Barras
        fig_rutas = go.Figure(go.Bar(
            x=act_grupos[COL_ZONA], 
            y=act_grupos["Visitas"],
            marker_color=[PALETA_RUTAS[i % len(PALETA_RUTAS)] for i in range(len(act_grupos))],
            text=act_grupos["Visitas"],
            textposition='outside',
            cliponaxis=False
        ))
        fig_rutas.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=20, b=20, l=10, r=10),
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor="rgba(128,128,128,0.2)"),
            height=340
        )
        
        # Rows Table
        for i, row in act_grupos.iterrows():
            zona = row[COL_ZONA]
            vis = row["Visitas"]
            pct = round((vis / total_visitas_zonas * 100), 2) if total_visitas_zonas > 0 else 0
            est = obtener_estado(vis, df_estado_usar, num_periodos)
            
            est_u = str(est).upper()
            if "ACTIVO" in est_u:
                color_est = "#10b981" 
            elif "REGULAR" in est_u:
                color_est = "#f59e0b" 
            elif "BAJO" in est_u:
                color_est = "#ef4444" 
            else:
                color_est = "var(--text-color)"
                
            html_rutas_table += f"""
<tr>
<td style="font-weight: 700;">{zona}</td>
<td>{vis}</td>
<td>{pct}%</td>
<td style="color: {color_est}; font-weight: 700;">{est}</td>
</tr>"""

    html_rutas_table += "</tbody></table>"

    with col_chart:
        st.markdown(f"""
<div class="dashboard-panel" style="margin-top: 15px; height: 100%;">
<div class="panel-title">Distribución de Visitas por Ruta</div>
<div class="panel-subtitle">{rutas_label}</div>
""", unsafe_allow_html=True)
        if not act_grupos.empty:
            st.plotly_chart(fig_rutas, use_container_width=True, config={'displayModeBar': False})
            st.markdown(f"""<div class="rutas-container" style="margin-top: 10px;">
{html_rutas}
</div>""", unsafe_allow_html=True)
        else:
            st.info("Sin registros para graficar en este periodo.")
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col_table:
        st.markdown(f"""
<div class="dashboard-panel" style="margin-top: 15px; height: 100%;">
<div class="panel-title">Detalle por Ruta</div>
<div class="panel-subtitle">Visitas y participación</div>
{html_rutas_table}
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PESTAÑAS (Por Región)
# ══════════════════════════════════════════════════════════════════════════════
tab_todos, tab_provincia, tab_lima = st.tabs(["🌎 Todos", "🏞️ Provincia", "🏙️ Lima"])

with tab_todos:
    if dff.empty:
        st.info("No hay datos para mostrar.")
    else:
        render_region_dashboard(dff, is_todos=True)

with tab_provincia:
    if COL_REGION in dff.columns:
        df_prov = dff[dff[COL_REGION].str.upper() == "PROVINCIA"]
        if df_prov.empty:
            st.info("No hay datos de Provincia para los filtros actuales.")
        else:
            render_region_dashboard(df_prov, is_todos=False)
    else:
        st.warning("Falta la columna Región en la base para hacer el filtro.")

with tab_lima:
    if COL_REGION in dff.columns:
        df_lim = dff[dff[COL_REGION].str.upper() == "LIMA"]
        if df_lim.empty:
            st.info("No hay datos de Lima para los filtros actuales.")
        else:
            render_region_dashboard(df_lim, is_todos=False)
    else:
        st.warning("Falta la columna Región en la base para hacer el filtro.")


# ══════════════════════════════════════════════════════════════════════════════
# ─── LÓGICA DE EXPORTACIÓN PPTX (SIDEBAR) ──────────────────────────────────────
graficos_exportar = []
with st.sidebar:
    st.divider()
    st.markdown("### 📥 Exportar")
    if st.button("Presentación de pruebas (PPTX)", help="Solo funciona con gráficos Plotly", use_container_width=True):
        st.warning("La exportación está temporalmente deshabilitada.")

# SECCIÓN FINAL — TABLA DE DETALLE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## 📋 Detalle de Visitas")

tab_dt_todos, tab_dt_prov, tab_dt_lima = st.tabs(["🌎 Todos", "🏞️ Provincia", "🏙️ Lima"])

cols_mostrar = [COL_FECHA, COL_VENDEDOR, COL_ZONA, COL_REGION, COL_TIPO, COL_TIPO_VIS, COL_CLIENTE, COL_MOTIVO, COL_RESULTADO]

def prep_detalle(df):
    disponibles = [c for c in cols_mostrar if c in df.columns]
    d = df[disponibles].copy()
    if COL_FECHA in d.columns:
        d[COL_FECHA] = d[COL_FECHA].dt.date
        d = d.sort_values(COL_FECHA, ascending=False)
    return d

with tab_dt_todos:
    st.dataframe(prep_detalle(dff), use_container_width=True, hide_index=True)
with tab_dt_prov:
    if COL_REGION in dff.columns:
        st.dataframe(prep_detalle(dff[dff[COL_REGION].str.upper() == "PROVINCIA"]), use_container_width=True, hide_index=True)
with tab_dt_lima:
    if COL_REGION in dff.columns:
        st.dataframe(prep_detalle(dff[dff[COL_REGION].str.upper() == "LIMA"]), use_container_width=True, hide_index=True)

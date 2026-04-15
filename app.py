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
    /* Estilos para las tarjetas del KPI de la imagen */
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 20px;
        margin-top: 10px;
        margin-bottom: 25px;
        padding: 0 5%;
    }
    .kpi-card {
        flex: 1;
        background-color: var(--secondary-background-color) !important;
        border-radius: 8px;
        padding: 24px 10px;
        border-left: 4px solid #1c64f2; 
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        color: var(--text-color) !important;
    }
    .kpi-card:nth-child(2) { border-left-color: #3b82f6; }
    .kpi-card:nth-child(3) { border-left-color: #60a5fa; }
    .kpi-card:nth-child(4) { border-left-color: #2563eb; }

    .kpi-value {
        font-size: 28px;
        font-weight: 800;
        margin-bottom: 8px;
        color: #1e3a8a;
    }
    .kpi-label {
        font-size: 13px;
        opacity: 0.7;
    }
    
    /* Tab activo */
    [data-baseweb="tab"][aria-selected="true"] {
        border-bottom: 2px solid var(--primary-color) !important;
        color: var(--primary-color) !important;
    }
    
    /* Titulo top */
    .header-title-container {
        text-align: center;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    .header-title {
        color: #1e3a8a;
        font-size: 38px;
        font-weight: 800;
        margin-bottom: 5px;
        display: inline-flex;
        align-items: center;
        gap: 15px;
    }
    .header-subtitle {
        color: #6b7280;
        font-size: 16px;
    }
    
</style>
""", unsafe_allow_html=True)

_theme   = st.get_option("theme.base") or "dark"
_is_dark = (_theme == "dark")

if _is_dark:
    st.markdown("""
    <style>
        .header-title { color: #60a5fa !important; }
        .header-subtitle { color: #9ca3af !important; }
        .kpi-value { color: #bfc1c6; }
    </style>
    """, unsafe_allow_html=True)


# ─── CONSTANTES ────────────────────────────────────────────────────────────────
EXCEL_FILE = "visitas_ventas.xlsx"

COL_FECHA    = "Date"
COL_TIPO     = "Tipo" # PROSPECCIÓN o MANTENIMIENTO
COL_TIPO_CLI = "Giro"
COL_CLIENTE  = "Cliente o Prospecto"
COL_DISTRITO = "Distrito"
COL_MOTIVO   = "Task"
COL_RESULTADO= "Obs"

# Variables nuevas o cruzadas
COL_ZONA     = "Zona"
COL_REGION   = "Región"
COL_DEPARTA  = "Departamento"
COL_PROVINCIA= "Provincia"
COL_TIPO_VIS = "Tipo Visita" # FÍSICA o VIRTUAL
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

COLORES_PRINCIPALES = {
    "azul":    "#4f8ef7", "verde":   "#1fc98e", "naranja": "#f7954f",
    "rojo":    "#f75f4f", "morado":  "#9b74f7", "amarillo":"#f7d14f",
    "cyan":    "#4ff0f7",
}

# ─── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_datos(ruta) -> pd.DataFrame:
    df_log = pd.read_excel(ruta, sheet_name="Log", engine="openpyxl")
    df_users = pd.read_excel(ruta, sheet_name="Users", engine="openpyxl")
    df_zona = pd.read_excel(ruta, sheet_name="Zona", engine="openpyxl")
    
    # Limpiar columnas de posibles espacios
    df_log.columns = df_log.columns.str.strip()
    df_users.columns = df_users.columns.str.strip()
    df_zona.columns = df_zona.columns.str.strip()

    # Cruces (Name extraído de Users; Región extraída de Zona)
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

    # Variables de rango
    df_log["_semana_nro"] = df_log[COL_FECHA].dt.isocalendar().week.astype(str).str.zfill(2)
    df_log["_anio_semana"] = df_log[COL_FECHA].dt.isocalendar().year.astype(str)
    df_log["_sem_lbl"] = df_log["_anio_semana"] + "-S" + df_log["_semana_nro"]
    
    df_log["_mes_lbl"] = df_log[COL_FECHA].dt.strftime("%Y-%m")

    # Normalizar texto (eliminar espacios iniciales/finales extras)
    for col in [COL_VENDEDOR, COL_TIPO, COL_TIPO_CLI, COL_CLIENTE, COL_DISTRITO, COL_MOTIVO, COL_TIPO_VIS, COL_REGION]:
        if col in df_log.columns:
            df_log[col] = df_log[col].astype(str).str.strip()

    return df_log

def get_layout_base() -> dict:
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=60, b=40, l=10, r=10),
        yaxis=dict(automargin=True, title_standoff=10),
    )

LAYOUT_BASE = get_layout_base()

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
        df_raw = cargar_datos(ruta_usar)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    if df_raw.empty:
        st.error("El archivo está vacío o no tiene filas válidas.")
        st.stop()

    st.divider()
    # ── Filtros ───────────────────────────────────────────────────────────────
    st.markdown("### 🔽 Filtros")

    # Filtro Vendedor
    vendedores_list = ["Todos"] + sorted(df_raw[COL_VENDEDOR].dropna().unique().tolist())
    sel_vendedor = st.selectbox("Vendedor", vendedores_list)

    # Filtro Fechas (Mes / Semana)
    st.markdown("#### Periodo de Filtro")
    modo_fecha = st.radio("Agrupar por:", ["Mes", "Semana"], horizontal=True)
    
    df_raw = df_raw.sort_values(COL_FECHA)
    
    if modo_fecha == "Mes":
        meses_disp = sorted(df_raw["_mes_lbl"].dropna().unique().tolist())
        if not meses_disp:
            sel_rango = []
        else:
            sel_rango = st.select_slider(
                "Rango de Meses", 
                options=meses_disp, 
                value=(meses_disp[0], meses_disp[-1]) if len(meses_disp) > 1 else meses_disp[0]
            )
    else:
        sem_disp = sorted(df_raw["_sem_lbl"].dropna().unique().tolist())
        if not sem_disp:
            sel_rango = []
        else:
            sel_rango = st.select_slider(
                "Rango de Semanas", 
                options=sem_disp, 
                value=(sem_disp[0], sem_disp[-1]) if len(sem_disp) > 1 else sem_disp[0]
            )

    st.divider()
    if st.button("↻ Limpiar caché e ir a inicio"):
        st.cache_data.clear()
        st.rerun()

# ── Aplicar filtros ────────────────────────────────────────────────────────────
dff = df_raw.copy()

if sel_vendedor != "Todos":
    dff = dff[dff[COL_VENDEDOR] == sel_vendedor]

rango_label = ""
if modo_fecha == "Mes" and sel_rango:
    if isinstance(sel_rango, tuple) and len(sel_rango) == 2:
        mes_ini, mes_fin = sel_rango[0], sel_rango[1]
        dff = dff[(dff["_mes_lbl"] >= mes_ini) & (dff["_mes_lbl"] <= mes_fin)]
        rango_label = f"entre {mes_ini} y {mes_fin}"
    else:
        dff = dff[dff["_mes_lbl"] == sel_rango]
        rango_label = f"en {sel_rango}"
elif modo_fecha == "Semana" and sel_rango:
    if isinstance(sel_rango, tuple) and len(sel_rango) == 2:
        sem_ini, sem_fin = sel_rango[0], sel_rango[1]
        dff = dff[(dff["_sem_lbl"] >= sem_ini) & (dff["_sem_lbl"] <= sem_fin)]
        rango_label = f"de {sem_ini} a {sem_fin}"
    else:
        dff = dff[dff["_sem_lbl"] == sel_rango]
        rango_label = f"en {sel_rango}"

with st.sidebar:
    st.caption(f"**{len(dff):,}** registros filtrados")

# ══════════════════════════════════════════════════════════════════════════════
graficos_exportar = []

# ── HEADER PRINCIPAL ──────────────────────────────────────────────────────────
st.markdown("""
<div class="header-title-container">
    <div class="header-title">🗺️ Reporte de Visitas Comerciales</div>
</div>
""", unsafe_allow_html=True)
st.markdown(f"<div style='text-align: center; margin-bottom: 2rem;'><span class='header-subtitle'>Análisis de rutas, conversión y cobertura comercial {rango_label.lower()}</span></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PESTAÑAS (Por Región)
# ══════════════════════════════════════════════════════════════════════════════
tab_todos, tab_provincia, tab_lima = st.tabs(["🌎 Todos", "🏞️ Provincia", "🏙️ Lima"])

def render_summary_indicators(df_filtro):
    """Calcula y dibuja las tarjetas indicadoras superiores para la vista."""
    # 1) Visitas totales (PROSPECCIÓN o MANTENIMIENTO, solo FÍSICA)
    if COL_TIPO in df_filtro.columns and COL_TIPO_VIS in df_filtro.columns:
        visitas_fisicas = df_filtro[
            (df_filtro[COL_TIPO].str.upper().isin(["PROSPECCIÓN", "MANTENIMIENTO"])) & 
            (df_filtro[COL_TIPO_VIS].str.upper() == "FÍSICA")
        ]
        total_visitas = len(visitas_fisicas)
    else:
        total_visitas = 0
    
    # 2) Prospectos (únicos) -> Filtrando sobre la bolsa de PROSPECCIÓN
    if COL_TIPO in df_filtro.columns:
        df_pros = df_filtro[df_filtro[COL_TIPO].str.upper() == "PROSPECCIÓN"]
    else:
        df_pros = pd.DataFrame()
        
    total_prospectos = df_pros[COL_CLIENTE].nunique() if not df_pros.empty else 0
    
    # 3) Cierres (donde la etapa más avanzada calculada fue CIERRE)
    n_cierres = 0
    if not df_pros.empty and COL_MOTIVO in df_pros.columns:
        orden_etapa = {e: i for i, e in enumerate(ETAPAS_EMBUDO)}
        df_pros_etapas = df_pros[df_pros[COL_MOTIVO].str.upper().isin([e.upper() for e in ETAPAS_EMBUDO])].copy()
        
        if not df_pros_etapas.empty:
            df_pros_etapas["_orden"] = df_pros_etapas[COL_MOTIVO].str.upper().map(orden_etapa)
            ultima_etapa = (
                df_pros_etapas.sort_values("_orden")
                .groupby(COL_CLIENTE)
                .last()[[COL_MOTIVO]]
                .reset_index()
            )
            n_cierres = ultima_etapa[ultima_etapa[COL_MOTIVO].str.upper() == "CIERRE"].shape[0]

    # 4) Conversión
    tasa_conv = round((n_cierres / total_prospectos * 100), 2) if total_prospectos > 0 else 0.0

    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="kpi-value">{total_visitas}</div>
            <div class="kpi-label">Visitas Totales</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">{total_prospectos}</div>
            <div class="kpi-label">Prospectos</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">{n_cierres}</div>
            <div class="kpi-label">Cierres</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-value">{tasa_conv}%</div>
            <div class="kpi-label">Conversión</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Footer del KPI card
    st.markdown(f"""
    <div style='text-align: center; margin-bottom: 2rem; color: #3b82f6;'>
        <small>🗓️ <b>Periodo:</b> {sel_rango if isinstance(sel_rango, str) else ' a '.join(list(sel_rango))}</small>
    </div>
    """, unsafe_allow_html=True)

with tab_todos:
    if dff.empty:
        st.info("No hay datos para mostrar.")
    else:
        render_summary_indicators(dff)
        st.info("💡  Aquí verás más indicadores en futuras actualizaciones, incluyendo gráficos detallados de los movimientos comerciales del mes.")

with tab_provincia:
    if COL_REGION in dff.columns:
        df_prov = dff[dff[COL_REGION].str.upper() == "PROVINCIA"]
        if df_prov.empty:
            st.info("No hay datos de Provincia para los filtros actuales.")
        else:
            st.info("💡 Indicadores específicos de Provincia estarán disponibles aquí.")
    else:
        st.warning("Falta la columna Región en la data procesada.")

with tab_lima:
    if COL_REGION in dff.columns:
        df_lim = dff[dff[COL_REGION].str.upper() == "LIMA"]
        if df_lim.empty:
            st.info("No hay datos de Lima para los filtros actuales.")
        else:
            st.info("💡 Indicadores específicos de Lima estarán disponibles aquí.")
    else:
        st.warning("Falta la columna Región en la data procesada.")


# ══════════════════════════════════════════════════════════════════════════════
# ─── LÓGICA DE EXPORTACIÓN PPTX (SIDEBAR) ──────────────────────────────────────
with st.sidebar:
    st.divider()
    st.markdown("### 📥 Exportar")
    tema_export = st.radio("Tema de la presentación:", ["Oscuro", "Claro"], horizontal=True)
    export_dark = (tema_export == "Oscuro")
    
    if st.button("Generar Presentación (PPTX)", use_container_width=True):
        if not graficos_exportar:
            st.warning("No hay gráficos para exportar por el momento.")
        else:
            with st.spinner("Generando diapositivas..."):
                try:
                    prs = Presentation()
                    prs.slide_width = Inches(13.33)
                    prs.slide_height = Inches(7.5)
                    blank_slide_layout = prs.slide_layouts[6] 
                    
                    for titulo, fig in graficos_exportar:
                        if titulo == "SECCIÓN":
                            slide = prs.slides.add_slide(blank_slide_layout)
                            if export_dark:
                                slide.background.fill.solid()
                                slide.background.fill.fore_color.rgb = RGBColor(14, 17, 23)
                            
                            txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11.33), Inches(2))
                            p = txBox.text_frame.add_paragraph()
                            p.text = fig
                            p.font.size = Pt(64)
                            p.font.bold = True
                            if export_dark: p.font.color.rgb = RGBColor(250, 250, 250)
                            continue

                        slide = prs.slides.add_slide(blank_slide_layout)
                        if export_dark:
                            slide.background.fill.solid()
                            slide.background.fill.fore_color.rgb = RGBColor(14, 17, 23)
                        
                        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.33), Inches(0.8))
                        p = txBox.text_frame.add_paragraph()
                        p.text = titulo
                        p.font.size = Pt(28)
                        p.font.bold = True
                        if export_dark: p.font.color.rgb = RGBColor(250, 250, 250)
                        
                        img_bytes = io.BytesIO()
                        bg_color_fig = "#0e1117" if export_dark else "white"
                        font_color_fig = "white" if export_dark else "black"
                        
                        fig.update_layout(
                            paper_bgcolor=bg_color_fig, 
                            plot_bgcolor=bg_color_fig,
                            font=dict(color=font_color_fig),
                            margin=dict(t=60, b=130, l=40, r=40)
                        )
                        fig.update_xaxes(automargin=True)
                        fig.update_yaxes(automargin=True)
                        fig.write_image(img_bytes, format="png", width=1200, height=600, scale=2)
                        img_bytes.seek(0)
                        
                        slide.shapes.add_picture(img_bytes, Inches(1), Inches(1.2), width=Inches(11.33))
                    
                    pptx_out = io.BytesIO()
                    prs.save(pptx_out)
                    pptx_out.seek(0)
                    
                    st.session_state["pptx_data"] = pptx_out.getvalue()
                    st.success("¡Presentación lista!")
                    
                except Exception as e:
                    st.error(f"Error al generar la presentación: {e}")

    if "pptx_data" in st.session_state:
        st.download_button(
            label="Descargar Presentación",
            data=st.session_state["pptx_data"],
            file_name=f"Reporte_Export_{sel_vendedor}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary"
        )


# SECCIÓN FINAL — TABLA DE DETALLE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## 📋 Detalle de Visitas")

tab_dt_todos, tab_dt_prov, tab_dt_lima = st.tabs(["🌎 Todos", "🏞️ Provincia", "🏙️ Lima"])

cols_mostrar = [COL_FECHA, COL_VENDEDOR, COL_REGION, COL_TIPO, COL_TIPO_VIS, COL_CLIENTE,
                COL_DISTRITO, COL_MOTIVO, COL_RESULTADO]

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

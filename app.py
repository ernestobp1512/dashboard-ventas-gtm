import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard de Visitas de Ventas",
    page_icon="📊",
    layout="wide",
)

# ─── ESTILOS ───────────────────────────────────────────────────────────────────
# Variables CSS nativas de Streamlit (--text-color, --secondary-background-color, etc.)
# se actualizan automáticamente con el tema seleccionado por el usuario.
st.markdown("""
<style>
    /* Cards de métricas — usa variables CSS de Streamlit */
    [data-testid="metric-container"] {
        background-color: var(--secondary-background-color) !important;
        border: 1px solid rgba(128,128,128,0.25) !important;
        border-radius: 12px;
        padding: 16px 20px;
    }
    [data-testid="stMetricLabel"] { color: var(--text-color) !important; opacity: 0.7; font-size: 0.82rem !important; }
    [data-testid="stMetricValue"] { color: var(--text-color) !important; font-size: 1.8rem !important; font-weight: 700 !important; }
    [data-testid="stMetricDelta"] { font-size: 0.80rem !important; }

    /* Tab activo — usa el color primario del tema */
    [data-baseweb="tab"][aria-selected="true"] {
        border-bottom: 2px solid var(--primary-color) !important;
        color: var(--primary-color) !important;
    }
</style>
""", unsafe_allow_html=True)


# Variables para Plotly (se calculan una vez; solo afectan los gráficos)
_theme   = st.get_option("theme.base") or "dark"
_is_dark = (_theme == "dark")


# ─── CONSTANTES ────────────────────────────────────────────────────────────────
EXCEL_FILE = "visitas_ventas.xlsx"

COL_VENDEDOR = "VENDEDOR"
COL_FECHA    = "FECHA"
COL_TIPO     = "TIPO DE VISITA"
COL_TIPO_CLI = "TIPO DE CLIENTE"
COL_CLIENTE  = "RAZON SOCIAL CLIENTE"
COL_DISTRITO = "DISTRITO"
COL_CONTACTO = "CONTACTO"
COL_TELEFONO = "TELÉFONO"
COL_MOTIVO   = "MOTIVO VISITA"
COL_MOT_NRO  = "MOTIVO NRO"
COL_RESULTADO= "RESULTADO / OBS"

VAL_MANT = "MANTENIMIENTO"
VAL_PROS = "PROSPECCIÓN"

# Etapas del embudo de prospección en orden
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
    "azul":    "#4f8ef7",
    "verde":   "#1fc98e",
    "naranja": "#f7954f",
    "rojo":    "#f75f4f",
    "morado":  "#9b74f7",
    "amarillo":"#f7d14f",
    "cyan":    "#4ff0f7",
}

PALETA_EMBUDO = [
    "#4f8ef7", "#1fc98e", "#f7d14f",
    "#f7954f", "#9b74f7", "#f75f4f", "#8fa0c0",
]

# ─── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_datos(ruta: str) -> pd.DataFrame:
    df = pd.read_excel(ruta, sheet_name=0, engine="openpyxl")

    # Normalizar texto para evitar problemas de espacios / mayúsculas
    for col in [COL_VENDEDOR, COL_TIPO, COL_TIPO_CLI, COL_CLIENTE,
                COL_DISTRITO, COL_MOTIVO]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce")
    df = df.dropna(subset=[COL_FECHA])

    # Columnas derivadas útiles
    df["_semana"]  = df[COL_FECHA].dt.isocalendar().week.astype(str).str.zfill(2)
    df["_anio"]    = df[COL_FECHA].dt.year.astype(str)
    df["_sem_lbl"] = df["_anio"] + "-S" + df["_semana"]
    df["_esMant"]  = df[COL_TIPO].str.upper() == VAL_MANT.upper()

    return df


# ─── HELPERS ───────────────────────────────────────────────────────────────────
def color_plotly(nombre: str) -> str:
    return COLORES_PRINCIPALES.get(nombre, "#4f8ef7")

def get_layout_base() -> dict:
    """Fondos transparentes: Streamlit aplica su tema via theme='streamlit'."""
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=60, b=40, l=10, r=10),
        yaxis=dict(automargin=True, title_standoff=10),
    )

LAYOUT_BASE = get_layout_base()

def bar_scale(end_color: str) -> list:
    """Devuelve escala degradada para gráficos de barra, adaptada al tema."""
    start = "#b8c8e8" if not _is_dark else "#1e2d4a"
    return [start, end_color]



# ─── APP PRINCIPAL ─────────────────────────────────────────────────────────────

# Título
st.markdown("# 📊 Dashboard de Visitas de Ventas")
st.markdown(f"<p class='dash-subtitle'>Seguimiento de actividad de campo — Mantenimiento &amp; Prospección</p>",
            unsafe_allow_html=True)
st.divider()

# ── Selector de archivo ───────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    archivo = st.file_uploader(
        "Cargar Excel de visitas",
        type=["xlsx", "xls"],
        help="Sube tu archivo con las columnas estándar del registro de visitas."
    )

    ruta_default = Path(EXCEL_FILE)
    if archivo is not None:
        ruta_usar = archivo
        fuente_label = f"📂 {archivo.name}"
    elif ruta_default.exists():
        ruta_usar = str(ruta_default)
        fuente_label = f"📄 {EXCEL_FILE} (local)"
    else:
        st.warning("⚠️ Sube un archivo Excel para comenzar.")
        st.stop()

    st.caption(f"Fuente: **{fuente_label}**")
    st.divider()

    # Cargar datos
    try:
        df_raw = cargar_datos(ruta_usar)
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    if df_raw.empty:
        st.error("El archivo está vacío o no tiene filas válidas.")
        st.stop()

    # ── Filtros ───────────────────────────────────────────────────────────────
    st.markdown("### 🔽 Filtros")

    vendedores_list = ["Todos"] + sorted(df_raw[COL_VENDEDOR].dropna().unique().tolist())
    sel_vendedor = st.selectbox("Vendedor", vendedores_list)

    fecha_min = df_raw[COL_FECHA].min().date()
    fecha_max = df_raw[COL_FECHA].max().date()
    sel_rango = st.date_input(
        "Rango de fechas",
        value=(fecha_min, fecha_max),
        min_value=fecha_min,
        max_value=fecha_max,
        format="DD/MM/YYYY",
    )

    st.divider()
    if st.button("↻ Limpiar caché y recargar"):
        st.cache_data.clear()
        st.rerun()

# ── Aplicar filtros ────────────────────────────────────────────────────────────
dff = df_raw.copy()
if sel_vendedor != "Todos":
    dff = dff[dff[COL_VENDEDOR] == sel_vendedor]

# Rango de fechas (el widget puede devolver 1 o 2 fechas)
if isinstance(sel_rango, (list, tuple)) and len(sel_rango) == 2:
    fecha_ini = pd.Timestamp(sel_rango[0])
    fecha_fin = pd.Timestamp(sel_rango[1])
    dff = dff[(dff[COL_FECHA] >= fecha_ini) & (dff[COL_FECHA] <= fecha_fin)]

with st.sidebar:
    st.caption(f"**{len(dff):,}** registros filtrados")

if dff.empty:
    st.warning("No hay datos para los filtros seleccionados.")
    st.stop()

# ── Separar por tipo ───────────────────────────────────────────────────────────
df_mant = dff[dff["_esMant"]].copy()
df_pros = dff[~dff["_esMant"]].copy()

# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
# SECCIONES PRINCIPALES EN PESTAÑAS
# ══════════════════════════════════════════════════════════════════════════════
tab_pros, tab_mant = st.tabs(["🎯 Prospección", "🔧 Mantenimiento"])

with tab_pros:
    # SECCIÓN 1 — PROSPECCIÓN
    # ══════════════════════════════════════════════════════════════════════════════
    st.markdown("## 🎯 Prospección")

    if df_pros.empty:
        st.info("No hay datos de Prospección para los filtros actuales.")
    else:

        # ── Cálculos base ─────────────────────────────────────────────────────────
        # Total de prospectos únicos visitados en el rango
        total_prospectos = df_pros[COL_CLIENTE].nunique()

        # Etapa más avanzada por prospecto (según el orden del embudo)
        orden_etapa = {e: i for i, e in enumerate(ETAPAS_EMBUDO)}
        df_pros_etapas = df_pros[df_pros[COL_MOTIVO].isin(ETAPAS_EMBUDO)].copy()
        df_pros_etapas["_orden"] = df_pros_etapas[COL_MOTIVO].map(orden_etapa)
        ultima_etapa = (
            df_pros_etapas.sort_values("_orden")
            .groupby(COL_CLIENTE)
            .last()[[COL_MOTIVO]]
            .reset_index()
            .rename(columns={COL_MOTIVO: "Etapa"})
        )

        # Conteo del embudo (por última etapa de cada prospecto)
        embudo = (
            ultima_etapa.groupby("Etapa")
            .size()
            .reindex(ETAPAS_EMBUDO, fill_value=0)
            .reset_index(name="Prospectos")
        )
        embudo["% del Total"] = (embudo["Prospectos"] / total_prospectos * 100).round(1)

        # Tasa de conversión: prospectos en CIERRE / total prospectos únicos
        n_cierre = ultima_etapa[ultima_etapa["Etapa"] == "CIERRE"].shape[0]
        tasa_conv = round(n_cierre / total_prospectos * 100, 1) if total_prospectos else 0.0

        # ── KPI Cards ─────────────────────────────────────────────────────────────
        kp1, kp2, kp3 = st.columns(3)
        kp1.metric("Prospectos Visitados", f"{total_prospectos:,}",
                   help="Total de prospectos únicos visitados en el rango de fechas seleccionado.")
        kp2.metric("Cierres", f"{n_cierre:,}",
                   help="Cantidad de prospectos cuya etapa más avanzada registrada es CIERRE.")
        kp3.metric("Tasa de Conversión", f"{tasa_conv}%",
                   help="Prospectos que alcanzaron la etapa CIERRE / total prospectos únicos visitados.")

        # ── Ranking por Vendedor (solo cuando se muestran todos) ──────────────────
        if sel_vendedor == "Todos":
            st.markdown("### 🏆 Ranking por Vendedor")
            st.caption("Prospectos únicos visitados y cierres por vendedor en el rango de fechas seleccionado.")

            # Prospectos únicos por vendedor
            rank_pros = (
                df_pros.groupby(COL_VENDEDOR)[COL_CLIENTE]
                .nunique()
                .reset_index(name="Prospectos Únicos")
            )

            # Cierres por vendedor (etapa más avanzada = CIERRE)
            if not df_pros_etapas.empty:
                ultima_etapa_vend = (
                    df_pros_etapas.sort_values("_orden")
                    .groupby([COL_VENDEDOR, COL_CLIENTE])
                    .last()["_orden"]  # solo necesitamos saber si llegó a CIERRE
                    .reset_index()
                )
                idx_cierre = orden_etapa.get("CIERRE", 999)
                cierres_vend = (
                    ultima_etapa_vend[ultima_etapa_vend["_orden"] == idx_cierre]
                    .groupby(COL_VENDEDOR)
                    .size()
                    .reset_index(name="Cierres")
                )
            else:
                cierres_vend = pd.DataFrame(columns=[COL_VENDEDOR, "Cierres"])

            ranking_df = (
                rank_pros
                .merge(cierres_vend, on=COL_VENDEDOR, how="left")
                .fillna({"Cierres": 0})
                .sort_values("Prospectos Únicos", ascending=False)
            )
            ranking_df["Cierres"] = ranking_df["Cierres"].astype(int)

            col_rank1, col_rank2 = st.columns([3, 2])
            with col_rank1:
                fig_rank = go.Figure()
                fig_rank.add_trace(go.Bar(
                    x=ranking_df[COL_VENDEDOR],
                    y=ranking_df["Prospectos Únicos"],
                    name="Prospectos Únicos",
                    marker_color=COLORES_PRINCIPALES["azul"],
                    text=ranking_df["Prospectos Únicos"],
                    textposition="outside", cliponaxis=False,
                ))
                fig_rank.add_trace(go.Bar(
                    x=ranking_df[COL_VENDEDOR],
                    y=ranking_df["Cierres"],
                    name="Cierres",
                    marker_color=COLORES_PRINCIPALES["verde"],
                    text=ranking_df["Cierres"],
                    textposition="outside", cliponaxis=False,
                ))
                fig_rank.update_layout(
                    **LAYOUT_BASE,
                    barmode="group",
                    height=420,
                    xaxis_title="",
                    yaxis_title="",
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
                )
                st.plotly_chart(fig_rank, use_container_width=True, theme="streamlit")

            with col_rank2:
                tabla_rank = ranking_df[[COL_VENDEDOR, "Prospectos Únicos", "Cierres"]].copy()
                tabla_rank["Tasa Cierre"] = (
                    tabla_rank["Cierres"] / tabla_rank["Prospectos Únicos"] * 100
                ).round(1).astype(str) + "%"
                st.dataframe(tabla_rank, use_container_width=True, hide_index=True)

        st.divider()


        # ── Indicador 2: Embudo de Prospección ────────────────────────────────────
        st.markdown("### 🔽 Embudo de Prospección")
        st.caption("Cada prospecto se cuenta una sola vez, en su etapa más avanzada registrada dentro del rango de fechas.")

        col_emb1, col_emb2 = st.columns([1, 1])

        with col_emb1:
            fig_funnel = go.Figure(go.Funnel(
                y=embudo["Etapa"],
                x=embudo["Prospectos"],
                text=[f"{p}  ({pct}%)" for p, pct in zip(embudo["Prospectos"], embudo["% del Total"])],
                textinfo="text",
                marker=dict(color=PALETA_EMBUDO[:len(embudo)]),
                connector=dict(line=dict(color="#2e3a55", width=2)),
            ))
            fig_funnel.update_layout(**LAYOUT_BASE, height=380)
            st.plotly_chart(fig_funnel, use_container_width=True, theme="streamlit")

        with col_emb2:
            # Tabla del embudo
            st.dataframe(
                embudo[embudo["Prospectos"] > 0].style.format({"% del Total": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # ── Indicador 4: Actividad por Ciudad ─────────────────────────────────────
        st.markdown("### 🏙️ Actividad por Ciudad")

        col_ciu1, col_ciu2 = st.columns([3, 2])

        ciudad_counts = (
            df_pros.groupby(COL_DISTRITO)
            .size()
            .reset_index(name="Visitas")
            .sort_values("Visitas", ascending=False)
        )
        total_vis_ciu = len(df_pros)
        ciudad_counts["% del Total"] = (ciudad_counts["Visitas"] / total_vis_ciu * 100).round(1)

        with col_ciu1:
            fig_ciudad = px.bar(
                ciudad_counts,
                x=COL_DISTRITO, y="Visitas",
                color="Visitas",
                color_continuous_scale=bar_scale("#4f8ef7"),
                text="Visitas",
            )
            fig_ciudad.update_traces(textposition="outside", cliponaxis=False)
            fig_ciudad.update_layout(**LAYOUT_BASE, height=320,
                                      xaxis_title="", yaxis_title="Nº Visitas",
                                      coloraxis_showscale=False)
            st.plotly_chart(fig_ciudad, use_container_width=True, theme="streamlit")

        with col_ciu2:
            st.dataframe(
                ciudad_counts.style.format({"% del Total": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # ── Indicador 5: Actividad por Giro ───────────────────────────────────────
        st.markdown("### 🏷️ Actividad por Giro")

        giro_counts = (
            df_pros.groupby(COL_TIPO_CLI)
            .size()
            .reset_index(name="Visitas")
            .sort_values("Visitas", ascending=False)
        )
        total_vis_pros_n = len(df_pros)
        giro_counts["% del Total"] = (giro_counts["Visitas"] / total_vis_pros_n * 100).round(1)

        col_giro1, col_giro2 = st.columns([3, 2])

        with col_giro1:
            fig_giro = px.pie(
                giro_counts,
                names=COL_TIPO_CLI, values="Visitas",
                hole=0.45,
                color_discrete_sequence=list(COLORES_PRINCIPALES.values()),
            )
            fig_giro.update_traces(textinfo="percent+label", textposition="outside")
            fig_giro.update_layout(**LAYOUT_BASE, height=320, showlegend=False)
            st.plotly_chart(fig_giro, use_container_width=True, theme="streamlit")

        with col_giro2:
            st.dataframe(
                giro_counts.style.format({"% del Total": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # ── Indicador 6: Frecuencia de Visitas ────────────────────────────────────
        st.markdown("### 📊 Frecuencia de Visitas por Prospecto")

        freq = (
            df_pros.groupby(COL_CLIENTE)
            .size()
            .reset_index(name="Nº Visitas")
            .sort_values("Nº Visitas", ascending=False)
        )

        col_frq1, col_frq2 = st.columns([3, 2])

        with col_frq1:
            fig_freq = px.bar(
                freq,
                x=COL_CLIENTE, y="Nº Visitas",
                color="Nº Visitas",
                color_continuous_scale=bar_scale("#1fc98e"),
                text="Nº Visitas",
            )
            fig_freq.update_traces(textposition="outside", cliponaxis=False)
            fig_freq.update_layout(**LAYOUT_BASE, height=340,
                                    xaxis_title="", yaxis_title="Nº Visitas",
                                    xaxis_tickangle=-35, coloraxis_showscale=False)
            st.plotly_chart(fig_freq, use_container_width=True, theme="streamlit")

        with col_frq2:
            st.dataframe(freq, use_container_width=True, hide_index=True)

        st.divider()

        # ── Indicador 7: Visitas por Semana ───────────────────────────────────────
        st.markdown("### 📅 Visitas por Semana")
        st.caption("Semanas en orden relativo al rango de fechas. Los extremos se recortan según las fechas del filtro.")

        # Construir etiquetas de semana con rango de fechas (lunes–domingo, recortado por el filtro)
        def etiqueta_semana(sem_lbl: str, fecha_ini_filtro, fecha_fin_filtro) -> str:
            """Devuelve 'Semana N (DD/MM/YY-DD/MM/YY)' con inicio=lunes, fin=domingo,
               recortado por las fechas del filtro activo."""
            # sem_lbl tiene formato YYYY-SWW
            anio, sw = sem_lbl.split("-S")
            # Primer día (lunes) de la ISO-week
            lunes = pd.Timestamp.fromisocalendar(int(anio), int(sw), 1)
            domingo = lunes + pd.Timedelta(days=6)
            # Recortar por el rango del filtro
            inicio = max(lunes, pd.Timestamp(fecha_ini_filtro))
            fin    = min(domingo, pd.Timestamp(fecha_fin_filtro))
            return f"{inicio.strftime('%d/%m/%y')}-{fin.strftime('%d/%m/%y')}"

        # Obtener fechas del filtro activo (o rango global si no se aplicó)
        if isinstance(sel_rango, (list, tuple)) and len(sel_rango) == 2:
            fi_filtro, ff_filtro = sel_rango[0], sel_rango[1]
        else:
            fi_filtro = df_pros[COL_FECHA].min().date()
            ff_filtro = df_pros[COL_FECHA].max().date()

        semanas_ordenadas = sorted(df_pros["_sem_lbl"].unique())
        sem_mapa = {
            s: f"Semana {i+1} ({etiqueta_semana(s, fi_filtro, ff_filtro)})"
            for i, s in enumerate(semanas_ordenadas)
        }

        sem_pros = (
            df_pros.groupby("_sem_lbl")
            .size()
            .reset_index(name="Visitas")
            .sort_values("_sem_lbl")
        )
        sem_pros["Semana"] = sem_pros["_sem_lbl"].map(sem_mapa)

        fig_sem = px.bar(
            sem_pros,
            x="Semana", y="Visitas",
            color="Visitas",
            color_continuous_scale=bar_scale("#9b74f7"),
            text="Visitas",
        )
        fig_sem.update_traces(textposition="outside", cliponaxis=False)
        fig_sem.update_layout(**LAYOUT_BASE, height=340,
                               xaxis_title="", yaxis_title="Nº Visitas",
                               coloraxis_showscale=False)
        st.plotly_chart(fig_sem, use_container_width=True, theme="streamlit")

    st.divider()

    # ══════════════════════════════════════════════════════════════════════════════

with tab_mant:
    # SECCIÓN 2 — MANTENIMIENTO
    # ══════════════════════════════════════════════════════════════════════════════
    st.markdown("## 🔧 Mantenimiento")

    if df_mant.empty:
        st.info("No hay datos de Mantenimiento para los filtros actuales.")
    else:

        # ── Cálculos base ────────────────────────────────────────────────────
        total_vis_mant   = len(df_mant)
        n_pedido         = (df_mant[COL_MOTIVO] == "TOMAR PEDIDO").sum()
        tasa_conv_mant   = round(n_pedido / total_vis_mant * 100, 1) if total_vis_mant else 0.0

        # ── KPI Cards ─────────────────────────────────────────────────────
        km1, km2, km3 = st.columns(3)
        km1.metric("Total Visitas a Clientes", f"{total_vis_mant:,}",
                   help="Total de visitas (incluyendo múltiples visitas al mismo cliente) en el rango de fechas.")
        km2.metric("Visitas con Pedido", f"{n_pedido:,}",
                   help="Cantidad de visitas registradas con motivo TOMAR PEDIDO.")
        km3.metric("Tasa de Conversión", f"{tasa_conv_mant}%",
                   help="Visitas con motivo TOMAR PEDIDO / total visitas de mantenimiento.")

        # ── Ranking por Vendedor (solo cuando se muestran todos) ──────────────────
        if sel_vendedor == "Todos":
            st.markdown("### 🏆 Ranking por Vendedor")
            st.caption("Total de visitas y visitas con pedido por vendedor en el rango de fechas seleccionado.")

            # Total visitas por vendedor
            rank_mant = (
                df_mant.groupby(COL_VENDEDOR)
                .size()
                .reset_index(name="Total Visitas")
            )

            # Visitas con pedido por vendedor
            pedidos_vend = (
                df_mant[df_mant[COL_MOTIVO] == "TOMAR PEDIDO"]
                .groupby(COL_VENDEDOR)
                .size()
                .reset_index(name="Con Pedido")
            )

            ranking_m_df = (
                rank_mant
                .merge(pedidos_vend, on=COL_VENDEDOR, how="left")
                .fillna({"Con Pedido": 0})
                .sort_values("Total Visitas", ascending=False)
            )
            ranking_m_df["Con Pedido"] = ranking_m_df["Con Pedido"].astype(int)

            col_mrank1, col_mrank2 = st.columns([3, 2])
            with col_mrank1:
                fig_mrank = go.Figure()
                fig_mrank.add_trace(go.Bar(
                    x=ranking_m_df[COL_VENDEDOR],
                    y=ranking_m_df["Total Visitas"],
                    name="Total Visitas",
                    marker_color=COLORES_PRINCIPALES["azul"],
                    text=ranking_m_df["Total Visitas"],
                    textposition="outside",
                    cliponaxis=False,
                ))
                fig_mrank.add_trace(go.Bar(
                    x=ranking_m_df[COL_VENDEDOR],
                    y=ranking_m_df["Con Pedido"],
                    name="Con Pedido",
                    marker_color=COLORES_PRINCIPALES["amarillo"],
                    text=ranking_m_df["Con Pedido"],
                    textposition="outside",
                    cliponaxis=False,
                ))
                fig_mrank.update_layout(
                    **LAYOUT_BASE,
                    barmode="group",
                    height=420,
                    xaxis_title="",
                    yaxis_title="",
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
                )
                st.plotly_chart(fig_mrank, use_container_width=True, theme="streamlit")

            with col_mrank2:
                tabla_mrank = ranking_m_df[[COL_VENDEDOR, "Total Visitas", "Con Pedido"]].copy()
                tabla_mrank["Tasa Conv."] = (
                    tabla_mrank["Con Pedido"] / tabla_mrank["Total Visitas"] * 100
                ).round(1).astype(str) + "%"
                st.dataframe(tabla_mrank, use_container_width=True, hide_index=True)

        st.divider()

        # ── Indicador 2: Mix de Mantenimiento ──────────────────────────────────
        st.markdown("### 🧩 Mix de Mantenimiento")
        st.caption("Distribución de visitas según motivo. Se contabilizan todas las visitas del rango.")

        MOT_MANT_ORDEN = ["TOMAR PEDIDO", "CAPACITACIÓN", "LANZAMIENTO", "COBRANZA", "RECLAMO", "OTROS"]
        mix_counts = (
            df_mant.groupby(COL_MOTIVO)
            .size()
            .reindex(MOT_MANT_ORDEN, fill_value=0)
            .reset_index(name="Visitas")
        )
        mix_counts["% del Total"] = (mix_counts["Visitas"] / total_vis_mant * 100).round(1)

        col_mix1, col_mix2 = st.columns([1, 1])

        with col_mix1:
            fig_mix = px.pie(
                mix_counts,
                names=COL_MOTIVO, values="Visitas",
                hole=0.45,
                color_discrete_sequence=list(COLORES_PRINCIPALES.values()),
            )
            fig_mix.update_traces(textinfo="percent+label", textposition="outside")
            fig_mix.update_layout(**LAYOUT_BASE, height=360, showlegend=False)
            st.plotly_chart(fig_mix, use_container_width=True, theme="streamlit")

        with col_mix2:
            st.dataframe(
                mix_counts.style.format({"% del Total": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # ── Indicador 4: Actividad por Ciudad ─────────────────────────────────
        st.markdown("### 🏙️ Actividad por Ciudad")

        col_mciu1, col_mciu2 = st.columns([3, 2])

        mciu_counts = (
            df_mant.groupby(COL_DISTRITO)
            .size()
            .reset_index(name="Visitas")
            .sort_values("Visitas", ascending=False)
        )
        mciu_counts["% del Total"] = (mciu_counts["Visitas"] / total_vis_mant * 100).round(1)

        with col_mciu1:
            fig_mciu = px.bar(
                mciu_counts,
                x=COL_DISTRITO, y="Visitas",
                color="Visitas",
                color_continuous_scale=bar_scale("#4f8ef7"),
                text="Visitas",
            )
            fig_mciu.update_traces(textposition="outside", cliponaxis=False)
            fig_mciu.update_layout(**LAYOUT_BASE, height=320,
                                    xaxis_title="", yaxis_title="Nº Visitas",
                                    coloraxis_showscale=False)
            st.plotly_chart(fig_mciu, use_container_width=True, theme="streamlit")

        with col_mciu2:
            st.dataframe(
                mciu_counts.style.format({"% del Total": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # ── Indicador 5: Actividad por Giro ──────────────────────────────────
        st.markdown("### 🏷️ Actividad por Giro")

        mgiro_counts = (
            df_mant.groupby(COL_TIPO_CLI)
            .size()
            .reset_index(name="Visitas")
            .sort_values("Visitas", ascending=False)
        )
        mgiro_counts["% del Total"] = (mgiro_counts["Visitas"] / total_vis_mant * 100).round(1)

        col_mgiro1, col_mgiro2 = st.columns([3, 2])

        with col_mgiro1:
            fig_mgiro = px.pie(
                mgiro_counts,
                names=COL_TIPO_CLI, values="Visitas",
                hole=0.45,
                color_discrete_sequence=list(COLORES_PRINCIPALES.values()),
            )
            fig_mgiro.update_traces(textinfo="percent+label", textposition="outside")
            fig_mgiro.update_layout(**LAYOUT_BASE, height=320, showlegend=False)
            st.plotly_chart(fig_mgiro, use_container_width=True, theme="streamlit")

        with col_mgiro2:
            st.dataframe(
                mgiro_counts.style.format({"% del Total": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # ── Indicador 6: Frecuencia de Visitas ──────────────────────────────
        st.markdown("### 📊 Frecuencia de Visitas por Cliente")

        mfreq = (
            df_mant.groupby(COL_CLIENTE)
            .size()
            .reset_index(name="Nº Visitas")
            .sort_values("Nº Visitas", ascending=False)
        )

        col_mfrq1, col_mfrq2 = st.columns([3, 2])

        with col_mfrq1:
            fig_mfreq = px.bar(
                mfreq,
                x=COL_CLIENTE, y="Nº Visitas",
                color="Nº Visitas",
                color_continuous_scale=bar_scale("#f7954f"),
                text="Nº Visitas",
            )
            fig_mfreq.update_traces(textposition="outside", cliponaxis=False)
            fig_mfreq.update_layout(**LAYOUT_BASE, height=340,
                                     xaxis_title="", yaxis_title="Nº Visitas",
                                     xaxis_tickangle=-35, coloraxis_showscale=False)
            st.plotly_chart(fig_mfreq, use_container_width=True, theme="streamlit")

        with col_mfrq2:
            st.dataframe(mfreq, use_container_width=True, hide_index=True)

        st.divider()

        # ── Indicador 7: Visitas por Semana ────────────────────────────────
        st.markdown("### 📅 Visitas por Semana")
        st.caption("Semanas en orden relativo al rango de fechas. Los extremos se recortan según las fechas del filtro.")

        if isinstance(sel_rango, (list, tuple)) and len(sel_rango) == 2:
            fi_m, ff_m = sel_rango[0], sel_rango[1]
        else:
            fi_m = df_mant[COL_FECHA].min().date()
            ff_m = df_mant[COL_FECHA].max().date()

        semanas_m_ord = sorted(df_mant["_sem_lbl"].unique())
        sem_mapa_m = {
            s: f"Semana {i+1} ({etiqueta_semana(s, fi_m, ff_m)})"
            for i, s in enumerate(semanas_m_ord)
        }

        sem_mant_df = (
            df_mant.groupby("_sem_lbl")
            .size()
            .reset_index(name="Visitas")
            .sort_values("_sem_lbl")
        )
        sem_mant_df["Semana"] = sem_mant_df["_sem_lbl"].map(sem_mapa_m)

        fig_sem_m = px.bar(
            sem_mant_df,
            x="Semana", y="Visitas",
            color="Visitas",
            color_continuous_scale=bar_scale("#f7d14f"),
            text="Visitas",
        )
        fig_sem_m.update_traces(textposition="outside", cliponaxis=False)
        fig_sem_m.update_layout(**LAYOUT_BASE, height=340,
                                 xaxis_title="", yaxis_title="Nº Visitas",
                                 coloraxis_showscale=False)
        st.plotly_chart(fig_sem_m, use_container_width=True, theme="streamlit")

    st.divider()

    # ══════════════════════════════════════════════════════════════════════════════


# SECCIÓN 5 — TABLA DE DETALLE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("## 📋 Detalle de Visitas")

tab1, tab2, tab3 = st.tabs(["📁 Todos los registros", "🔧 Mantenimiento", "🎯 Prospección"])

cols_mostrar = [COL_VENDEDOR, COL_FECHA, COL_TIPO, COL_CLIENTE,
                COL_DISTRITO, COL_MOTIVO, COL_RESULTADO]

def prep_detalle(df):
    """Prepara el df para mostrar: convierte FECHA a solo fecha."""
    d = df[cols_mostrar].copy()
    d[COL_FECHA] = d[COL_FECHA].dt.date
    return d.sort_values(COL_FECHA, ascending=False)

with tab1:
    st.dataframe(prep_detalle(dff), use_container_width=True, hide_index=True)
with tab2:
    st.dataframe(prep_detalle(df_mant), use_container_width=True, hide_index=True)
with tab3:
    st.dataframe(prep_detalle(df_pros), use_container_width=True, hide_index=True)

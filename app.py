"""
app.py — Portal de acceso Go To Market SAC
Pantalla de login + bienvenida con acceso a los dashboards por región.
"""
import streamlit as st
import sys, os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import auth

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Go To Market SAC — Portal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS GLOBAL (bienvenida) ──────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

/* Fondo degradado suave */
.main {
    background: linear-gradient(145deg, #F0F4FF 0%, #F8FAFC 50%, #FFF5F5 100%) !important;
    min-height: 100vh;
}

/* ── Header ── */
.gtm-header {
    background: white;
    border-radius: 20px;
    padding: 30px 40px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 4px 24px rgba(0,0,0,0.06);
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.gtm-brand-title {
    font-size: 46px;
    font-weight: 900;
    color: #0F172A;
    letter-spacing: -2.5px;
    line-height: 1;
}
.gtm-brand-sub {
    font-size: 20px;
    font-weight: 700;
    color: #0F172A;
    margin-top: 4px;
}

/* ── Barra de estado de usuario ── */
.user-bar {
    background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
    border-radius: 14px;
    padding: 16px 24px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    color: white;
}

/* ── Tarjetas de dashboard ── */
.dash-card-wrap {
    animation: fadeUp 0.4s ease both;
}
@keyframes fadeUp {
    from { opacity: 0; transform: translateY(16px); }
    to   { opacity: 1; transform: translateY(0); }
}

/* Botones de las tarjetas */
div[data-testid="stButton"] > button[kind="primary"] {
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    padding: 10px !important;
    border: none !important;
    transition: transform 0.15s, box-shadow 0.15s !important;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
}

/* Logout sidebar */
div[data-testid="stButton"] > button[kind="secondary"] {
    border-radius: 8px !important;
    font-weight: 600 !important;
}

/* Quitar doble borde en cualquier input de esta página */
.stTextInput [data-baseweb="input"],
.stTextInput [data-baseweb="base-input"] {
    box-shadow: none !important;
    border: none !important;
}
</style>
""", unsafe_allow_html=True)

# ─── LOGIN GATE ───────────────────────────────────────────────────────────────
if not st.session_state.get("auth_ok"):
    if not auth.show_login():
        st.stop()

# ─── SESIÓN VÁLIDA — PANTALLA DE BIENVENIDA ──────────────────────────────────
user      = auth.get_current_user()
user_info = auth.USERS[user]
pages     = auth.get_user_pages()

# ── Sidebar: usuario + logout ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg,#FFF1F1,#FFE4E4);
        border: 1px solid #FECACA;
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 10px;
    ">
        <div style="font-size:10px;color:#DC2626;font-weight:700;
                    text-transform:uppercase;letter-spacing:1px;">SESIÓN ACTIVA</div>
        <div style="font-size:15px;font-weight:700;color:#1E293B;margin-top:4px;">
            👤 {user_info['name']}
        </div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("🚪 Cerrar sesión", use_container_width=True, key="_welcome_logout"):
        auth.logout()

# ── Header principal ──────────────────────────────────────────────────────────
st.markdown(f"""
<div class="gtm-header">
    <div style="display:flex; align-items:center; gap:24px;">
        <div style="display:flex; align-items:flex-end; gap:7px;">
            <div style="width:13px;height:22px;background:#DC2626;border-radius:3px;opacity:0.6;"></div>
            <div style="width:13px;height:38px;background:#DC2626;border-radius:3px;opacity:0.8;"></div>
            <div style="width:13px;height:54px;background:#DC2626;border-radius:3px;"></div>
        </div>
        <div>
            <div class="gtm-brand-title">Go To Market SAC</div>
            <div class="gtm-brand-sub">
                go<span style="color:#DC2626;">to</span>market
            </div>
        </div>
    </div>
    <div style="text-align:right;">
        <div style="font-size:11px;color:#94A3B8;font-weight:600;
                    letter-spacing:1.5px;text-transform:uppercase;">
            Portal Comercial
        </div>
        <div style="font-size:22px;font-weight:800;color:#1E293B;margin-top:2px;">
            Dashboard 2026
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Bienvenida personalizada ──────────────────────────────────────────────────
num_pages  = len(pages)
label_dash = "dashboard" if num_pages == 1 else "dashboards"

st.markdown(f"""
<div style="
    background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
    border-radius: 16px;
    padding: 22px 28px;
    margin-bottom: 30px;
    display: flex;
    align-items: center;
    justify-content: space-between;
">
    <div>
        <div style="font-size:22px;font-weight:800;color:white;letter-spacing:-0.5px;">
            👋 Bienvenido, {user_info['name']}
        </div>
        <div style="font-size:13px;color:#94A3B8;margin-top:5px;">
            Tienes acceso a <b style="color:#FCA5A5;">{num_pages} {label_dash}</b>.
            Selecciona el que deseas abrir.
        </div>
    </div>
    <div style="
        background: rgba(220,38,38,0.15);
        border: 1px solid rgba(220,38,38,0.3);
        border-radius: 50px;
        padding: 8px 18px;
        font-size: 13px;
        font-weight: 700;
        color: #FCA5A5;
    ">
        📅 2026
    </div>
</div>
""", unsafe_allow_html=True)

# ── Tarjetas de dashboards ────────────────────────────────────────────────────
CARD_CONFIG = {
    "lima": {
        "icon":       "🏙️",
        "title":      "Dashboard Lima",
        "subtitle":   "REGIÓN LIMA",
        "desc":       "Seguimiento de visitas, prospección y mantenimiento de la región Lima.",
        "color":      "#DC2626",
        "bg":         "#FFF1F1",
        "btn_color":  "#DC2626",
        "page":       "pages/1_Lima.py",
        "delay":      "0.0s",
    },
    "norte": {
        "icon":       "🌄",
        "title":      "Dashboard Norte",
        "subtitle":   "REGIÓN NORTE",
        "desc":       "Gestión comercial y embudo de ventas para la región Norte del país.",
        "color":      "#2563EB",
        "bg":         "#EFF6FF",
        "btn_color":  "#2563EB",
        "page":       "pages/2_Norte.py",
        "delay":      "0.1s",
    },
    "provincia": {
        "icon":       "🗺️",
        "title":      "Dashboard Provincia",
        "subtitle":   "REGIÓN PROVINCIA",
        "desc":       "Control de indicadores y alertas de cumplimiento para Provincia.",
        "color":      "#16A34A",
        "bg":         "#F0FDF4",
        "btn_color":  "#16A34A",
        "page":       "pages/3_Provincia.py",
        "delay":      "0.2s",
    },
}

if pages:
    cols = st.columns(len(pages), gap="large")
    for i, page_key in enumerate(pages):
        cfg = CARD_CONFIG[page_key]
        with cols[i]:
            # Animación escalonada por tarjeta
            st.markdown(f"""
            <style>
            .card-{page_key} {{
                animation: fadeUp 0.45s ease {cfg['delay']} both;
            }}
            @keyframes fadeUp {{
                from {{ opacity:0; transform:translateY(20px); }}
                to   {{ opacity:1; transform:translateY(0); }}
            }}
            </style>
            <div class="card-{page_key}"></div>
            """, unsafe_allow_html=True)

            with st.container(border=True):
                st.markdown(f"""
                <div style="padding: 6px 0 16px 0;">
                    <!-- Chip de región -->
                    <div style="
                        display: inline-block;
                        background: {cfg['bg']};
                        color: {cfg['color']};
                        font-size: 10px;
                        font-weight: 700;
                        letter-spacing: 1.2px;
                        text-transform: uppercase;
                        border-radius: 20px;
                        padding: 4px 12px;
                        margin-bottom: 16px;
                    ">{cfg['subtitle']}</div>

                    <!-- Icono -->
                    <div style="
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        width: 72px; height: 72px;
                        background: {cfg['bg']};
                        border-radius: 18px;
                        font-size: 36px;
                        margin-bottom: 16px;
                        border: 2px solid {cfg['color']}22;
                    ">{cfg['icon']}</div>

                    <!-- Título -->
                    <div style="
                        font-size: 20px;
                        font-weight: 800;
                        color: {cfg['color']};
                        margin-bottom: 8px;
                        letter-spacing: -0.3px;
                    ">{cfg['title']}</div>

                    <!-- Descripción -->
                    <div style="
                        font-size: 13px;
                        color: #64748B;
                        line-height: 1.6;
                        margin-bottom: 20px;
                    ">{cfg['desc']}</div>
                </div>
                """, unsafe_allow_html=True)

                # Botón de navegación
                if st.button(
                    f"Abrir {cfg['title']} →",
                    key=f"goto_{page_key}",
                    use_container_width=True,
                    type="primary",
                ):
                    st.switch_page(cfg["page"])

# ── Separador + Footer ────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-top: 50px; padding-top: 20px; border-top: 1px solid #E2E8F0;
            display: flex; align-items: center; justify-content: space-between;">
    <div style="font-size:12px; color:#94A3B8;">
        Go To Market SAC · Dashboard Comercial · 2026
    </div>
    <div style="font-size:11px; color:#CBD5E1;">
        Acceso solo para personal autorizado
    </div>
</div>
""", unsafe_allow_html=True)
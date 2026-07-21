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

# ─── CSS GLOBAL ──────────────────────────────────────────────────────────────
# IMPORTANTE: definir TODOS los estilos aquí.
# Los st.markdown() del cuerpo usan clases cortas sin estilos inline multilínea.
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
.main { background: linear-gradient(145deg,#F0F4FF 0%,#F8FAFC 55%,#FFF5F5 100%) !important; }

/* ── Header principal ── */
.gtm-header {
    background: white; border-radius: 20px; padding: 28px 36px;
    border: 1px solid #E2E8F0; box-shadow: 0 4px 24px rgba(0,0,0,0.06);
    margin-bottom: 10px; display: flex; align-items: center; justify-content: space-between;
}
.gtm-brand-title { font-size: 44px; font-weight: 900; color: #0F172A; letter-spacing: -2.5px; line-height: 1; }
.gtm-brand-sub   { font-size: 19px; font-weight: 700; color: #0F172A; margin-top: 4px; }
.gtm-bars        { display: flex; align-items: flex-end; gap: 7px; }
.gtm-bar1 { width:13px; height:22px; background:#DC2626; border-radius:3px; opacity:0.55; }
.gtm-bar2 { width:13px; height:38px; background:#DC2626; border-radius:3px; opacity:0.78; }
.gtm-bar3 { width:13px; height:54px; background:#DC2626; border-radius:3px; }
.gtm-header-right { text-align: right; }
.gtm-header-label { font-size: 11px; color: #94A3B8; font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; }
.gtm-header-year  { font-size: 22px; font-weight: 800; color: #1E293B; margin-top: 2px; }

/* ── Barra de bienvenida ── */
.welcome-bar {
    background: linear-gradient(135deg,#1E293B,#0F172A); border-radius: 16px;
    padding: 20px 28px; margin-bottom: 28px;
    display: flex; align-items: center; justify-content: space-between;
}
.welcome-name { font-size: 22px; font-weight: 800; color: white; letter-spacing: -0.4px; }
.welcome-sub  { font-size: 13px; color: #94A3B8; margin-top: 5px; }
.welcome-badge {
    background: rgba(220,38,38,0.15); border: 1px solid rgba(220,38,38,0.35);
    border-radius: 50px; padding: 8px 18px; font-size: 13px; font-weight: 700; color: #FCA5A5;
}

/* ── Tarjetas ── */
.card-chip {
    display: inline-block; border-radius: 20px; padding: 4px 12px;
    font-size: 10px; font-weight: 700; letter-spacing: 1.2px;
    text-transform: uppercase; margin-bottom: 14px;
}
.chip-lima     { background:#FFF1F1; color:#DC2626; }
.chip-norte    { background:#EFF6FF; color:#2563EB; }
.chip-provincia{ background:#F0FDF4; color:#16A34A; }

.card-icon {
    display: flex; align-items: center; justify-content: center;
    width: 72px; height: 72px; border-radius: 18px;
    font-size: 36px; margin-bottom: 16px;
}
.icon-lima     { background:#FFF1F1; border: 2px solid rgba(220,38,38,0.13); }
.icon-norte    { background:#EFF6FF; border: 2px solid rgba(37,99,235,0.13); }
.icon-provincia{ background:#F0FDF4; border: 2px solid rgba(22,163,74,0.13); }

.card-title { font-size: 20px; font-weight: 800; margin-bottom: 8px; letter-spacing: -0.3px; }
.title-lima     { color: #DC2626; }
.title-norte    { color: #2563EB; }
.title-provincia{ color: #16A34A; }

.card-desc { font-size: 13px; color: #64748B; line-height: 1.6; margin-bottom: 20px; }

/* ── Botones de tarjeta (primary) ── */
div[data-testid="stButton"] > button[kind="primary"] {
    border-radius: 10px !important; font-weight: 700 !important;
    font-size: 14px !important; padding: 10px !important; border: none !important;
    transition: transform 0.15s, box-shadow 0.15s !important;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    transform: translateY(-1px) !important; opacity: 0.9 !important;
}

/* ── Footer ── */
.gtm-footer {
    margin-top: 48px; padding-top: 20px; border-top: 1px solid #E2E8F0;
    display: flex; align-items: center; justify-content: space-between;
}
.gtm-footer-left  { font-size: 12px; color: #94A3B8; }
.gtm-footer-right { font-size: 11px; color: #CBD5E1; }

/* ── Quitar doble borde Streamlit en inputs ── */
.stTextInput [data-baseweb="input"],
.stTextInput [data-baseweb="base-input"] { box-shadow: none !important; border: none !important; }
</style>
""", unsafe_allow_html=True)

# ─── LOGIN GATE ───────────────────────────────────────────────────────────────
if not st.session_state.get("auth_ok"):
    if not auth.show_login():
        st.stop()

# ─── SESIÓN VÁLIDA ────────────────────────────────────────────────────────────
user      = auth.get_current_user()
user_info = auth.USERS[user]
pages     = auth.get_user_pages()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        f'<div style="background:linear-gradient(135deg,#FFF1F1,#FFE4E4);'
        f'border:1px solid #FECACA;border-radius:12px;padding:14px 16px;margin-bottom:10px;">'
        f'<div style="font-size:10px;color:#DC2626;font-weight:700;text-transform:uppercase;letter-spacing:1px;">SESIÓN ACTIVA</div>'
        f'<div style="font-size:15px;font-weight:700;color:#1E293B;margin-top:4px;">👤 {user_info["name"]}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    if st.button("🚪 Cerrar sesión", use_container_width=True, key="_welcome_logout"):
        auth.logout()

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="gtm-header">
<div style="display:flex;align-items:center;gap:24px;">
<div class="gtm-bars"><div class="gtm-bar1"></div><div class="gtm-bar2"></div><div class="gtm-bar3"></div></div>
<div><div class="gtm-brand-title">Go To Market SAC</div>
<div class="gtm-brand-sub">go<span style="color:#DC2626;">to</span>market</div></div>
</div>
<div class="gtm-header-right">
<div class="gtm-header-label">Portal Comercial</div>
<div class="gtm-header-year">Dashboard 2026</div>
</div>
</div>
""", unsafe_allow_html=True)

# ── Bienvenida ────────────────────────────────────────────────────────────────
num   = len(pages)
label = "dashboard" if num == 1 else "dashboards"
st.markdown(
    f'<div class="welcome-bar">'
    f'<div><div class="welcome-name">👋 Bienvenido, {user_info["name"]}</div>'
    f'<div class="welcome-sub">Tienes acceso a <b style="color:#FCA5A5;">{num} {label}</b>. Selecciona el que deseas abrir.</div></div>'
    f'<div class="welcome-badge">📅 2026</div>'
    f'</div>',
    unsafe_allow_html=True,
)

# ── Tarjetas ──────────────────────────────────────────────────────────────────
CARDS = {
    "lima": {
        "chip_class":  "chip-lima",
        "icon_class":  "icon-lima",
        "title_class": "title-lima",
        "subtitle":    "REGIÓN LIMA",
        "icon":        "🏙️",
        "title":       "Dashboard Lima",
        "desc":        "Seguimiento de visitas, prospección y mantenimiento de la región Lima.",
        "page":        "pages/1_Lima.py",
    },
    "norte": {
        "chip_class":  "chip-norte",
        "icon_class":  "icon-norte",
        "title_class": "title-norte",
        "subtitle":    "REGIÓN NORTE",
        "icon":        "🌄",
        "title":       "Dashboard Norte",
        "desc":        "Gestión comercial y embudo de ventas para la región Norte del país.",
        "page":        "pages/2_Norte.py",
    },
    "provincia": {
        "chip_class":  "chip-provincia",
        "icon_class":  "icon-provincia",
        "title_class": "title-provincia",
        "subtitle":    "REGIÓN PROVINCIA",
        "icon":        "🗺️",
        "title":       "Dashboard Provincia",
        "desc":        "Control de indicadores y alertas de cumplimiento para Provincia.",
        "page":        "pages/3_Provincia.py",
    },
}

if pages:
    cols = st.columns(len(pages), gap="large")
    for i, page_key in enumerate(pages):
        c = CARDS[page_key]
        with cols[i]:
            with st.container(border=True):
                # Chip de región — HTML simple, una sola línea
                st.markdown(
                    f'<div class="card-chip {c["chip_class"]}">{c["subtitle"]}</div>',
                    unsafe_allow_html=True,
                )
                # Icono
                st.markdown(
                    f'<div class="card-icon {c["icon_class"]}">{c["icon"]}</div>',
                    unsafe_allow_html=True,
                )
                # Título
                st.markdown(
                    f'<div class="card-title {c["title_class"]}">{c["title"]}</div>',
                    unsafe_allow_html=True,
                )
                # Descripción
                st.markdown(
                    f'<div class="card-desc">{c["desc"]}</div>',
                    unsafe_allow_html=True,
                )
                # Botón de navegación
                if st.button(
                    f"Abrir {c['title']} →",
                    key=f"goto_{page_key}",
                    use_container_width=True,
                    type="primary",
                ):
                    st.switch_page(c["page"])

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="gtm-footer">
<div class="gtm-footer-left">Go To Market SAC · Dashboard Comercial · 2026</div>
<div class="gtm-footer-right">Acceso solo para personal autorizado</div>
</div>
""", unsafe_allow_html=True)
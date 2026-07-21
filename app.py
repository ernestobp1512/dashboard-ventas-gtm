"""
app.py — Portal de acceso Go To Market SAC
Pantalla de login + bienvenida con acceso a los dashboards por región.
"""
import streamlit as st
import sys, os

# Asegurar que auth.py se encuentre en el path
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
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
.main { background-color: #F8FAFC; }

/* ── Tarjetas de dashboard ────────────────────────── */
.dash-card {
    background: white;
    border-radius: 16px;
    padding: 28px 24px 20px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    transition: transform 0.18s, box-shadow 0.18s;
    height: 100%;
}
.dash-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 12px 30px rgba(0,0,0,0.08);
}
.card-icon   { font-size: 48px; margin-bottom: 14px; }
.card-title  { font-size: 20px; font-weight: 800; margin-bottom: 6px; }
.card-desc   { font-size: 13px; color: #64748B; margin-bottom: 20px; line-height: 1.5; }

/* ── Botones de tarjeta ──────────────────────────── */
div[data-testid="stButton"] > button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
    padding: 9px !important;
    width: 100% !important;
    transition: opacity 0.18s !important;
}
div[data-testid="stButton"] > button:hover { opacity: 0.85 !important; }

/* ── Header card ─────────────────────────────────── */
.header-card {
    background: white;
    border-radius: 15px;
    padding: 28px 35px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    margin-bottom: 30px;
    display: flex;
    align-items: center;
    justify-content: space-between;
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

# ── Sidebar ───────────────────────────────────────────────────────────────────
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

# ── Header principal ─────────────────────────────────────────────────────────
st.markdown("""
<div class="header-card">
    <div style="display:flex;align-items:center;gap:22px;">
        <div style="display:flex;align-items:flex-end;gap:6px;">
            <div style="width:13px;height:22px;background:#334155;border-radius:3px;"></div>
            <div style="width:13px;height:38px;background:#334155;border-radius:3px;"></div>
            <div style="width:13px;height:54px;background:#334155;border-radius:3px;"></div>
        </div>
        <div>
            <div style="font-size:48px;font-weight:900;color:#0F172A;
                        letter-spacing:-2px;line-height:1;">Go To Market SAC</div>
            <div style="font-size:22px;font-weight:800;color:#0F172A;margin-top:3px;">
                go<span style="color:#DC2626;">to</span>market
            </div>
        </div>
    </div>
    <div style="text-align:right;color:#64748B;font-size:12px;">
        <b>PORTAL COMERCIAL</b><br>Dashboard Gestión 2026
    </div>
</div>
""", unsafe_allow_html=True)

# ── Saludo personalizado ──────────────────────────────────────────────────────
st.markdown(f"""
<div style="margin-bottom:28px;">
    <div style="font-size:26px;font-weight:800;color:#1E293B;">
        👋 Bienvenido, {user_info['name']}
    </div>
    <div style="font-size:14px;color:#64748B;margin-top:5px;">
        Tienes acceso a <b>{len(pages)}</b>
        {'dashboard' if len(pages) == 1 else 'dashboards'}.
        Selecciona el que deseas ver.
    </div>
</div>
""", unsafe_allow_html=True)

# ── Tarjetas de dashboards accesibles ────────────────────────────────────────
CARD_STYLE = {
    "lima":      {"border": "#DC2626", "bg_icon": "#FFF1F1"},
    "norte":     {"border": "#2563EB", "bg_icon": "#EFF6FF"},
    "provincia": {"border": "#16A34A", "bg_icon": "#F0FDF4"},
}

# CSS extra para los botones de cada tarjeta
CARD_BTN_CSS = {
    "lima":      "background: #DC2626;",
    "norte":     "background: #2563EB;",
    "provincia": "background: #16A34A;",
}

if pages:
    cols = st.columns(len(pages), gap="large")
    for i, page_key in enumerate(pages):
        info  = auth.PAGE_INFO[page_key]
        style = CARD_STYLE[page_key]
        with cols[i]:
            with st.container(border=True):
                # Icono + título + descripción
                st.markdown(f"""
                <div style="padding: 6px 0 14px 0;">
                    <div style="
                        display: inline-flex;
                        align-items: center;
                        justify-content: center;
                        width: 64px; height: 64px;
                        background: {style['bg_icon']};
                        border-radius: 16px;
                        font-size: 32px;
                        margin-bottom: 14px;
                    ">{info['icon']}</div>
                    <div style="font-size:20px; font-weight:800;
                                color:{info['color']}; margin-bottom:6px;">
                        {info['label']}
                    </div>
                    <div style="font-size:13px; color:#64748B;
                                line-height:1.5; margin-bottom:18px;">
                        {info['desc']}
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Botón de navegación — st.button + st.switch_page es
                # más confiable que st.page_link en todas las versiones
                if st.button(
                    f"Abrir {info['label']} →",
                    key=f"goto_{page_key}",
                    use_container_width=True,
                    type="primary",
                ):
                    st.switch_page(info["page"])

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;color:#CBD5E1;font-size:12px;
            margin-top:50px;padding-top:20px;border-top:1px solid #E2E8F0;">
    Go To Market SAC · Dashboard Comercial · 2026
</div>
""", unsafe_allow_html=True)
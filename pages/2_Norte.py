"""
pages/2_Norte.py — Wrapper para Dashboard Norte
Lee y ejecuta Dashboard Norte.py sin modificarlo.
La autenticación y el control de acceso se manejan aquí.
"""
import streamlit as st
import sys, os

# ─── PATH AL DIRECTORIO RAÍZ DEL PROYECTO ────────────────────────────────────
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)
import auth

# ─── PAGE CONFIG (debe ser la primera llamada a Streamlit) ───────────────────
st.set_page_config(
    page_title="GTM SAC - REGIÓN NORTE",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── VERIFICAR SESIÓN ─────────────────────────────────────────────────────────
if not st.session_state.get("auth_ok"):
    st.switch_page("app.py")
    st.stop()

# ─── VERIFICAR PERMISO ────────────────────────────────────────────────────────
if not auth.has_access("norte"):
    st.error("🚫 No tienes permiso para ver esta página.")
    pages_accesibles = [auth.PAGE_INFO[p]["label"] for p in auth.get_user_pages()]
    if pages_accesibles:
        st.info(f"Tu acceso está limitado a: **{', '.join(pages_accesibles)}**")
    if st.button("← Volver al inicio", type="primary"):
        st.switch_page("app.py")
    st.stop()

# ─── INFO DE USUARIO EN SIDEBAR ───────────────────────────────────────────────
auth.show_sidebar_user()

# ─── EJECUTAR DASHBOARD SIN MODIFICAR EL ARCHIVO ORIGINAL ───────────────────
_orig_set_page_config = st.set_page_config
st.set_page_config = lambda *args, **kwargs: None

_dash_path = os.path.join(_ROOT, "Dashboard Norte.py")
with open(_dash_path, "r", encoding="utf-8") as _f:
    _code = _f.read()

exec(compile(_code, _dash_path, "exec"), {"__file__": _dash_path, "__name__": "__main__"})

st.set_page_config = _orig_set_page_config

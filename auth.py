"""
auth.py — Módulo de autenticación compartido para GTM Dashboard
Importado por app.py y por cada wrapper en pages/.
NO llama a st.set_page_config.
"""
import streamlit as st
import hmac

# ══════════════════════════════════════════════════════════════════════════════
# USUARIOS, CONTRASEÑAS Y PERMISOS
# ══════════════════════════════════════════════════════════════════════════════

USERS: dict = {
    "admin": {
        "password": "Admin@GTM2026",
        "name":     "Administrador",
        "access":   ["lima", "norte", "provincia"],
    },
    "lima": {
        "password": "Lima@GTM2026",
        "name":     "Responsable Lima",
        "access":   ["lima"],
    },
    "norte": {
        "password": "Norte@GTM2026",
        "name":     "Responsable Norte",
        "access":   ["norte"],
    },
    "provincia": {
        "password": "Prov@GTM2026",
        "name":     "Responsable Provincia",
        "access":   ["provincia"],
    },
}

# Información de cada página para mostrar en la UI
PAGE_INFO: dict = {
    "lima": {
        "label": "Dashboard Lima",
        "icon":  "🏙️",
        "color": "#DC2626",
        "desc":  "Gestión comercial · Región Lima",
        "page":  "pages/1_Lima.py",
    },
    "norte": {
        "label": "Dashboard Norte",
        "icon":  "🌄",
        "color": "#2563EB",
        "desc":  "Gestión comercial · Región Norte",
        "page":  "pages/2_Norte.py",
    },
    "provincia": {
        "label": "Dashboard Provincia",
        "icon":  "🗺️",
        "color": "#16A34A",
        "desc":  "Gestión comercial · Región Provincia",
        "page":  "pages/3_Provincia.py",
    },
}

# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA DE LOGIN
# ══════════════════════════════════════════════════════════════════════════════

def _do_login() -> None:
    """Callback para el formulario de login."""
    user = st.session_state.get("_auth_user", "").strip().lower()
    pwd  = st.session_state.get("_auth_pwd", "")

    if user in USERS and hmac.compare_digest(USERS[user]["password"], pwd):
        st.session_state["auth_ok"]   = True
        st.session_state["auth_user"] = user
        st.session_state.pop("_auth_pwd", None)
    else:
        st.session_state["auth_ok"]    = False
        st.session_state["auth_error"] = True


def show_login() -> bool:
    """
    Renderiza la pantalla de login.
    Devuelve True si ya hay sesión válida, False si no.
    """
    if st.session_state.get("auth_ok"):
        return True

    # ── CSS del login ─────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

    /* Ocultar sidebar en la pantalla de login */
    [data-testid="stSidebar"]       { display: none !important; }
    [data-testid="collapsedControl"]{ display: none !important; }

    /* Centrar y limitar ancho del contenido */
    .main .block-container {
        max-width: 460px !important;
        padding-top: 60px !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    /* Inputs */
    .stTextInput input {
        border-radius: 10px !important;
        border: 1.5px solid #E2E8F0 !important;
        padding: 10px 14px !important;
        font-size: 14px !important;
        background: #FAFAFA !important;
        transition: border 0.2s, box-shadow 0.2s;
    }
    .stTextInput input:focus {
        border-color: #DC2626 !important;
        box-shadow: 0 0 0 3px rgba(220,38,38,0.1) !important;
        background: white !important;
    }

    /* Botón de login */
    div[data-testid="stButton"] > button[kind="primary"] {
        background: linear-gradient(135deg, #DC2626, #991B1B) !important;
        color: white !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        padding: 12px !important;
        border: none !important;
        width: 100% !important;
        margin-top: 6px !important;
        letter-spacing: 0.3px;
        transition: opacity 0.2s !important;
    }
    div[data-testid="stButton"] > button[kind="primary"]:hover {
        opacity: 0.88 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Logo y marca ──────────────────────────────────────────────────────────
    st.markdown("""
    <div style="text-align:center; margin-bottom:36px;">
        <div style="display:inline-flex; align-items:flex-end; gap:6px; margin-bottom:14px;">
            <div style="width:11px;height:22px;background:#DC2626;border-radius:3px;"></div>
            <div style="width:11px;height:38px;background:#DC2626;border-radius:3px;"></div>
            <div style="width:11px;height:52px;background:#DC2626;border-radius:3px;"></div>
        </div>
        <div style="font-size:44px;font-weight:900;color:#0F172A;letter-spacing:-2px;line-height:1.05;">
            Go To Market SAC
        </div>
        <div style="font-size:18px;font-weight:700;color:#0F172A;margin-top:5px;">
            go<span style="color:#DC2626;">to</span>market
        </div>
        <div style="font-size:12px;color:#94A3B8;margin-top:8px;letter-spacing:1px;">
            DASHBOARD COMERCIAL 2026
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Tarjeta de login ──────────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("""
        <div style="text-align:center;margin-bottom:20px;">
            <div style="font-size:19px;font-weight:700;color:#1E293B;">🔒 Acceso Restringido</div>
            <div style="font-size:13px;color:#64748B;margin-top:4px;">
                Ingresa tus credenciales para continuar
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.text_input("Usuario", key="_auth_user", placeholder="Tu usuario")
        st.text_input(
            "Contraseña", type="password",
            key="_auth_pwd", placeholder="Tu contraseña",
            on_change=_do_login,
        )
        st.button("Ingresar →", on_click=_do_login,
                  use_container_width=True, type="primary")

        if st.session_state.get("auth_error"):
            st.error("❌ Usuario o contraseña incorrectos")

    st.markdown("""
    <div style="text-align:center;color:#CBD5E1;font-size:11px;margin-top:18px;">
        Go To Market SAC · Acceso solo para personal autorizado
    </div>
    """, unsafe_allow_html=True)

    return False


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS DE SESIÓN Y PERMISOS
# ══════════════════════════════════════════════════════════════════════════════

def get_current_user() -> str | None:
    """Devuelve el usuario autenticado o None."""
    if st.session_state.get("auth_ok"):
        return st.session_state.get("auth_user")
    return None


def has_access(page_key: str) -> bool:
    """True si el usuario activo tiene acceso a la página indicada."""
    user = get_current_user()
    return user is not None and page_key in USERS.get(user, {}).get("access", [])


def get_user_pages() -> list[str]:
    """Devuelve la lista de page_keys accesibles para el usuario activo."""
    user = get_current_user()
    return USERS[user]["access"] if user else []


def logout() -> None:
    """Cierra la sesión y recarga la app."""
    for key in ["auth_ok", "auth_user", "auth_error",
                "_auth_user", "_auth_pwd", "password_correct"]:
        st.session_state.pop(key, None)
    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR CON INFORMACIÓN DE USUARIO (para páginas wrapper)
# ══════════════════════════════════════════════════════════════════════════════

def show_sidebar_user() -> None:
    """
    Muestra en el sidebar: tarjeta de usuario + botón de logout.
    Llamar desde cada página wrapper ANTES de exec() del dashboard.
    """
    user = get_current_user()
    if not user:
        return

    info = USERS[user]
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
                        text-transform:uppercase;letter-spacing:1px;">
                SESIÓN ACTIVA
            </div>
            <div style="font-size:15px;font-weight:700;color:#1E293B;margin-top:4px;">
                👤 {info['name']}
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("🚪 Cerrar sesión", use_container_width=True,
                     key="_sidebar_logout"):
            logout()

        st.divider()

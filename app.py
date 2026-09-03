import base64
import os
import streamlit as st

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Sistema Avícola - HUPA",
    page_icon="🐔",
    layout="centered",
    initial_sidebar_state="expanded"
    if st.session_state.get("authenticated", False)
    else "collapsed",
)

# Credenciales de acceso
USER_CORRECTO = "ADMIN_HUPA"
PASS_CORRECTO = "PASCUAL2026"

# Inicializar estado de autenticación
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False


# --- FUNCIÓN PARA CONVERTIR IMAGEN A BASE64 PARA CSS ---
def obtener_base64_imagen(ruta_imagen):
    if os.path.exists(ruta_imagen):
        with open(ruta_imagen, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        return f"data:image/jpeg;base64,{encoded_string}"
    return None


def login():
    ruta_fondo = os.path.join("DATA", "fondo.jpg")
    bg_image = obtener_base64_imagen(ruta_fondo)

    # CSS de fondo e interfaz de login
    css_fondo = f"""
    <style>
    /* Fondo de pantalla completa con overlay translúcido */
    .stApp {{
        background: linear-gradient(rgba(15, 23, 42, 0.65), rgba(15, 23, 42, 0.65)), url("{bg_image}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
    }}
    
    /* Contenedor tipo tarjeta flotante */
    div[data-testid="stForm"] {{
        background-color: rgba(255, 255, 255, 0.95) !important;
        padding: 30px !important;
        border-radius: 16px !important;
        border: 1px solid rgba(226, 232, 240, 0.8) !important;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.3), 0 10px 10px -5px rgba(0, 0, 0, 0.2) !important;
    }}

    .login-title {{
        text-align: center;
        color: #0f172a;
        font-size: 20px;
        font-weight: 700;
        margin-bottom: 15px;
    }}
    </style>
    """
    st.markdown(css_fondo, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([0.5, 2, 0.5])

    with col2:
        with st.form("form_login"):
            # Mostrar logo HUPA si existe en DATA
            ruta_logo = os.path.join("DATA", "logo hupa.png")
            if os.path.exists(ruta_logo):
                st.image(ruta_logo, use_container_width=True)

            st.markdown(
                "<div class='login-title'>🐔 Panel de Acceso Avícola</div>",
                unsafe_allow_html=True,
            )

            usuario = st.text_input(
                "Usuario:", placeholder="ADMIN_HUPA", key="input_user"
            )
            password = st.text_input(
                "Contraseña:",
                type="password",
                placeholder="••••••••",
                key="input_pass",
            )
            submit_button = st.form_submit_button(
                "🔑 Iniciar Sesión", use_container_width=True
            )

            if submit_button:
                if usuario == USER_CORRECTO and password == PASS_CORRECTO:
                    st.session_state["authenticated"] = True
                    st.session_state["user"] = usuario
                    st.success("¡Autenticación exitosa! Accediendo...")
                    st.rerun()
                else:
                    st.error("⚠️ Usuario o contraseña incorrectos.")


# --- NAVEGACIÓN Y ACCESO ---
if not st.session_state["authenticated"]:
    # Ocultar la barra lateral hasta iniciar sesión
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] {display: none;}
        </style>
        """,
        unsafe_allow_html=True,
    )
    login()
else:
    st.sidebar.markdown(
        f"👤 **Usuario Activo:** `{st.session_state.get('user', 'ADMIN')}`"
    )
    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["user"] = None
        st.rerun()

    st.title("🐔 Panel de Control Avícola - HUPA")
    st.success(
        "Bienvenido al sistema. Selecciona una opción en la barra lateral para explorar las consultas."
    )

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.info(
            "📋 **Reporte Consolidado**\n\nResumen agrupado por fecha, granja y lote con filtros unificados."
        )
    with c2:
        st.info(
            "📊 **Histórico por Lote**\n\nAuditoría detallada día por día con extracción de comentarios y facturas."
        )
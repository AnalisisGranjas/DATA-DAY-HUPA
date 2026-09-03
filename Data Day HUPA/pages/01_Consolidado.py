import io
import os
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# --- VALIDACIÓN DE SESIÓN (LOGIN SECURITY) ---
if not st.session_state.get("authenticated", False):
    st.warning("⚠️ Debes iniciar sesión para acceder a este reporte.")
    st.stop()

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Consolidado General - Avícola",
    page_icon="🐔",
    layout="wide",
)

# --- LOGO EN LA BARRA LATERAL ---
ruta_logo = os.path.join("DATA", "logo hupa.png")
if os.path.exists(ruta_logo):
    st.sidebar.image(ruta_logo, use_container_width=True)
    st.sidebar.divider()

# --- ESTILOS CSS PARA REDISEÑAR FILTROS Y CONTENEDOR CON SCROLL ---
st.markdown(
    """
    <style>
    span[data-baseweb="tag"] {
        background-color: #f1f5f9 !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        color: #334155 !important;
        font-size: 11px !important;
        padding: 2px 6px !important;
        margin: 2px !important;
    }
    
    span[data-baseweb="tag"] span[role="button"] {
        color: #64748b !important;
    }
    span[data-baseweb="tag"] span[role="button"]:hover {
        color: #0f172a !important;
    }

    div[data-baseweb="select"] > div {
        max-height: 42px !important;
        overflow-y: auto !important;
        border-radius: 8px !important;
        border-color: #cbd5e1 !important;
        background-color: #ffffff !important;
    }

    .filter-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 20px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🐔 Consulta General y Reporte Consolidado por Fecha")
st.markdown(
    "Visualización y consulta consolidada con filtros de **Razón Social, Granja de Producción, Lotes y Rango Rápido de Fechas**."
)
st.divider()

RUTA_REPORTE_LOCAL = os.path.join("DATA", "REPORTE_AVITRACK_FINAL.xlsx")


# --- FUNCIÓN DE CARGA ---
@st.cache_data(ttl=60)
def cargar_datos_locales():
    if os.path.exists(RUTA_REPORTE_LOCAL):
        try:
            return pd.read_excel(RUTA_REPORTE_LOCAL)
        except Exception as e:
            st.error(f"Error al abrir el archivo consolidado local: {e}")
            return None
    return None


df_base = cargar_datos_locales()

if df_base is not None and not df_base.empty:

    # 1. Normalizar columna de Fecha
    col_fecha_origen = (
        "Fecha" if "Fecha" in df_base.columns else df_base.columns[1]
    )
    df_base["Fecha_dt"] = pd.to_datetime(
        df_base[col_fecha_origen], dayfirst=True, errors="coerce"
    )

    # 2. Identificar columnas de control
    col_rs = [
        c
        for c in df_base.columns
        if "razon social" in c.lower() or "empresa" in c.lower()
    ]
    nombre_col_rs = col_rs[0] if col_rs else "Razon Social"

    col_granja_p = [c for c in df_base.columns if "Nombre de Granja (P)" in c]
    nombre_col_granja = (
        col_granja_p[0] if col_granja_p else "Nombre de Granja (L) :"
    )

    col_lote = [
        c for c in df_base.columns if "Número de Lote" in c or "Lote" in c
    ]
    nombre_col_lote = col_lote[0] if col_lote else "Archivo"

    # 3. Opciones para multiselect
    opciones_rs = sorted(
        [
            str(x)
            for x in df_base[nombre_col_rs].dropna().unique()
            if str(x).strip() != ""
        ]
    )
    opciones_granjas = sorted(
        [
            str(x)
            for x in df_base[nombre_col_granja].dropna().unique()
            if str(x).strip() != ""
        ]
    )
    opciones_lotes = sorted(
        [
            str(x)
            for x in df_base[nombre_col_lote].dropna().unique()
            if str(x).strip() != ""
        ]
    )

    # --- INICIALIZACIÓN DE SESSION STATE ---
    st.session_state.setdefault("sel_rs", [])
    st.session_state.setdefault("sel_granjas", [])
    st.session_state.setdefault("sel_lotes", [])

    # Configuración de fecha por defecto (Hoy - 7 días hasta la fecha máxima registrada)
    fechas_validas = df_base["Fecha_dt"].dropna()
    if not fechas_validas.empty:
        max_fecha_dt = fechas_validas.max().date()
        min_fecha_dt = fechas_validas.min().date()
        default_inicio = max(min_fecha_dt, max_fecha_dt - timedelta(days=7))
        st.session_state.setdefault("rango_fechas", [default_inicio, max_fecha_dt])

    # --- BARRA LATERAL ---
    st.sidebar.header("⚙️ Acciones de Filtro")

    col_btn1, col_btn2 = st.sidebar.columns(2)
    with col_btn1:
        if st.sidebar.button("✅ Seleccionar Todo", use_container_width=True):
            st.session_state["sel_rs"] = opciones_rs
            st.session_state["sel_granjas"] = opciones_granjas
            st.session_state["sel_lotes"] = opciones_lotes
            st.rerun()

    with col_btn2:
        if st.sidebar.button("🧹 Limpiar Filtros", use_container_width=True):
            st.session_state["sel_rs"] = []
            st.session_state["sel_granjas"] = []
            st.session_state["sel_lotes"] = []
            st.rerun()

    st.sidebar.divider()
    st.sidebar.header("📂 Estado de Datos")
    st.sidebar.success(
        f"**Origen:** `{RUTA_REPORTE_LOCAL}`\n\n**Total Registros:** {len(df_base):,}"
    )

    if st.sidebar.button("🔄 Recargar Datos"):
        st.cache_data.clear()
        st.rerun()

    # --- LIMPIEZA DE VALORES NUMÉRICOS ---
    cols_a_limpiar = [
        "Mort.",
        "Otros",
        "Selec.",
        "Trasl Ventas",
        "Saldo Aves",
        "Costo Alimento",
        "Ingreso B X 40 K",
        "Consumo B X 40 K",
        "Traslado B X 40 K",
        "Saldo B X 40 K",
        "Producción Huevos Día",
        "Salida Huevos dia",
        "Saldo de Huevo",
        "Saldo de Huevos",
        "Ingreso",
        "Consumo",
        "Traslado",
        "Traslados",
        "Saldo",
    ]

    for col in df_base.columns:
        if any(
            c_key.lower() == col.lower() or c_key.lower() in col.lower()
            for c_key in cols_a_limpiar
        ):
            if df_base[col].dtype == "object":
                df_base[col] = (
                    df_base[col]
                    .astype(str)
                    .str.replace("$", "", regex=False)
                    .str.replace(" ", "", regex=False)
                    .str.replace(",", ".", regex=False)
                    .str.strip()
                )
            df_base[col] = pd.to_numeric(
                df_base[col], errors="coerce"
            ).fillna(0)

    # --- FILTROS DE CONSULTA MODERNIZADOS CON RANGOS RÁPIDOS ---
    with st.expander("🔍 **Panel de Filtros de Consulta**", expanded=True):
        f_col1, f_col2, f_col3, f_col4 = st.columns(4)

        with f_col1:
            rs_sel = st.multiselect(
                "Razón Social / Empresa:",
                options=opciones_rs,
                key="sel_rs",
                placeholder="Todas las empresas",
            )

        with f_col2:
            granja_sel = st.multiselect(
                "Granja (Producción):",
                options=opciones_granjas,
                key="sel_granjas",
                placeholder="Todas las granjas",
            )

        with f_col3:
            lote_sel = st.multiselect(
                "Lote:",
                options=opciones_lotes,
                key="sel_lotes",
                placeholder="Todos los lotes",
            )

        with f_col4:
            if not fechas_validas.empty:
                max_f = fechas_validas.max().date()
                min_f = fechas_validas.min().date()

                st.markdown("**Acceso Rápido por Días:**")
                b_c1, b_c2, b_c3 = st.columns(3)
                with b_c1:
                    if st.button("7 días", use_container_width=True):
                        st.session_state["rango_fechas"] = [max(min_f, max_f - timedelta(days=7)), max_f]
                        st.rerun()
                with b_c2:
                    if st.button("15 días", use_container_width=True):
                        st.session_state["rango_fechas"] = [max(min_f, max_f - timedelta(days=15)), max_f]
                        st.rerun()
                with b_c3:
                    if st.button("30 días", use_container_width=True):
                        st.session_state["rango_fechas"] = [max(min_f, max_f - timedelta(days=30)), max_f]
                        st.rerun()

                rango_fechas = st.date_input(
                    "Rango de Fechas Seleccionado:",
                    key="rango_fechas",
                    min_value=min_f,
                    max_value=max_f,
                )
            else:
                rango_fechas = []

    # --- APLICACIÓN DE FILTROS ---
    df_filtrado = df_base.copy()

    if rs_sel:
        df_filtrado = df_filtrado[
            df_filtrado[nombre_col_rs].astype(str).isin(rs_sel)
        ]

    if granja_sel:
        df_filtrado = df_filtrado[
            df_filtrado[nombre_col_granja].astype(str).isin(granja_sel)
        ]

    if lote_sel:
        df_filtrado = df_filtrado[
            df_filtrado[nombre_col_lote].astype(str).isin(lote_sel)
        ]

    if len(rango_fechas) == 2:
        df_filtrado = df_filtrado[
            (df_filtrado["Fecha_dt"].dt.date >= rango_fechas[0])
            & (df_filtrado["Fecha_dt"].dt.date <= rango_fechas[1])
        ]
    elif len(rango_fechas) == 1:
        df_filtrado = df_filtrado[
            df_filtrado["Fecha_dt"].dt.date == rango_fechas[0]
        ]

    st.divider()

    # --- MAPEO DE COLUMNAS MÉTRICAS ---
    def obtener_columna_exacta(patron, indice_bloque=0):
        cols = [c for c in df_filtrado.columns if patron.lower() in c.lower()]
        if cols and len(cols) > indice_bloque:
            return cols[indice_bloque]
        return cols[0] if cols else None

    c_fecha = col_fecha_origen
    c_mort = obtener_columna_exacta("mort")
    c_otros = obtener_columna_exacta("otros")
    c_selec = obtener_columna_exacta("selec")
    c_trasl_ventas = obtener_columna_exacta("trasl ventas") or obtener_columna_exacta("ventas")
    c_saldo_aves = obtener_columna_exacta("saldo aves")

    c_costo_alim = obtener_columna_exacta("costo alimento")
    c_ingreso_b = obtener_columna_exacta("ingreso b x 40")
    c_consumo_b = obtener_columna_exacta("consumo b x 40")
    c_traslado_b = obtener_columna_exacta("traslado b x 40")
    c_saldo_b = obtener_columna_exacta("saldo b x 40")

    c_prod_huevo = obtener_columna_exacta("producción huevos") or obtener_columna_exacta("prod")
    c_salida_huevo = obtener_columna_exacta("salida huevos") or obtener_columna_exacta("salida")
    c_saldo_huevo = obtener_columna_exacta("saldo de huevo")

    cols_sin_comentarios = [c for c in df_filtrado.columns if "comentario" not in c.lower()]

    def buscar_en_limpias(patron):
        c_matches = [c for c in cols_sin_comentarios if patron.lower() in c.lower()]
        return c_matches[-1] if c_matches else None

    c_ing_band = buscar_en_limpias("ingreso")
    c_cons_band = buscar_en_limpias("consumo")
    c_tras_band = buscar_en_limpias("traslado")
    c_sal_band = buscar_en_limpias("saldo")

    columnas_ordenadas = [
        ("Mort.", c_mort, "aves", "int"),
        ("Otros", c_otros, "aves", "int"),
        ("Selec.", c_selec, "aves", "int"),
        ("Trasl Ventas", c_trasl_ventas, "aves", "int"),
        ("Saldo Aves", c_saldo_aves, "aves", "int"),
        ("Costo Alimento", c_costo_alim, "alimento", "currency"),
        ("Ingreso B X 40 K", c_ingreso_b, "alimento", "float"),
        ("Consumo B X 40 K", c_consumo_b, "alimento", "float"),
        ("Traslado B X 40 K", c_traslado_b, "alimento", "float"),
        ("Saldo B X 40 K", c_saldo_b, "alimento", "float"),
        ("Producción Huevos Día", c_prod_huevo, "huevos", "int"),
        ("Salida Huevos dia", c_salida_huevo, "huevos", "int"),
        ("Saldo de Huevos", c_saldo_huevo, "huevos", "int"),
        ("Ingreso", c_ing_band, "bandejas", "int"),
        ("Consumo", c_cons_band, "bandejas", "int"),
        ("Traslados", c_tras_band, "bandejas", "int"),
        ("Saldo", c_sal_band, "bandejas", "int"),
    ]

    st.subheader("📋 Consolidado Avícola por Fecha")

    fechas_unicas_dt = sorted(df_filtrado["Fecha_dt"].dropna().unique())
    fechas_unicas_str = [pd.to_datetime(f).strftime("%d/%m/%Y") for f in fechas_unicas_dt]

    col_vis1, col_vis2 = st.columns([1, 2])
    with col_vis1:
        mostrar_desglose = st.checkbox("🔍 Desglosar Granja & Lote por Fecha", value=False)
    
    fecha_a_desglosar = None
    if mostrar_desglose and fechas_unicas_str:
        with col_vis2:
            fecha_a_desglosar = st.selectbox(
                "Selecciona la fecha a desglosar:",
                options=["Todas las Fechas"] + fechas_unicas_str,
                index=0
            )

    filas_unificadas = []

    for f_dt in fechas_unicas_dt:
        df_sub_fecha = df_filtrado[df_filtrado["Fecha_dt"] == f_dt]
        fecha_str = pd.to_datetime(f_dt).strftime("%d/%m/%Y")

        debe_desglosar_esta = mostrar_desglose and (
            fecha_a_desglosar == "Todas las Fechas" or fecha_a_desglosar == fecha_str
        )

        label_fecha = fecha_str if not debe_desglosar_esta else f"📅 {fecha_str} (TOTAL DÍA)"
        fila_fecha = {"Fecha / Concepto": label_fecha}
        for nombre_final, col_orig, _, _ in columnas_ordenadas:
            if col_orig in df_sub_fecha.columns:
                fila_fecha[nombre_final] = df_sub_fecha[col_orig].sum()
        filas_unificadas.append(fila_fecha)

        if debe_desglosar_esta:
            df_sub_fecha_agrup = (
                df_sub_fecha.groupby([nombre_col_granja, nombre_col_lote], as_index=False)
                .sum(numeric_only=True)
            )
            
            for _, row_gl in df_sub_fecha_agrup.iterrows():
                g_val = row_gl[nombre_col_granja]
                l_val = row_gl[nombre_col_lote]
                fila_gl = {"Fecha / Concepto": f"   └─ {g_val} - Lote: {l_val}"}
                for nombre_final, col_orig, _, _ in columnas_ordenadas:
                    if col_orig in row_gl.index:
                        fila_gl[nombre_final] = row_gl[col_orig]
                filas_unificadas.append(fila_gl)

    if filas_unificadas:
        df_unificado = pd.DataFrame(filas_unificadas)

        # Fila TOTAL GENERAL
        fila_total_general = {"Fecha / Concepto": "TOTAL GENERAL"}
        for nombre_final, _, _, _ in columnas_ordenadas:
            if nombre_final in df_unificado.columns:
                if "saldo" in nombre_final.lower():
                    ultima_fecha_str = fechas_unicas_str[-1]
                    if mostrar_desglose:
                        filas_max_dia = df_unificado[df_unificado["Fecha / Concepto"].str.contains(ultima_fecha_str)]
                    else:
                        filas_max_dia = df_unificado[df_unificado["Fecha / Concepto"] == ultima_fecha_str]
                    
                    if not filas_max_dia.empty:
                        fila_total_general[nombre_final] = filas_max_dia[nombre_final].iloc[0]
                    else:
                        fila_total_general[nombre_final] = df_unificado[nombre_final].iloc[-1]
                else:
                    if mostrar_desglose:
                        filas_base = df_unificado[~df_unificado["Fecha / Concepto"].str.contains("└─")]
                    else:
                        filas_base = df_unificado
                    fila_total_general[nombre_final] = filas_base[nombre_final].sum()

        df_display = pd.concat([df_unificado, pd.DataFrame([fila_total_general])], ignore_index=True)

        # HTML / CSS CON SCROLL VERTICAL (~12 FILAS) Y STICKY HEADER
        html_code = """
        <style>
            .scroll-table-container {
                width: 100%;
                max-height: 460px;
                overflow-y: auto;
                overflow-x: hidden;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                font-family: system-ui, -apple-system, sans-serif;
            }
            .custom-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px; }
            .custom-table th {
                position: sticky;
                top: 0;
                z-index: 2;
                padding: 6px 2px;
                text-align: center;
                font-weight: 700;
                word-wrap: break-word;
                line-height: 1.1;
                border: 1px solid #cbd5e1;
                font-size: 10px;
                box-shadow: 0 1px 2px rgba(0,0,0,0.05);
            }
            .custom-table td { padding: 5px 2px; text-align: right; border: 1px solid #e2e8f0; white-space: nowrap; font-size: 10.5px; }
            .th-fecha { background-color: #f1f5f9; width: 14%; text-align: left !important; }
            .th-aves { background-color: #dbeafe; color: #1e40af; }
            .th-alimento { background-color: #fef3c7; color: #92400e; }
            .th-costo { background-color: #fef3c7; color: #92400e; width: 7.5%; }
            .th-huevos { background-color: #d1fae5; color: #065f46; }
            .th-bandejas { background-color: #ffe4e6; color: #9f1239; }
            .td-aves { background-color: #ebf3fe; }
            .td-alimento { background-color: #fef8ea; }
            .td-huevos { background-color: #eaf8f0; }
            .td-bandejas { background-color: #fdf0ed; }
            .row-total td {
                position: sticky;
                bottom: 0;
                z-index: 2;
                font-weight: bold;
                background-color: #cbd5e1 !important;
                border-top: 2px solid #94a3b8;
            }
            .row-header-day { font-weight: bold; background-color: #f1f5f9; }
            .td-concept { text-align: left !important; overflow: hidden; text-overflow: ellipsis; }
        </style>
        <div class="scroll-table-container">
        <table class="custom-table">
            <thead>
                <tr>
                    <th class="th-fecha">Fecha / Granja & Lote</th>
        """

        for nombre_final, _, bloque, tipo_dato in columnas_ordenadas:
            clase_th = "th-costo" if tipo_dato == "currency" else f"th-{bloque}"
            html_code += f'<th class="{clase_th}">{nombre_final}</th>'

        html_code += "</tr></thead><tbody>"

        for _, row in df_display.iterrows():
            concepto = str(row["Fecha / Concepto"])
            is_total = "TOTAL GENERAL" in concepto
            is_day_total = "TOTAL DÍA" in concepto
            
            row_class = "row-total" if is_total else ("row-header-day" if is_day_total else "")
            
            html_code += f'<tr class="{row_class}">'
            html_code += f'<td class="td-concept" title="{concepto}">{concepto}</td>'

            for nombre_final, _, bloque, tipo_dato in columnas_ordenadas:
                val = row.get(nombre_final, 0)
                if pd.isna(val):
                    val = 0
                
                if tipo_dato == "int":
                    val_str = f"{int(round(val)):,}"
                elif tipo_dato == "currency":
                    val_str = f"$ {int(round(val)):,}"
                else:
                    val_str = f"{val:,.2f}"
                
                html_code += f'<td class="{f"td-{bloque}"}">{val_str}</td>'

            html_code += "</tr>"

        html_code += "</tbody></table></div>"

        st.html(html_code)

        # Botón de Descarga Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_display.to_excel(
                writer, index=False, sheet_name="Consolidado_Avicola"
            )
            df_filtrado.to_excel(
                writer, index=False, sheet_name="Detalle_Filtrado"
            )
        buffer.seek(0)

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📥 Descargar Reporte Consolidado (.xlsx)",
            data=buffer,
            file_name=f"CONSULTA_CONSOLIDADA_AVICOLA_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No se encontraron registros con los filtros seleccionados.")
else:
    st.warning(
        f"⚠️ No se encontró el archivo consolidado en `{RUTA_REPORTE_LOCAL}`."
    )
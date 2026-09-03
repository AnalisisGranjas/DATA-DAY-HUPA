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
    page_title="Histórico por Lote - Avícola",
    page_icon="📊",
    layout="wide",
)

# --- LOGO EN LA BARRA LATERAL ---
ruta_logo = os.path.join("DATA", "logo hupa.png")
if os.path.exists(ruta_logo):
    st.sidebar.image(ruta_logo, use_container_width=True)
    st.sidebar.divider()

# --- ESTILOS CSS ---
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
    div[data-baseweb="select"] > div {
        max-height: 42px !important;
        overflow-y: auto !important;
        border-radius: 8px !important;
        border-color: #cbd5e1 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Consulta del Histórico Día a Día por Lote")
st.markdown(
    "Audita la evolución detallada día a día."
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
    col_granja_p = [c for c in df_base.columns if "Nombre de Granja (P)" in c]
    nombre_col_granja = (
        col_granja_p[0] if col_granja_p else "Nombre de Granja (L) :"
    )

    col_lote = [
        c for c in df_base.columns if "Número de Lote" in c or "Lote" in c
    ]
    nombre_col_lote = col_lote[0] if col_lote else "Archivo"

    # --- CONTROLES DE SELECCIÓN DE LOTE Y FECHAS ---
    st.subheader("🎯 Selección de Lote y Rango de Fechas")
    c_f1, c_f2, c_f3 = st.columns([1.5, 1.5, 2])

    opciones_granjas = sorted(
        [
            str(x)
            for x in df_base[nombre_col_granja].dropna().unique()
            if str(x).strip() != ""
        ]
    )
    with c_f1:
        granja_sel = st.selectbox("Selecciona la Granja:", options=opciones_granjas)

    df_sub_granja = df_base[df_base[nombre_col_granja].astype(str) == granja_sel]
    opciones_lotes = sorted(
        [
            str(x)
            for x in df_sub_granja[nombre_col_lote].dropna().unique()
            if str(x).strip() != ""
        ]
    )

    with c_f2:
        lote_sel = st.selectbox("Selecciona el Lote:", options=opciones_lotes)

    df_lote = df_sub_granja[df_sub_granja[nombre_col_lote].astype(str) == lote_sel]
    fechas_validas = df_lote["Fecha_dt"].dropna()

    if not fechas_validas.empty:
        min_f, max_f = fechas_validas.min().date(), fechas_validas.max().date()
        
        # Configurar por defecto: Últimos 7 días del lote
        default_inicio = max(min_f, max_f - timedelta(days=7))
        st.session_state.setdefault("rango_fechas_lote", [default_inicio, max_f])

        with c_f3:
            st.markdown("**Acceso Rápido por Días:**")
            b_c1, b_c2, b_c3 = st.columns(3)
            with b_c1:
                if st.button("7 días", key="btn_7_lote", use_container_width=True):
                    st.session_state["rango_fechas_lote"] = [max(min_f, max_f - timedelta(days=7)), max_f]
                    st.rerun()
            with b_c2:
                if st.button("15 días", key="btn_15_lote", use_container_width=True):
                    st.session_state["rango_fechas_lote"] = [max(min_f, max_f - timedelta(days=15)), max_f]
                    st.rerun()
            with b_c3:
                if st.button("30 días", key="btn_30_lote", use_container_width=True):
                    st.session_state["rango_fechas_lote"] = [max(min_f, max_f - timedelta(days=30)), max_f]
                    st.rerun()

            rango_fechas = st.date_input(
                "Rango de Fechas a Consultar:",
                key="rango_fechas_lote",
                min_value=min_f,
                max_value=max_f,
            )
    else:
        rango_fechas = []

    st.divider()

    # --- MAPEO SEGURO Y EXACTO DE COLUMNAS DE EXCEL ---
    cols_excel = list(df_lote.columns)

    def buscar_col_exacta_o_patron(patrones, omitir=[]):
        for c in cols_excel:
            c_low = c.lower()
            if any(p.lower() in c_low for p in patrones):
                if not any(o.lower() in c_low for o in omitir):
                    return c
        return None

    def obtener_col_siguiente(col_ref):
        if col_ref in cols_excel:
            idx = cols_excel.index(col_ref)
            if idx + 1 < len(cols_excel):
                return cols_excel[idx + 1]
        return None

    c_fecha = col_fecha_origen
    c_edad = buscar_col_exacta_o_patron(["edad", "sem"])
    c_mort = buscar_col_exacta_o_patron(["mort"])
    c_trasl_ventas = buscar_col_exacta_o_patron(["trasl ventas", "ventas"], omitir=["comentario", "fac", "obs"])
    c_saldo_aves = buscar_col_exacta_o_patron(["saldo aves"])

    # BLOQUE ALIMENTO
    c_obs_alim = buscar_col_exacta_o_patron(["observaciones alimento", "obs alimento"])
    c_costo_alim = buscar_col_exacta_o_patron(["costo alimento"])
    c_ingreso_b = buscar_col_exacta_o_patron(["ingreso b x 40"])
    
    # Comentario / Factura de Ingreso de Alimento
    c_com_ing_alim = buscar_col_exacta_o_patron(["comentario_ingreso_alimento", "comentario_ingreso_aliment"]) or obtener_col_siguiente(c_ingreso_b)
    
    c_consumo_b = buscar_col_exacta_o_patron(["consumo b x 40"])
    c_traslado_b = buscar_col_exacta_o_patron(["traslado b x 40"])
    c_saldo_b = buscar_col_exacta_o_patron(["saldo b x 40"])

    # BLOQUE HUEVOS
    c_prod_huevo = buscar_col_exacta_o_patron(["producción huevos", "prod huevos"])
    c_salida_huevo = buscar_col_exacta_o_patron(["salida huevos"], omitir=["comentario", "fac", "obs"])
    
    # Comentario / Factura de Salida de Huevos
    c_com_sal_huevo = buscar_col_exacta_o_patron(["comentario_salida_huevo", "comentario_salida"]) or obtener_col_siguiente(c_salida_huevo)
    
    c_saldo_huevo = buscar_col_exacta_o_patron(["saldo de huevo", "saldo huevos"])

    # BLOQUE BANDEJAS
    c_ing_band = buscar_col_exacta_o_patron(["ingreso"], omitir=["b x 40", "comentario", "fac", "aliment", "obs"])
    
    # Comentario / Factura Entrada de Bandeja
    c_com_ent_band = buscar_col_exacta_o_patron(["comentario_entrada_bandeja", "comentario_entrada"]) or obtener_col_siguiente(c_ing_band)
    
    c_cons_band = buscar_col_exacta_o_patron(["consumo"], omitir=["b x 40", "comentario", "fac", "obs"])
    c_tras_band = buscar_col_exacta_o_patron(["traslado", "traslados"], omitir=["b x 40", "comentario", "ventas", "fac", "obs"])
    
    c_com_tras_ventas = buscar_col_exacta_o_patron(["comentario_trasl_ventas", "comentario_trasl"]) or obtener_col_siguiente(c_tras_band)
    c_sal_band = buscar_col_exacta_o_patron(["saldo"], omitir=["b x 40", "aves", "huevo"])

    # Lista de columnas que son estrictamente TEXTO (PROTEGIDAS CONTRA CONVERSIÓN A FLOAT)
    cols_texto_relacion = [
        c_fecha, c_edad, c_obs_alim, c_com_ing_alim, c_com_sal_huevo, c_com_ent_band, c_com_tras_ventas
    ]
    cols_texto_relacion = [c for c in cols_texto_relacion if c is not None]

    # --- LIMPIEZA NUMÉRICA PROTEGIENDO COLUMNAS DE TEXTO ---
    cols_a_limpiar = [
        "Mort.",
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
        if col not in cols_texto_relacion and "comentario" not in col.lower() and "obs" not in col.lower() and "fac" not in col.lower():
            if any(c_key.lower() in col.lower() for c_key in cols_a_limpiar):
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

    # Re-filtrar el DataFrame del Lote
    df_lote = df_base[(df_base[nombre_col_granja].astype(str) == granja_sel) & (df_base[nombre_col_lote].astype(str) == lote_sel)]

    # ESTRUCTURA DEFINITIVA DE LAS 22 COLUMNAS
    columnas_totales_def = [
        ("Fecha", c_fecha, "info", "text"),
        ("Edad Sem + Días", c_edad, "info", "text"),
        ("Mort.", c_mort, "aves", "int"),
        ("Trasl Ventas", c_trasl_ventas, "aves", "int"),
        ("Saldo Aves", c_saldo_aves, "aves", "int"),
        ("Observaciones Alimento", c_obs_alim, "alimento", "text"),
        ("Costo Alimento", c_costo_alim, "alimento", "currency"),
        ("Ingreso B X 40 K", c_ingreso_b, "alimento", "float"),
        ("Fac_Ingreso_Alimento", c_com_ing_alim, "alimento", "text"),
        ("Consumo B X 40 K", c_consumo_b, "alimento", "float"),
        ("Traslado B X 40 K", c_traslado_b, "alimento", "float"),
        ("Saldo B X 40 K", c_saldo_b, "alimento", "float"),
        ("Producción Huevos Día", c_prod_huevo, "huevos", "int"),
        ("Salida Huevos dia", c_salida_huevo, "huevos", "int"),
        ("Fac_Salida_Huevo", c_com_sal_huevo, "huevos", "text"),
        ("Saldo de Huevos", c_saldo_huevo, "huevos", "int"),
        ("Ingreso", c_ing_band, "bandejas", "int"),
        ("Fac_Entrada_Bandeja", c_com_ent_band, "bandejas", "text"),
        ("Consumo", c_cons_band, "bandejas", "int"),
        ("Traslados", c_tras_band, "bandejas", "int"),
        ("Fac_Trasl_Ventas", c_com_tras_ventas, "bandejas", "text"),
        ("Saldo", c_sal_band, "bandejas", "int"),
    ]

    # --- SELECCIONADOR DINÁMICO DE COLUMNAS ---
    with st.expander("👁️ **Personalizar Columnas a Mostrar**", expanded=False):
        nombres_todas_cols = [n for n, _, _, _ in columnas_totales_def]
        cols_seleccionadas_nombres = st.multiselect(
            "Selecciona o remueve las columnas que deseas consultar:",
            options=nombres_todas_cols,
            default=nombres_todas_cols,
        )

    columnas_ordenadas = [item for item in columnas_totales_def if item[0] in cols_seleccionadas_nombres]

    # --- FILTRAR DÍA A DÍA ---
    df_lote_filtrado = df_lote.copy()

    if len(rango_fechas) == 2:
        df_lote_filtrado = df_lote_filtrado[
            (df_lote_filtrado["Fecha_dt"].dt.date >= rango_fechas[0])
            & (df_lote_filtrado["Fecha_dt"].dt.date <= rango_fechas[1])
        ]
    elif len(rango_fechas) == 1:
        df_lote_filtrado = df_lote_filtrado[
            df_lote_filtrado["Fecha_dt"].dt.date == rango_fechas[0]
        ]

    def unir_textos(series):
        textos_validos = []
        for x in series.dropna():
            s = str(x).strip()
            if s and s.lower() not in ["nan", "none", "0.0", "0", "null", ""]:
                textos_validos.append(s)
        return " | ".join(dict.fromkeys(textos_validos)) if textos_validos else ""

    if not df_lote_filtrado.empty and columnas_ordenadas:
        mapa_agg = {}
        for nombre_final, col_orig, _, tipo_dato in columnas_ordenadas:
            if col_orig and col_orig in df_lote_filtrado.columns:
                if tipo_dato == "text":
                    mapa_agg[col_orig] = unir_textos
                else:
                    mapa_agg[col_orig] = "sum"

        df_diario = (
            df_lote_filtrado.groupby("Fecha_dt", as_index=False)
            .agg(mapa_agg)
            .sort_values("Fecha_dt")
        )

        df_diario["Fecha / Concepto"] = df_diario["Fecha_dt"].dt.strftime("%d/%m/%Y")

        dict_renombrar = {col_orig: nombre_final for nombre_final, col_orig, _, _ in columnas_ordenadas if col_orig}
        df_diario = df_diario.rename(columns=dict_renombrar)

        # Fila TOTALES
        fila_totales = {"Fecha / Concepto": "TOTALES / ACUMULADO"}
        for nombre_final, _, _, tipo_dato in columnas_ordenadas:
            if nombre_final in df_diario.columns:
                if tipo_dato == "text":
                    fila_totales[nombre_final] = "-"
                elif "saldo" in nombre_final.lower():
                    fila_totales[nombre_final] = df_diario[nombre_final].iloc[-1]
                else:
                    fila_totales[nombre_final] = df_diario[nombre_final].sum()

        df_display = pd.concat([df_diario, pd.DataFrame([fila_totales])], ignore_index=True)

        st.subheader(f"📅 Histórico Diario: Granja {granja_sel} - Lote: {lote_sel}")

        # HTML / CSS RESPONSIVO Y SCROLL
        html_code = """
        <style>
            .scroll-table-container {
                width: 100%;
                max-height: 460px;
                overflow-y: auto;
                overflow-x: auto;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                font-family: system-ui, -apple-system, sans-serif;
            }
            .custom-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 11px;
            }
            .custom-table th {
                position: sticky;
                top: 0;
                z-index: 2;
                padding: 6px 4px;
                text-align: center;
                font-weight: 700;
                word-wrap: break-word;
                line-height: 1.1;
                border: 1px solid #cbd5e1;
                font-size: 10px;
                box-shadow: 0 1px 2px rgba(0,0,0,0.05);
            }
            .custom-table td {
                padding: 5px 4px;
                border: 1px solid #e2e8f0;
                white-space: nowrap;
                font-size: 10.5px;
            }
            .th-info { background-color: #f8fafc; color: #334155; }
            .th-aves { background-color: #dbeafe; color: #1e40af; }
            .th-alimento { background-color: #fef3c7; color: #92400e; }
            .th-huevos { background-color: #d1fae5; color: #065f46; }
            .th-bandejas { background-color: #ffe4e6; color: #9f1239; }
            
            .td-info { background-color: #ffffff; text-align: left; }
            .td-aves { background-color: #ebf3fe; text-align: right; }
            .td-alimento { background-color: #fef8ea; text-align: right; }
            .td-huevos { background-color: #eaf8f0; text-align: right; }
            .td-bandejas { background-color: #fdf0ed; text-align: right; }
            .td-text { text-align: left !important; font-style: italic; color: #334155; font-weight: 500; }
            
            .row-total td {
                position: sticky;
                bottom: 0;
                z-index: 2;
                font-weight: bold;
                background-color: #cbd5e1 !important;
                border-top: 2px solid #94a3b8;
            }
        </style>
        <div class="scroll-table-container">
        <table class="custom-table">
            <thead>
                <tr>
        """

        for nombre_final, _, bloque, _ in columnas_ordenadas:
            html_code += f'<th class="th-{bloque}">{nombre_final}</th>'

        html_code += "</tr></thead><tbody>"

        for _, row in df_display.iterrows():
            concepto = str(row.get("Fecha / Concepto", row.get("Fecha", "")))
            is_total = "TOTALES" in concepto
            row_class = "row-total" if is_total else ""

            html_code += f'<tr class="{row_class}">'

            for nombre_final, _, bloque, tipo_dato in columnas_ordenadas:
                val = row.get(nombre_final, "")
                
                if tipo_dato == "text":
                    val_str = str(val) if pd.notna(val) and str(val).lower() not in ["nan", "none", "0.0", "0", ""] else ""
                    html_code += f'<td class="td-{bloque} td-text" title="{val_str}">{val_str}</td>'
                else:
                    val_num = pd.to_numeric(val, errors="coerce")
                    val_num = 0 if pd.isna(val_num) else val_num

                    if tipo_dato == "int":
                        val_str = f"{int(round(val_num)):,}"
                    elif tipo_dato == "currency":
                        val_str = f"$ {int(round(val_num)):,}"
                    else:
                        val_str = f"{val_num:,.2f}"

                    html_code += f'<td class="td-{bloque}">{val_str}</td>'

            html_code += "</tr>"

        html_code += "</tbody></table></div>"

        st.html(html_code)

        # Botón de Descarga Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            cols_export = [n for n, _, _, _ in columnas_ordenadas if n in df_display.columns]
            df_display[cols_export].to_excel(
                writer, index=False, sheet_name=f"Lote_{lote_sel}"
            )
        buffer.seek(0)

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label=f"📥 Descargar Histórico de Lote {lote_sel} (.xlsx)",
            data=buffer,
            file_name=f"HISTORICO_LOTE_{lote_sel}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif not columnas_ordenadas:
        st.warning("Selecciona al menos una columna para visualizar.")
    else:
        st.info("No hay registros para este lote en el rango de fechas seleccionado.")
else:
    st.warning(
        f"⚠️ No se encontró el archivo consolidado en `{RUTA_REPORTE_LOCAL}`."
    )
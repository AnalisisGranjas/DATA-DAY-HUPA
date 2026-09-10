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

# --- ESTILOS CSS UNIFICADOS Y REAJUSTADOS ---
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
    
    /* Estilo de Tarjetas KPI centradas con borde */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 12px 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    div[data-testid="stMetric"] label {
        justify-content: center !important;
        font-weight: 600;
        color: #475569;
    }
    
    .scroll-table-container {
        width: 100%;
        max-height: 460px;
        overflow-y: auto;
        overflow-x: auto;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        font-family: system-ui, -apple-system, sans-serif;
    }
    .custom-table { width: max-content; min-width: 100%; border-collapse: collapse; table-layout: auto; }
    
    /* ENCABEZADOS ESTRECHOS Y ALTOS (MULTILÍNEA) */
    .custom-table th {
        position: sticky; top: 0; z-index: 2; padding: 5px 3px; text-align: center !important; font-weight: 700;
        white-space: normal !important; word-break: break-word !important; word-wrap: break-word !important; 
        line-height: 1.15; border: 1px solid #cbd5e1; font-size: 10px !important; 
        max-width: 65px !important; min-width: 55px !important; box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    
    /* Celdas con tamaño legible de 12px */
    .custom-table td { 
        padding: 6px 6px; text-align: center !important; border: 1px solid #e2e8f0; 
        white-space: nowrap !important; font-size: 12px !important; width: auto; 
        background-color: #ffffff; color: #1e293b;
    }
    .custom-table tr:nth-child(even) td { background-color: #f8fafc; }
    
    /* COLORES DE ENCABEZADOS POR BLOQUE */
    .th-info { background-color: #f8fafc; color: #334155; text-align: left !important; min-width: 120px !important; }
    .th-aves { background-color: #dbeafe; color: #1e40af; }
    .th-alimento { background-color: #fef3c7; color: #92400e; }
    .th-costo { background-color: #fef3c7; color: #92400e; }
    .th-huevos { background-color: #d1fae5; color: #065f46; }
    .th-bandejas { background-color: #ffe4e6; color: #9f1239; }
    
    .td-text { text-align: left !important; font-style: italic; color: #334155; font-weight: 500; }
    
    .row-total td { 
        position: sticky; bottom: 0; z-index: 2; font-weight: bold; 
        background-color: #cbd5e1 !important; border-top: 2px solid #94a3b8; 
        font-size: 12.5px !important; color: #0f172a !important; 
    }
    
    .val-pos { color: #15803d !important; font-weight: bold; }
    .val-neg { color: #b91c1c !important; font-weight: bold; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Consulta del Histórico Día a Día por Lote")
st.markdown(
    "Audita la evolución detallada día a día con la estructura completa de **columnas de datos, desviaciones de tabla y facturas**."
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

    DIAS_ESPANOL = {
        "Monday": "Lun",
        "Tuesday": "Mar",
        "Wednesday": "Mié",
        "Thursday": "Jue",
        "Friday": "Vie",
        "Saturday": "Sáb",
        "Sunday": "Dom",
    }

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
    st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 4px; color: #1e293b;'>🎯 Selección de Lote y Rango de Fechas</p>", unsafe_allow_html=True)
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
        max_registrada = fechas_validas.max().date()
        max_f = min(max_registrada, datetime.now().date() - timedelta(days=1))
        min_f = fechas_validas.min().date()

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

    # --- MAPEO DE COLUMNAS DE EXCEL ---
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

    c_cons_gr_ave = buscar_col_exacta_o_patron(["Consumo Gr. A. D."])
    c_gr_ave_tabla = buscar_col_exacta_o_patron(["Gr. A. D. Tabla"])

    c_obs_alim = buscar_col_exacta_o_patron(["observaciones alimento", "obs alimento"])
    c_costo_alim = buscar_col_exacta_o_patron(["costo alimento"])
    c_ingreso_b = buscar_col_exacta_o_patron(["ingreso b x 40"])
    c_com_ing_alim = buscar_col_exacta_o_patron(["comentario_ingreso_alimento", "comentario_ingreso_aliment"]) or obtener_col_siguiente(c_ingreso_b)
    
    c_consumo_b = buscar_col_exacta_o_patron(["consumo b x 40"])
    c_traslado_b = buscar_col_exacta_o_patron(["traslado b x 40"])
    c_saldo_b = buscar_col_exacta_o_patron(["saldo b x 40"])

    c_pct_prod_dia = buscar_col_exacta_o_patron(["% Diario de Prod."])
    c_pct_prod_tabla = buscar_col_exacta_o_patron(["% Dia Prod. Tab"])

    c_prod_huevo = buscar_col_exacta_o_patron(["producción huevos", "prod huevos"])
    c_salida_huevo = buscar_col_exacta_o_patron(["salida huevos"], omitir=["comentario", "fac", "obs"])
    c_com_sal_huevo = buscar_col_exacta_o_patron(["comentario_salida_huevo", "comentario_salida"]) or obtener_col_siguiente(c_salida_huevo)
    c_saldo_huevo = buscar_col_exacta_o_patron(["saldo de huevo", "saldo huevos"])

    c_ing_band = buscar_col_exacta_o_patron(["ingreso"], omitir=["b x 40", "comentario", "fac", "aliment", "obs"])
    c_com_ent_band = buscar_col_exacta_o_patron(["comentario_entrada_bandeja", "comentario_entrada"]) or obtener_col_siguiente(c_ing_band)
    c_cons_band = buscar_col_exacta_o_patron(["consumo"], omitir=["b x 40", "comentario", "fac", "obs"])
    c_tras_band = buscar_col_exacta_o_patron(["traslado", "traslados"], omitir=["b x 40", "comentario", "ventas", "fac", "obs"])
    c_com_tras_ventas = buscar_col_exacta_o_patron(["comentario_trasl_ventas", "comentario_trasl"]) or obtener_col_siguiente(c_tras_band)
    c_sal_band = buscar_col_exacta_o_patron(["saldo"], omitir=["b x 40", "aves", "huevo"])

    cols_texto_relacion = [
        c_fecha, c_edad, c_obs_alim, c_com_ing_alim, c_com_sal_huevo, c_com_ent_band, c_com_tras_ventas
    ]
    cols_texto_relacion = [c for c in cols_texto_relacion if c is not None]

    cols_a_limpiar = [
        "Mort.", "Trasl Ventas", "Saldo Aves", "Costo Alimento", "Ingreso B X 40 K",
        "Consumo B X 40 K", "Traslado B X 40 K", "Saldo B X 40 K", "Producción Huevos Día",
        "Salida Huevos dia", "Saldo de Huevo", "Saldo de Huevos", "Ingreso", "Consumo",
        "Traslado", "Traslados", "Saldo", "Consumo Gr. A. D.", "Gr. A. D. Tabla",
        "% Diario de Prod.", "% Dia Prod. Tab"
    ]

    for col in df_base.columns:
        if col not in cols_texto_relacion and "comentario" not in col.lower() and "obs" not in col.lower() and "fac" not in col.lower():
            if any(c_key.lower() in col.lower() for c_key in cols_a_limpiar):
                if df_base[col].dtype == "object":
                    df_base[col] = (
                        df_base[col]
                        .astype(str)
                        .str.replace("$", "", regex=False)
                        .str.replace("%", "", regex=False)
                        .str.replace(" ", "", regex=False)
                        .str.replace(",", ".", regex=False)
                        .str.strip()
                    )
                df_base[col] = pd.to_numeric(
                    df_base[col], errors="coerce"
                ).fillna(0)

    df_lote = df_base[(df_base[nombre_col_granja].astype(str) == granja_sel) & (df_base[nombre_col_lote].astype(str) == lote_sel)]

    # ESTRUCTURA COMPLETA DE COLUMNAS CON COMPARATIVOS MULTILÍNEA
    columnas_totales_def = [
        ("Fecha", c_fecha, "info", "text"),
        ("Edad Sem<br>+ Días", c_edad, "info", "text"),
        ("Mort.", c_mort, "aves", "int"),
        ("Trasl<br>Ventas", c_trasl_ventas, "aves", "int"),
        ("Saldo<br>Aves", c_saldo_aves, "aves", "int"),
        ("Consumo<br>Gr. A. D.", c_cons_gr_ave, "alimento", "avg_float"),
        ("Gr. A. D.<br>Tabla", c_gr_ave_tabla, "alimento", "avg_float"),
        ("Dif. Cons.<br>(g)", None, "alimento", "diff_gramos"),
        ("Observaciones<br>Alimento", c_obs_alim, "alimento", "text"),
        ("Costo<br>Alimento", c_costo_alim, "alimento", "currency"),
        ("Ingreso B<br>X 40 K", c_ingreso_b, "alimento", "float"),
        ("Fac_Ingreso<br>Alimento", c_com_ing_alim, "alimento", "text"),
        ("Consumo B<br>X 40 K", c_consumo_b, "alimento", "float"),
        ("Traslado B<br>X 40 K", c_traslado_b, "alimento", "float"),
        ("Saldo B<br>X 40 K", c_saldo_b, "alimento", "float"),
        ("Producción<br>Huevos Día", c_prod_huevo, "huevos", "int"),
        ("% Diario<br>de Prod.", c_pct_prod_dia, "huevos", "avg_pct"),
        ("% Dia<br>Prod. Tab", c_pct_prod_tabla, "huevos", "avg_pct"),
        ("Dif. %<br>Prod", None, "huevos", "diff_pct"),
        ("Salida<br>Huevos dia", c_salida_huevo, "huevos", "int"),
        ("Fac_Salida<br>Huevo", c_com_sal_huevo, "huevos", "text"),
        ("Saldo de<br>Huevos", c_saldo_huevo, "huevos", "int"),
        ("Costo Huevo/<br>Alimento", None, "huevos", "currency_dec"),
        ("Ingreso", c_ing_band, "bandejas", "int"),
        ("Fac_Entrada<br>Bandeja", c_com_ent_band, "bandejas", "text"),
        ("Consumo", c_cons_band, "bandejas", "int"),
        ("Traslados", c_tras_band, "bandejas", "int"),
        ("Fac_Trasl<br>Ventas", c_com_tras_ventas, "bandejas", "text"),
        ("Saldo", c_sal_band, "bandejas", "int"),
    ]

    # --- SELECCIONADOR DINÁMICO DE COLUMNAS ---
    with st.expander("👁️ **Personalizar Columnas a Mostrar / Ocultar**", expanded=False):
        nombres_todas_cols = [n for n, _, _, _ in columnas_totales_def]
        cols_seleccionadas_nombres = st.multiselect(
            "Selecciona o desmarca las columnas para la vista y exportación:",
            options=nombres_todas_cols,
            default=nombres_todas_cols,
            key="multiselect_cols_lote",
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

        # --- SECCIÓN KPI COMPARATIVO SUPERIOR ---
        if len(rango_fechas) == 2:
            f_start, f_end = rango_fechas[0], rango_fechas[1]
            num_dias = (f_end - f_start).days + 1
            f_start_prev = f_start - timedelta(days=num_dias)
            f_end_prev = f_start - timedelta(days=1)

            df_curr_kpi = df_lote[(df_lote["Fecha_dt"].dt.date >= f_start) & (df_lote["Fecha_dt"].dt.date <= f_end)]
            df_prev_kpi = df_lote[(df_lote["Fecha_dt"].dt.date >= f_start_prev) & (df_lote["Fecha_dt"].dt.date <= f_end_prev)]

            c_costo_cur = df_curr_kpi[c_costo_alim].sum() if c_costo_alim in df_curr_kpi.columns else 0
            c_prod_cur = df_curr_kpi[c_prod_huevo].sum() if c_prod_huevo in df_curr_kpi.columns else 0
            c_mort_cur = df_curr_kpi[c_mort].sum() if c_mort in df_curr_kpi.columns else 0
            c_cons_cur = df_curr_kpi[c_consumo_b].sum() if c_consumo_b in df_curr_kpi.columns else 0
            costo_huevo_cur = (c_costo_cur / c_prod_cur) if c_prod_cur > 0 else 0.0

            c_costo_prev = df_prev_kpi[c_costo_alim].sum() if c_costo_alim in df_prev_kpi.columns else 0
            c_prod_prev = df_prev_kpi[c_prod_huevo].sum() if c_prod_huevo in df_prev_kpi.columns else 0
            c_mort_prev = df_prev_kpi[c_mort].sum() if c_mort in df_prev_kpi.columns else 0
            c_cons_prev = df_prev_kpi[c_consumo_b].sum() if c_consumo_b in df_prev_kpi.columns else 0
            costo_huevo_prev = (c_costo_prev / c_prod_prev) if c_prod_prev > 0 else 0.0

            delta_costo_huevo = costo_huevo_cur - costo_huevo_prev
            delta_prod = c_prod_cur - c_prod_prev
            delta_mort = c_mort_cur - c_mort_prev
            delta_cons = c_cons_cur - c_cons_prev

            st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 2px; color: #1e293b;'>📈 Resumen Gerencial de Lote y Deltas Comparativos</p>", unsafe_allow_html=True)
            st.caption(f"Comparando el lote {lote_sel} en el periodo seleccionado ({f_start.strftime('%d/%m')} - {f_end.strftime('%d/%m')}) vs. el periodo anterior equivalente ({f_start_prev.strftime('%d/%m')} - {f_end_prev.strftime('%d/%m')}).")

            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            with kpi1:
                st.metric(
                    label="Costo Huevo/Alimento",
                    value=f"$ {costo_huevo_cur:,.2f}",
                    delta=f"{delta_costo_huevo:+,.2f} $/huevo",
                    delta_color="inverse",
                    help="Costo de alimento requerido para producir 1 huevo en este lote."
                )
            with kpi2:
                st.metric(
                    label="Producción Huevos",
                    value=f"{int(round(c_prod_cur)):,} u.",
                    delta=f"{int(round(delta_prod)):,}",
                    help="Suma total de unidades de huevo recolectadas en el rango seleccionado."
                )
            with kpi3:
                st.metric(
                    label="Mortalidad Aves",
                    value=f"{int(round(c_mort_cur)):,} aves",
                    delta=f"{int(round(delta_mort)):,}",
                    delta_color="inverse",
                    help="Cantidad acumulada de bajas de aves en este lote."
                )
            with kpi4:
                st.metric(
                    label="Consumo Alimento",
                    value=f"{c_cons_cur:,.1f} bultos",
                    delta=f"{delta_cons:+,.1f}",
                    help="Bultos de alimento (40 kg) consumidos por el lote."
                )
            st.divider()

        mapa_agg = {}
        for nombre_final, col_orig, _, tipo_dato in columnas_totales_def:
            if col_orig and col_orig in df_lote_filtrado.columns:
                if tipo_dato == "text":
                    mapa_agg[col_orig] = unir_textos
                elif "avg" in tipo_dato:
                    mapa_agg[col_orig] = "mean"
                else:
                    mapa_agg[col_orig] = "sum"

        df_diario = (
            df_lote_filtrado.groupby("Fecha_dt", as_index=False)
            .agg(mapa_agg)
            .sort_values("Fecha_dt")
        )

        # FECHA CON DÍA DE LA SEMANA
        df_diario["Fecha / Concepto"] = df_diario["Fecha_dt"].apply(
            lambda f: f"{DIAS_ESPANOL.get(f.strftime('%A'), '')} {f.strftime('%d/%m/%Y')}"
        )

        dict_renombrar = {col_orig: nombre_final for nombre_final, col_orig, _, _ in columnas_totales_def if col_orig}
        df_diario = df_diario.rename(columns=dict_renombrar)

        # CÁLCULOS COMPARATIVOS POR DÍA
        if "Consumo<br>Gr. A. D." in df_diario.columns and "Gr. A. D.<br>Tabla" in df_diario.columns:
            df_diario["Dif. Cons.<br>(g)"] = df_diario["Consumo<br>Gr. A. D."] - df_diario["Gr. A. D.<br>Tabla"]
        
        if "% Diario<br>de Prod." in df_diario.columns and "% Dia<br>Prod. Tab" in df_diario.columns:
            df_diario["Dif. %<br>Prod"] = df_diario["% Diario<br>de Prod."] - df_diario["% Dia<br>Prod. Tab"]

        # CÁLCULO COSTO HUEVO/ALIMENTO POR DÍA
        val_costo = df_diario["Costo<br>Alimento"] if "Costo<br>Alimento" in df_diario.columns else 0
        val_prod = df_diario["Producción<br>Huevos Día"] if "Producción<br>Huevos Día" in df_diario.columns else 0
        df_diario["Costo Huevo/<br>Alimento"] = (val_costo / val_prod).fillna(0).replace([float('inf'), -float('inf')], 0)

        # Fila TOTALES / ACUMULADO EXACTA
        fila_totales = {"Fecha / Concepto": "TOTALES / ACUMULADO"}
        for nombre_final, col_orig, _, tipo_dato in columnas_totales_def:
            if nombre_final in df_diario.columns:
                if tipo_dato == "text":
                    fila_totales[nombre_final] = "-"
                elif nombre_final in ["Costo Huevo/<br>Alimento", "Dif. Cons.<br>(g)", "Dif. %<br>Prod"]:
                    continue
                elif "avg" in tipo_dato:
                    if col_orig and col_orig in df_lote_filtrado.columns:
                        fila_totales[nombre_final] = df_lote_filtrado[col_orig].mean()
                    else:
                        fila_totales[nombre_final] = df_diario[nombre_final].mean()
                elif "saldo" in nombre_final.lower():
                    # Saldo exacto del último día del rango
                    max_f_lote = df_lote_filtrado["Fecha_dt"].max()
                    df_ult_lote = df_lote_filtrado[df_lote_filtrado["Fecha_dt"] == max_f_lote]
                    if col_orig and col_orig in df_ult_lote.columns:
                        fila_totales[nombre_final] = df_ult_lote[col_orig].sum()
                    else:
                        fila_totales[nombre_final] = df_diario[nombre_final].iloc[-1]
                else:
                    fila_totales[nombre_final] = df_diario[nombre_final].sum()

        tot_costo = fila_totales.get("Costo<br>Alimento", 0)
        tot_prod = fila_totales.get("Producción<br>Huevos Día", 0)
        fila_totales["Costo Huevo/<br>Alimento"] = (tot_costo / tot_prod) if tot_prod > 0 else 0.0

        cons_r = fila_totales.get("Consumo<br>Gr. A. D.", 0)
        cons_t = fila_totales.get("Gr. A. D.<br>Tabla", 0)
        fila_totales["Dif. Cons.<br>(g)"] = cons_r - cons_t

        pct_r = fila_totales.get("% Diario<br>de Prod.", 0)
        pct_t = fila_totales.get("% Dia<br>Prod. Tab", 0)
        fila_totales["Dif. %<br>Prod"] = pct_r - pct_t

        df_display = pd.concat([df_diario, pd.DataFrame([fila_totales])], ignore_index=True)

        st.markdown(f"<p style='font-size: 14px; font-weight: bold; margin-bottom: 4px; color: #1e293b;'>📅 Histórico Diario: Granja {granja_sel} - Lote: {lote_sel}</p>", unsafe_allow_html=True)

        # RENDER HTML
        html_code = """
        <div class="scroll-table-container">
        <table class="custom-table">
            <thead>
                <tr>
        """

        for nombre_final, _, bloque, tipo_dato in columnas_ordenadas:
            clase_th = "th-costo" if "currency" in tipo_dato else f"th-{bloque}"
            html_code += f'<th class="{clase_th}">{nombre_final}</th>'

        html_code += "</tr></thead><tbody>"

        for _, row in df_display.iterrows():
            concepto = str(row.get("Fecha / Concepto", row.get("Fecha", "")))
            is_total = "TOTALES" in concepto
            prod_val = row.get("Producción<br>Huevos Día", 0)
            
            es_sin_datos = (not is_total) and (pd.to_numeric(prod_val, errors="coerce") == 0)
            row_class = "row-total" if is_total else ""

            html_code += f'<tr class="{row_class}">'

            for nombre_final, _, bloque, tipo_dato in columnas_ordenadas:
                val = row.get(nombre_final, "")
                
                if es_sin_datos and ("diff" in tipo_dato or "avg" in tipo_dato):
                    html_code += '<td>-</td>'
                    continue

                if tipo_dato == "text":
                    val_str = str(val) if pd.notna(val) and str(val).lower() not in ["nan", "none", "0.0", "0", ""] else ""
                    html_code += f'<td class="td-text" title="{val_str}">{val_str}</td>'
                else:
                    val_num = pd.to_numeric(val, errors="coerce")
                    val_num = 0 if pd.isna(val_num) else val_num

                    if tipo_dato == "int":
                        val_str = f"{int(round(val_num)):,}"
                        html_code += f'<td>{val_str}</td>'
                    elif tipo_dato == "currency":
                        val_str = f"$ {int(round(val_num)):,}"
                        html_code += f'<td>{val_str}</td>'
                    elif tipo_dato == "currency_dec":
                        val_str = f"$ {val_num:,.2f}"
                        html_code += f'<td>{val_str}</td>'
                    elif tipo_dato == "avg_pct":
                        val_str = f"{val_num:.2f}%"
                        html_code += f'<td>{val_str}</td>'
                    elif tipo_dato == "avg_float":
                        val_str = f"{val_num:.2f} g"
                        html_code += f'<td>{val_str}</td>'
                    elif tipo_dato == "diff_pct":
                        val_str = f"{val_num:+.2f}%"
                        cls_color = "val-neg" if val_num < 0 else "val-pos"
                        html_code += f'<td class="{cls_color}">{val_str}</td>'
                    elif tipo_dato == "diff_gramos":
                        val_str = f"{val_num:+.2f} g"
                        cls_color = "val-neg" if val_num > 5 else "val-pos"
                        html_code += f'<td class="{cls_color}">{val_str}</td>'
                    else:
                        val_str = f"{val_num:,.2f}"
                        html_code += f'<td>{val_str}</td>'

            html_code += "</tr>"

        html_code += "</tbody></table></div>"

        st.html(html_code)

        # Botón de Descarga Excel (Limpia <br> en los encabezados exportados)
        buffer = io.BytesIO()
        cols_export = [n for n, _, _, _ in columnas_ordenadas if n in df_display.columns]
        df_export = df_display[cols_export].rename(columns=lambda c: str(c).replace("<br>", " "))
        
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_export.to_excel(
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
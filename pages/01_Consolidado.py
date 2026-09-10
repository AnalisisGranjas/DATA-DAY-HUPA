import io
import os
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# --- VALIDACIÓN DE SESIÓN ---
if not st.session_state.get("authenticated", False):
    st.warning("⚠️ Debes iniciar sesión para acceder a este reporte.")
    st.stop()

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Consolidado General - Avícola",
    page_icon="🐔",
    layout="wide",
)

# --- LOGO EN BARRA LATERAL ---
ruta_logo = os.path.join("DATA", "logo hupa.png")
if os.path.exists(ruta_logo):
    st.sidebar.image(ruta_logo, use_container_width=True)
    st.sidebar.divider()

# --- ESTILOS CSS REAJUSTADOS ---
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
        background-color: #ffffff !important;
    }
    
    div[data-testid="stSubheader"] h3 {
        font-size: 14px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin-bottom: 2px !important;
    }
    
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
    
    .scroll-table-7 { max-height: 340px; overflow-y: auto; overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 8px; }
    .scroll-table-10 { max-height: 460px; overflow-y: auto; overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 8px; }
    
    .custom-table-container { width: 100%; border: 1px solid #cbd5e1; border-radius: 8px; overflow-y: auto; overflow-x: auto; }
    .custom-table { width: max-content; min-width: 100%; border-collapse: collapse; font-family: system-ui, -apple-system, sans-serif; table-layout: auto; }
    
    /* ENCABEZADOS ESTRECHOS Y ALTOS (FORZADOS EN 2 O 3 RENGLONES) */
    .custom-table th {
        position: sticky; top: 0; z-index: 2; padding: 5px 3px; text-align: center !important; font-weight: 700;
        white-space: normal !important; word-break: break-word !important; word-wrap: break-word !important; 
        line-height: 1.15; border: 1px solid #cbd5e1; font-size: 10px !important; 
        max-width: 65px !important; min-width: 55px !important; box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    
    .custom-table td { 
        padding: 6px 6px; text-align: center !important; border: 1px solid #e2e8f0; 
        white-space: nowrap !important; font-size: 12px !important; width: auto; 
        background-color: #ffffff; color: #1e293b;
    }
    .custom-table tr:nth-child(even) td { background-color: #f8fafc; }
    
    .th-fecha { background-color: #f1f5f9; color: #334155; text-align: left !important; min-width: 130px !important; max-width: 180px !important; }
    .th-aves { background-color: #dbeafe; color: #1e40af; }
    .th-alimento { background-color: #fef3c7; color: #92400e; }
    .th-costo { background-color: #fef3c7; color: #92400e; }
    .th-huevos { background-color: #d1fae5; color: #065f46; }
    .th-bandejas { background-color: #ffe4e6; color: #9f1239; }
    
    .row-total td { position: sticky; bottom: 0; z-index: 2; font-weight: bold; background-color: #cbd5e1 !important; border-top: 2px solid #94a3b8; font-size: 12.5px !important; color: #0f172a !important; }
    .td-concept { text-align: left !important; font-weight: 500; }
    
    .row-sin-datos td { background-color: #fee2e2 !important; color: #991b1b !important; }
    .badge-sin-datos { background-color: #ef4444; color: white; padding: 2px 5px; border-radius: 4px; font-size: 10px; font-weight: bold; margin-left: 4px; }
    
    .val-pos { color: #15803d !important; font-weight: bold; }
    .val-neg { color: #b91c1c !important; font-weight: bold; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🐔 Consulta General y Reporte Consolidado por Fecha")
st.markdown("Visualización en 3 cuadros de auditoría con **métricas KPI comparativas**, colores por bloque y títulos compactos.")
st.divider()

RUTA_REPORTE_LOCAL = os.path.join("DATA", "REPORTE_AVITRACK_FINAL.xlsx")


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

    col_fecha_origen = "Fecha" if "Fecha" in df_base.columns else df_base.columns[1]
    df_base["Fecha_dt"] = pd.to_datetime(df_base[col_fecha_origen], dayfirst=True, errors="coerce")

    DIAS_ESPANOL = {
        "Monday": "Lun", "Tuesday": "Mar", "Wednesday": "Mié",
        "Thursday": "Jue", "Friday": "Vie", "Saturday": "Sáb", "Sunday": "Dom"
    }

    col_rs = [c for c in df_base.columns if "razon social" in c.lower() or "empresa" in c.lower()]
    nombre_col_rs = col_rs[0] if col_rs else "Razon Social"

    col_granja_p = [c for c in df_base.columns if "Nombre de Granja (P)" in c]
    nombre_col_granja = col_granja_p[0] if col_granja_p else "Nombre de Granja (L) :"

    col_lote = [c for c in df_base.columns if "Número de Lote" in c or "Lote" in c]
    nombre_col_lote = col_lote[0] if col_lote else "Archivo"

    opciones_rs = sorted([str(x) for x in df_base[nombre_col_rs].dropna().unique() if str(x).strip() != ""])
    opciones_granjas = sorted([str(x) for x in df_base[nombre_col_granja].dropna().unique() if str(x).strip() != ""])
    opciones_lotes = sorted([str(x) for x in df_base[nombre_col_lote].dropna().unique() if str(x).strip() != ""])

    st.session_state.setdefault("sel_rs", [])
    st.session_state.setdefault("sel_granjas", [])
    st.session_state.setdefault("sel_lotes", [])

    fechas_validas = df_base["Fecha_dt"].dropna()
    if not fechas_validas.empty:
        max_registrada = fechas_validas.max().date()
        max_fecha_dt = min(max_registrada, datetime.now().date() - timedelta(days=1))
        min_fecha_dt = fechas_validas.min().date()
        default_inicio = max(min_fecha_dt, max_fecha_dt - timedelta(days=7))

        st.session_state.setdefault("rango_c1", [default_inicio, max_fecha_dt])
        st.session_state.setdefault("rango_c2", [default_inicio, max_fecha_dt])
        st.session_state.setdefault("rango_c3", [default_inicio, max_fecha_dt])

    # BARRA LATERAL
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
    st.sidebar.success(f"**Origen:** `{RUTA_REPORTE_LOCAL}`\n\n**Total Registros:** {len(df_base):,}")

    if st.sidebar.button("🔄 Recargar Datos"):
        st.cache_data.clear()
        st.rerun()

    # LIMPIEZA NUMÉRICA
    cols_a_limpiar = [
        "Mort.", "Otros", "Selec.", "Trasl Ventas", "Saldo Aves", "Costo Alimento",
        "Ingreso B X 40 K", "Consumo B X 40 K", "Traslado B X 40 K", "Ajuste Alimento", "Saldo B X 40 K",
        "Producción Huevos Día", "Salida Huevos dia", "Ajuste Huevo", "Saldo de Huevo", "Saldo de Huevos",
        "Ingreso", "Consumo", "Traslado", "Traslados", "Ajuste Bandeja", "Saldo",
        "Consumo Gr. A. D.", "Gr. A. D. Tabla", "% Diario de Prod.", "% Dia Prod. Tab"
    ]

    for col in df_base.columns:
        if any(c_key.lower() == col.lower() or c_key.lower() in col.lower() for c_key in cols_a_limpiar):
            if df_base[col].dtype == "object":
                df_base[col] = (
                    df_base[col].astype(str).str.replace("$", "", regex=False)
                    .str.replace("%", "", regex=False).str.replace(" ", "", regex=False)
                    .str.replace(",", ".", regex=False).str.strip()
                )
            df_base[col] = pd.to_numeric(df_base[col], errors="coerce").fillna(0)

    def obtener_columna_exacta(patron, indice_bloque=0):
        cols = [c for c in df_base.columns if patron.lower() in c.lower()]
        if cols and len(cols) > indice_bloque:
            return cols[indice_bloque]
        return cols[0] if cols else None

    c_fecha = col_fecha_origen
    c_mort = obtener_columna_exacta("mort")
    c_otros = obtener_columna_exacta("otros")
    c_selec = obtener_columna_exacta("selec")
    c_trasl_ventas = obtener_columna_exacta("trasl ventas") or obtener_columna_exacta("ventas")
    c_saldo_aves = obtener_columna_exacta("saldo aves")

    c_cons_gr_ave = obtener_columna_exacta("Consumo Gr. A. D.")
    c_gr_ave_tabla = obtener_columna_exacta("Gr. A. D. Tabla")

    c_costo_alim = obtener_columna_exacta("costo alimento")
    c_ingreso_b = obtener_columna_exacta("ingreso b x 40")
    c_consumo_b = obtener_columna_exacta("consumo b x 40")
    c_traslado_b = obtener_columna_exacta("traslado b x 40")
    c_saldo_b = obtener_columna_exacta("saldo b x 40")
    c_ajuste_alim = obtener_columna_exacta("ajuste alimento")

    c_pct_prod_dia = obtener_columna_exacta("% Diario de Prod.")
    c_pct_prod_tabla = obtener_columna_exacta("% Dia Prod. Tab")

    c_prod_huevo = obtener_columna_exacta("producción huevos") or obtener_columna_exacta("prod")
    c_salida_huevo = obtener_columna_exacta("salida huevos") or obtener_columna_exacta("salida")
    c_saldo_huevo = obtener_columna_exacta("saldo de huevo")
    c_ajuste_huevo = obtener_columna_exacta("ajuste huevo")

    cols_sin_comentarios = [c for c in df_base.columns if "comentario" not in c.lower()]

    def buscar_en_limpias(patron, omitir=[]):
        c_matches = [c for c in cols_sin_comentarios if patron.lower() in c.lower() and not any(o.lower() in c.lower() for o in omitir)]
        return c_matches[-1] if c_matches else None

    c_ing_band = buscar_en_limpias("ingreso", omitir=["b x 40", "alimento"])
    c_cons_band = buscar_en_limpias("consumo", omitir=["b x 40"])
    c_tras_band = buscar_en_limpias("traslado", omitir=["b x 40", "ventas"])
    c_sal_band = buscar_en_limpias("saldo", omitir=["b x 40", "aves", "huevo"])
    c_ajuste_band = obtener_columna_exacta("ajuste bandeja")

    # ENCABEZADOS CON SALTO DE LÍNEA MULTILÍNEA
    columnas_totales_def = [
        ("Mort.", c_mort, "aves", "int"),
        ("Otros", c_otros, "aves", "int"),
        ("Selec.", c_selec, "aves", "int"),
        ("Trasl<br>Ventas", c_trasl_ventas, "aves", "int"),
        ("Saldo<br>Aves", c_saldo_aves, "aves", "int"),
        ("Consumo<br>Gr. A. D.", c_cons_gr_ave, "alimento", "avg_float"),
        ("Gr. A. D.<br>Tabla", c_gr_ave_tabla, "alimento", "avg_float"),
        ("Dif. Cons.<br>(g)", None, "alimento", "diff_gramos"),
        ("Costo<br>Alimento", c_costo_alim, "alimento", "currency"),
        ("Ingreso B<br>X 40 K", c_ingreso_b, "alimento", "float"),
        ("Consumo B<br>X 40 K", c_consumo_b, "alimento", "float"),
        ("Traslado B<br>X 40 K", c_traslado_b, "alimento", "float"),
        ("Ajuste<br>Alimento", c_ajuste_alim, "alimento", "float"),
        ("Saldo B<br>X 40 K", c_saldo_b, "alimento", "float"),
        ("Producción<br>Huevos Día", c_prod_huevo, "huevos", "int"),
        ("% Diario<br>de Prod.", c_pct_prod_dia, "huevos", "avg_pct"),
        ("% Dia<br>Prod. Tab", c_pct_prod_tabla, "huevos", "avg_pct"),
        ("Dif. %<br>Prod", None, "huevos", "diff_pct"),
        ("Salida<br>Huevos dia", c_salida_huevo, "huevos", "int"),
        ("Ajuste<br>Huevo", c_ajuste_huevo, "huevos", "int"),
        ("Saldo de<br>Huevos", c_saldo_huevo, "huevos", "int"),
        ("Costo Huevo/<br>Alimento", None, "huevos", "currency_dec"),
        ("Ingreso", c_ing_band, "bandejas", "int"),
        ("Consumo", c_cons_band, "bandejas", "int"),
        ("Traslados", c_tras_band, "bandejas", "int"),
        ("Ajuste<br>Bandeja", c_ajuste_band, "bandejas", "int"),
        ("Saldo", c_sal_band, "bandejas", "int"),
    ]

    # PANEL UNIFICADO
    with st.expander("🔍 **Filtros Maestros y Personalización de Columnas**", expanded=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            rs_sel = st.multiselect("Razón Social / Empresa:", options=opciones_rs, key="sel_rs", placeholder="Todas")
        with f_col2:
            granja_sel = st.multiselect("Granja (Producción):", options=opciones_granjas, key="sel_granjas", placeholder="Todas")
        with f_col3:
            lote_sel = st.multiselect("Lote:", options=opciones_lotes, key="sel_lotes", placeholder="Todos")

        st.markdown("---")
        nombres_todas = [n for n, _, _, _ in columnas_totales_def]
        cols_seleccionadas_nombres = st.multiselect(
            "👁️ **Selecciona o desmarca las columnas que deseas consultar en los 3 cuadros:**",
            options=nombres_todas, default=nombres_todas, key="multiselect_cols_consolidado"
        )

    columnas_ordenadas = [item for item in columnas_totales_def if item[0] in cols_seleccionadas_nombres]

    df_base_filtrado = df_base.copy()
    if rs_sel:
        df_base_filtrado = df_base_filtrado[df_base_filtrado[nombre_col_rs].astype(str).isin(rs_sel)]
    if granja_sel:
        df_base_filtrado = df_base_filtrado[df_base_filtrado[nombre_col_granja].astype(str).isin(granja_sel)]
    if lote_sel:
        df_base_filtrado = df_base_filtrado[df_base_filtrado[nombre_col_lote].astype(str).isin(lote_sel)]

    st.divider()

    # CÁLCULO DE PROMEDIOS Y DIFERENCIAS POR GRUPO
    def calcular_metricas_comparativas_grupo(fila_dict, df_grupo):
        c_real = df_grupo[c_cons_gr_ave].mean() if c_cons_gr_ave and c_cons_gr_ave in df_grupo.columns else 0.0
        c_tab = df_grupo[c_gr_ave_tabla].mean() if c_gr_ave_tabla and c_gr_ave_tabla in df_grupo.columns else 0.0
        fila_dict["Consumo<br>Gr. A. D."] = c_real
        fila_dict["Gr. A. D.<br>Tabla"] = c_tab
        fila_dict["Dif. Cons.<br>(g)"] = c_real - c_tab

        p_real = df_grupo[c_pct_prod_dia].mean() if c_pct_prod_dia and c_pct_prod_dia in df_grupo.columns else 0.0
        p_tab = df_grupo[c_pct_prod_tabla].mean() if c_pct_prod_tabla and c_pct_prod_tabla in df_grupo.columns else 0.0
        fila_dict["% Diario<br>de Prod."] = p_real
        fila_dict["% Dia<br>Prod. Tab"] = p_tab
        fila_dict["Dif. %<br>Prod"] = p_real - p_tab

        return fila_dict

    # FUNCIÓN CORREGIDA PARA CÁLCULO DE TOTALES GENERALES EXACTOS
    def calcular_fila_total_general(df_cuadro_disp, df_datos_base_filtrados, col_concepto="TOTAL GENERAL"):
        tot_dict = {"Fecha / Concepto": col_concepto}
        
        for nombre_final, col_orig, _, tipo_dato in columnas_totales_def:
            if nombre_final in df_cuadro_disp.columns:
                if nombre_final in ["Costo Huevo/<br>Alimento", "Dif. Cons.<br>(g)", "Dif. %<br>Prod"]:
                    continue
                elif "avg" in tipo_dato:
                    if col_orig and col_orig in df_datos_base_filtrados.columns:
                        tot_dict[nombre_final] = df_datos_base_filtrados[col_orig].mean()
                    else:
                        tot_dict[nombre_final] = df_cuadro_disp[nombre_final].mean()
                elif "saldo" in nombre_final.lower():
                    # Obtener la suma exacta de los saldos del ÚLTIMO DÍA del rango filtrado
                    max_f = df_datos_base_filtrados["Fecha_dt"].max()
                    df_ult_dia = df_datos_base_filtrados[df_datos_base_filtrados["Fecha_dt"] == max_f]
                    if col_orig and col_orig in df_ult_dia.columns:
                        tot_dict[nombre_final] = df_ult_dia[col_orig].sum()
                    else:
                        tot_dict[nombre_final] = df_cuadro_disp[nombre_final].iloc[-1]
                else:
                    tot_dict[nombre_final] = df_cuadro_disp[nombre_final].sum()

        # Recálculo exacto de métricas compuestas para la fila TOTAL GENERAL
        tot_costo = tot_dict.get("Costo<br>Alimento", 0)
        tot_prod = tot_dict.get("Producción<br>Huevos Día", 0)
        tot_dict["Costo Huevo/<br>Alimento"] = (tot_costo / tot_prod) if tot_prod > 0 else 0.0

        cons_r = tot_dict.get("Consumo<br>Gr. A. D.", 0)
        cons_t = tot_dict.get("Gr. A. D.<br>Tabla", 0)
        tot_dict["Dif. Cons.<br>(g)"] = cons_r - cons_t

        pct_r = tot_dict.get("% Diario<br>de Prod.", 0)
        pct_t = tot_dict.get("% Dia<br>Prod. Tab", 0)
        tot_dict["Dif. %<br>Prod"] = pct_r - pct_t

        return tot_dict

    # RENDERIZADOR HTML
    def render_tabla_html(df_input, col_concepto_nombre, css_class_scroll="scroll-table-10", audit_sin_datos=False):
        html_code = f"""
        <div class="{css_class_scroll} custom-table-container">
        <table class="custom-table">
            <thead>
                <tr>
                    <th class="th-fecha">{col_concepto_nombre}</th>
        """
        for nombre_final, _, bloque, tipo_dato in columnas_ordenadas:
            clase_th = "th-costo" if "currency" in tipo_dato else f"th-{bloque}"
            html_code += f'<th class="{clase_th}">{nombre_final}</th>'

        html_code += "</tr></thead><tbody>"

        for _, row in df_input.iterrows():
            concepto = str(row["Fecha / Concepto"])
            is_total = "TOTAL GENERAL" in concepto
            prod_val = row.get("Producción<br>Huevos Día", 0)
            
            es_sin_datos = audit_sin_datos and (not is_total) and (prod_val == 0)
            row_class = "row-total" if is_total else ("row-sin-datos" if es_sin_datos else "")
            label_concepto = concepto if not es_sin_datos else f"{concepto} <span class='badge-sin-datos'>⚠️ SIN DATOS</span>"

            html_code += f'<tr class="{row_class}">'
            html_code += f'<td class="td-concept" title="{concepto}">{label_concepto}</td>'

            for nombre_final, _, _, tipo_dato in columnas_ordenadas:
                val = row.get(nombre_final, 0)
                val_num = pd.to_numeric(val, errors="coerce")
                val_num = 0 if pd.isna(val_num) else val_num
                
                if es_sin_datos and ("diff" in tipo_dato or "avg" in tipo_dato):
                    html_code += '<td>-</td>'
                    continue

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
        return html_code

    def filtrar_por_fecha(df_in, clave_state):
        r_f = st.session_state.get(clave_state, [])
        if len(r_f) == 2:
            return df_in[(df_in["Fecha_dt"].dt.date >= r_f[0]) & (df_in["Fecha_dt"].dt.date <= r_f[1])]
        elif len(r_f) == 1:
            return df_in[df_in["Fecha_dt"].dt.date == r_f[0]]
        return df_in

    if columnas_ordenadas:

        # ==========================================
        # SECCIÓN DE MÉTRICAS KPI COMPARATIVAS
        # ==========================================
        rango_c1_sel = st.session_state.get("rango_c1", [])
        if len(rango_c1_sel) == 2:
            f_start, f_end = rango_c1_sel[0], rango_c1_sel[1]
            num_dias = (f_end - f_start).days + 1
            f_start_prev = f_start - timedelta(days=num_dias)
            f_end_prev = f_start - timedelta(days=1)

            df_curr = df_base_filtrado[(df_base_filtrado["Fecha_dt"].dt.date >= f_start) & (df_base_filtrado["Fecha_dt"].dt.date <= f_end)]
            df_prev = df_base_filtrado[(df_base_filtrado["Fecha_dt"].dt.date >= f_start_prev) & (df_base_filtrado["Fecha_dt"].dt.date <= f_end_prev)]

            c_costo_cur = df_curr[c_costo_alim].sum() if c_costo_alim in df_curr.columns else 0
            c_prod_cur = df_curr[c_prod_huevo].sum() if c_prod_huevo in df_curr.columns else 0
            c_mort_cur = df_curr[c_mort].sum() if c_mort in df_curr.columns else 0
            c_cons_cur = df_curr[c_consumo_b].sum() if c_consumo_b in df_curr.columns else 0
            costo_huevo_cur = (c_costo_cur / c_prod_cur) if c_prod_cur > 0 else 0.0

            c_costo_prev = df_prev[c_costo_alim].sum() if c_costo_alim in df_prev.columns else 0
            c_prod_prev = df_prev[c_prod_huevo].sum() if c_prod_huevo in df_prev.columns else 0
            c_mort_prev = df_prev[c_mort].sum() if c_mort in df_prev.columns else 0
            c_cons_prev = df_prev[c_consumo_b].sum() if c_consumo_b in df_prev.columns else 0
            costo_huevo_prev = (c_costo_prev / c_prod_prev) if c_prod_prev > 0 else 0.0

            delta_costo_huevo = costo_huevo_cur - costo_huevo_prev
            delta_prod = c_prod_cur - c_prod_prev
            delta_mort = c_mort_cur - c_mort_prev
            delta_cons = c_cons_cur - c_cons_prev

            st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 2px; color: #1e293b;'>📈 Resumen Gerencial y Deltas Comparativos</p>", unsafe_allow_html=True)
            st.caption(f"Comparando el periodo seleccionado en el Cuadro 1 ({f_start.strftime('%d/%m')} - {f_end.strftime('%d/%m')}) vs. el periodo anterior equivalente ({f_start_prev.strftime('%d/%m')} - {f_end_prev.strftime('%d/%m')}).")

            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            with kpi1:
                st.metric(
                    label="Costo Huevo/Alimento",
                    value=f"$ {costo_huevo_cur:,.2f}",
                    delta=f"{delta_costo_huevo:+,.2f} $/huevo",
                    delta_color="inverse",
                    help="Costo de alimento requerido para producir 1 huevo. Se calcula como: Costo Alimento Total / Producción Huevos Día."
                )
            with kpi2:
                st.metric(
                    label="Producción Huevos",
                    value=f"{int(round(c_prod_cur)):,} u.",
                    delta=f"{int(round(delta_prod)):,}",
                    help="Suma total de unidades de huevo recolectadas durante el rango de fechas seleccionado."
                )
            with kpi3:
                st.metric(
                    label="Mortalidad Aves",
                    value=f"{int(round(c_mort_cur)):,} aves",
                    delta=f"{int(round(delta_mort)):,}",
                    delta_color="inverse",
                    help="Cantidad acumulada de bajas de aves registradas en el periodo."
                )
            with kpi4:
                st.metric(
                    label="Consumo Alimento",
                    value=f"{c_cons_cur:,.1f} bultos",
                    delta=f"{delta_cons:+,.1f}",
                    help="Bultos de alimento (40 kg) consumidos por el lote en el rango de fechas."
                )
            st.divider()

        # ==========================================
        # CUADRO 1: GENERAL POR FECHA
        # ==========================================
        st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 4px; color: #1e293b;'>1️⃣ Consolidado General por Fecha</p>", unsafe_allow_html=True)
        with st.expander("📅 **Filtro de Fecha - Cuadro 1 (General)**", expanded=True):
            c1_1, c1_2 = st.columns([1, 2])
            with c1_1:
                st.markdown("**Acceso Rápido:**")
                b1, b2, b3 = st.columns(3)
                with b1:
                    if st.button("7 días", key="b7_c1", use_container_width=True):
                        st.session_state["rango_c1"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=7)), max_fecha_dt]
                        st.rerun()
                with b2:
                    if st.button("15 días", key="b15_c1", use_container_width=True):
                        st.session_state["rango_c1"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=15)), max_fecha_dt]
                        st.rerun()
                with b3:
                    if st.button("30 días", key="b30_c1", use_container_width=True):
                        st.session_state["rango_c1"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=30)), max_fecha_dt]
                        st.rerun()
            with c1_2:
                st.date_input("Rango de Fechas (Cuadro 1):", key="rango_c1", min_value=min_fecha_dt, max_value=max_fecha_dt)

        df_c1_filt = filtrar_por_fecha(df_base_filtrado, "rango_c1")
        fechas_c1 = sorted(df_c1_filt["Fecha_dt"].dropna().unique())

        if fechas_c1:
            filas_c1 = []
            for f_dt in fechas_c1:
                df_sub = df_c1_filt[df_c1_filt["Fecha_dt"] == f_dt]
                dia_nom = DIAS_ESPANOL.get(pd.to_datetime(f_dt).strftime("%A"), "")
                fecha_str = f"{dia_nom} {pd.to_datetime(f_dt).strftime('%d/%m/%Y')}"

                fila = {"Fecha / Concepto": fecha_str}
                for nombre_final, col_orig, _, tipo_dato in columnas_totales_def:
                    if col_orig and col_orig in df_sub.columns and "avg" not in tipo_dato and "diff" not in tipo_dato:
                        fila[nombre_final] = df_sub[col_orig].sum()

                fila = calcular_metricas_comparativas_grupo(fila, df_sub)

                c_val = df_sub[c_costo_alim].sum() if c_costo_alim in df_sub.columns else 0
                p_val = df_sub[c_prod_huevo].sum() if c_prod_huevo in df_sub.columns else 0
                fila["Costo Huevo/<br>Alimento"] = (c_val / p_val) if p_val > 0 else 0.0
                filas_c1.append(fila)

            df_c1 = pd.DataFrame(filas_c1)
            tot_c1 = calcular_fila_total_general(df_c1, df_c1_filt, "TOTAL GENERAL")
            df_c1_disp = pd.concat([df_c1, pd.DataFrame([tot_c1])], ignore_index=True)

            st.html(render_tabla_html(df_c1_disp, "Fecha", css_class_scroll="scroll-table-7"))
        else:
            st.info("No hay registros en el rango seleccionado para el Cuadro 1.")

        st.divider()

        # ==========================================
        # CUADRO 2: POR FECHA Y GRANJA
        # ==========================================
        st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 4px; color: #1e293b;'>2️⃣ Consolidado por Fecha y Granja</p>", unsafe_allow_html=True)
        with st.expander("📅 **Filtro de Fecha - Cuadro 2 (Por Granja)**", expanded=True):
            c2_1, c2_2 = st.columns([1, 2])
            with c2_1:
                st.markdown("**Acceso Rápido:**")
                b1, b2, b3 = st.columns(3)
                with b1:
                    if st.button("7 días", key="b7_c2", use_container_width=True):
                        st.session_state["rango_c2"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=7)), max_fecha_dt]
                        st.rerun()
                with b2:
                    if st.button("15 días", key="b15_c2", use_container_width=True):
                        st.session_state["rango_c2"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=15)), max_fecha_dt]
                        st.rerun()
                with b3:
                    if st.button("30 días", key="b30_c2", use_container_width=True):
                        st.session_state["rango_c2"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=30)), max_fecha_dt]
                        st.rerun()
            with c2_2:
                st.date_input("Rango de Fechas (Cuadro 2):", key="rango_c2", min_value=min_fecha_dt, max_value=max_fecha_dt)

        df_c2_filt = filtrar_por_fecha(df_base_filtrado, "rango_c2")
        fechas_c2 = sorted(df_c2_filt["Fecha_dt"].dropna().unique())

        if fechas_c2:
            filas_c2 = []
            for f_dt in fechas_c2:
                df_sub = df_c2_filt[df_c2_filt["Fecha_dt"] == f_dt]
                dia_nom = DIAS_ESPANOL.get(pd.to_datetime(f_dt).strftime("%A"), "")
                fecha_str = f"{dia_nom} {pd.to_datetime(f_dt).strftime('%d/%m/%Y')}"

                granjas_unicas = df_sub[nombre_col_granja].dropna().unique()
                for g_nombre in granjas_unicas:
                    df_sub_g = df_sub[df_sub[nombre_col_granja] == g_nombre]
                    fila = {"Fecha / Concepto": f"{fecha_str} - {g_nombre}"}
                    
                    for nombre_final, col_orig, _, tipo_dato in columnas_totales_def:
                        if col_orig and col_orig in df_sub_g.columns and "avg" not in tipo_dato and "diff" not in tipo_dato:
                            fila[nombre_final] = df_sub_g[col_orig].sum()

                    fila = calcular_metricas_comparativas_grupo(fila, df_sub_g)

                    cg_val = df_sub_g[c_costo_alim].sum() if c_costo_alim and c_costo_alim in df_sub_g.columns else 0
                    pg_val = df_sub_g[c_prod_huevo].sum() if c_prod_huevo and c_prod_huevo in df_sub_g.columns else 0
                    fila["Costo Huevo/<br>Alimento"] = (cg_val / pg_val) if pg_val > 0 else 0.0
                    filas_c2.append(fila)

            df_c2 = pd.DataFrame(filas_c2)
            tot_c2 = calcular_fila_total_general(df_c2, df_c2_filt, "TOTAL GENERAL")
            df_c2_disp = pd.concat([df_c2, pd.DataFrame([tot_c2])], ignore_index=True)

            st.html(render_tabla_html(df_c2_disp, "Fecha - Granja", css_class_scroll="scroll-table-10", audit_sin_datos=True))
        else:
            st.info("No hay registros en el rango seleccionado para el Cuadro 2.")

        st.divider()

        # ==========================================
        # CUADRO 3: POR FECHA, GRANJA Y LOTE
        # ==========================================
        st.markdown("<p style='font-size: 14px; font-weight: bold; margin-bottom: 4px; color: #1e293b;'>3️⃣ Consolidado por Fecha, Granja y Lote</p>", unsafe_allow_html=True)
        with st.expander("📅 **Filtro de Fecha - Cuadro 3 (Por Granja & Lote)**", expanded=True):
            c3_1, c3_2 = st.columns([1, 2])
            with c3_1:
                st.markdown("**Acceso Rápido:**")
                b1, b2, b3 = st.columns(3)
                with b1:
                    if st.button("7 días", key="b7_c3", use_container_width=True):
                        st.session_state["rango_c3"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=7)), max_fecha_dt]
                        st.rerun()
                with b2:
                    if st.button("15 días", key="b15_c3", use_container_width=True):
                        st.session_state["rango_c3"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=15)), max_fecha_dt]
                        st.rerun()
                with b3:
                    if st.button("30 días", key="b30_c3", use_container_width=True):
                        st.session_state["rango_c3"] = [max(min_fecha_dt, max_fecha_dt - timedelta(days=30)), max_fecha_dt]
                        st.rerun()
            with c3_2:
                st.date_input("Rango de Fechas (Cuadro 3):", key="rango_c3", min_value=min_fecha_dt, max_value=max_fecha_dt)

        df_c3_filt = filtrar_por_fecha(df_base_filtrado, "rango_c3")
        fechas_c3 = sorted(df_c3_filt["Fecha_dt"].dropna().unique())

        if fechas_c3:
            filas_c3 = []
            for f_dt in fechas_c3:
                df_sub = df_c3_filt[df_c3_filt["Fecha_dt"] == f_dt]
                dia_nom = DIAS_ESPANOL.get(pd.to_datetime(f_dt).strftime("%A"), "")
                fecha_str = f"{dia_nom} {pd.to_datetime(f_dt).strftime('%d/%m/%Y')}"

                grupos_gl = df_sub.groupby([nombre_col_granja, nombre_col_lote])
                for (g_nombre, l_nombre), df_sub_gl in grupos_gl:
                    fila = {"Fecha / Concepto": f"{fecha_str} - {g_nombre} (Lote: {l_nombre})"}
                    
                    for nombre_final, col_orig, _, tipo_dato in columnas_totales_def:
                        if col_orig and col_orig in df_sub_gl.columns and "avg" not in tipo_dato and "diff" not in tipo_dato:
                            fila[nombre_final] = df_sub_gl[col_orig].sum()

                    fila = calcular_metricas_comparativas_grupo(fila, df_sub_gl)

                    cgl_val = df_sub_gl[c_costo_alim].sum() if c_costo_alim and c_costo_alim in df_sub_gl.columns else 0
                    pgl_val = df_sub_gl[c_prod_huevo].sum() if c_prod_huevo and c_prod_huevo in df_sub_gl.columns else 0
                    fila["Costo Huevo/<br>Alimento"] = (cgl_val / pgl_val) if pgl_val > 0 else 0.0
                    filas_c3.append(fila)

            df_c3 = pd.DataFrame(filas_c3)
            tot_c3 = calcular_fila_total_general(df_c3, df_c3_filt, "TOTAL GENERAL")
            df_c3_disp = pd.concat([df_c3, pd.DataFrame([tot_c3])], ignore_index=True)

            st.html(render_tabla_html(df_c3_disp, "Fecha - Granja - Lote", css_class_scroll="scroll-table-10", audit_sin_datos=True))
        else:
            st.info("No hay registros en el rango seleccionado para el Cuadro 3.")

        # DESCARGA DE REPORTE EXCEL
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            def limpiar_cols(df_in):
                return df_in.rename(columns=lambda c: str(c).replace("<br>", " "))

            if "df_c1_disp" in locals():
                cols_e1 = ["Fecha / Concepto"] + [n for n, _, _, _ in columnas_ordenadas if n in df_c1_disp.columns]
                limpiar_cols(df_c1_disp[cols_e1]).to_excel(writer, index=False, sheet_name="General_Fecha")
            if "df_c2_disp" in locals():
                cols_e2 = ["Fecha / Concepto"] + [n for n, _, _, _ in columnas_ordenadas if n in df_c2_disp.columns]
                limpiar_cols(df_c2_disp[cols_e2]).to_excel(writer, index=False, sheet_name="Por_Granja")
            if "df_c3_disp" in locals():
                cols_e3 = ["Fecha / Concepto"] + [n for n, _, _, _ in columnas_ordenadas if n in df_c3_disp.columns]
                limpiar_cols(df_c3_disp[cols_e3]).to_excel(writer, index=False, sheet_name="Por_Granja_Lote")
        buffer.seek(0)

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📥 Descargar Reporte de 3 Cuadros (.xlsx)",
            data=buffer,
            file_name=f"CONSULTA_AVICOLA_3_NIVELES_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif not columnas_ordenadas:
        st.warning("Selecciona al menos una columna para mostrar.")
    else:
        st.info("No se encontraron registros con los filtros seleccionados.")
else:
    st.warning(f"⚠️ No se encontró el archivo consolidado en `{RUTA_REPORTE_LOCAL}`.")
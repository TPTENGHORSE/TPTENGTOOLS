import os
import re
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import base64
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def render_box(label, value):
    return f"""
    <div style='font-weight:bold; margin-bottom:8px; font-size:13px;'>{label}</div>
    <div style='padding:3px 5px; border:1px solid #eee; border-radius:4px; background:#fafafa; width:100%; max-width:none; min-width:120px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; font-size:12px;' title='{value}'>{value}</div>
    """

def _coerce_to_int(val):
    """Attempt to coerce various cell formats to an integer.
    Handles None/NaN, numeric strings with punctuation, and floats.
    Returns 0 on failure.
    """
    try:
        if pd.isna(val):
            return 0
    except Exception:
        pass
    # direct numeric
    if isinstance(val, (int, float)):
        try:
            return int(round(float(val)))
        except Exception:
            return 0
    # strings with numbers (e.g., "12 días", "~ 3.5", "4,0")
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return 0
        m = re.search(r"[-+]?\d+(?:[\.,]\d+)?", s)
        if m:
            try:
                num = float(m.group(0).replace(',', '.'))
                return int(round(num))
            except Exception:
                return 0
    # fallback
    try:
        return int(val)
    except Exception:
        return 0

# Load data from new Excel (VTT DATA.xlsx)
vtt_data_path = os.path.join(os.path.dirname(__file__), "VTT DATA.xlsx")
df_vtt = pd.read_excel(vtt_data_path)

# --- STREAMLIT INTERFACE ---
st.set_page_config(layout="wide")

st.markdown(
    """
    <style>
    .main .block-container {
        padding-top: 0.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 80% !important;
    }
    header[data-testid="stHeader"] {
        height: 0px !important;
        min-height: 0px !important;
        padding: 0 !important;
    }
    /* vertical text utility for date headers */
    .vtt-vertical-text {
        display: inline-block;
        writing-mode: vertical-rl;
        text-orientation: upright;
        white-space: nowrap;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

col_info, col_timeline = st.columns([1, 4], gap="small")

with col_info:
    # Filtros POL y POD en una fila
    col_pol, col_pod = st.columns([1, 1], gap="medium")
    pol_options = df_vtt['POL'].dropna().astype(str).unique().tolist() if 'POL' in df_vtt.columns else []
    pod_options = df_vtt['POD'].dropna().astype(str).unique().tolist() if 'POD' in df_vtt.columns else []
    if 'pol_select' not in st.session_state:
        st.session_state['pol_select'] = pol_options[0] if pol_options else ''
    if 'pod_select' not in st.session_state:
        st.session_state['pod_select'] = pod_options[0] if pod_options else ''
    with col_pol:
        # Label grande para POL y ocultar la etiqueta por defecto del widget
        st.markdown("<div style='font-size:28px; font-weight:700; line-height:1; margin:0 0 6px;'>POL</div>", unsafe_allow_html=True)
        selected_pol = st.selectbox("POL", pol_options, key="pol_select", label_visibility="collapsed")
    # Limitar PODs según el POL seleccionado
    filtered_pod_options = (
        df_vtt[df_vtt['POL'].astype(str) == st.session_state['pol_select']]['POD']
        .dropna().astype(str).unique().tolist()
        if 'POD' in df_vtt.columns else []
    )
    with col_pod:
        # Label grande para POD y ocultar la etiqueta por defecto del widget
        st.markdown("<div style='font-size:28px; font-weight:700; line-height:1; margin:0 0 6px;'>POD</div>", unsafe_allow_html=True)
        selected_pod = st.selectbox("POD", filtered_pod_options, key="pod_select", label_visibility="collapsed")
    # Mantener POD válido cuando cambie POL
    if selected_pod not in filtered_pod_options and filtered_pod_options:
        st.session_state['pod_select'] = filtered_pod_options[0]
    # Filtrar por POL y POD seleccionados
    filtered_df = df_vtt[(df_vtt['POL'].astype(str) == st.session_state['pol_select']) & (df_vtt['POD'].astype(str) == st.session_state['pod_select'])]
    # Si hay múltiples filas para el mismo par POL/POD permitir elegir el registro específico
    if not filtered_df.empty:
        if len(filtered_df) > 1:
            def build_label(r):
                parts = []
                if 'ID' in r and pd.notnull(r['ID']):
                    parts.append(f"ID:{r['ID']}")
                if 'Carrier' in r and pd.notnull(r['Carrier']):
                    parts.append(f"Carrier:{r['Carrier']}")
                if 'Name Destin Site' in r and pd.notnull(r['Name Destin Site']):
                    parts.append(f"Plant:{r['Name Destin Site']}")
                if 'Expiration Date' in r and pd.notnull(r['Expiration Date']):
                    exp_val = r['Expiration Date']
                    if isinstance(exp_val, (pd.Timestamp, datetime)):
                        exp_str = exp_val.strftime('%d/%m/%Y')
                    else:
                        try:
                            exp_str = pd.to_datetime(exp_val).strftime('%d/%m/%Y')
                        except Exception:
                            exp_str = str(exp_val)
                    parts.append(f"Exp:{exp_str}")
                return " | ".join(parts) if parts else str(r.name)

            option_indices = list(filtered_df.index)
            option_labels = [build_label(filtered_df.loc[idx]) for idx in option_indices]
            # Valor por defecto en session_state
            if 'record_select' not in st.session_state or st.session_state['record_select'] not in option_indices:
                st.session_state['record_select'] = option_indices[0]
            selected_label = st.selectbox(
                "Registro (varios coincidieron)",
                options=option_indices,
                format_func=lambda x: option_labels[option_indices.index(x)],
                key='record_select'
            )
            row = filtered_df.loc[selected_label]
        else:
            row = filtered_df.iloc[0]
    else:
        row = None

    # E/D se muestra más abajo del timeline y antes del botón Generate files

# KPIs movidos al final

# --- TIMELINE (Gantt stays here; controls will be rendered below) ---
st.markdown("<hr style='margin:16px 0;'>", unsafe_allow_html=True)

# Render the info row (ID, Carrier, Shipper, ILN, PLANT) in the wide column
with col_timeline:
    st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
    info_cols = st.columns([1.0, 1.2, 1.2, 1.0, 1.0], gap="medium")
    with info_cols[0]:
        if row is not None and 'ID' in df_vtt.columns:
            st.markdown(render_box('ID', row['ID']), unsafe_allow_html=True)
        else:
            st.info("No existe la columna ID o no hay coincidencia.")
    with info_cols[1]:
        if row is not None and 'Carrier' in df_vtt.columns:
            st.markdown(render_box('Carrier', row['Carrier']), unsafe_allow_html=True)
        else:
            st.info("No existe la columna Carrier o no hay coincidencia.")
    with info_cols[2]:
        if row is not None and len(df_vtt.columns) > 10:
            try:
                col_k = df_vtt.columns[10]
                st.markdown(render_box('Shipper', row.get(col_k, "")), unsafe_allow_html=True)
            except Exception:
                st.info("No se pudo leer la columna K (Shipper) o no hay coincidencia.")
        else:
            st.info("No se pudo leer la columna K (Shipper) o no hay coincidencia.")
    with info_cols[3]:
        if row is not None and len(df_vtt.columns) > 8:
            try:
                col_i = df_vtt.columns[8]
                st.markdown(render_box('ILN', row.get(col_i, "")), unsafe_allow_html=True)
            except Exception:
                st.info("No se pudo leer la columna I (ILN) o no hay coincidencia.")
        else:
            st.info("No se pudo leer la columna I (ILN) o no hay coincidencia.")
    with info_cols[4]:
        if row is not None and 'Name Destin Site' in df_vtt.columns:
            st.markdown(render_box('PLANT', row['Name Destin Site']), unsafe_allow_html=True)
        else:
            st.info("No existe la columna Name Destin Site o no hay coincidencia.")

time_cols_fixed = 4
today = datetime.today()
# Calcular el lunes de la semana actual
start_date = today - timedelta(days=today.weekday())
# Leer el valor del slider desde session_state (el control se renderiza al final)
num_days = int(st.session_state.get("days_slider_timeline", 110))
timeline_days = [start_date + timedelta(days=i) for i in range(num_days)]
time_cols = time_cols_fixed + num_days

# Encabezados fijos y dinámicos
headers = ["Steps", "Day", "Day+", "Final Day"]
table_html = """
<table class='timeline-table' style='width:100%; border-collapse:collapse; margin-top:8px;'>
    <thead>"""
# Fila de semana combinada
table_html += "<tr>"
for idx_h, h in enumerate(headers):
    if idx_h == 0:
        # Steps column: wider and no wrapping
        table_html += "<th style='border:none; background:none; min-width:80px; white-space:nowrap;'></th>"
    else:
        table_html += "<th style='border:none; background:none'></th>"
# Agrupar días por semana
semana_actual = None
colspan = 0
for idx, day in enumerate(timeline_days):
    semana = day.isocalendar()[1]
    if semana_actual is None:
        semana_actual = semana
        colspan = 1
    elif semana == semana_actual:
        colspan += 1
    else:
        # Imprimir celda combinada para la semana anterior
        table_html += f"<th colspan='{colspan}' style='padding:0 1px; border:1px solid #eee; min-width:28px; text-align:center; background:#fffbe6; font-size:13.5px; font-weight:bold;'>W{semana_actual}</th>"
        semana_actual = semana
        colspan = 1
# Imprimir la última semana
if semana_actual is not None:
    table_html += f"<th colspan='{colspan}' style='padding:0 1px; border:1px solid #eee; min-width:28px; text-align:center; background:#fffbe6; font-size:13.5px; font-weight:bold;'>W{semana_actual}</th>"
table_html += "</tr>"
# Fila de encabezados de fechas
table_html += "<tr>"
for idx_h, h in enumerate(headers):
    if idx_h == 0:
        table_html += f"<th style='padding:5px 7px; border:1px solid #eee; min-width:200px; text-align:center; background:#f5f5f5; white-space:nowrap'>{h}</th>"
    else:
        table_html += f"<th style='padding:5px 7px; border:1px solid #eee; min-width:50px; width:50px; text-align:center; background:#f5f5f5'>{h}</th>"
for idx, day in enumerate(timeline_days):
    # Colorear sábados y domingos
    if day.weekday() in [5, 6]:
        th_style = "padding:0 1px; border:1px solid #eee; min-width:15px; width:18px; height:50px; text-align:center; background:#ffd6d6; font-size:12px; vertical-align:bottom;"
    else:
        th_style = "padding:0 1px; border:1px solid #eee; min-width:20px; width:20px; height:50px; text-align:center; background:#e3eafc; font-size:12px; vertical-align:bottom;"
    # Mostrar solo la letra inicial del día en mayúscula
    vertical_label = day.strftime('%a')[0].upper()  # M, T, W, etc.
    # Centrar verticalmente la letra inicial
    table_html += f"<th style='{th_style}'><span class='vtt-vertical-text' style='display:flex;align-items:center;justify-content:center;height:100%;'>{vertical_label}</span></th>"
table_html += "</tr></thead><tbody>"

# Etiquetas de filas
time_labels = [
    "1. Day Customer Order",
    "2. Day ILN Order",
    "3. First Receipt Days",
    "4. Pack. prep. & load",
    "5. Transport ILN to POL",
    "6. First Day to POL",
    "7. Cut off",
    "8. ETD",
    "9. TT (ETD> ETA)",
    "10. Days flexibility 1",
    "11. Days flexibility 2",
    "12. Customs clearence",
    "13. Transport to plant",
    "14. Rounding",
    "15. Due Date"
]

time_rows = len(time_labels)
for i in range(time_rows):
    # Reduce row height ~35% (15px -> ~10px)
    table_html += "<tr style='height:15px;'>"
    for j in range(time_cols):
        cell_content = ""
        # Alinear la primera columna (etiquetas) a la izquierda
        if j == 0:
            # Steps column: make it wider and prevent wrapping
            cell_style = "padding:4px 6px; border:1px solid #eee; text-align:left; font-weight:bold; background:#f5f5f5; min-width:200px; white-space:nowrap;"
        else:
            cell_style = "padding:4px 6px; border:1px solid #eee; text-align:center;"
        # Compactar altura y padding en todas las celdas de steps (≈ -35%)
        cell_style += "height:15px; line-height:15px; padding:1px 4px;"
        # Colorear sábados y domingos en las celdas de fechas
        if j >= 4:
            fecha_actual = timeline_days[j-4] if (j-4) < len(timeline_days) else None
            if fecha_actual is not None and fecha_actual.weekday() in [5, 6]:
                cell_style += "background-color:#ffd6d6;"
        if i == 0:  # 1. Day Customer Order
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '1 Day Customer Order' in df_vtt.columns:
                    cell_content = row['1 Day Customer Order']
                else:
                    cell_content = "-"
            elif j == 2:
                cell_content = "0"
            elif j == 3:
                if row is not None and '1 Day Customer Order' in df_vtt.columns:
                    cell_content = row['1 Day Customer Order']
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['1 Day Customer Order']) if row is not None and '1 Day Customer Order' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                paint_len = 1  # Day+ = 0 -> solo último día
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 1:  # 2. Day ILN Order
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '2 Day ILN Order' in df_vtt.columns:
                    val = row['2 Day ILN Order']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 2:
                cell_content = "0"
            elif j == 3:
                if row is not None and '2 Day ILN Order' in df_vtt.columns:
                    val = row['2 Day ILN Order']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['2 Day ILN Order']) if row is not None and '2 Day ILN Order' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                paint_len = 1  # Day+ = 0 -> solo último día
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 2:  # 3. First Receipt Days
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '3 First Receipt Days' in df_vtt.columns:
                    val = row['3 First Receipt Days']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "No hay datos para la combinación POL/POD seleccionada"
            elif j == 2:
                if row is not None and '3 .1 Time of Recept in ILN' in df_vtt.columns:
                    val = row['3 .1 Time of Recept in ILN']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:
                if row is not None and '3.2 First Receipt Days' in df_vtt.columns:
                    val = row['3.2 First Receipt Days']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['3.2 First Receipt Days']) if row is not None and '3.2 First Receipt Days' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                day_plus_val = _coerce_to_int(row['3 .1 Time of Recept in ILN']) if row is not None and '3 .1 Time of Recept in ILN' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 4:  # 5. Transport ILN to POL
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '5.1 Transport ILN to POL' in df_vtt.columns:
                    val = row['5.1 Transport ILN to POL']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 2:
                if row is not None and '5.2 Transport ILN to POL' in df_vtt.columns:
                    val = row['5.2 Transport ILN to POL']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:
                if row is not None and '5.3 Transport ILN to POL' in df_vtt.columns:
                    val = row['5.3 Transport ILN to POL']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['5.3 Transport ILN to POL']) if row is not None and '5.3 Transport ILN to POL' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                day_plus_val = _coerce_to_int(row['5.2 Transport ILN to POL']) if row is not None and '5.2 Transport ILN to POL' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 5:  # 6. First Day to POL
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '6 First Day to POL' in df_vtt.columns:
                    val = row['6 First Day to POL']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 2:
                cell_content = "0"
            elif j == 3:
                if row is not None and '6 First Day to POL' in df_vtt.columns:
                    val = row['6 First Day to POL']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['6 First Day to POL']) if row is not None and '6 First Day to POL' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                paint_len = 1  # Day+ = 0 -> solo último día
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 6:  # 7. Cut off
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '7 Cutt off' in df_vtt.columns:
                    val = row['7 Cutt off']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 2:
                cell_content = "0"
            elif j == 3:
                if row is not None and '7 Cutt off' in df_vtt.columns:
                    val = row['7 Cutt off']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['7 Cutt off']) if row is not None and '7 Cutt off' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                paint_len = 1  # Day+ = 0
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 7:  # 8. ETD
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '8 ETD' in df_vtt.columns:
                    val = row['8 ETD']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 2:
                cell_content = "0"
            elif j == 3:
                if row is not None and '8 ETD' in df_vtt.columns:
                    val = row['8 ETD']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['8 ETD']) if row is not None and '8 ETD' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                paint_len = 1  # Day+ = 0
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 8:  # 9. TT (ETD> ETA)
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '8 ETD' in df_vtt.columns:
                    val = row['8 ETD']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 2:
                if row is not None and 'Transit time' in df_vtt.columns:
                    val = row['Transit time']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:
                final_col = None
                if row is not None:
                    if '9 ETD> ETA' in df_vtt.columns:
                        final_col = '9 ETD> ETA'
                    elif '9 ETD>ETA' in df_vtt.columns:
                        final_col = '9 ETD>ETA'
                if final_col is not None:
                    val = row[final_col]
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                dias_final_day = 0
                try:
                    if row is not None and '9 ETD> ETA' in df_vtt.columns:
                        dias_final_day = int(row['9 ETD> ETA'])
                    elif row is not None and '9 ETD>ETA' in df_vtt.columns:
                        dias_final_day = int(row['9 ETD>ETA'])
                except Exception:
                    dias_final_day = 0
                day_plus_val = _coerce_to_int(row['Transit time']) if row is not None and 'Transit time' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#00008b;"
        elif i == 9:  # 10. Days flexibility 1
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:  # Day = Final day of step 9 + 1
                # Base: '9 ETD> ETA' o '9 ETD>ETA'
                base_val = None
                if row is not None:
                    candidate = None
                    if '9 ETD> ETA' in df_vtt.columns:
                        candidate = row.get('9 ETD> ETA')
                    if (candidate is None or pd.isna(candidate)) and '9 ETD>ETA' in df_vtt.columns:
                        candidate = row.get('9 ETD>ETA')
                    if pd.isna(candidate) if isinstance(candidate, (int, float, pd.Series, pd.Timestamp)) else (candidate is None):
                        cell_content = "-"
                    else:
                        # Convertir a número y sumar 1
                        try:
                            num_val = pd.to_numeric(candidate, errors='coerce')
                            if pd.isna(num_val):
                                # regex fallback
                                matches = re.findall(r"[-+]?\d*\.?\d+", str(candidate))
                                num_val = float(matches[0]) if matches else float('nan')
                            if pd.isna(num_val):
                                cell_content = "-"
                            else:
                                cell_content = str(int(float(num_val)) + 1)
                        except Exception:
                            cell_content = "-"
                else:
                    cell_content = "-"
            elif j == 2:  # Day+
                if row is not None and 'Time for security' in df_vtt.columns:
                    val = row['Time for security']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:  # Final Day
                # Usar columna '10 Days flexibility 1' si existe; si no, derivar Day + Day+
                if row is not None and '10 Days flexibility 1' in df_vtt.columns:
                    val = row['10 Days flexibility 1']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    # derivado
                    cell_content = "-"
                    try:
                        # compute from base (9 ETD>ETA) + 1 + buffer
                        base = None
                        if row is not None and '9 ETD> ETA' in df_vtt.columns:
                            base = row['9 ETD> ETA']
                        elif row is not None and '9 ETD>ETA' in df_vtt.columns:
                            base = row['9 ETD>ETA']
                        bnum = pd.to_numeric(base, errors='coerce') if base is not None else float('nan')
                        if pd.isna(bnum):
                            m = re.findall(r"[-+]?\d*\.?\d+", str(base)) if base is not None else []
                            bnum = float(m[0]) if m else float('nan')
                        plus = _coerce_to_int(row['Time for security']) if row is not None and 'Time for security' in df_vtt.columns else 0
                        if not pd.isna(bnum):
                            cell_content = str(int(float(bnum)) + 1 + int(plus))
                    except Exception:
                        cell_content = "-"
            elif j >= 4:
                # pintar últimos Day+ días hasta el Final Day
                try:
                    dias_final_day = 0
                    if row is not None and '10 Days flexibility 1' in df_vtt.columns:
                        dias_final_day = int(row['10 Days flexibility 1'])
                except Exception:
                    dias_final_day = 0
                day_plus_val = _coerce_to_int(row['Time for security']) if row is not None and 'Time for security' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 10:  # 11. Days flexibility 2
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:  # Day = Final day of step 10 + 1
                if row is not None and '10 Days flexibility 1' in df_vtt.columns:
                    base_val = row['10 Days flexibility 1']
                    num_val = pd.to_numeric(base_val, errors='coerce')
                    if pd.isna(num_val):
                        try:
                            matches = re.findall(r"[-+]?\d*\.?\d+", str(base_val))
                            num_val = float(matches[0]) if matches else float('nan')
                        except Exception:
                            num_val = float('nan')
                    if pd.isna(num_val):
                        cell_content = "-"
                    else:
                        try:
                            cell_content = str(int(float(num_val)) + 1)
                        except Exception:
                            cell_content = "-"
                else:
                    cell_content = "-"
            elif j == 2:  # Day+ usa Time for security2 buffer
                if row is not None and 'Time for security2 buffer' in df_vtt.columns:
                    val = row['Time for security2 buffer']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "0"
            elif j == 3:  # Final Day
                if row is not None and '11 Days flexibility 2' in df_vtt.columns:
                    val = row['11 Days flexibility 2']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['11 Days flexibility 2']) if row is not None and '11 Days flexibility 2' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                # Pintado basado en Time for security2 buffer
                day_plus_val = _coerce_to_int(row['Time for security2 buffer']) if row is not None and 'Time for security2 buffer' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 11:  # 12. Customs clearence
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '11 Days flexibility 2' in df_vtt.columns:
                    base_val = row['11 Days flexibility 2']
                    num_val = pd.to_numeric(base_val, errors='coerce')
                    if pd.isna(num_val):
                        try:
                            matches = re.findall(r"[-+]?\d*\.?\d+", str(base_val))
                            num_val = float(matches[0]) if matches else float('nan')
                        except Exception:
                            num_val = float('nan')
                    if pd.isna(num_val):
                        cell_content = "-"
                    else:
                        try:
                            cell_content = str(int(float(num_val)) + 1)
                        except Exception:
                            cell_content = "-"
                else:
                    cell_content = "-"
            elif j == 2:
                if row is not None and 'Cust.' in df_vtt.columns:
                    val = row['Cust.']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:
                final_col = None
                if row is not None:
                    if '12 Customs Clearance' in df_vtt.columns:
                        final_col = '12 Customs Clearance'
                    elif '12 Customs clearence' in df_vtt.columns:
                        final_col = '12 Customs clearence'
                if final_col is not None:
                    val = row[final_col]
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                dias_final_day = 0
                try:
                    if row is not None and '12 Customs Clearance' in df_vtt.columns:
                        dias_final_day = int(row['12 Customs Clearance'])
                    elif row is not None and '12 Customs clearence' in df_vtt.columns:
                        dias_final_day = int(row['12 Customs clearence'])
                except Exception:
                    dias_final_day = 0
                day_plus_val = _coerce_to_int(row['Cust.']) if row is not None and 'Cust.' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 12:  # 13. Transport to plant
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                base_val = None
                if row is not None:
                    if '12 Customs Clearance' in df_vtt.columns and pd.notna(row.get('12 Customs Clearance')):
                        base_val = row.get('12 Customs Clearance')
                    elif '12 Customs clearence' in df_vtt.columns and pd.notna(row.get('12 Customs clearence')):
                        base_val = row.get('12 Customs clearence')
                if base_val is not None:
                    num_val = pd.to_numeric(base_val, errors='coerce')
                    if pd.isna(num_val):
                        try:
                            matches = re.findall(r"[-+]?\d*\.?\d+", str(base_val))
                            num_val = float(matches[0]) if matches else float('nan')
                        except Exception:
                            num_val = float('nan')
                    if pd.isna(num_val):
                        cell_content = "-"
                    else:
                        try:
                            cell_content = str(int(float(num_val)) + 1)
                        except Exception:
                            cell_content = "-"
                else:
                    cell_content = "-"
            elif j == 2:
                if row is not None and 'Trpt POD/PFI vers Usine' in df_vtt.columns:
                    val = row['Trpt POD/PFI vers Usine']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:
                if row is not None and '13 Transport to Plant' in df_vtt.columns:
                    val = row['13 Transport to Plant']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['13 Transport to Plant']) if row is not None and '13 Transport to Plant' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                day_plus_val = _coerce_to_int(row['Trpt POD/PFI vers Usine']) if row is not None and 'Trpt POD/PFI vers Usine' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 13:  # 14. Rounding
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '13 Transport to Plant' in df_vtt.columns:
                    base_val = row['13 Transport to Plant']
                    num_val = pd.to_numeric(base_val, errors='coerce')
                    if pd.isna(num_val):
                        try:
                            matches = re.findall(r"[-+]?\d*\.?\d+", str(base_val))
                            num_val = float(matches[0]) if matches else float('nan')
                        except Exception:
                            num_val = float('nan')
                    if pd.isna(num_val):
                        cell_content = "-"
                    else:
                        try:
                            cell_content = str(int(float(num_val)) + 1)
                        except Exception:
                            cell_content = "-"
                else:
                    cell_content = "-"
            elif j == 2:
                val = None
                if row is not None:
                    if 'Round.' in df_vtt.columns:
                        val = row['Round.']
                    elif 'Round' in df_vtt.columns:
                        val = row['Round']
                if val is not None:
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:
                if row is not None and '14 Rounding' in df_vtt.columns:
                    val = row['14 Rounding']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['14 Rounding']) if row is not None and '14 Rounding' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                day_plus_val = None
                if row is not None:
                    if 'Round.' in df_vtt.columns:
                        day_plus_val = _coerce_to_int(row['Round.'])
                    elif 'Round' in df_vtt.columns:
                        day_plus_val = _coerce_to_int(row['Round'])
                day_plus_val = day_plus_val if day_plus_val is not None else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 14:  # 15. Due Date
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '14 Rounding' in df_vtt.columns:
                    base_val = row['14 Rounding']
                    num_val = pd.to_numeric(base_val, errors='coerce')
                    if pd.isna(num_val):
                        try:
                            matches = re.findall(r"[-+]?\d*\.?\d+", str(base_val))
                            num_val = float(matches[0]) if matches else float('nan')
                        except Exception:
                            num_val = float('nan')
                    if pd.isna(num_val):
                        cell_content = "-"
                    else:
                        try:
                            cell_content = str(int(float(num_val)) + 1)
                        except Exception:
                            cell_content = "-"
                else:
                    cell_content = "-"
            elif j == 2:
                cell_content = "7"
            elif j == 3:
                if row is not None and '15 Due Date' in df_vtt.columns:
                    val = row['15 Due Date']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['15 Due Date']) if row is not None and '15 Due Date' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                day_plus_val = 7
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 15:  # 16. Manufacturing
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '15 Due Date' in df_vtt.columns:
                    base_val = row['15 Due Date']
                    num_val = pd.to_numeric(base_val, errors='coerce')
                    if pd.isna(num_val):
                        try:
                            matches = re.findall(r"[-+]?\d*\.?\d+", str(base_val))
                            num_val = float(matches[0]) if matches else float('nan')
                        except Exception:
                            num_val = float('nan')
                    if pd.isna(num_val):
                        cell_content = "-"
                    else:
                        try:
                            cell_content = str(int(float(num_val)) + 1)
                        except Exception:
                            cell_content = "-"
                else:
                    cell_content = "-"
            elif j == 2:
                cell_content = "7"
            elif j == 3:
                if row is not None and '16 Manufacturing' in df_vtt.columns:
                    val = row['16 Manufacturing']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['16 Manufacturing']) if row is not None and '16 Manufacturing' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                day_plus_val = 7
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        elif i == 3:  # 4. Pack. prep. & load
            if j == 0:
                cell_content = time_labels[i]
            elif j == 1:
                if row is not None and '4.1 Packaging préparation & loading' in df_vtt.columns:
                    val = row['4.1 Packaging préparation & loading']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 2:
                if row is not None and '4.2 Packaging préparation & loading' in df_vtt.columns:
                    val = row['4.2 Packaging préparation & loading']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j == 3:
                if row is not None and '4.3 Packaging préparation & loading' in df_vtt.columns:
                    val = row['4.3 Packaging préparation & loading']
                    if pd.isna(val):
                        cell_content = "-"
                    elif val == 0:
                        cell_content = "0"
                    else:
                        cell_content = str(val)
                else:
                    cell_content = "-"
            elif j >= 4:
                try:
                    dias_final_day = int(row['4.3 Packaging préparation & loading']) if row is not None and '4.3 Packaging préparation & loading' in df_vtt.columns else 0
                except Exception:
                    dias_final_day = 0
                day_plus_val = _coerce_to_int(row['4.2 Packaging préparation & loading']) if row is not None and '4.2 Packaging préparation & loading' in df_vtt.columns else 0
                paint_len = day_plus_val if (day_plus_val and day_plus_val > 0) else 1
                start_idx = max(1, dias_final_day - paint_len + 1)
                if start_idx <= (j-3) <= dias_final_day:
                    cell_content = ""
                    cell_style += "background-color:#90ee90;"
        else:
            if j == 0:
                cell_content = time_labels[i]
            else:
                cell_content = ""
        table_html += f"<td style='{cell_style}'>{cell_content}</td>"
    table_html += "</tr>"

table_html += "</tbody></table>"
# Render visible table as before, but with a distinct id to avoid capture conflicts
wrapped_html_visible = f"<div id='timeline_capture_table' style='display:inline-block'>{table_html}</div>"
st.markdown(wrapped_html_visible, unsafe_allow_html=True)



# --- KPIs al final ---
st.markdown("<hr style='margin:32px 0;'>", unsafe_allow_html=True)

# Parts Vanning
col_label, col_value = st.columns([1, 3], gap="small")
with col_label:
    st.markdown(
        "<div style='font-weight:bold; font-size:25px; margin-bottom:8px;'>Parts Vanning</div>",
        unsafe_allow_html=True,
    )
with col_value:
    if row is not None and "Parts Vanning" in df_vtt.columns:
        st.markdown(
            f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa; font-size:28px;'>{row['Parts Vanning']}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.info("No existe la columna Parts Vanning o no hay coincidencia.")

# Transit Time
col_label_tt, col_value_tt = st.columns([1, 3], gap="small")
with col_label_tt:
    st.markdown(
        "<div style='font-weight:bold; font-size:25px; margin-bottom:8px;'>Transit Time</div>",
        unsafe_allow_html=True,
    )
with col_value_tt:
    total_tt = None
    if row is not None:
        t1 = pd.to_numeric(row.get("Transit time", None), errors="coerce") if "Transit time" in df_vtt.columns else None
        t2 = pd.to_numeric(row.get("Time for security", None), errors="coerce") if "Time for security" in df_vtt.columns else None
        parts = [v for v in (t1, t2) if v is not None and pd.notna(v)]
        if parts:
            total_tt = float(sum(parts))
    if total_tt is not None and pd.notna(total_tt):
        display_tt = int(total_tt) if abs(total_tt - int(total_tt)) < 1e-9 else round(total_tt, 2)
        st.markdown(
            f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa; font-size:28px;'>{display_tt}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.info(
            "No existe 'Transit time' o 'Time for security' o no hay datos para sumarlos."
        )

# Customer Leadtime
col_label2, col_value2 = st.columns([1, 3], gap="small")
with col_label2:
    st.markdown(
        "<div style='font-weight:bold; font-size:25px; margin-bottom:8px;'>Customer Leadtime</div>",
        unsafe_allow_html=True,
    )
with col_value2:
    if row is not None and 'Cust. Leadtime' in df_vtt.columns:
        st.markdown(
            f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa; font-size:28px;'>{row['Cust. Leadtime']}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.info("No existe la columna Cust. Leadtime o no hay coincidencia.")

# Customer Safety STOCK
col_label3, col_value3 = st.columns([1, 3], gap="small")
with col_label3:
    st.markdown("<div style='font-weight:bold; font-size:25px; margin-bottom:8px;'>Customer Safety STOCK</div>", unsafe_allow_html=True)
with col_value3:
    if row is not None and 'Safety stock' in df_vtt.columns:
        st.markdown(
            f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa; font-size:28px;'>{row['Safety stock']}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.info("No existe la columna Safety stock o no hay coincidencia.")
# Controles de Timeline al final (sin mover la tabla de gantt)
st.subheader("Timeline")
st.slider(
    "Days to Show",
    min_value=7,
    max_value=150,
    value=st.session_state.get("days_slider_timeline", 100),
    step=1,
    key="days_slider_timeline",
)

# Build an off-screen composite capture area that includes table + KPIs + selection context
capture_pol = st.session_state.get('pol_select','')
capture_pod = st.session_state.get('pod_select','')
capture_days = st.session_state.get('days_slider_timeline', 100)
composite_html = ""
composite_html += "<div id='timeline_capture' style='position:absolute; left:-100000px; top:0; background:#fff; padding:8px; font-family:Arial, sans-serif; display:inline-block; width:max-content; max-width:none; overflow:visible;'>"
composite_html += "<div style='font-size:22px; font-weight:700; margin-bottom:8px;'>VTT View</div>"
composite_html += f"<div style='margin-bottom:8px;'><b>POL:</b> {capture_pol} &nbsp;&nbsp; <b>POD:</b> {capture_pod} &nbsp;&nbsp; <b>Days to Show:</b> {capture_days}</div>"

# Add ID, Carrier, Shipper, ILN, PLANT row (E/D irá abajo del timeline)
_id_val = _carrier_val = _shipper_val = _iln_val = _plant_val = _ed_val = ""
if row is not None:
    try:
        _id_val = str(row.get('ID', '')) if 'ID' in df_vtt.columns else ''
    except Exception:
        _id_val = ''
    try:
        _carrier_val = str(row.get('Carrier', '')) if 'Carrier' in df_vtt.columns else ''
    except Exception:
        _carrier_val = ''
    # Shipper from column K (index 10) if present
    try:
        _shipper_col = df_vtt.columns[10] if len(df_vtt.columns) > 10 else None
        _shipper_val = str(row.get(_shipper_col, '')) if _shipper_col and _shipper_col in df_vtt.columns else ''
    except Exception:
        _shipper_val = ''
    # ILN from column I (index 8) if present
    try:
        _iln_col = df_vtt.columns[8] if len(df_vtt.columns) > 8 else None
        _iln_val = str(row.get(_iln_col, '')) if _iln_col and _iln_col in df_vtt.columns else ''
    except Exception:
        _iln_val = ''
    try:
        _plant_val = str(row.get('Name Destin Site', '')) if 'Name Destin Site' in df_vtt.columns else ''
    except Exception:
        _plant_val = ''

composite_html += "<div style='display:grid; grid-template-columns: max-content 1fr max-content 1fr max-content 1fr; gap:6px 12px; align-items:center; margin:6px 0 10px 0;'>"
composite_html += f"<div style='font-weight:bold;'>ID:</div><div>{_id_val}</div>"
composite_html += f"<div style='font-weight:bold;'>Carrier:</div><div>{_carrier_val}</div>"
composite_html += f"<div style='font-weight:bold;'>Shipper:</div><div>{_shipper_val}</div>"
composite_html += f"<div style='font-weight:bold;'>ILN:</div><div>{_iln_val}</div>"
composite_html += f"<div style='font-weight:bold;'>PLANT:</div><div>{_plant_val}</div>"
composite_html += "</div>"

# Wrap the table to allow full-width capture (no fixed width)
composite_html += f"<div style='display:inline-block; width:max-content; overflow:visible;'>{table_html}</div>"

# E/D debajo del timeline en el PNG
try:
    if row is not None and 'Expiration Date' in df_vtt.columns:
        _exp_date = row.get('Expiration Date', '')
        if pd.notnull(_exp_date):
            if isinstance(_exp_date, (pd.Timestamp, datetime)):
                _ed_val = _exp_date.strftime('%d/%m/%Y')
            else:
                try:
                    _ed_val = pd.to_datetime(_exp_date).strftime('%d/%m/%Y')
                except Exception:
                    _ed_val = str(_exp_date)
        else:
            _ed_val = ''
    composite_html += f"<div style='margin:8px 0 4px 0;'>{render_box('E/D', _ed_val)}</div>"
except Exception:
    pass
composite_html += "<hr style='margin:16px 0;'>"
composite_html += "<div style='font-size:18px; font-weight:700; margin-bottom:8px;'>KPIs</div>"
composite_html += "<div style='display:grid; grid-template-columns: 220px 1fr; row-gap:6px; column-gap:12px;'>"
composite_html += "<div style='font-weight:bold;'>Parts Vanning</div>"
if row is not None and "Parts Vanning" in df_vtt.columns:
    composite_html += f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa;'>{row['Parts Vanning']}</div>"
else:
    composite_html += "<div>-</div>"
composite_html += "<div style='font-weight:bold;'>Transit Time</div>"
_tt_display = None
if row is not None:
    _t1 = pd.to_numeric(row.get('Transit time', None), errors='coerce') if 'Transit time' in df_vtt.columns else None
    _t2 = pd.to_numeric(row.get('Time for security', None), errors='coerce') if 'Time for security' in df_vtt.columns else None
    _parts = [v for v in (_t1, _t2) if v is not None and pd.notna(v)]
    if _parts:
        _sum = float(sum(_parts))
        _tt_display = int(_sum) if abs(_sum - int(_sum)) < 1e-9 else round(_sum, 2)
composite_html += f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa;'>{_tt_display if _tt_display is not None else '-'}</div>"
composite_html += "<div style='font-weight:bold;'>Customer Leadtime</div>"
if row is not None and 'Cust. Leadtime' in df_vtt.columns:
    composite_html += f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa;'>{row['Cust. Leadtime']}</div>"
else:
    composite_html += "<div>-</div>"
composite_html += "<div style='font-weight:bold;'>Customer Safety STOCK</div>"
if row is not None and 'Safety stock' in df_vtt.columns:
    composite_html += f"<div style='padding:4px 8px; border:1px solid #eee; border-radius:4px; background:#fafafa;'>{row['Safety stock']}</div>"
else:
    composite_html += "<div>-</div>"
composite_html += "</div>"  # end grid
composite_html += "</div>"  # end capture root
st.markdown(composite_html, unsafe_allow_html=True)

# --- Descargar Excel con la visualización completa ---
def _hex_to_fill(hex_color):
    if not hex_color:
        return None
    h = hex_color.lstrip('#')
    if len(h) == 6:
        h = 'FF' + h.upper()
    return PatternFill(fill_type='solid', start_color=h, end_color=h)

def _compute_week_spans(days):
    spans = []
    current_week = None
    count = 0
    for d in days:
        w = d.isocalendar()[1]
        if current_week is None:
            current_week = w
            count = 1
        elif w == current_week:
            count += 1
        else:
            spans.append((current_week, count))
            current_week = w
            count = 1
    if current_week is not None:
        spans.append((current_week, count))
    return spans

def _get_value_safe(val):
    if pd.isna(val):
        return "-"
    try:
        return str(val)
    except Exception:
        return "-"

def _final_day_for_step(i, row, df_vtt):
    try:
        if i == 0:
            return int(row['1 Day Customer Order']) if row is not None and '1 Day Customer Order' in df_vtt.columns else 0
        if i == 1:
            return int(row['2 Day ILN Order']) if row is not None and '2 Day ILN Order' in df_vtt.columns else 0
        if i == 2:
            return int(row['3.2 First Receipt Days']) if row is not None and '3.2 First Receipt Days' in df_vtt.columns else 0
        if i == 3:
            return int(row['4.3 Packaging préparation & loading']) if row is not None and '4.3 Packaging préparation & loading' in df_vtt.columns else 0
        if i == 4:
            return int(row['5.3 Transport ILN to POL']) if row is not None and '5.3 Transport ILN to POL' in df_vtt.columns else 0
        if i == 5:
            return int(row['6 First Day to POL']) if row is not None and '6 First Day to POL' in df_vtt.columns else 0
        if i == 6:
            return int(row['7 Cutt off']) if row is not None and '7 Cutt off' in df_vtt.columns else 0
        if i == 7:
            return int(row['8 ETD']) if row is not None and '8 ETD' in df_vtt.columns else 0
        if i == 8:
            if row is not None and '9 ETD> ETA' in df_vtt.columns:
                return int(row['9 ETD> ETA'])
            if row is not None and '9 ETD>ETA' in df_vtt.columns:
                return int(row['9 ETD>ETA'])
            return 0
        if i == 9:
            if row is not None and '10 Days flexibility 1' in df_vtt.columns and pd.notna(row['10 Days flexibility 1']):
                return int(row['10 Days flexibility 1'])
            # derive: base(9) + 1 + buffer
            base = None
            if row is not None and '9 ETD> ETA' in df_vtt.columns:
                base = row['9 ETD> ETA']
            elif row is not None and '9 ETD>ETA' in df_vtt.columns:
                base = row['9 ETD>ETA']
            bnum = pd.to_numeric(base, errors='coerce') if base is not None else float('nan')
            if pd.isna(bnum):
                m = re.findall(r"[-+]?\d*\.?\d+", str(base)) if base is not None else []
                bnum = float(m[0]) if m else float('nan')
            plus = _coerce_to_int(row['Time for security']) if row is not None and 'Time for security' in df_vtt.columns else 0
            return int(float(bnum)) + 1 + int(plus) if not pd.isna(bnum) else 0
        if i == 10:
            return int(row['11 Days flexibility 2']) if row is not None and '11 Days flexibility 2' in df_vtt.columns else 0
        if i == 11:
            if row is not None and '12 Customs Clearance' in df_vtt.columns:
                return int(row['12 Customs Clearance'])
            if row is not None and '12 Customs clearence' in df_vtt.columns:
                return int(row['12 Customs clearence'])
            return 0
        if i == 12:
            return int(row['13 Transport to Plant']) if row is not None and '13 Transport to Plant' in df_vtt.columns else 0
        if i == 13:
            return int(row['14 Rounding']) if row is not None and '14 Rounding' in df_vtt.columns else 0
        if i == 14:
            return int(row['15 Due Date']) if row is not None and '15 Due Date' in df_vtt.columns else 0
        if i == 15:
            return int(row['16 Manufacturing']) if row is not None and '16 Manufacturing' in df_vtt.columns else 0
        return 0
    except Exception:
        return 0

def _day_plus_for_step(i, row, df_vtt):
    if i in (0,1,5,6,7):
        return 0
    if i == 2:
        return _coerce_to_int(row['3 .1 Time of Recept in ILN']) if row is not None and '3 .1 Time of Recept in ILN' in df_vtt.columns else 0
    if i == 3:
        return _coerce_to_int(row['4.2 Packaging préparation & loading']) if row is not None and '4.2 Packaging préparation & loading' in df_vtt.columns else 0
    if i == 4:
        return _coerce_to_int(row['5.2 Transport ILN to POL']) if row is not None and '5.2 Transport ILN to POL' in df_vtt.columns else 0
    if i == 8:
        return _coerce_to_int(row['Transit time']) if row is not None and 'Transit time' in df_vtt.columns else 0
    if i == 9:
        return _coerce_to_int(row['Time for security']) if row is not None and 'Time for security' in df_vtt.columns else 0
    if i == 10:
        return _coerce_to_int(row['Time for security2 buffer']) if row is not None and 'Time for security2 buffer' in df_vtt.columns else 0
    if i == 11:
        return _coerce_to_int(row['Cust.']) if row is not None and 'Cust.' in df_vtt.columns else 0
    if i == 12:
        return _coerce_to_int(row['Trpt POD/PFI vers Usine']) if row is not None and 'Trpt POD/PFI vers Usine' in df_vtt.columns else 0
    if i == 13:
        if row is not None and 'Round.' in df_vtt.columns:
            return _coerce_to_int(row['Round.'])
        if row is not None and 'Round' in df_vtt.columns:
            return _coerce_to_int(row['Round'])
        return 0
    if i == 14:
        return 7
    if i == 15:
        return 7
    return 0

def _day_value_for_step(i, row, df_vtt):
    # Returns the Day column value per step as string
    try:
        if i == 0:
            return _get_value_safe(row['1 Day Customer Order']) if row is not None and '1 Day Customer Order' in df_vtt.columns else '-'
        if i == 1:
            return _get_value_safe(row['2 Day ILN Order']) if row is not None and '2 Day ILN Order' in df_vtt.columns else '-'
        if i == 2:
            return _get_value_safe(row['3 First Receipt Days']) if row is not None and '3 First Receipt Days' in df_vtt.columns else 'No hay datos para la combinación POL/POD seleccionada'
        if i == 3:
            return _get_value_safe(row['4.1 Packaging préparation & loading']) if row is not None and '4.1 Packaging préparation & loading' in df_vtt.columns else '-'
        if i == 4:
            return _get_value_safe(row['5.1 Transport ILN to POL']) if row is not None and '5.1 Transport ILN to POL' in df_vtt.columns else '-'
        if i == 5:
            return _get_value_safe(row['6 First Day to POL']) if row is not None and '6 First Day to POL' in df_vtt.columns else '-'
        if i == 6:
            return _get_value_safe(row['7 Cutt off']) if row is not None and '7 Cutt off' in df_vtt.columns else '-'
        if i == 7:
            return _get_value_safe(row['8 ETD']) if row is not None and '8 ETD' in df_vtt.columns else '-'
        if i == 8:
            # Day for step 9 is ETD
            return _get_value_safe(row['8 ETD']) if row is not None and '8 ETD' in df_vtt.columns else '-'
        if i == 9:
            # base final of 9 + 1
            base = _final_day_for_step(8, row, df_vtt)
            return str(base + 1) if base else '-'
        if i == 10:
            base = _final_day_for_step(9, row, df_vtt)
            return str(base + 1) if base else '-'
        if i == 11:
            base = _final_day_for_step(10, row, df_vtt)
            return str(base + 1) if base else '-'
        if i == 12:
            base = _final_day_for_step(11, row, df_vtt)
            return str(base + 1) if base else '-'
        if i == 13:
            base = _final_day_for_step(12, row, df_vtt)
            return str(base + 1) if base else '-'
        if i == 14:
            base = _final_day_for_step(13, row, df_vtt)
            return str(base + 1) if base else '-'
        if i == 15:
            base = _final_day_for_step(14, row, df_vtt)
            return str(base + 1) if base else '-'
        return '-'
    except Exception:
        return '-'

def build_excel_workbook(row, df_vtt, selected_pol, selected_pod, time_labels, headers, timeline_days):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Timeline'

    # styles
    bold = Font(bold=True)
    hfill = _hex_to_fill('#f5f5f5')
    weekfill = _hex_to_fill('#fffbe6')
    weekendfill = _hex_to_fill('#ffd6d6')
    weekdayfill = _hex_to_fill('#e3eafc')
    paintfill = _hex_to_fill('#90ee90')
    darkbluefill = _hex_to_fill('#00008b')
    border = Border(left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'), top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD'))

    r = 1
    ws.cell(row=r, column=1, value='POL:').font = bold; ws.cell(row=r, column=2, value=selected_pol)
    ws.cell(row=r, column=3, value='POD:').font = bold; ws.cell(row=r, column=4, value=selected_pod)
    r += 1
    # Info row if available
    if row is not None:
        info_pairs = []
        for label, colname in [('ID','ID'), ('Carrier','Carrier'), ('Shipper', df_vtt.columns[10] if len(df_vtt.columns) > 10 else None), ('ILN', df_vtt.columns[8] if len(df_vtt.columns) > 8 else None), ('PLANT','Name Destin Site'), ('E/D','Expiration Date')]:
            val = row.get(colname, '') if (colname and colname in df_vtt.columns) else ''
            info_pairs.append((label, val))
        c = 1
        for label, val in info_pairs:
            ws.cell(row=r, column=c, value=f'{label}:').font = bold
            ws.cell(row=r, column=c+1, value=str(val))
            c += 2
        r += 1

    r += 1
    # Week header merged cells
    start_col = 5  # dates start at column 5
    ws.cell(row=r, column=1, value='')
    # leave placeholders for Steps/Day/Day+/Final Day
    spans = _compute_week_spans(timeline_days)
    c = start_col
    for week, span in spans:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c+span-1)
        cell = ws.cell(row=r, column=c, value=f'W{week}')
        cell.fill = weekfill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        # borders
        for cc in range(c, c+span):
            ws.cell(row=r, column=cc).border = border
        c += span
    r += 1

    # Header row (Steps, Day, Day+, Final Day, then dates)
    for ci, h in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=ci, value=h)
        cell.fill = hfill
        cell.font = bold
        cell.border = border
        cell.alignment = Alignment(horizontal='center') if ci > 1 else Alignment(horizontal='left')
    for idx, d in enumerate(timeline_days):
        ci = start_col + idx
        cell = ws.cell(row=r, column=ci, value=d.strftime('%d-%b'))
        cell.fill = weekendfill if d.weekday() in (5,6) else weekdayfill
        cell.border = border
        # Rotate text vertically (top-to-bottom)
        cell.alignment = Alignment(horizontal='center', vertical='bottom', textRotation=90)
    r += 1

    # Row content
    for i, label in enumerate(time_labels):
        # Reduce Excel row height for step rows (~35% smaller than default ~15pt)
        try:
            ws.row_dimensions[r+i].height = 10.5
        except Exception:
            pass
        ws.cell(row=r+i, column=1, value=label).fill = hfill
        ws.cell(row=r+i, column=1).font = bold
        ws.cell(row=r+i, column=1).border = border
        ws.cell(row=r+i, column=1).alignment = Alignment(horizontal='left')

        # Day
        day_val = _day_value_for_step(i, row, df_vtt)
        ws.cell(row=r+i, column=2, value=day_val).border = border
        ws.cell(row=r+i, column=2).alignment = Alignment(horizontal='center')

        # Day+
        day_plus = _day_plus_for_step(i, row, df_vtt)
        ws.cell(row=r+i, column=3, value=str(day_plus) if day_plus != 0 else ("0" if i in (3,4,8,9,11,12,13,14,15) else "0" if day_plus==0 else "-")).border = border
        ws.cell(row=r+i, column=3).alignment = Alignment(horizontal='center')

        # Final Day
        fday = _final_day_for_step(i, row, df_vtt)
        ws.cell(row=r+i, column=4, value=str(fday) if fday != 0 else "-").border = border
        ws.cell(row=r+i, column=4).alignment = Alignment(horizontal='center')

        # Paint date cells
        paint_len = day_plus if (isinstance(day_plus, int) and day_plus > 0) else 1
        if i in (0,1,5,6,7):
            paint_len = 1  # Day+ = 0 -> only final day for these steps
        start_idx = max(1, fday - paint_len + 1) if fday else 0
        for idx, d in enumerate(timeline_days, start=0):
            ci = start_col + idx
            cell = ws.cell(row=r+i, column=ci, value="")
            cell.border = border
            # weekend shading
            if d.weekday() in (5,6):
                cell.fill = weekendfill
            # paint range overrides shade
            if fday and start_idx <= (idx+1) <= fday:
                if i == 8:
                    cell.fill = darkbluefill
                else:
                    cell.fill = paintfill

    # Column widths
    ws.column_dimensions['A'].width = 36
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 12
    for k in range(start_col, start_col + len(timeline_days)):
        ws.column_dimensions[get_column_letter(k)].width = 4

    # KPIs block under the table
    rr = r + len(time_labels) + 2
    ws.cell(row=rr, column=1, value='Parts Vanning').font = bold
    if row is not None and 'Parts Vanning' in df_vtt.columns:
        ws.cell(row=rr, column=2, value=str(row['Parts Vanning']))
    rr += 1
    ws.cell(row=rr, column=1, value='Transit Time').font = bold
    # recompute total_tt similar to UI
    total_tt = None
    if row is not None:
        t1 = pd.to_numeric(row.get('Transit time', None), errors='coerce') if 'Transit time' in df_vtt.columns else None
        t2 = pd.to_numeric(row.get('Time for security', None), errors='coerce') if 'Time for security' in df_vtt.columns else None
        parts = [v for v in (t1, t2) if v is not None and pd.notna(v)]
        if parts:
            total_tt = float(sum(parts))
    if total_tt is not None and pd.notna(total_tt):
        ws.cell(row=rr, column=2, value=int(total_tt) if abs(total_tt - int(total_tt)) < 1e-9 else round(total_tt,2))
    rr += 1
    ws.cell(row=rr, column=1, value='Customer Leadtime').font = bold
    if row is not None and 'Cust. Leadtime' in df_vtt.columns:
        ws.cell(row=rr, column=2, value=str(row['Cust. Leadtime']))
    rr += 1
    ws.cell(row=rr, column=1, value='Customer Safety STOCK').font = bold
    if row is not None and 'Safety stock' in df_vtt.columns:
        ws.cell(row=rr, column=2, value=str(row['Safety stock']))

    # Return bytes
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()



# --- Mostrar E/D debajo del timeline y antes del botón Generate files ---
st.markdown("<div style='height: 8px'></div>", unsafe_allow_html=True)
try:
    _ed_display = ""
    if row is not None and 'Expiration Date' in df_vtt.columns:
        _exp_date = row.get('Expiration Date', '')
        if pd.notnull(_exp_date):
            if isinstance(_exp_date, (pd.Timestamp, datetime)):
                _ed_display = _exp_date.strftime('%d/%m/%Y')
            else:
                try:
                    _ed_display = pd.to_datetime(_exp_date).strftime('%d/%m/%Y')
                except Exception:
                    _ed_display = str(_exp_date)
    st.markdown(render_box('E/D', _ed_display), unsafe_allow_html=True)
except Exception:
    pass

# --- Single 'Generate files' button, then show download buttons in English ---
st.markdown("<hr style='margin:32px 0;'>", unsafe_allow_html=True)
if st.button("Generate files", key="generate_files"):
    excel_bytes = build_excel_workbook(
        row=row,
        df_vtt=df_vtt,
        selected_pol=st.session_state.get('pol_select',''),
        selected_pod=st.session_state.get('pod_select',''),
        time_labels=time_labels,
        headers=headers,
        timeline_days=timeline_days,
    )
    excel_b64 = base64.b64encode(excel_bytes).decode('utf-8') if excel_bytes else ''
    st.markdown(f"""
    <div style='width:100%; display:flex; justify-content:center; align-items:center; margin:32px 0;'>
        <a id='excelBtn' href='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}' download='VTT_FULL_VIEW.xlsx' style='display:inline-block;background:#1f77b4;color:#fff;border:none;border-radius:6px;padding:10px 16px;font-size:18px;cursor:pointer;text-decoration:none;margin-right:24px;'>Excel file</a>
        <button id='imgBtn' style='display:inline-block;background:#1f77b4;color:#fff;border:none;border-radius:6px;padding:10px 16px;font-size:18px;cursor:pointer;'>Image</button>
    </div>
    """, unsafe_allow_html=True)
    components.html(
        """
        <script src='https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js'></script>
        <script>
        (function(){
            function parentDoc(){
                try { return window.parent && window.parent.document ? window.parent.document : document; } catch(e){ return document; }
            }
            function getBtn(){ return parentDoc().getElementById('imgBtn'); }
            function getArea(){
                var d = parentDoc();
                // Prefer off-screen composite that includes all data and KPIs
                return d.getElementById('timeline_capture') || d.getElementById('timeline_capture_table') || d.body || document.body;
            }
            function ensureHtml2CanvasReady(cb){
                if (window.html2canvas) return cb();
                var tries = 0; (function waitLib(){
                    if (window.html2canvas) return cb();
                    if (++tries > 50) { alert('html2canvas no cargó.'); return; }
                    setTimeout(waitLib, 100);
                })();
            }
            function bind(){
                var button = getBtn();
                if (!button) { setTimeout(bind, 250); return; }
                button.addEventListener('click', function(){
                    ensureHtml2CanvasReady(function(){
                        var area = getArea();
                        if (!area) { alert('No se encontró el área visual para capturar.'); return; }
                        window.html2canvas(area, { backgroundColor:'#fff', useCORS:true, allowTaint:true, scale:2 })
                        .then(function(canvas){
                            canvas.toBlob(function(blob){
                                if(!blob){ alert('No se pudo generar la imagen'); return; }
                                var d = parentDoc();
                                var url = URL.createObjectURL(blob);
                                var a = d.createElement('a');
                                var ts = new Date().toISOString().slice(0,19).replace(/[.:T]/g,'-');
                                a.href = url;
                                a.download = 'VTTFULL_VIEW_' + ts + '.png';
                                d.body.appendChild(a);
                                a.click();
                                setTimeout(function(){ d.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
                            }, 'image/png', 0.95);
                        })
                        .catch(function(err){ alert('Error capturando imagen: ' + err); });
                    });
                });
            }
            bind();
        })();
        </script>
        """,
        height=10,
    )
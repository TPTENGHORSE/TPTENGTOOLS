import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# Load data from Excel with correct path
excel_path = os.path.join(os.path.dirname(__file__), "Tool_VTT_Horse_Unlocked.xlsm")
df = pd.read_excel(excel_path, sheet_name="VTT actif")

# Renombrar columnas clave
df = df.rename(columns={
    df.columns[5]: "ID",           # Columna F
    df.columns[14]: "AILN",        # Columna O
    df.columns[23]: "Shipper",     # Columna X
    df.columns[11]: "Consignee",   # Columna L
    df.columns[37]: "CDD_PFI"      # Columna AL
})

# Filtrar y limpiar IDs
df = df[df["ID"].astype(str).str.contains("-Actif", na=False)].copy()
df["ID_limpio"] = df["ID"].str.replace("-Actif", "", regex=False)

# --- STREAMLIT INTERFACE ---
st.set_page_config(layout="wide")

# Reducir el padding superior de la app para pegar el UI a la barra del navegador
st.markdown("""
    <style>
    .main .block-container {
        padding-top: 0.5rem !important;
    }
    header[data-testid="stHeader"] {
        height: 0px !important;
        min-height: 0px !important;
        padding: 0 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# First: technical sheet (left), then timeline (right)
col_info, col_timeline = st.columns([2, 3], gap="large")

# TECHNICAL SHEET BY ID
with col_info:
    # Usar columnas para alinear selectbox y datos técnicos en la parte superior
    col_select, col_tech = st.columns([1, 2], gap="small")
    with col_select:
        # Mostrar solo el primer ID como opción al iniciar
        id_list = [i for i in df["ID_limpio"].unique().tolist() if i != "Key(ID)"]
        if 'id_select' not in st.session_state:
            st.session_state['id_select'] = id_list[0]
        selected_id = st.selectbox("Select an ID", id_list, key="id_select", index=0, label_visibility="collapsed")
    with col_tech:
        row = df[df["ID_limpio"] == selected_id].iloc[0]
        tech_labels = [
            ("Name ILN", row['AILN']),
            ("Name Shipper", row['Shipper']),
            ("Name PFI/CDD", row['CDD_PFI']),
            ("Name Site", row['Consignee'])
        ]
        table_html = """
        <table style='width:100%; border-collapse:collapse; margin-top:0;'>
            <tbody>
        """
        for custom, value in tech_labels:
            table_html += f"<tr><td style='font-weight:bold; padding:4px 8px; border-bottom:1px solid #eee; width:140px;'>{custom}</td>"
            table_html += f"<td style='padding:4px 8px; border-bottom:1px solid #eee;'>{value}</td></tr>"
        table_html += """
            </tbody>
        </table>
        """
        st.markdown(table_html, unsafe_allow_html=True)

    # Zona de tiempos: 4 columnas fijas + columnas dinámicas según timeline
    time_rows = 19
    time_cols_fixed = 4
    # Timeline dinámico
    start_date = datetime.today()
    num_days = st.slider("Number of days to display", min_value=7, max_value=120, value=min(30, 120), step=1, key="days_slider_info")
    timeline_days = [start_date + timedelta(days=i) for i in range(num_days)]
    time_cols = time_cols_fixed + num_days
    
    time_table_html = """
    <table style='width:100%; border-collapse:collapse; margin-top:24px;'>
        <thead><tr>"""
    # Encabezados fijos
    headers = ["Vanning", "Day", "Day+", "Final Day"]
    for h in headers:
        time_table_html += f"<th style='padding:6px 8px; border:1px solid #eee; min-width:90px; width:90px; text-align:center; background:#f5f5f5'>{h}</th>"
    # Encabezados dinámicos (timeline)
    for day in timeline_days:
        time_table_html += f"<th style='padding:0 1px; border:1px solid #eee; min-width:28px; width:28px; text-align:center; background:#e3eafc; font-size:9px'>{day.strftime('%d-%b')}</th>"
    time_table_html += "</tr></thead><tbody>"

    time_labels = [
        "Day Customer Order",
        "", # blank because needs formula
        "", # blank because needs formula
        "Packaging préparation & loading",
        "Transport ILN to POL",
        "First Day to POL",
        "Cut off",
        "ETD",
        "Transit time (ETD => ETA)",
        "Days flexibility 1",
        "Days flexibility 2",
        "Customs clearence",
        "",  # blank because needs formula
        "",  # blank because needs formula
        "Transport to plant",
        "Rounding",
        "Due Date",
        "Manufacturing"
    ]
    time_rows = len(time_labels)
    while len(time_labels) < time_rows:
        time_labels.append("")

    zona_col_widths = [90]*4 + [28]*(time_cols-4)
    row_values = [[None]*time_cols for _ in range(time_rows)]
    for i in range(time_rows):
        time_table_html += "<tr style='height:30px;'>"
        for j in range(time_cols):
            if j < 4:
                # Suma para la celda 'Final Day' de la fila 'Day Customer Order'
                if j == 3 and i == 0:
                    try:
                        day = float(row_values[i][1]) if row_values[i][1] not in (None, "", "-") else 0
                    except:
                        day = 0
                    try:
                        dayplus = float(row_values[i][2]) if row_values[i][2] not in (None, "", "-") else 0
                    except:
                        dayplus = 0
                    suma = day + dayplus
                    cell_content = str(int(suma)) if suma == int(suma) else str(suma)
                # Suma para la celda 'Final Day' de la fila 'Day ILN Order'
                elif j == 3 and i == 1:
                    try:
                        day = float(row_values[i][1]) if row_values[i][1] not in (None, "", "-") else 0
                    except:
                        day = 0
                    try:
                        dayplus = float(row_values[i][2]) if row_values[i][2] not in (None, "", "-") else 0
                    except:
                        dayplus = 0
                    suma = day + dayplus
                    cell_content = str(int(suma)) if suma == int(suma) else str(suma)
                elif j == 3 and i in [0, 1]:
                    cell_content = "0"
                elif i == 2 and j == 0:
                    if str(row['AILN']).strip() == "SFK Freight Forwarder":
                        cell_content = ""
                    else:
                        cell_content = "Time of recept in AILN"
                elif i == 3 and j == 0:
                    cell_content = "Packaging préparation & loading"
                elif i == 0 and j == 1:  # Packaging préparation & loading, Day column

 # Implement BUSCARX logic: find in df where ID == selected_id, return column N (index 13)
                    match_row = df[df['ID_limpio'] == selected_id]
                    if not match_row.empty:
                        value = match_row.iloc[0, 13]  # Column N
                        cell_content = value if pd.notnull(value) else "-"
                    else:
                        cell_content = "-"
                elif i == 1 and j == 1:  # Transport ILN to POL, Day column
                    if str(row['AILN']).strip() == "SFK Freight Forwarder":
                        cell_content = "0"
                    else:
                        match_row = df[df['ID_limpio'] == selected_id]
                        if not match_row.empty:
                            value = match_row.iloc[0, 16]  # Columna Q
                            cell_content = value if pd.notnull(value) else "-"
                        else:
                            cell_content = "-"
                elif i == 1 and j == 0:  # Vanning, segunda fila
                    if str(row['AILN']).strip() == "SFK Freight Forwarder":
                        cell_content = ""
                    else:
                        cell_content = "Day ILN Order"
                elif i == 2 and j == 0:  # Vanning, fila 2
                    if str(row['AILN']).strip() == "SFK Freight Forwarder":
                        cell_content = ""
                    else:
                        cell_content = "Time of recept in AILN"
                elif j == 0:
                    cell_content = time_labels[i]
                else:
                    cell_content = ""
                # Guardar el valor de la celda para sumas posteriores
                row_values[i][j] = cell_content
            else:
                # Dinámicas: celdas vacías por defecto (puedes personalizar)
                cell_content = ""
            # Ajustar estilos: columnas de fechas mucho más angostas
            if j < 4:
                cell_style = f"padding:0 2px; border:1px solid #eee; min-width:90px; width:90px; height:30px; text-align:center; vertical-align:middle; font-size:10px;"
            else:
                cell_style = f"padding:0 1px; border:1px solid #eee; min-width:28px; width:28px; height:30px; text-align:center; vertical-align:middle; font-size:9px;"
            time_table_html += f"<td style='{cell_style}'>{cell_content}</td>"
        time_table_html += "</tr>"
    time_table_html += """
        </tbody>
    </table>
    """
    st.markdown(time_table_html, unsafe_allow_html=True)

# TIMELINE HORIZONTAL (user can select range)
with col_timeline:
    start_date = datetime.today()
    # Usar el mismo num_days definido en col_info para evitar duplicidad de slider
    timeline_days = [start_date + timedelta(days=i) for i in range(num_days)]

    # Placeholder for future transit bars
    st.markdown("<div style='height: 60px;'></div>", unsafe_allow_html=True)

# Prevent Streamlit from jumping to the top when changing selectbox
st.markdown("""
    <style>
    section.main > div:has(div[data-testid='stVerticalBlock']) {
        scroll-behavior: auto !important;
    }
    </style>
    """, unsafe_allow_html=True)

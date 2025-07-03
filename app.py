# app.py
import streamlit as st
import importlib.util
import sys
import os

st.set_page_config(page_title="Transport Engineering Tools", layout="centered")

# Sidebar menu as buttons with session state to persist selection
st.sidebar.image("logo.png", width=120)
menu_options = ["Main menu", "Empower3D", "VTT", "Quotation Tool"]
if "active_menu" not in st.session_state:
    st.session_state["active_menu"] = menu_options[0]
for option in menu_options:
    if st.sidebar.button(option, key=f"menu_{option}"):
        st.session_state["active_menu"] = option
menu = st.session_state["active_menu"]

if menu == "Main menu":
    # Use Streamlit layout only, no custom divs, for perfect centering and no scroll
    st.empty()  # Clear any previous content
    st.write("")  # Spacer
    # Center logo above the title using columns for perfect horizontal centering
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.image("logo.png", width=400)
    st.markdown("<h1 style='text-align: center;'>Transport Engineering Tools</h1>", unsafe_allow_html=True)
    st.write("")  # Spacer
    st.write("")  # More space if needed

elif menu == "Empower3D":
    empower3d_path = os.path.join(os.path.dirname(__file__), "Packaging", "Empower3D.py")
    if os.path.exists(empower3d_path):
        spec = importlib.util.spec_from_file_location("Empower3D", empower3d_path)
        empower3d = importlib.util.module_from_spec(spec)
        sys.modules["Empower3D"] = empower3d
        spec.loader.exec_module(empower3d)
        empower3d.run()  # Run the function to show the app
    else:
        st.error("Empower3D.py not found")

elif menu == "Empower3D+":
    empower3dplus_path = os.path.join(os.path.dirname(__file__), "Packaging", "Empower3D+.py")
    if os.path.exists(empower3dplus_path):
        spec = importlib.util.spec_from_file_location("Empower3Dplus_module", empower3dplus_path)
        empower3dplus_module = importlib.util.module_from_spec(spec)
        sys.modules["Empower3Dplus_module"] = empower3dplus_module
        spec.loader.exec_module(empower3dplus_module)
        empower3dplus_module.main()
    else:
        st.error("Empower3D+.py not found")

elif menu == "VTT":
    st.markdown("""
        <div style='display:flex; align-items:center; gap:12px; margin-bottom:0.5rem;'>
            <img src='https://img.icons8.com/ios-filled/50/000000/cargo-ship.png' width='36' height='36' style='margin-bottom:0;'>
            <span style='font-size:2rem; font-weight:700;'>VTT Tool</span>
        </div>
    """, unsafe_allow_html=True)
    vtt_path = os.path.join(os.path.dirname(__file__), "VTT Tool", "VTT.py")
    if os.path.exists(vtt_path):
        spec = importlib.util.spec_from_file_location("VTT", vtt_path)
        vtt = importlib.util.module_from_spec(spec)
        sys.modules["VTT"] = vtt
        spec.loader.exec_module(vtt)
    else:
        st.error("VTT.py not found")

elif menu == "HorseLuis":
    horseluis_path = os.path.join(os.path.dirname(__file__), "ChatbotIA", "HorseLuis.py")
    if os.path.exists(horseluis_path):
        spec = importlib.util.spec_from_file_location("HorseLuis", horseluis_path)
        horseluis = importlib.util.module_from_spec(spec)
        sys.modules["HorseLuis"] = horseluis
        spec.loader.exec_module(horseluis)
        horseluis.run()
    else:
        st.error("HorseLuis.py not found")

elif menu == "Quotation Tool":
    st.markdown("<h2>Quotation Tool</h2>", unsafe_allow_html=True)
    from Quotations.Quotation_tool import procesar_quotation
    import pandas as pd
    import os

    # Cargar archivos backend una sola vez, con verificación robusta
    import os
    import pandas as pd
    import streamlit as st
    @st.cache_data
    def load_backend_files():
        base_path = os.path.join(os.path.dirname(__file__), "Quotations", "Dataframe")
        files = {
            "Base_EMB": os.path.join(base_path, "Base_EMB.xlsx"),
            "Inland": os.path.join(base_path, "cifrados Overseas-Inland.xlsx"),
            "Rates": os.path.join(base_path, "RATES_04_2025.xlsx")
        }
        data = {}
        for name, path in files.items():
            if not os.path.exists(path):
                st.error(f"❌ File not found: {path}")
                continue
            try:
                data[name] = pd.read_excel(path)
            except Exception as e:
                st.error(f"⚠️ Error loading {name}: {e}")
        return data

    backend_data = load_backend_files()
    if not all(k in backend_data for k in ["Base_EMB", "Inland", "Rates"]):
        st.stop()
    base_emb_df = backend_data["Base_EMB"]
    inland_df = backend_data["Inland"]
    rates_df = backend_data["Rates"]

    st.write("Sube el archivo de Plantilla_Quotation:")
    plantilla_file = st.file_uploader("Plantilla_Quotation.xlsx", type=["xlsx"])

    if plantilla_file:
        plantilla_df = pd.read_excel(plantilla_file)
        st.write("Columnas detectadas en Plantilla_Quotation:")
        st.write(list(plantilla_df.columns))
    else:
        plantilla_df = None

    if st.button("Procesar cotización"):
        if plantilla_df is None:
            st.error("Debes subir el archivo Plantilla_Quotation.")
        else:
            df_result = procesar_quotation(plantilla_df, base_emb_df, inland_df, rates_df)
            st.success("Cotización procesada.")
            st.dataframe(df_result)
            st.download_button("Descargar resultado", data=df_result.to_excel(index=False, engine='openpyxl'), file_name="Quotation_Result.xlsx")


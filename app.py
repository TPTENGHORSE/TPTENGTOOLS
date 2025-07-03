# app.py
import streamlit as st
import importlib.util
import sys
import os

st.set_page_config(page_title="Transport Engineering Tools", layout="centered")

# Sidebar menu as buttons with session state to persist selection
st.sidebar.image("logo.png", width=120)
menu_options = ["Main menu", "Empower3D"]
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



# app.py
import streamlit as st
import importlib.util
import sys
import os

st.set_page_config(page_title="Transport Engineering Tools", layout="wide")

# Sidebar menu as buttons with session state to persist selection
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
st.sidebar.image(logo_path, width=120)
menu_options = ["VTTs", "Empower3D"]
if "active_menu" not in st.session_state:
    st.session_state["active_menu"] = menu_options[0]
for option in menu_options:
    if st.sidebar.button(option, key=f"menu_{option}"):
        st.session_state["active_menu"] = option
menu = st.session_state["active_menu"]

if menu == "Empower3D":
    empower3d_path = os.path.join(os.path.dirname(__file__), "Packaging", "Empower3D.py")
    if os.path.exists(empower3d_path):
        spec = importlib.util.spec_from_file_location("Empower3D", empower3d_path)
        empower3d = importlib.util.module_from_spec(spec)
        sys.modules["Empower3D"] = empower3d
        spec.loader.exec_module(empower3d)
        empower3d.run()  # Run the function to show the app
    else:
        st.error("Empower3D.py not found")

elif menu == "VTTs":
    # Integrate the VTT timeline app (VTT Tool/VTT2.py) into this main app
    vtt2_path = os.path.join(os.path.dirname(__file__), "VTT Tool", "VTT2.py")
    if os.path.exists(vtt2_path):
        try:
            # Avoid multiple set_page_config calls by temporarily no-op'ing it
            orig_set_pc = getattr(st, "set_page_config", None)
            if callable(orig_set_pc):
                st.set_page_config = lambda *args, **kwargs: None  # type: ignore
            spec = importlib.util.spec_from_file_location("VTT2", vtt2_path)
            vtt2 = importlib.util.module_from_spec(spec)
            sys.modules["VTT2"] = vtt2
            assert spec and spec.loader
            spec.loader.exec_module(vtt2)
        except Exception as e:
            st.error(f"Error loading VTT2 app: {e}")
            st.exception(e)
        finally:
            # Restore original set_page_config
            if 'orig_set_pc' in locals() and callable(orig_set_pc):
                st.set_page_config = orig_set_pc  # type: ignore
    else:
        st.error("VTT2.py not found in 'VTT Tool' folder")

    # No other menu entries
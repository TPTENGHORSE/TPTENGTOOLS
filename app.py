# app.py
import streamlit as st
import importlib.util
import sys
import os
import re
import tempfile
import urllib.request
from datetime import datetime

from Quotations.qtool_loader import load_input_template
from Quotations import generate_quote as gq
from Quotations.generate_quote import build_output

st.set_page_config(page_title="Transport Engineering Tools", layout="wide")

# Sidebar menu as buttons with session state to persist selection
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
st.sidebar.image(logo_path, width=120)
menu_options = ["VTTs", "Empower3D", "MyQuotes"]
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

elif menu == "MyQuotes":
    st.title("MyQuotes")
    st.caption("Upload Quotation Template _INPUT and download the generated Horse_Quotation.")

    def _secret_or_env(name: str) -> str:
        try:
            if name in st.secrets:
                return str(st.secrets[name]).strip()
        except Exception:
            pass
        return str(os.environ.get(name, "")).strip()

    def ensure_qtool_data(runtime_dir: str) -> tuple[bool, str, str | None]:
        """Ensure QUOTATION TOOL DATA.xlsx exists in runtime_dir.
        Priority:
        1) Repository-bundled file: Quotations/QUOTATION TOOL DATA.xlsx
        2) SharePoint URL from secrets/env
        3) Existing local file in generate_quote.QTOOL_DIR
        """
        os.makedirs(runtime_dir, exist_ok=True)
        dst = os.path.join(runtime_dir, "QUOTATION TOOL DATA.xlsx")

        # 1) Repository-bundled file (recommended when deploying with GitHub)
        repo_data = os.path.join(os.path.dirname(__file__), "Quotations", "QUOTATION TOOL DATA.xlsx")
        if os.path.exists(repo_data):
            return True, "Database loaded from Quotations/QUOTATION TOOL DATA.xlsx.", repo_data

        # 2) SharePoint source (recommended for Streamlit Cloud)
        sp_url = _secret_or_env("QTOOL_DATA_SHAREPOINT_URL")
        sp_token = _secret_or_env("QTOOL_DATA_BEARER_TOKEN")
        if sp_url:
            try:
                if not os.path.exists(dst):
                    headers = {"User-Agent": "QFLOW/1.0"}
                    if sp_token:
                        headers["Authorization"] = f"Bearer {sp_token}"
                    req = urllib.request.Request(sp_url, headers=headers)
                    with urllib.request.urlopen(req, timeout=90) as r:
                        content = r.read()
                    with open(dst, "wb") as f:
                        f.write(content)
                return True, "Database loaded from SharePoint.", dst
            except Exception as e:
                return False, f"Could not download QUOTATION TOOL DATA from SharePoint: {e}", None

        # 3) Local fallback (desktop execution)
        local_qtool_dir = getattr(gq, "QTOOL_DIR", "")
        local_data = os.path.join(local_qtool_dir, "QUOTATION TOOL DATA.xlsx") if local_qtool_dir else ""
        if local_data and os.path.exists(local_data):
            return True, "Using local QUOTATION TOOL DATA.", local_data

        return False, (
            "QUOTATION TOOL DATA was not found. Add Quotations/QUOTATION TOOL DATA.xlsx to the repository or configure "
            "QTOOL_DATA_SHAREPOINT_URL in Streamlit secrets (and optionally QTOOL_DATA_BEARER_TOKEN if authentication is required)."
        ), None

    def next_qflow_output_path(directory: str) -> str:
        """Build next output file path: Horse_Quotation_YYYYMMDD_descarga_N.xlsx."""
        os.makedirs(directory, exist_ok=True)
        date_tag = datetime.now().strftime("%Y%m%d")
        pattern = re.compile(rf"^Horse_Quotation_{date_tag}_descarga_(\d+)\.xlsx$", re.IGNORECASE)
        max_n = 0
        try:
            for name in os.listdir(directory):
                m = pattern.match(name)
                if m:
                    try:
                        n = int(m.group(1))
                        if n > max_n:
                            max_n = n
                    except Exception:
                        pass
        except FileNotFoundError:
            pass
        next_n = max_n + 1
        return os.path.join(directory, f"Horse_Quotation_{date_tag}_descarga_{next_n}.xlsx")

    uploaded = st.file_uploader(
        "Select Quotation Template _INPUT (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
    )

    if uploaded is not None:
        input_name = uploaded.name or ""
        if "Quotation Template _INPUT" not in input_name:
            st.warning("The file should be named Quotation Template _INPUT (or contain that text in the filename).")

        if st.button("Generate Horse_Quotation", type="primary"):
            try:
                runtime_qtool_dir = os.path.join(tempfile.gettempdir(), "qflow_qtool")
                ok_db, db_msg, db_path = ensure_qtool_data(runtime_qtool_dir)
                if not ok_db:
                    st.error(db_msg)
                    st.stop()

                # Point generator to a writable runtime directory where QTOOL data exists.
                gq.QTOOL_DIR = runtime_qtool_dir

                # If local fallback was used from a different folder, copy once to runtime dir.
                if db_path and os.path.abspath(db_path) != os.path.abspath(os.path.join(runtime_qtool_dir, "QUOTATION TOOL DATA.xlsx")):
                    with open(db_path, "rb") as srcf:
                        with open(os.path.join(runtime_qtool_dir, "QUOTATION TOOL DATA.xlsx"), "wb") as dstf:
                            dstf.write(srcf.read())

                # Save uploaded file to a temporary path so existing loader can consume it.
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
                    tmp_in.write(uploaded.getbuffer())
                    in_path = tmp_in.name

                input_df = load_input_template(in_path, sheet="Input")
                out_path = next_qflow_output_path(runtime_qtool_dir)
                build_output(input_df, out_path)

                with open(out_path, "rb") as f:
                    out_bytes = f.read()

                st.success(f"File generated: {os.path.basename(out_path)}")
                st.caption(db_msg)
                st.download_button(
                    label="Download Horse_Quotation",
                    data=out_bytes,
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Error while generating file: {e}")
            finally:
                try:
                    if "in_path" in locals() and os.path.exists(in_path):
                        os.remove(in_path)
                except Exception:
                    pass
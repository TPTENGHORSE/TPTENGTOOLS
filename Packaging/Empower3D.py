import streamlit as st 
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# Dimensiones internas (volumen bruto)
DIMENSIONES_INTERNAS = {
    "Container 40 HC": (12032, 2352, 2700),
    "Container 20 Ft Std": (5898, 2352, 2393),
    "Trailer 40m3": (7000, 2400, 2400),
    "Mega Trailer 90m3": (13620, 2480, 2900)
}

# Dimensiones operativas reales para cálculo de UCM
DIMENSIONES_OPERATIVAS = {
    "Container 20 Ft Std": (5898, 2352, 2243),
    "Container 40 HC": (12032, 2352, 2550),
    "Trailer 40m3": (7000, 2400, 2300),
    "Mega Trailer 90m3": (13620, 2480, 2900)
}

# Peso máximo por contenedor
PESOS_MAXIMOS = {
    "Container 20 Ft Std": 25200,
    "Container 40 HC": 24750,
    "Trailer 40m3": 12000,
    "Mega Trailer 90m3": 25000
}

# Solo rotaciones permitidas en eje X e Y (altura fija)
def rotaciones_caja(l, w, h):
    return [
        (l, w, h),
        (w, l, h),
    ]

# Cálculo de cajas con límite de apilamiento (stockage)

# Cálculo mixto de cajas: llena con orientación principal y usa el espacio sobrante para rotar cajas en planta
def calcula_cajas(contenedor, caja, stacking):
    Lc, Wc, Hc = contenedor
    l1, w1, h = caja
    l2, w2 = w1, l1  # Rotación en planta
    nh = min(Hc // h, stacking)

    # Opción 1: principal (l1, w1)
    nl1 = Lc // l1
    nw1 = Wc // w1
    sobrante_w = Wc - (nw1 * w1)
    sobrante_l = Lc - (nl1 * l1)

    # 1. Cajas rotadas en el espacio sobrante del ancho (a lo largo de todo el largo principal)
    nl2 = nl1
    nw2 = sobrante_w // w2 if sobrante_w >= w2 else 0

    # 2. Cajas rotadas en el espacio sobrante del largo (usando ancho de la caja rotada w2)
    nl3 = sobrante_l // l2 if sobrante_l >= l2 else 0
    nw3 = Wc // w2 if sobrante_l >= l2 else 0

    # 3. Cajas rotadas en la esquina sobrante (si cabe)
    nl4 = sobrante_l // l2 if sobrante_l >= l2 else 0
    nw4 = sobrante_w // w2 if sobrante_w >= w2 else 0

    total1 = (nl1 * nw1 + nl2 * nw2 + nl3 * nw3 + nl4 * nw4) * nh
    distribucion1 = ((nl1, nw1, nh), (nl2, nw2, nh), (nl3, nw3, nh), (nl4, nw4, nh))

    # Opción 2: principal (w1, l1)
    nl1b = Lc // w1
    nw1b = Wc // l1
    sobrante_wb = Wc - (nw1b * l1)
    sobrante_lb = Lc - (nl1b * w1)

    # Opción 2: el complemento es (l1 a lo largo de L, w1 a lo largo de W)
    nl2b = nl1b
    nw2b = sobrante_wb // w1 if sobrante_wb >= w1 else 0
    nl3b = sobrante_lb // l1 if sobrante_lb >= l1 else 0
    nw3b = Wc // w1 if sobrante_lb >= l1 else 0
    nl4b = sobrante_lb // l1 if sobrante_lb >= l1 else 0
    nw4b = sobrante_wb // w1 if sobrante_wb >= w1 else 0

    total2 = (nl1b * nw1b + nl2b * nw2b + nl3b * nw3b + nl4b * nw4b) * nh
    distribucion2 = ((nl1b, nw1b, nh), (nl2b, nw2b, nh), (nl3b, nw3b, nh), (nl4b, nw4b, nh))

    if total1 >= total2:
        mejor_cantidad = total1
        mejor_rotacion = (l1, w1, h)
        mejor_distribucion = distribucion1
    else:
        mejor_cantidad = total2
        mejor_rotacion = (w1, l1, h)
        mejor_distribucion = distribucion2

    return mejor_cantidad, mejor_rotacion, mejor_distribucion

# Dibujo de contenedor con cajas

# Dibuja ambas orientaciones de cajas (principal y rotada) con diferentes colores
def dibuja_cajas_3d(contenedor, caja_dim, distribuciones, max_cajas=None, titulo=""):
    (nl1, nw1, nh1) = distribuciones[0]
    Lc, Wc, Hc = contenedor
    l, w, h = caja_dim
    l_rot, w_rot = w, l

    fig = plt.figure(figsize=(12, 8), facecolor='#F8FAFB')
    ax = fig.add_subplot(111, projection='3d')
    ax.set_facecolor('#F8FAFB')
    ax.set_box_aspect((Lc, Wc, Hc))
    ax.xaxis.pane.fill = False
    ax.yaxis.pane.fill = False
    ax.zaxis.pane.fill = False
    ax.xaxis.pane.set_edgecolor('#D0D7DE')
    ax.yaxis.pane.set_edgecolor('#D0D7DE')
    ax.zaxis.pane.set_edgecolor('#D0D7DE')
    ax.grid(True, alpha=0.2, linestyle='--', linewidth=0.5)
    draw_box(ax, (0, 0, 0), Lc, Wc, Hc, '#B8D4F0', alpha=0.07)

    cajas_dibujadas = 0
    bloques = [
        (distribuciones[0], (l, w),         '#C9956C'),  # principal  – warm tan
        (distribuciones[1], (l_rot, w_rot), '#4A90D9'),  # rotated W  – blue
        (distribuciones[2], (l_rot, w_rot), '#5BAD6F'),  # rotated L  – green
        (distribuciones[3], (l_rot, w_rot), '#D4546A'),  # corner     – red
    ]
    total_cajas = max_cajas if max_cajas is not None else sum(nl * nw * nh for (nl, nw, nh), _, _ in bloques)
    for idx, (dist, (lx, wx), color) in enumerate(bloques):
        nl, nw, nh = dist
        for z in range(nh):
            for x in range(nl):
                for y in range(nw):
                    if cajas_dibujadas >= total_cajas:
                        break
                    if idx == 0:
                        x_offset = x * lx
                        y_offset = y * wx
                    elif idx == 1:
                        x_offset = x * lx
                        y_offset = nw1 * w + y * wx
                    elif idx == 2:
                        x_offset = nl1 * l + x * lx
                        y_offset = y * wx
                    else:
                        x_offset = nl1 * l + x * lx
                        y_offset = nw1 * w + y * wx
                    if (x_offset + lx > Lc) or (y_offset + wx > Wc) or (z * h + h > Hc):
                        continue
                    draw_box(ax, (x_offset, y_offset, z * h), lx, wx, h, color, alpha=0.88)
                    cajas_dibujadas += 1

    ax.set_xlabel('Length (mm)', fontsize=8, color='#718096', labelpad=8)
    ax.set_ylabel('Width (mm)', fontsize=8, color='#718096', labelpad=8)
    ax.set_zlabel('Height (mm)', fontsize=8, color='#718096', labelpad=8)
    ax.tick_params(axis='both', labelsize=7, colors='#A0AEC0')
    ax.set_xlim(0, Lc)
    ax.set_ylim(0, Wc)
    ax.set_zlim(0, Hc)
    ax.view_init(elev=25, azim=45)
    fig.tight_layout(pad=1.5)
    st.pyplot(fig)
    plt.close(fig)

def draw_box(ax, origin, l, w, h, color='orange', alpha=1.0):
    x, y, z = origin
    vertices = np.array([
        [x, y, z],
        [x + l, y, z],
        [x + l, y + w, z],
        [x, y + w, z],
        [x, y, z + h],
        [x + l, y, z + h],
        [x + l, y + w, z + h],
        [x, y + w, z + h]
    ])
    faces = [
        [vertices[j] for j in [0, 1, 2, 3]],
        [vertices[j] for j in [4, 5, 6, 7]],
        [vertices[j] for j in [0, 1, 5, 4]],
        [vertices[j] for j in [2, 3, 7, 6]],
        [vertices[j] for j in [1, 2, 6, 5]],
        [vertices[j] for j in [4, 7, 3, 0]],
    ]
    ax.add_collection3d(Poly3DCollection(faces, facecolors=color, linewidths=0.5, edgecolors='black', alpha=alpha))

def main():
    st.set_page_config(
        page_title="Empower³ · UCM Optimizer",
        page_icon="📦",
        layout="wide",
        initial_sidebar_state="collapsed"
    )

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Inter', sans-serif !important;
        background-color: #EEF2F7 !important;
    }

    /* ── Header ─────────────────────────────── */
    .e3d-header {
        background: linear-gradient(135deg, #1B2E4B 0%, #2C5282 100%);
        border-radius: 14px;
        padding: 20px 30px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 20px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.18);
    }
    .e3d-title { color: white; font-size: 1.8rem; font-weight: 800; letter-spacing: -0.5px; margin: 0; }
    .e3d-subtitle { color: rgba(255,255,255,0.55); font-size: 0.78rem; margin-top: 3px; }
    .e3d-badge {
        background: rgba(255,107,43,0.9);
        color: white; padding: 5px 14px;
        border-radius: 20px; font-size: 0.7rem;
        font-weight: 700; letter-spacing: 0.6px; text-transform: uppercase;
    }

    /* ── Section labels ──────────────────────── */
    .e3d-sec {
        font-size: 0.67rem; font-weight: 700; color: #718096;
        text-transform: uppercase; letter-spacing: 0.8px;
        margin: 14px 0 5px 2px;
    }

    /* ── Result hero card ────────────────────── */
    .e3d-hero {
        background: linear-gradient(135deg, #E8590C 0%, #FF8C42 100%);
        border-radius: 12px; padding: 16px 20px;
        margin: 12px 0 10px 0;
        box-shadow: 0 4px 16px rgba(232,89,12,0.35);
    }
    .e3d-hero-lbl { color: rgba(255,255,255,0.82); font-size: 0.68rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; }
    .e3d-hero-val { color: white; font-size: 2.6rem; font-weight: 800; line-height: 1.1; }
    .e3d-hero-sub { color: rgba(255,255,255,0.72); font-size: 0.76rem; margin-top: 2px; }

    /* ── Info card ───────────────────────────── */
    .e3d-info {
        background: white; border-radius: 10px;
        padding: 13px 17px; margin-bottom: 10px;
        box-shadow: 0 1px 5px rgba(0,0,0,0.06);
        border-left: 3px solid #2C5282;
    }
    .e3d-info-lbl { font-size: 0.67rem; font-weight: 700; color: #A0AEC0; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
    .e3d-info-row { font-size: 0.82rem; color: #2D3748; margin: 3px 0; }
    .e3d-info-row b { color: #1B2E4B; }

    /* ── Progress bars ───────────────────────── */
    .e3d-pw { margin-bottom: 9px; }
    .e3d-ph { display: flex; justify-content: space-between; font-size: 0.71rem; color: #718096; font-weight: 500; margin-bottom: 4px; }
    .e3d-bg { background: #E2E8F0; border-radius: 6px; height: 7px; overflow: hidden; }
    .e3d-fill { height: 7px; border-radius: 6px; }

    /* ── Stackability display ────────────────── */
    .e3d-stack {
        text-align: center; padding: 9px 0;
        border: 1.5px solid #E2E8F0; border-radius: 8px;
        background: white; font-weight: 700; font-size: 1rem;
        color: #1B2E4B; box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }

    /* ── Right panel placeholder ─────────────── */
    .e3d-ph-box {
        display: flex; flex-direction: column;
        align-items: center; justify-content: center;
        height: 440px; background: white;
        border-radius: 14px; color: #A0AEC0;
        box-shadow: 0 1px 5px rgba(0,0,0,0.06);
    }
    .e3d-ph-icon { font-size: 3.5rem; margin-bottom: 12px; }
    .e3d-ph-txt  { font-size: 0.88rem; font-weight: 500; }

    /* ── Chart header ────────────────────────── */
    .e3d-ch {
        background: white; border-radius: 14px 14px 0 0;
        padding: 15px 22px 10px 22px;
        border-bottom: 1px solid #EEF2F7;
        margin-bottom: 0;
    }
    .e3d-ch-title { font-size: 1rem; font-weight: 700; color: #1B2E4B; margin: 0; }
    .e3d-ch-sub   { font-size: 0.76rem; color: #A0AEC0; margin-top: 2px; }

    /* ── Streamlit overrides ─────────────────── */
    .stButton > button {
        border-radius: 9px !important; font-weight: 600 !important;
        font-size: 0.87rem !important; transition: all 0.2s ease !important;
    }
    .stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.14) !important; }
    div[data-testid="stNumberInput"] label,
    div[data-testid="stSelectbox"] label {
        font-size: 0.77rem !important; font-weight: 600 !important; color: #4A5568 !important;
    }
    div[data-testid="stMetric"] { background: white; border-radius: 10px; padding: 10px 14px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
    div[data-testid="stMetricLabel"] { font-size: 0.72rem !important; color: #718096 !important; }
    div[data-testid="stMetricValue"] { font-size: 1.15rem !important; font-weight: 700 !important; color: #1B2E4B !important; }
    #MainMenu { visibility: hidden; }
    footer     { visibility: hidden; }
    header     { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)

    # ── Header ──────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="e3d-header">
        <div>
            <div class="e3d-title">📦 Empower<sup style="font-size:1rem;font-weight:600">3D</sup></div>
            <div class="e3d-subtitle">UCM Packaging Optimization Tool</div>
        </div>
        <span class="e3d-badge">v 1.0</span>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1.05, 1.35])

    with col_left:
        # Transport type
        st.markdown('<div class="e3d-sec">🚛 Transport Type</div>', unsafe_allow_html=True)
        container_sel = st.selectbox("", list(DIMENSIONES_INTERNAS.keys()), label_visibility="collapsed")

        # Dimensions
        st.markdown('<div class="e3d-sec">📐 UCM Dimensions</div>', unsafe_allow_html=True)
        dim_col1, dim_col2, dim_col3 = st.columns(3)
        with dim_col1:
            box_length = st.number_input("Length (mm)", min_value=1, value=2200)
        with dim_col2:
            box_width = st.number_input("Width (mm)", min_value=1, value=1000)
        with dim_col3:
            box_height = st.number_input("Height (mm)", min_value=1, value=975)

        # Weights
        st.markdown('<div class="e3d-sec">⚖️ Weight Parameters</div>', unsafe_allow_html=True)
        peso_col1, peso_col2, peso_col3 = st.columns(3)
        with peso_col1:
            box_weight_pn = st.number_input("Weight PN (kg)", min_value=0.001, value=10.000, format="%.1f")
        with peso_col2:
            box_weight_ucm = st.number_input("UCM (kg)", min_value=0.01, value=100.0, format="%.1f")
        with peso_col3:
            pn_ucm = st.number_input("PN/UCM", min_value=0.01, value=100.0, format="%.2f")
        box_weight = (box_weight_pn * pn_ucm) + box_weight_ucm

        # Stackability
        operative_dim = DIMENSIONES_OPERATIVAS[container_sel]
        box_dim = (box_length, box_width, box_height)
        _, _, (dist1, dist2, dist3, dist4) = calcula_cajas(operative_dim, box_dim, 9999)
        nl, nw, nh = dist1
        max_stacking_possible = nh

        st.markdown('<div class="e3d-sec">📚 Stackability</div>', unsafe_allow_html=True)
        col_stack1, col_stack2, col_stack3 = st.columns([1, 2, 1])

        if "stackability_value" not in st.session_state:
            st.session_state.stackability_value = 0
        if st.session_state.stackability_value > 99:
            st.session_state.stackability_value = 99

        with col_stack1:
            if st.button("➖", key="stack_minus"):
                if st.session_state.stackability_value > 0:
                    st.session_state.stackability_value -= 1
        with col_stack2:
            st.markdown(
                f"<div class='e3d-stack'>{st.session_state.stackability_value}/1</div>",
                unsafe_allow_html=True
            )
        with col_stack3:
            if st.button("➕", key="stack_plus"):
                if st.session_state.stackability_value < 99:
                    st.session_state.stackability_value += 1

        max_stacking = st.session_state.stackability_value + 1
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        calculate = st.button("🔍  Calculate", type="primary", use_container_width=True)

        # ── Results ─────────────────────────────────────────────────────────
        if calculate:
            operative_dim = DIMENSIONES_OPERATIVAS[container_sel]
            external_dim  = DIMENSIONES_INTERNAS[container_sel]
            max_container_weight = PESOS_MAXIMOS[container_sel]
            box_dim = (box_length, box_width, box_height)
            _, _, (dist1, dist2, dist3, dist4) = calcula_cajas(operative_dim, box_dim, 9999)
            nl, nw, nh = dist1
            max_stacking_possible = nh
            total_by_volume, rotation, distribuciones = calcula_cajas(
                operative_dim, box_dim, min(max_stacking, max_stacking_possible)
            )
            box_volume           = (box_dim[0]/1000) * (box_dim[1]/1000) * (box_dim[2]/1000)
            total_external_volume = (external_dim[0]/1000) * (external_dim[1]/1000) * (external_dim[2]/1000)
            max_ucm_by_weight    = int(max_container_weight // box_weight)
            realistic_ucm        = min(total_by_volume, max_ucm_by_weight)
            realistic_volume     = box_volume * realistic_ucm
            realistic_volume_sat = realistic_volume / total_external_volume * 100
            realistic_weight     = box_weight * realistic_ucm
            total_pn_ut          = pn_ucm * realistic_ucm
            densidad             = realistic_weight / realistic_volume if realistic_volume > 0 else 0
            w_pct                = (realistic_weight / max_container_weight) * 100

            (nl1, nw1, nh_d), (nl2, nw2, nh2), (nl3, nw3, nh3), (nl4, nw4, nh4) = distribuciones
            texto_dist = f"Main: {nl1} × {nw1} × {nh_d}"
            if nl2 > 0 and nw2 > 0:
                texto_dist += f"  +  Rot.W: {nl2} × {nw2} × {nh2}"
            if nl3 > 0 and nw3 > 0:
                texto_dist += f"  +  Rot.L: {nl3} × {nw3} × {nh3}"
            if nl4 > 0 and nw4 > 0:
                texto_dist += f"  +  Corner: {nl4} × {nw4} × {nh4}"

            limited_by = "Weight" if realistic_ucm == max_ucm_by_weight else "Volume"

            # Hero metric
            st.markdown(f"""
            <div class="e3d-hero">
                <div class="e3d-hero-lbl">Realistic UCM · {limited_by} Limited</div>
                <div class="e3d-hero-val">{realistic_ucm}</div>
                <div class="e3d-hero-sub">Max by volume: {total_by_volume} &nbsp;·&nbsp; Max by weight: {max_ucm_by_weight}</div>
            </div>
            """, unsafe_allow_html=True)

            # Config card
            st.markdown(f"""
            <div class="e3d-info">
                <div class="e3d-info-lbl">Configuration</div>
                <div class="e3d-info-row"><b>Best Rotation (L×W×H):</b> {rotation[0]} × {rotation[1]} × {rotation[2]} mm</div>
                <div class="e3d-info-row"><b>Layout:</b> {texto_dist}</div>
                <div class="e3d-info-row"><b>Volume / UCM:</b> {box_volume:.3f} m³</div>
            </div>
            """, unsafe_allow_html=True)

            # KPI row 1
            kpi1, kpi2, kpi3 = st.columns(3)
            with kpi1:
                st.metric("Total Volume",  f"{realistic_volume:.2f} m³")
            with kpi2:
                st.metric("Total Weight",  f"{realistic_weight:,.0f} kg")
            with kpi3:
                st.metric("Total PN",      f"{total_pn_ut:,.0f}")

            # KPI row 2
            kpi4, kpi5, kpi6 = st.columns(3)
            with kpi4:
                st.metric("Vol. Saturation",    f"{realistic_volume_sat:.1f}%")
            with kpi5:
                st.metric("Weight Saturation",  f"{w_pct:.1f}%")
            with kpi6:
                st.metric("Density",            f"{densidad:.0f} kg/m³")

            # Progress bars
            def bar_color(pct):
                if pct >= 90: return "background:linear-gradient(90deg,#E53E3E,#C53030)"
                if pct >= 70: return "background:linear-gradient(90deg,#DD6B20,#C05621)"
                return "background:linear-gradient(90deg,#38A169,#276749)"

            v_cap = min(realistic_volume_sat, 100)
            w_cap = min(w_pct, 100)
            st.markdown(f"""
            <div style="margin-top:10px">
              <div class="e3d-pw">
                <div class="e3d-ph"><span>📦 Volume Saturation</span><span>{realistic_volume_sat:.1f}%</span></div>
                <div class="e3d-bg"><div class="e3d-fill" style="width:{v_cap:.1f}%;{bar_color(realistic_volume_sat)}"></div></div>
              </div>
              <div class="e3d-pw">
                <div class="e3d-ph"><span>⚖️ Weight Saturation</span><span>{w_pct:.1f}%</span></div>
                <div class="e3d-bg"><div class="e3d-fill" style="width:{w_cap:.1f}%;{bar_color(w_pct)}"></div></div>
              </div>
            </div>
            """, unsafe_allow_html=True)

    with col_right:
        if 'calculate' in locals() and calculate:
            st.markdown("""
            <div class="e3d-ch">
                <div class="e3d-ch-title">3D UCM Distribution</div>
                <div class="e3d-ch-sub">Optimized packing · Limited by volume &amp; weight</div>
            </div>
            """, unsafe_allow_html=True)
            dibuja_cajas_3d(operative_dim, rotation, distribuciones, max_cajas=realistic_ucm)
        else:
            st.markdown("""
            <div class="e3d-ph-box">
                <div class="e3d-ph-icon">📦</div>
                <div class="e3d-ph-txt">Configure parameters and click <b>Calculate</b></div>
                <div style="font-size:0.76rem;color:#CBD5E0;margin-top:6px;">3D visualization will appear here</div>
            </div>
            """, unsafe_allow_html=True)

def run():
    main()

if __name__ == "__main__":
    main()
#python -m streamlit run Container3D.py
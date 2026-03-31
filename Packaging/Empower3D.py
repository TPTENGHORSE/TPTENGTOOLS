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
    "Mega Trailer 90m3": (13620, 2480, 2800)
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

    # 2. Cajas rotadas en el espacio sobrante del largo (a lo ancho de todo el ancho principal)
    nl3 = sobrante_l // l2 if sobrante_l >= l2 else 0
    nw3 = nw1

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

    nl2b = nl1b
    nw2b = sobrante_wb // w2 if sobrante_wb >= w2 else 0
    nl3b = sobrante_lb // l2 if sobrante_lb >= l2 else 0
    nw3b = nw1b
    nl4b = sobrante_lb // l2 if sobrante_lb >= l2 else 0
    nw4b = sobrante_wb // w2 if sobrante_wb >= w2 else 0

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
def dibuja_cajas_3d(contenedor, caja_dim, distribuciones, max_cajas=None, titulo="3D Distribution"):
    # Extraer la distribución principal para usar nl1 y nw1 en los offsets
    (nl1, nw1, nh1) = distribuciones[0]
    Lc, Wc, Hc = contenedor
    l, w, h = caja_dim
    l_rot, w_rot = w, l

    fig = plt.figure(figsize=(12, 8))
    ax = fig.add_subplot(111, projection='3d')
    ax.set_box_aspect((Lc, Wc, Hc))
    draw_box(ax, (0, 0, 0), Lc, Wc, Hc, 'lightblue', alpha=0.1)

    cajas_dibujadas = 0
    # Desempaquetar todas las distribuciones
    bloques = [
        (distribuciones[0], (l, w), 'burlywood'),  # principal
        (distribuciones[1], (l_rot, w_rot), 'orange'),  # rotadas en ancho
        (distribuciones[2], (l_rot, w_rot), 'green'),   # rotadas en largo
        (distribuciones[3], (l_rot, w_rot), 'red'),     # rotadas en esquina
    ]
    total_cajas = max_cajas if max_cajas is not None else sum(nl * nw * nh for (nl, nw, nh), _, _ in bloques)
    for idx, (dist, (lx, wx), color) in enumerate(bloques):
        nl, nw, nh = dist
        for z in range(nh):
            for x in range(nl):
                for y in range(nw):
                    if cajas_dibujadas >= total_cajas:
                        break
                    # Calcular offset según bloque
                    if idx == 0:
                        x_offset = x * lx
                        y_offset = y * wx
                    elif idx == 1:
                        x_offset = x * lx
                        y_offset = nw1 * w + y * wx
                    elif idx == 2:
                        x_offset = nl1 * l + x * lx
                        y_offset = y * wx
                    else:  # esquina
                        x_offset = nl1 * l + x * lx
                        y_offset = nw1 * w + y * wx
                    # Limitar para que no desborde el contenedor
                    if (x_offset + lx > Lc) or (y_offset + wx > Wc) or (z * h + h > Hc):
                        continue
                    draw_box(ax, (x_offset, y_offset, z * h), lx, wx, h, color, alpha=0.8)
                    cajas_dibujadas += 1

    ax.set_xlabel('Length (mm)')
    ax.set_ylabel('Width (mm)')
    ax.set_zlabel('Height (mm)')
    ax.set_xlim(0, Lc)
    ax.set_ylim(0, Wc)
    ax.set_zlim(0, Hc)
    ax.view_init(elev=25, azim=45)
    plt.title(titulo)
    st.pyplot(fig)

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
    # logo = Image.open("logo.png")
    # st.image(logo, width=110)
    st.markdown("<h1 style='text-align: center;'>📦 Empower<sup>3D</sup></h1>  v1.0", unsafe_allow_html=True)

    col_left, col_right = st.columns([1.1, 1.3])

    with col_left:
        container_sel = st.selectbox("Select Transport Type", list(DIMENSIONES_INTERNAS.keys()))
        # Primera fila: dimensiones
        dim_col1, dim_col2, dim_col3 = st.columns(3)
        with dim_col1:
            box_length = st.number_input("Length (mm)", min_value=1, value=1140)
        with dim_col2:
            box_width = st.number_input("Width (mm)", min_value=1, value=900)
        with dim_col3:
            box_height = st.number_input("Height (mm)", min_value=1, value=850)

        # Segunda fila: pesos y PN/UCM
        peso_col1, peso_col2, peso_col3 = st.columns(3)
        with peso_col1:
            box_weight_pn = st.number_input("Weight PN (kg)", min_value=0.01, value=10.0, format="%.1f")
        with peso_col2:
            box_weight_ucm = st.number_input("UCM (kg)", min_value=0.01, value=100.0, format="%.1f")
        with peso_col3:
            pn_ucm = st.number_input("PN/UCM", min_value=0.01, value=100.0, format="%.2f")
        # Nuevo cálculo de peso unitario de la caja
        box_weight = (box_weight_pn * pn_ucm) + box_weight_ucm
        
        # Calculate maximum possible stackability based on current dimensions
        operative_dim = DIMENSIONES_OPERATIVAS[container_sel]
        box_dim = (box_length, box_width, box_height)
        _, _, (dist1, dist2, dist3, dist4) = calcula_cajas(operative_dim, box_dim, 9999)
        nl, nw, nh = dist1
        max_stacking_possible = nh
        
        # Stackability with custom display format
        st.write("Stackability")
        col_stack1, col_stack2, col_stack3 = st.columns([1, 2, 1])
        
        # Initialize session state for stackability (starting with 0/1)
        if "stackability_value" not in st.session_state:
            st.session_state.stackability_value = 0
        
        # Ensure stackability doesn't exceed 99/1
        if st.session_state.stackability_value > 99:
            st.session_state.stackability_value = 99
        
        with col_stack1:
            if st.button("➖", key="stack_minus"):
                if st.session_state.stackability_value > 0:
                    st.session_state.stackability_value -= 1
        
        with col_stack2:
            # Display format: 0/1, 1/1, 2/1, 3/1, etc.
            display_text = f"{st.session_state.stackability_value}/1"
            st.markdown(f"<div style='text-align: center; padding: 8px; border: 1px solid #ccc; border-radius: 4px; background-color: white;'>{display_text}</div>", unsafe_allow_html=True)
        
        with col_stack3:
            if st.button("➕", key="stack_plus"):
                if st.session_state.stackability_value < 99:  # Maximum limit is 99/1
                    st.session_state.stackability_value += 1
        
        max_stacking = st.session_state.stackability_value + 1  # Convert from #/1 format to actual stacking value
        calculate = st.button("Calculate")

        # Show results between the inputs and the chart
        if 'calculate' not in locals():
            calculate = False

        if calculate:
            operative_dim = DIMENSIONES_OPERATIVAS[container_sel]
            external_dim = DIMENSIONES_INTERNAS[container_sel]
            max_container_weight = PESOS_MAXIMOS[container_sel]
            box_dim = (box_length, box_width, box_height)
            # Calculate the real maximum stacking possible
            _, _, (dist1, dist2, dist3, dist4) = calcula_cajas(operative_dim, box_dim, 9999)
            nl, nw, nh = dist1
            max_stacking_possible = nh
            total_by_volume, rotation, distribuciones = calcula_cajas(operative_dim, box_dim, min(max_stacking, max_stacking_possible))
            box_volume = (box_dim[0] / 1000) * (box_dim[1] / 1000) * (box_dim[2] / 1000)
            total_usable_volume = box_volume * total_by_volume
            total_external_volume = (external_dim[0] / 1000) * (external_dim[1] / 1000) * (external_dim[2] / 1000)
            volume_saturation = total_usable_volume / total_external_volume * 100
            total_weight = total_by_volume * box_weight
            weight_saturation = total_weight / max_container_weight * 100
            max_ucm_by_weight = int(max_container_weight // box_weight)
            realistic_ucm = min(total_by_volume, max_ucm_by_weight)
            realistic_volume = box_volume * realistic_ucm
            realistic_volume_saturation = realistic_volume / total_external_volume * 100
            realistic_weight = box_weight * realistic_ucm
            # Mostrar detalles de ambas distribuciones
            st.success(f"🔢 Realistic UCM (weight limited): **{realistic_ucm}**")
            st.write(f"Best main rotation (LxWxH): **{rotation}**")
            # Mostrar cómo se distribuyen las cajas (cuatro bloques)
            (nl1, nw1, nh), (nl2, nw2, nh2), (nl3, nw3, nh3), (nl4, nw4, nh4) = distribuciones
            texto_dist = f"Main: {nl1} x {nw1} x {nh}"
            if nl2 > 0 and nw2 > 0:
                texto_dist += f" + Rotated W: {nl2} x {nw2} x {nh2}"
            if nl3 > 0 and nw3 > 0:
                texto_dist += f" + Rotated L: {nl3} x {nw3} x {nh3}"
            if nl4 > 0 and nw4 > 0:
                texto_dist += f" + Corner: {nl4} x {nw4} x {nh4}"
            st.write(f"Distribution (UCM): **{texto_dist}**")
            st.write(f"📦 Volume per UCM: **{box_volume:.3f} m³**")
            st.write(f"📏 Total volume (realistic): **{realistic_volume:.2f} m³**")
            st.write(f"🧱 Volume saturation (realistic): **{realistic_volume_saturation:.2f}%**")
            st.write(f"⚖️ Total weight (realistic): **{realistic_weight:,.0f} kg**")
            st.write(f"🏋️ Weight saturation: **{(realistic_weight/max_container_weight)*100:.2f}%**")
            # Nuevo: Total PN/UT y Densidad
            total_pn_ut = pn_ucm * realistic_ucm
            st.write(f"🧮 Total PN/Transport Type: **{total_pn_ut:,.0f}**")
            densidad = realistic_weight / realistic_volume if realistic_volume > 0 else 0
            st.write(f"🧪 Densidad: **{densidad:.0f} kg/m³**")

    with col_right:
        if 'calculate' in locals() and calculate:
            st.markdown("<h3 style='text-align: center;'>3D UMs Distribution</h3>", unsafe_allow_html=True)
            dibuja_cajas_3d(operative_dim, box_dim, distribuciones, max_cajas=realistic_ucm, titulo="Limited by max volume & weight")

def run():
    main()

if __name__ == "__main__":
    main()
#python -m streamlit run Container3D.py
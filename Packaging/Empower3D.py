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
    "Mega Trailer 90m3": 32800
}

# Solo rotaciones permitidas en eje X e Y (altura fija)
def rotaciones_caja(l, w, h):
    return [
        (l, w, h),
        (w, l, h),
    ]

# Cálculo de cajas con límite de apilamiento (stockage)
def calcula_cajas(contenedor, caja, stacking):
    Lc, Wc, Hc = contenedor
    mejor_cantidad = 0
    mejor_rotacion = None
    mejor_distribucion = (0, 0, 0)

    for (l, w, h) in rotaciones_caja(*caja):
        nl = Lc // l
        nw = Wc // w
        nh = min(Hc // h, stacking)
        total = nl * nw * nh
        if total > mejor_cantidad:
            mejor_cantidad = total
            mejor_rotacion = (l, w, h)
            mejor_distribucion = (nl, nw, nh)

    return mejor_cantidad, mejor_rotacion, mejor_distribucion

# Dibujo de contenedor con cajas
def dibuja_cajas_3d(contenedor, caja_dim, distribucion, max_cajas=None, titulo="3D Distribution"):
    Lc, Wc, Hc = contenedor
    nl, nw, nh = distribucion
    l, w, h = caja_dim

    fig = plt.figure(figsize=(12, 8))
    ax = fig.add_subplot(111, projection='3d')
    ax.set_box_aspect((Lc, Wc, Hc))
    draw_box(ax, (0, 0, 0), Lc, Wc, Hc, 'lightblue', alpha=0.1)

    cajas_dibujadas = 0
    total_cajas = max_cajas if max_cajas is not None else nl * nw * nh
    for z in range(nh):
        cajas_restantes = total_cajas - cajas_dibujadas
        if cajas_restantes <= 0:
            break
        cajas_en_este_nivel = min(cajas_restantes, nl * nw)
        for i in range(cajas_en_este_nivel):
            x = i // nw
            y = i % nw
            draw_box(ax, (x * l, y * w, z * h), l, w, h, 'burlywood', alpha=0.8)
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
        _, _, (nl, nw, nh) = calcula_cajas(operative_dim, box_dim, 9999)
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
            _, _, (nl, nw, nh) = calcula_cajas(operative_dim, box_dim, 9999)
            max_stacking_possible = nh
            total_by_volume, rotation, distribution = calcula_cajas(operative_dim, box_dim, min(max_stacking, max_stacking_possible))
            box_volume = (rotation[0] / 1000) * (rotation[1] / 1000) * (rotation[2] / 1000)
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
            # Calcular niveles y última capa para el mensaje usando la distribución real
            dist_nl, dist_nw, dist_nh = distribution
            dist_levels = realistic_ucm // (dist_nl * dist_nw) if dist_nl * dist_nw > 0 else 0
            dist_last_level = realistic_ucm % (dist_nl * dist_nw) if dist_nl * dist_nw > 0 else 0
            if (dist_levels == 0 and dist_last_level > 0) and max_stacking > 1:
                st.warning("⚠️ The maximum stacking for this configuration is 0/1. Value adjusted.")
            elif max_stacking > max_stacking_possible:
                if max_stacking_possible <= 1:
                    st.warning("⚠️ The maximum stacking for this configuration is 0/1. Value adjusted.")
                else:
                    max_stackability_display = max_stacking_possible - 1
                    st.warning(f"⚠️ The maximum stacking for this configuration is {max_stackability_display}/1. Value adjusted.")

            st.success(f"🔢 Realistic UCM (weight limited): **{realistic_ucm}**")
            st.write(f"Best rotation (LxWxH): **{rotation}**")
            # Calculate the actual distribution used in the drawing
            levels = realistic_ucm // (distribution[0] * distribution[1])
            last_level = realistic_ucm % (distribution[0] * distribution[1])
            if levels == 0 and last_level > 0:
                used_distribution = f"{last_level} in one level"
            elif last_level == 0:
                used_distribution = f"{distribution[0]} x {distribution[1]} x {levels}"
            else:
                used_distribution = f"{distribution[0]} x {distribution[1]} x {levels} + {last_level} on the last level"
            st.write(f"Distribution (UCM): **{used_distribution}**")
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
            dibuja_cajas_3d(operative_dim, rotation, distribution, max_cajas=realistic_ucm, titulo="Limited by max volume & weight")

def run():
    main()

if __name__ == "__main__":
    main()
#python -m streamlit run Container3D.py
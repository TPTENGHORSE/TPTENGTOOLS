import streamlit as st 
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# Dimensiones internas (volumen bruto)
DIMENSIONES_INTERNAS = {
    "20 Ft Std": (5898, 2352, 2393),
    "40 HC": (12032, 2352, 2700)
}

# Dimensiones operativas reales para cálculo de UCM
DIMENSIONES_OPERATIVAS = {
    "20 Ft Std": (5898, 2352, 1993),
    "40 HC": (12032, 2352, 2300)
}

# Peso máximo por contenedor
PESOS_MAXIMOS = {
    "20 Ft Std": 25200,
    "40 HC": 24750
}

# Solo rotaciones permitidas en eje X e Y (altura fija)
def rotaciones_caja(l, w, h):
    return [
        (l, w, h),
        (w, l, h),
    ]

# Cálculo de cajas con límite de apilamiento (stockage)
def calcula_cajas(contenedor, caja, stockage):
    Lc, Wc, Hc = contenedor
    mejor_cantidad = 0
    mejor_rotacion = None
    mejor_distribucion = (0, 0, 0)

    for (l, w, h) in rotaciones_caja(*caja):
        nl = Lc // l
        nw = Wc // w
        nh = min(Hc // h, stockage)
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
    for x in range(nl):
        for y in range(nw):
            for z in range(nh):
                if max_cajas is not None and cajas_dibujadas >= max_cajas:
                    break
                draw_box(ax, (x * l, y * w, z * h), l, w, h, 'orange', alpha=0.8)
                cajas_dibujadas += 1
            if max_cajas is not None and cajas_dibujadas >= max_cajas:
                break
        if max_cajas is not None and cajas_dibujadas >= max_cajas:
            break

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
    logo = Image.open("logo.png")
    st.image(logo, width=110)
    st.markdown("<h1 style='text-align: center;'>📦 Empower<sup>3D</sup></h1>  v1.0", unsafe_allow_html=True)

    contenedor_sel = st.selectbox("Select container type", list(DIMENSIONES_INTERNAS.keys()))

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        caja_largo = st.number_input("Length (mm)", min_value=1, value=1140)
    with col2:
        caja_ancho = st.number_input("Width (mm)", min_value=1, value=900)
    with col3:
        caja_alto = st.number_input("Height (mm)", min_value=1, value=850)
    with col4:
        caja_peso = st.number_input("Weight PN + UCM (kg)", min_value=0.01, value=150.0, format="%.1f")

    stockage = st.number_input("Stockage (max stacking UCM)", min_value=1, value=2)

    dim_operativa = DIMENSIONES_OPERATIVAS[contenedor_sel]
    dim_externa = DIMENSIONES_INTERNAS[contenedor_sel]
    peso_max_contenedor = PESOS_MAXIMOS[contenedor_sel]
    caja_dim = (caja_largo, caja_ancho, caja_alto)

    if st.button("Calculate"):
        total_por_volumen, rotacion, distribucion = calcula_cajas(dim_operativa, caja_dim, stockage)
        volumen_caja = (rotacion[0] / 1000) * (rotacion[1] / 1000) * (rotacion[2] / 1000)
        volumen_total_operativo = volumen_caja * total_por_volumen
        volumen_total_externo = (dim_externa[0] / 1000) * (dim_externa[1] / 1000) * (dim_externa[2] / 1000)
        porcentaje_volumen = volumen_total_operativo / volumen_total_externo * 100

        peso_total = total_por_volumen * caja_peso
        porcentaje_peso = peso_total / peso_max_contenedor * 100

        max_ucm_por_peso = int(peso_max_contenedor // caja_peso)
        ucm_final = min(total_por_volumen, max_ucm_por_peso)
        volumen_realista = volumen_caja * ucm_final
        porcentaje_volumen_realista = volumen_realista / volumen_total_externo * 100
        peso_realista = caja_peso * ucm_final

        st.success(f"🔢 Maximum UCM that fit by volume (usable space): **{total_por_volumen}**")

        if ucm_final < total_por_volumen:
            st.warning(f"🚨 Limited by weight: Only **{ucm_final}** UCM can be loaded to respect max weight.")
        else:
            st.info("✅ All UCM fit within volume & weight limit.")

        st.write(f"Better rotation (LxWxH): **{rotacion}**")
        st.write(f"Distribution (UCM): **{distribucion[0]} x {distribucion[1]} x {distribucion[2]}**")
        st.write(f"📦 Volume per UCM: **{volumen_caja:.3f} m³**")
        st.write(f"📏 Total volume of all UCM (usable): **{volumen_total_operativo:.2f} m³**")
        st.write(f"🧱 Volume saturation (vs full internal volume): **{porcentaje_volumen:.2f}%**")
        st.write(f"⚖️ Total weight of all UCM: **{peso_total:,.2f} kg**")
        st.write(f"🏋️ Weight saturation: **{porcentaje_peso:.2f}%**")

        st.markdown("### 🔲 Full load by volume (usable space):")
        dibuja_cajas_3d(dim_operativa, rotacion, distribucion, titulo="All boxes that fit by volume")

        if ucm_final < total_por_volumen:
            st.markdown("### 🧰 Realistic load (limited by weight):")
            dibuja_cajas_3d(dim_operativa, rotacion, distribucion, max_cajas=ucm_final, titulo="Limited by max weight")
            st.info(f"⚖️ Maximum UCM limited by weight ({peso_max_contenedor} kg): **{max_ucm_por_peso}**")
            st.write(f"📏 Total volume (realistic): **{volumen_realista:.2f} m³**")
            st.write(f"🧱 Volume saturation (realistic): **{porcentaje_volumen_realista:.2f}%**")
            st.write(f"⚖️ Total weight (realistic): **{peso_realista:,.2f} kg**")

if __name__ == "__main__":
    main()




#python -m streamlit run Container3D.py
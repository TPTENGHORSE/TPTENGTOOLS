import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
from mpl_toolkits.mplot3d.art3d import Poly3DCollection

# Dimensiones contenedores en mm
CONTENEDORES = {
    "20 ft std": (5898, 2352, 2393),
    "40 HC": (12032, 2352, 2700)
}

def rotaciones_caja(l, w, h):
    # Devuelve las 6 rotaciones posibles (largo, ancho, alto)
    return [
        (l, w, h),
        (l, h, w),
        (w, l, h),
        (w, h, l),
        (h, l, w),
        (h, w, l),
    ]

def calcula_cajas(contenedor, caja):
    Lc, Wc, Hc = contenedor
    mejor_cantidad = 0
    mejor_rotacion = None
    mejor_distribucion = (0,0,0)

    for (l, w, h) in rotaciones_caja(*caja):
        nl = Lc // l
        nw = Wc // w
        nh = Hc // h
        total = nl * nw * nh

        if total > mejor_cantidad:
            mejor_cantidad = total
            mejor_rotacion = (l, w, h)
            mejor_distribucion = (nl, nw, nh)

    return mejor_cantidad, mejor_rotacion, mejor_distribucion

def dibuja_cajas_3d(contenedor, caja_dim, distribucion):
    Lc, Wc, Hc = contenedor
    nl, nw, nh = distribucion
    l, w, h = caja_dim

    fig = plt.figure(figsize=(12,8))
    ax = fig.add_subplot(111, projection='3d')

    # Dibuja contenedor
    draw_box(ax, (0,0,0), Lc, Wc, Hc, 'lightblue', alpha=0.1)

    # Dibuja cajas
    for x in range(nl):
        for y in range(nw):
            for z in range(nh):
                draw_box(ax, (x*l, y*w, z*h), l, w, h, 'orange', alpha=0.8)

    ax.set_xlabel('Length (mm)')
    ax.set_ylabel('Wight (mm)')
    ax.set_zlabel('Height (mm)')

    ax.set_xlim(0, Lc)
    ax.set_ylim(0, Wc)
    ax.set_zlim(0, Hc)
    ax.view_init(elev=25, azim=45)
    plt.title(f"UCM container distribution")
    st.pyplot(fig)

def draw_box(ax, origin, l, w, h, color='orange', alpha=1.0):
    # Dibuja un cubo con esquina en origin y dimensiones l,w,h
    x, y, z = origin

    vertices = np.array([
        [x, y, z],
        [x+l, y, z],
        [x+l, y+w, z],
        [x, y+w, z],
        [x, y, z+h],
        [x+l, y, z+h],
        [x+l, y+w, z+h],
        [x, y+w, z+h]
    ])

    faces = [
        [vertices[j] for j in [0,1,2,3]],
        [vertices[j] for j in [4,5,6,7]],
        [vertices[j] for j in [0,1,5,4]],
        [vertices[j] for j in [2,3,7,6]],
        [vertices[j] for j in [1,2,6,5]],
        [vertices[j] for j in [4,7,3,0]],
    ]

    ax.add_collection3d(Poly3DCollection(faces, facecolors=color, linewidths=0.5, edgecolors='black', alpha=alpha))

def main():
    # Carga y muestra el logo con PIL y Streamlit
    logo = Image.open("logo.png")
    st.image(logo, width=100)

    st.title("📦 Container 3D")

    contenedor_sel = st.selectbox("Select container type", list(CONTENEDORES.keys()))
    caja_largo = st.number_input("Length (mm)", min_value=1, value=1140)
    caja_ancho = st.number_input("Width (mm)", min_value=1, value=900)
    caja_alto = st.number_input("Height (mm)", min_value=1, value=850)

    contenedor_dim = CONTENEDORES[contenedor_sel]
    caja_dim = (caja_largo, caja_ancho, caja_alto)

    if st.button("Calculate"):
        total, rotacion, distribucion = calcula_cajas(contenedor_dim, caja_dim)

        st.success(f"Maximum UCM that fit: {total}")
        st.write(f"Better rotation (LxWxH): {rotacion}")
        st.write(f"Distribution (UCM): Lenght x Width x Height = {distribucion}")

        dibuja_cajas_3d(contenedor_dim, rotacion, distribucion)

if __name__ == "__main__":
    main()

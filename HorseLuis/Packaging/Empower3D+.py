import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import os
from py3dbp import Packer, Bin, Item
from binpacking3d import Packer

# --- Container configuration ---
CONTAINERS = {
    "20 Ft Std": {"length": 5898, "width": 2352, "height": 2393, "max_weight": 25200},
    "40 HC": {"length": 12032, "width": 2352, "height": 2700, "max_weight": 24750},
}

# --- Load data from Excel ---
def load_packaging_db():
    import openpyxl
    file = os.path.join(os.path.dirname(__file__), 'Base_EMB.xlsx')
    sheet = 'Informe 1'
    df = pd.read_excel(file, sheet_name=sheet, dtype=str)
    # Ensure all column names are strings and strip whitespace
    df.columns = [str(col).strip() for col in df.columns]
    # Rename columns for internal use
    rename = {
        'Packaging Code': 'Packaging Code',
        'Nb pieces par UC': 'Nb pieces per UC',
        'Qté / UC': 'Qty per UC',
        'Length (mm)': 'Length',
        'Width (mm)': 'Width',
        'Height (mm)': 'Height',
        'Folded Height (mm)': 'Folded Height',
        'Weight EMPTY (kg)': 'Weight EMPTY',
        'Part Weight (kg)': 'Part Weight',
        'Reference': 'Reference',  # <-- Añadido para el filtro
    }
    df = df.rename(columns=rename)
    # Only relevant columns
    cols = ['Reference','Packaging Code','Nb pieces per UC','Qty per UC','Length','Width','Height','Folded Height','Weight EMPTY','Part Weight']
    missing = [c for c in cols if c not in df.columns]
    if missing:
        st.error(f"Missing columns in Excel: {missing}. Please check the file and column names.\nColumns found: {list(df.columns)}")
        st.stop()
    df = df[cols]  # Only select columns that exist (now guaranteed)
    # Clean nulls and types
    for c in ['Length','Width','Height','Folded Height','Weight EMPTY','Part Weight','Nb pieces per UC','Qty per UC']:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    df['Packaging Code'] = df['Packaging Code'].astype(str)
    return df

# --- Main UI ---
def main():
    st.title('Container Loading Optimizer')
    # Load database
    db = load_packaging_db()
    # Container selection
    container_type = st.selectbox('Container type', list(CONTAINERS.keys()))
    container = CONTAINERS[container_type]
    st.write(f"Container dimensions: {container['length']} x {container['width']} x {container['height']} mm, Max weight: {container['max_weight']} kg")
    # Session state for packaging list
    if 'packaging_list' not in st.session_state:
        st.session_state['packaging_list'] = []
    # --- Packaging selection ---
    # Packaging Code, Reference y Folded? en la misma fila, ambos como filtros dependientes
    col1, col2, col3 = st.columns(3)
    # Inicializar estado
    if 'selected_reference' not in st.session_state:
        st.session_state['selected_reference'] = 'All'
    if 'selected_code' not in st.session_state:
        st.session_state['selected_code'] = 'Manual'
    references = sorted(['All'] + db['Reference'].dropna().unique().tolist()) if 'Reference' in db.columns else ['All']
    codes = sorted(['Manual'] + [c for c in db['Packaging Code'].dropna().unique().tolist() if c != 'Manual'])
    # Filtrar según selección
    filtered_db = db.copy()
    if st.session_state['selected_reference'] != 'All':
        filtered_db = filtered_db[filtered_db['Reference'] == st.session_state['selected_reference']]
    if st.session_state['selected_code'] != 'Manual':
        filtered_db = filtered_db[filtered_db['Packaging Code'] == st.session_state['selected_code']]
    filtered_references = sorted(['All'] + filtered_db['Reference'].dropna().unique().tolist()) if 'Reference' in filtered_db.columns else ['All']
    filtered_codes = sorted(['Manual'] + [c for c in filtered_db['Packaging Code'].dropna().unique().tolist() if c != 'Manual'])
    # Mostrar selectores dependientes
    with col1:
        selected_code = st.selectbox('Packaging Code', filtered_codes, index=filtered_codes.index(st.session_state['selected_code']) if st.session_state['selected_code'] in filtered_codes else 0, key='code_select')
    with col2:
        selected_reference = st.selectbox('Reference', filtered_references, index=filtered_references.index(st.session_state['selected_reference']) if st.session_state['selected_reference'] in filtered_references else 0, key='ref_select')
    with col3:
        folded = st.selectbox('Folded?', ['No','Yes'])
    # Actualizar estado si cambia selección
    if selected_code != st.session_state['selected_code']:
        st.session_state['selected_code'] = selected_code
        # Si el código cambia, actualizar referencias posibles
        if selected_code != 'Manual':
            possible_refs = db[db['Packaging Code'] == selected_code]['Reference'].dropna().unique().tolist()
            if possible_refs:
                st.session_state['selected_reference'] = possible_refs[0]
        st.rerun()
    elif selected_reference != st.session_state['selected_reference']:
        st.session_state['selected_reference'] = selected_reference
        # Si la referencia cambia, actualizar códigos posibles
        if selected_reference != 'All':
            possible_codes = db[db['Reference'] == selected_reference]['Packaging Code'].dropna().unique().tolist()
            if possible_codes:
                st.session_state['selected_code'] = possible_codes[0]
            else:
                st.session_state['selected_code'] = 'Manual'
        else:
            st.session_state['selected_code'] = 'Manual'
        st.rerun()
    code_sel = st.session_state['selected_code']
    # Show data or allow manual input
    if code_sel != 'Manual':
        row = db[db['Packaging Code'] == code_sel].iloc[0]
        length = row['Length']
        width = row['Width']
        height = row['Folded Height'] if folded=='Yes' else row['Height']
        weight_empty = row['Weight EMPTY']
        part_weight = row['Part Weight']
        stacking = 2  # Default stacking for database selection
        qty = 1       # Default quantity for database selection
        # Show as a table for better UI (remove Part Weight, Nb pieces per UC, Qty per UC)
        props = ['Length (mm)', 'Width (mm)', 'Height (mm)', 'Weight EMPTY (kg)']
        vals = [length, width, height, weight_empty]
        cols = st.columns(len(props))
        for i, (col, prop, val) in enumerate(zip(cols, props, vals)):
            with col:
                st.markdown(f"<div style='text-align:center; font-weight:bold'>{prop}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='text-align:center'>{val}</div>", unsafe_allow_html=True)
    else:
        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            length = st.number_input('Length (mm)', min_value=1, value=1200)
        with col_b:
            width = st.number_input('Width (mm)', min_value=1, value=800)
        with col_c:
            height = st.number_input('Height (mm)', min_value=1, value=600)
        with col_d:
            weight_empty = st.number_input('Weight EMPTY (kg)', min_value=0.0, value=20.0)
        col_e, col_f, col_g = st.columns(3)
        with col_e:
            stacking = st.number_input('Stacking (max levels)', min_value=1, value=2)
        with col_f:
            qty = st.number_input('Number of packagings', min_value=1, value=1)
        with col_g:
            part_weight = st.number_input('Part Weight (kg)', min_value=0.0, value=10.0)
    if st.button('Add packaging'):
        st.session_state['packaging_list'].append({
            'Packaging Code': code_sel,
            'Folded': folded,
            'Length': length,
            'Width': width,
            'Height': height,
            'Weight EMPTY': weight_empty,
            'Part Weight': part_weight,
            'Stacking': stacking,
            'Quantity': qty
        })
    # Show added packaging table
    if st.session_state['packaging_list']:
        st.markdown('### Added Packaging')
        df_added = pd.DataFrame(st.session_state['packaging_list'])
        for idx, row in df_added.iterrows():
            total_weight = (row['Weight EMPTY'] + row['Part Weight']) * row['Quantity']
            total_volume = (row['Length'] * row['Width'] * row['Height'] / 1e9) * row['Quantity']
            cols = st.columns([8,1,1])
            with cols[0]:
                st.write(f"**{row['Packaging Code']}** | Length: {row['Length']} | Width: {row['Width']} | Height: {row['Height']} | Qty: {row['Quantity']} | Total weight: {total_weight:.2f} kg | Total volume: {total_volume:.2f} m³")
            with cols[1]:
                if st.button('Edit', key=f'edit_{idx}'):
                    st.session_state['edit_idx'] = idx
            with cols[2]:
                if st.button('Delete', key=f'del_{idx}'):
                    st.session_state['packaging_list'].pop(idx)
                    st.rerun()
        # Inline edit form
        if 'edit_idx' in st.session_state:
            edit_idx = st.session_state['edit_idx']
            edit_row = st.session_state['packaging_list'][edit_idx]
            st.markdown(f"#### Edit Packaging: {edit_row['Packaging Code']}")
            new_length = st.number_input('Length (mm)', min_value=1, value=int(edit_row['Length']), key='edit_length')
            new_width = st.number_input('Width (mm)', min_value=1, value=int(edit_row['Width']), key='edit_width')
            new_height = st.number_input('Height (mm)', min_value=1, value=int(edit_row['Height']), key='edit_height')
            new_weight_empty = st.number_input('Weight EMPTY (kg)', min_value=0.0, value=float(edit_row['Weight EMPTY']), key='edit_weight_empty')
            new_part_weight = st.number_input('Part Weight (kg)', min_value=0.0, value=float(edit_row['Part Weight']), key='edit_part_weight')
            new_stacking = st.number_input('Stacking (max levels)', min_value=1, value=int(edit_row['Stacking']), key='edit_stacking')
            new_quantity = st.number_input('Number of packagings', min_value=1, value=int(edit_row['Quantity']), key='edit_quantity')
            if st.button('Save changes', key='save_edit'):
                st.session_state['packaging_list'][edit_idx] = {
                    'Packaging Code': edit_row['Packaging Code'],
                    'Folded': edit_row['Folded'],
                    'Length': new_length,
                    'Width': new_width,
                    'Height': new_height,
                    'Weight EMPTY': new_weight_empty,
                    'Part Weight': new_part_weight,
                    'Stacking': new_stacking,
                    'Quantity': new_quantity
                }
                del st.session_state['edit_idx']
                st.rerun()
            if st.button('Cancel', key='cancel_edit'):
                del st.session_state['edit_idx']
                st.rerun()
        # Calculate total weight and volume (do NOT use Nb pieces per UC or Qty per UC)
        total_weight = sum((p['Weight EMPTY']+p['Part Weight'])*p['Quantity'] for p in st.session_state['packaging_list'])
        total_volume = sum((p['Length']*p['Width']*p['Height']/1e9)*p['Quantity'] for p in st.session_state['packaging_list'])
        st.info(f"Total weight: {total_weight:.2f} kg, Total volume: {total_volume:.2f} m³")
        if st.button('Calculate'):
            # --- Packing con 3dbinpacking (stacking real) ---
            pkgs = st.session_state['packaging_list']
            # Crear el contenedor
            container_dims = [container['length'], container['width'], container['height']]
            container_max_weight = container['max_weight']
            # Preparar lista de cajas para 3dbinpacking
            boxes = []
            for idx, p in enumerate(pkgs):
                l = int(p['Length'])
                w = int(p['Width'])
                h = int(p['Height'])
                stacking = int(p['Stacking'])
                weight = float(p['Weight EMPTY'])+float(p['Part Weight'])
                for i in range(int(p['Quantity'])):
                    boxes.append({
                        'id': f"{p['Packaging Code']}_{i}",
                        'length': l,
                        'width': w,
                        'height': h,
                        'weight': weight,
                        'stacking': stacking,
                        'type': p['Packaging Code']
                    })
            # Packing
            try:
                packer = Packer(
                    container_dims,
                    max_weight=container_max_weight,
                    support_surface_ratio=1.0, # solo apilar si hay soporte total
                    vertical_stacking=True
                )
                for box in boxes:
                    packer.add_box(box)
                packer.pack()
                # Obtener resultados
                placed_boxes = packer.bins[0]['items'] if packer.bins else []
                used_weight = sum(b['weight'] for b in placed_boxes)
                used_volume = sum(b['length']*b['width']*b['height']/1e9 for b in placed_boxes)
                st.success(f"Packed {len(placed_boxes)} boxes.")
                st.info(f"Weight saturation: {100*used_weight/container_max_weight:.1f}% | Volume saturation: {100*used_volume/(container['length']*container['width']*container['height']/1e9):.1f}%")
                st.info(f"Total packed weight: {used_weight:.2f} kg | Total packed volume: {used_volume:.2f} m³")
                if not placed_boxes:
                    st.warning("No boxes could be packed. Check dimensions and container size.")
                else:
                    plot_3d_boxes_3dbinpacking(pkgs, placed_boxes, container)
            except Exception as e:
                st.error(f"Packing failed: {e}")

def plot_3d_boxes_py3dbp(pkgs, placed_positions, container):
    import plotly.graph_objects as go
    fig = go.Figure()
    l, w, h = int(container['length']), int(container['width']), int(container['height'])
    # Transparent faces
    fig.add_trace(go.Mesh3d(
        x=[0, l, l, 0, 0, l, l, 0],
        y=[0, 0, w, w, 0, 0, w, w],
        z=[0, 0, 0, 0, h, h, h, h],
        i=[0,0,0,4,4,2], j=[1,2,3,5,6,3], k=[2,3,1,6,7,7],
        color='lightblue', opacity=0.08, name='Container', showscale=False
    ))
    # Borde del contenedor
    fig.add_trace(go.Scatter3d(
        x=[0,l,l,0,0,0,l,l,l,0,0,0,l,l],
        y=[0,0,w,w,0,0,0,0,w,w,w,0,0],
        z=[0,0,0,0,0,h,h,0,0,0,h,h,h,h],
        mode='lines', line=dict(color='black', width=5), showlegend=False
    ))
    # Colores para cada tipo
    colors = ['orange','green','red','blue','purple','yellow','brown','pink']
    for idx, items in enumerate(placed_positions):
        color = colors[idx%len(colors)]
        for item in items:
            # Now: x=width, y=length, z=height
            x0 = float(item.position[0])  # width (X) -> x
            y0 = float(item.position[2])  # depth (Z) -> y
            z0 = float(item.position[1])  # height (Y) -> z
            wx = float(item.width)        # Width
            ly = float(item.depth)        # Length
            hz = float(item.height)       # Height
            # Check if box is out of bounds
            out_of_bounds = (x0+wx > l) or (y0+ly > w) or (z0+hz > h) or (x0 < 0) or (y0 < 0) or (z0 < 0)
            box_color = 'red' if out_of_bounds else color
            # 8 vertices
            X = [x0, x0+wx, x0+wx, x0, x0, x0+wx, x0+wx, x0]
            Y = [y0, y0, y0+ly, y0+ly, y0, y0, y0+ly, y0+ly]
            Z = [z0, z0, z0, z0, z0+hz, z0+hz, z0+hz, z0+hz]
            faces = [
                [0,1,2,3], [4,5,6,7], [0,1,5,4], [2,3,7,6], [1,2,6,5], [0,3,7,4],
            ]
            for f in faces:
                fig.add_trace(go.Mesh3d(
                    x=[X[f[0]], X[f[1]], X[f[2]], X[f[3]]],
                    y=[Y[f[0]], Y[f[1]], Y[f[2]], Y[f[3]]],
                    z=[Z[f[0]], Z[f[1]], Z[f[2]], Z[f[3]]],
                    i=[0], j=[1], k=[2],
                    color=box_color, opacity=0.85, showscale=False, name=pkgs[idx]['Packaging Code'],
                ))
            fig.add_trace(go.Scatter3d(
                x=[X[0],X[1],X[2],X[3],X[0],X[4],X[5],X[1],X[5],X[6],X[2],X[6],X[7],X[3],X[7],X[4]],
                y=[Y[0],Y[1],Y[2],Y[3],Y[0],Y[4],Y[5],Y[1],Y[5],Y[6],Y[2],Y[6],Y[7],Y[3],Y[7],Y[4]],
                z=[Z[0],Z[1],Z[2],Z[3],Z[0],Z[4],Z[5],Z[1],Z[5],Z[6],Z[2],Z[6],Z[7],Z[3],Z[7],Z[4]],
                mode='lines', line=dict(color='red' if out_of_bounds else 'black', width=3), showlegend=False
            ))
    fig.update_layout(
        scene=dict(
            xaxis_title='Width (mm)',
            yaxis_title='Length (mm)',
            zaxis_title='Height (mm)',
            yaxis=dict(autorange='reversed'),
            aspectmode='data',
        ),
        margin=dict(l=0, r=0, b=0, t=30),
        height=700,
        title={
            'text': '3D Container Loading Visualization (Door at y=0)',
            'x': 0.5,
            'xanchor': 'center'
        }
    )
    st.plotly_chart(fig, use_container_width=True)

def plot_3d_boxes_3dbinpacking(pkgs, placed_boxes, container):
    import plotly.graph_objects as go
    fig = go.Figure()
    l, w, h = int(container['length']), int(container['width']), int(container['height'])
    # Transparent faces
    fig.add_trace(go.Mesh3d(
        x=[0, l, l, 0, 0, l, l, 0],
        y=[0, 0, w, w, 0, 0, w, w],
        z=[0, 0, 0, 0, h, h, h, h],
        i=[0,0,0,4,4,2], j=[1,2,3,5,6,3], k=[2,3,1,6,7,7],
        color='lightblue', opacity=0.08, name='Container', showscale=False
    ))
    # Borde del contenedor
    fig.add_trace(go.Scatter3d(
        x=[0,l,l,0,0,0,l,l,l,0,0,0,l,l],
        y=[0,0,w,w,0,0,0,0,w,w,w,0,0],
        z=[0,0,0,0,0,h,h,0,0,0,h,h,h,h],
        mode='lines', line=dict(color='black', width=5), showlegend=False
    ))
    # Colores para cada tipo
    colors = ['orange','green','red','blue','purple','yellow','brown','pink']
    type_to_color = {}
    color_idx = 0
    for box in placed_boxes:
        t = box['type']
        if t not in type_to_color:
            type_to_color[t] = colors[color_idx%len(colors)]
            color_idx += 1
        color = type_to_color[t]
        x0 = box['position'][0]
        y0 = box['position'][1]
        z0 = box['position'][2]
        lx = box['length']
        wx = box['width']
        hx = box['height']
        # 8 vertices
        X = [x0, x0+lx, x0+lx, x0, x0, x0+lx, x0+lx, x0]
        Y = [y0, y0, y0+wx, y0+wx, y0, y0, y0+wx, y0+wx]
        Z = [z0, z0, z0, z0, z0+hx, z0+hx, z0+hx, z0+hx]
        faces = [
            [0,1,2,3], [4,5,6,7], [0,1,5,4], [2,3,7,6], [1,2,6,5], [0,3,7,4],
        ]
        for f in faces:
            fig.add_trace(go.Mesh3d(
                x=[X[f[0]], X[f[1]], X[f[2]], X[f[3]]],
                y=[Y[f[0]], Y[f[1]], Y[f[2]], Y[f[3]]],
                z=[Z[f[0]], Z[f[1]], Z[f[2]], Z[f[3]]],
                i=[0], j=[1], k=[2],
                color=color, opacity=0.85, showscale=False, name=t,
            ))
        fig.add_trace(go.Scatter3d(
            x=[X[0],X[1],X[2],X[3],X[0],X[4],X[5],X[1],X[5],X[6],X[2],X[6],X[7],X[3],X[7],X[4]],
            y=[Y[0],Y[1],Y[2],Y[3],Y[0],Y[4],Y[5],Y[1],Y[5],Y[6],Y[2],Y[6],Y[7],Y[3],Y[7],Y[4]],
            z=[Z[0],Z[1],Z[2],Z[3],Z[0],Z[4],Z[5],Z[1],Z[5],Z[6],Z[2],Z[6],Z[7],Z[3],Z[7],Z[4]],
            mode='lines', line=dict(color='black', width=3), showlegend=False
        ))
    fig.update_layout(
        scene=dict(
            xaxis_title='Length (mm)',
            yaxis_title='Width (mm)',
            zaxis_title='Height (mm)',
            aspectmode='data',
        ),
        margin=dict(l=0, r=0, b=0, t=30),
        height=700,
        title={
            'text': '3D Container Loading Visualization (Stacking Real)',
            'x': 0.5,
            'xanchor': 'center'
        }
    )
    st.plotly_chart(fig, use_container_width=True)

if __name__ == '__main__':
    main()


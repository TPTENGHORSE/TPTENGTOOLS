import pandas as pd
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import time
import os

def get_location(place):
    geolocator = Nominatim(user_agent="geoapi")
    try:
        location = geolocator.geocode(place, timeout=10)
        time.sleep(1)
        return location
    except Exception as e:
        print(f"Error geocoding {place}: {e}")
        return None

def calcular_distancia(origen, destino):
    loc1 = get_location(origen)
    loc2 = get_location(destino)
    if loc1 is None or loc2 is None:
        return None
    coords_1 = (loc1.latitude, loc1.longitude)
    coords_2 = (loc2.latitude, loc2.longitude)
    distancia = geodesic(coords_1, coords_2).km
    return distancia * 1.3  # Ajuste 30% carretera

def calcular_volumen(largo, ancho, alto):
    return largo * ancho * alto

def calcular_cantidad_contenedor(volumen_unitario, volumen_contenedor=76):
    # 76 m3 es el volumen típico de un 40HC
    if volumen_unitario == 0:
        return 0
    return int(volumen_contenedor // volumen_unitario)

def procesar_quotation(plantilla_df, base_emb_df, inland_df, rates_df, lead_time_df=None):
    df = plantilla_df.copy()
    base_emb = base_emb_df.copy()
    inland = inland_df.copy()
    rates = rates_df.copy()

    # Limpia espacios en los nombres de columna
    base_emb.columns = base_emb.columns.str.strip()

    # Mostrar columnas detectadas
    print("Columnas detectadas en base_emb:", list(base_emb.columns))
    try:
        import streamlit as st
        st.write("Columnas detectadas en Base_EMB:", list(base_emb.columns))
    except ImportError:
        pass

    results = []
    for idx, row in df.iterrows():
        origin = f"{row['Origin City']}, {row['Origin Country']}"
        dest = f"{row['Destination City']}, {row['Destination Country']}"
        dist_total = calcular_distancia(origin, dest)
        dist_city_pol = dist_pod_city = None
        pol = row.get('POL', None)
        pod = row.get('POD', None)

        if pd.notnull(pol):
            dist_city_pol = calcular_distancia(origin, pol)
        if pd.notnull(pod):
            dist_pod_city = calcular_distancia(pod, dest)

        # Volumen y peso
        packaging_code = row['Packaging Code']
        emb_row = base_emb[base_emb['Packaging Code'] == packaging_code]
        if not emb_row.empty:
            try:
                largo = emb_row.iloc[0]['Largo'] / 1000
                ancho = emb_row.iloc[0]['Width (mm)'] / 1000
                alto = emb_row.iloc[0]['Height (mm)'] / 1000
            except KeyError as e:
                print(f"Error: Falta columna {e} en base_emb")
                largo = ancho = alto = 0
                volumen_unitario = 0
                cantidad_contenedor = 0
                peso_packaging = 0
            else:
                volumen_unitario = calcular_volumen(largo, ancho, alto)
                cantidad_contenedor = calcular_cantidad_contenedor(volumen_unitario)
                peso_packaging = emb_row.iloc[0].get('Weight EMPTY (kg)', 0)
        else:
            volumen_unitario = 0
            cantidad_contenedor = 0
            peso_packaging = 0

        peso_total = row['PN Weight'] + peso_packaging

        # Costos
        pais = row['Origin Country']
        inland_row = inland[inland['Country'] == pais]
        eur_km = inland_row.iloc[0]['Eur/km'] if not inland_row.empty else 0

        if pd.isnull(pol) and pd.isnull(pod):
            inland_cost = dist_total * eur_km if dist_total else 0
            overseas_cost = 0
        else:
            inland_cost = (dist_city_pol + dist_pod_city) * eur_km if dist_city_pol and dist_pod_city else 0
            rate_row = rates[(rates['POL'] == pol) & (rates['POD'] == pod)]
            overseas_cost = rate_row.iloc[0]['Rate 40ft all-in'] if not rate_row.empty else 0

        total_cost = inland_cost + overseas_cost

        results.append({
            **row,
            'Inland Cost': inland_cost,
            'Overseas Cost': overseas_cost,
            'Total Transportation Cost': total_cost,
            'Volumen Unitario': volumen_unitario,
            'Cantidad x 40HC': cantidad_contenedor,
            'Peso Total': peso_total
        })

    return pd.DataFrame(results)

if __name__ == "__main__":
    base_path = os.path.join(os.path.dirname(__file__), "Dataframe")
    files = {
        "Plantilla_Quotation.xlsx": os.path.join(base_path, "Plantilla_Quotation.xlsx"),
        "Base_EMB.xlsx": os.path.join(base_path, "Base_EMB.xlsx"),
        "cifrados Overseas-Inland.xlsx": os.path.join(base_path, "cifrados Overseas-Inland.xlsx"),
        "RATES_04_2025.xlsx": os.path.join(base_path, "RATES_04_2025.xlsx")
    }

    for fname, fpath in files.items():
        if not os.path.exists(fpath):
            print(f"ERROR: Archivo faltante: {os.path.abspath(fpath)}")
            exit(1)

    plantilla = pd.read_excel(files["Plantilla_Quotation.xlsx"])
    base_emb = pd.read_excel(files["Base_EMB.xlsx"])
    inland = pd.read_excel(files["cifrados Overseas-Inland.xlsx"])
    rates = pd.read_excel(files["RATES_04_2025.xlsx"])

    df_result = procesar_quotation(plantilla, base_emb, inland, rates)
    output_path = os.path.join(base_path, "Quotation_Result.xlsx")
    df_result.to_excel(output_path, index=False)
    print(f"Archivo Quotation_Result.xlsx generado en {output_path}.")

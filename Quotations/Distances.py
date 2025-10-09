from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import time

geolocator = Nominatim(user_agent="geoapi")  # Lo instanciamos una vez

def get_location(place):
    """
    Devuelve el objeto de ubicación de una ciudad/puerto.
    """
    try:
        location = geolocator.geocode(place, timeout=5)
        time.sleep(1)  # Evitar bloqueo de Nominatim
        return location
    except Exception as e:
        print(f"❌ Error geocoding '{place}': {e}")
        return None

def calcular_distancia(origen, destino, factor_ajuste=1.30):
    """
    Calcula la distancia geodésica entre dos lugares.
    Usa un factor opcional para simular ruta real (por carretera).
    """
    loc1 = get_location(origen)
    loc2 = get_location(destino)
    if not loc1 or not loc2:
        print("⚠️ No se encontró una de las ubicaciones.")
        return None

    coords1 = (loc1.latitude, loc1.longitude)
    coords2 = (loc2.latitude, loc2.longitude)

    distancia_km = geodesic(coords1, coords2).km
    return distancia_km * factor_ajuste  # Ruta estimada

# Ejemplo de uso
if __name__ == "__main__":
    origen = "Osaka, Japon"
    destino = "Tokyo, Japon"
    distancia = calcular_distancia(origen, destino)

    if distancia:
        print(f"📏 Distancia estimada entre {origen} y {destino}: {distancia:.2f} km")
    else:
        print("❌ No se pudo calcular la distancia.")


"""
Quotation Tool V0 - Basic Restored Version

Este script es una versión básica restaurada para cotizaciones.
Completa la lógica según tus necesidades.
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from time import sleep
from rapidfuzz import fuzz, process

st.title("Quotation Tool V0 (Restaurado)")

# PLANTILLAS
import sys
# Detectar la ruta base desde donde se ejecuta (compatible con .exe y script normal)
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # cuando está empaquetado con pyinstaller
else:
    try:
        # Cuando se ejecuta como script .py normal
        base_path = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        # Cuando se ejecuta en notebook o intérprete interactivo (sin __file__)
        base_path = os.getcwd()

# Construir la ruta al archivo Excel
Path_File = os.path.join(base_path, 'Plantilla_Quotation.xlsx')
# Cargar el Excel y transformar a mayúsculas
Template = pd.read_excel(Path_File, skiprows=1, header=0)
Template = Template.map(lambda x: x.upper() if isinstance(x, str) else x)

zip_cols = ['ZIP Code', 'ZIP Code.1']
Template[zip_cols] = Template[zip_cols].apply(pd.to_numeric, errors='coerce').astype('Int64')
Template["Maritimo"] = Template["POL"].notna() & Template["POD"].notna()
Template_Terrestre = Template[Template["Maritimo"] == False]
Template_Maritimo = Template[Template["Maritimo"] == True]

# DATABASES
Path_File = os.path.join(base_path, 'cifrados Overseas-Inland.xlsx')
CifradosInland = pd.read_excel(Path_File, sheet_name="INLAND")
CifradosInland["Picking_Country"] = CifradosInland["Picking_Country"].str.upper()
CifradosInland["Delivery_Country"] = CifradosInland["Delivery_Country"].str.upper()
CifradosInland["Origin_City"] = CifradosInland["Origin_City"].str.upper()
CifradosInland["Destination_City"] = CifradosInland["Destination_City"].str.upper()
Path_File = os.path.join(base_path, 'Distances_Costs Country_Port.xlsx')
RatesInland = pd.read_excel(Path_File)
Path_File = os.path.join(base_path, 'RATES_04_2025.xlsx')
CifradosOverseas = pd.read_excel(Path_File, sheet_name="MAIN PORTS")
PlantaPuerto = pd.read_excel(Path_File, sheet_name="HORSE-PUERTO")
PlantaPuerto["ZIP Code"] = PlantaPuerto["ZIP Code"].apply(lambda x: int(x) if pd.notnull(x) else pd.NA).astype("Int64")
Path_File = os.path.join(base_path, 'Reduced Packaging.xlsx')
Mapping_Motores = pd.read_excel(Path_File)
Mapping_Motores_reducido = Mapping_Motores[["Reference","Packaging Code","Qté / UC","Part Weight (kg)","Weight EMPTY (kg)","Hauteur hors tout","Largeur hors tout","Longueur hors tout"]]
Mapping_Motores_reducido = Mapping_Motores_reducido.rename(columns={
    "Hauteur hors tout": "Altura",
    "Largeur hors tout": "Anchura",
    "Longueur hors tout": "Largo"
})
Mapping_Motores_reducido = Mapping_Motores_reducido.drop_duplicates(subset=["Reference", "Packaging Code"])
Path_File = os.path.join(base_path, 'LEAD_TIME_FINAL.xlsx')
VTT_Final = pd.read_excel(Path_File)
Interes_Financiero = 0.078
# ...continúa la lógica del notebook...

# 🚛 QuotationTool - Herramienta de Cotización de Transporte

## 📋 Descripción

QuotationTool es una herramienta avanzada para el cálculo automático de costos de transporte que procesa tanto flujos marítimos como terrestres, integrando datos de embalaje, lead times y costos financieros.

## ✨ Características Principales

### 🚢 Flujo Marítimo
- **Tramo 1**: Origen/Planta → Puerto (POL)
- **Tramo 2**: Puerto origen (POL) → Puerto destino (POD) 
- **Tramo 3**: Puerto destino (POD) → Destino final
- Cálculos de saturación por peso/volumen
- Optimización de contenedores 40ft

### 🚛 Flujo Terrestre
- Rutas directas terrestre
- Optimización de camiones
- Cálculos de saturación adaptados

### 📦 Gestión de Embalajes
- Integración con base de datos de packaging
- Cálculos de volumen y peso
- Determinación automática de saturación (V/W)

### ⏰ Lead Times y Costos Financieros
- Integración con VTT (Vehicle Transit Time)
- Cálculo de Floating Stock
- Interés financiero configurable

## 📁 Archivos Requeridos

Coloca estos archivos en el directorio `Quotations/Dataframe/`:

### Archivos Obligatorios:
1. **`Plantilla_Quotation.xlsx`** - Datos de entrada con las cotizaciones
2. **`cifrados Overseas-Inland.xlsx`** - Tarifas terrestres por ruta
3. **`RATES_04_2025.xlsx`** - Tarifas marítimas y datos planta-puerto
4. **`Base_EMB.xlsx`** - Base de datos de embalajes
5. **`LEAD_TIME_FINAL.xlsx`** - Tiempos de tránsito (VTT)

### Archivos Opcionales (con alternativas):
6. **`Distances_Costs Country_Port.xlsx`** - Distancias y costos país-puerto *(si no existe, usa datos de RATES_04_2025)*
7. **`Reduced Packaging.xlsx`** - Base de datos de embalajes mejorada *(si no existe, usa Base_EMB.xlsx)*

## 🚀 Uso

### Desde la Interfaz Web (Streamlit)
1. Ejecuta `streamlit run app.py`
2. Selecciona "QuotationTool" en el menú lateral
3. Verifica que todos los archivos estén presentes
4. Haz clic en "▶️ Procesar Quotation"

### Desde Línea de Comandos
```bash
cd "Quotations"
python Quotation_toolV0.py
```

## 📊 Estructura de Archivos

```
Horse Luis/
├── app.py
├── Quotations/
│   ├── Quotation_toolV0.py
│   ├── QuotationTool.ipynb
│   ├── README_QuotationTool.md
│   ├── Dataframe/
│   │   ├── Plantilla_Quotation.xlsx
│   │   ├── cifrados Overseas-Inland.xlsx
│   │   ├── RATES_04_2025.xlsx
│   │   ├── Base_EMB.xlsx
│   │   ├── LEAD_TIME_FINAL.xlsx
│   │   ├── Distances_Costs Country_Port.xlsx (opcional)
│   │   └── Reduced Packaging.xlsx (opcional)
│   ├── Maritime Tool/          # Archivos generados
│   └── Land Tool/              # Archivos generados
```

## 📊 Estructura de Datos de Entrada

### Columnas Requeridas en `Plantilla_Quotation.xlsx`:

| Columna | Descripción |
|---------|-------------|
| `Country` | País de origen |
| `ZIP Code` | Código postal origen |
| `Name` | Nombre del proveedor origen |
| `City` | Ciudad origen |
| `Country.1` | País de destino |
| `ZIP Code.1` | Código postal destino |
| `Name.1` | Nombre del destinatario |
| `City.1` | Ciudad destino |
| `POL` | Puerto de carga (opcional) |
| `POD` | Puerto de descarga (opcional) |
| `Part Number (PN)` | Número de parte |
| `Packaging Code` | Código de embalaje |
| `Unit cost (€)` | Costo unitario |
| `Anual Needs` | Necesidades anuales |
| `Daily Need` | Necesidad diaria |

## 📈 Outputs Generados

### Archivos de Salida
- **`Maritime Tool/Maritime_Template_YYYY-MM-DD_X.xlsx`** - Resultados flujo marítimo
- **`Land Tool/Land_Template_YYYY-MM-DD_X.xlsx`** - Resultados flujo terrestre

### Columnas Calculadas
- **`LOG €/Part`** - Costo logístico por pieza
- **`TOTAL €/Part`** - Costo total por pieza (logística + material + financiero)
- **`Floating Stock €/Part`** - Costo financiero del stock en tránsito
- **`Annual weight`** - Peso anual en toneladas
- **`FCF Pipe`** - Free Cash Flow Pipeline
- **`Error Indicator`** - Indicadores de errores/datos faltantes

## ⚙️ Configuración

### Parámetros Configurables (en `Quotation_toolV0.py`):

```python
# Interés financiero anual
Interes_Financiero = 0.078  # 7.8%

# Saturación de contenedores/camiones
Filling_Weight = 24750  # kg máximo
Filling_Rate_Max_Terrestre = 85  # m³ para terrestre
Filling_Rate_Max_Maritimo = 62   # m³ para marítimo
```

## 🔧 Instalación de Dependencias

```bash
pip install pandas numpy openpyxl rapidfuzz streamlit
```

O usando requirements.txt:
```bash
pip install -r requirements.txt
```

## ❗ Manejo de Errores

El sistema detecta automáticamente:
- ✅ Combinaciones de rutas no encontradas
- ✅ Códigos de embalaje inexistentes
- ✅ Referencias de productos faltantes
- ✅ VTTs no disponibles

Los errores se reportan en la columna `Error Indicator` del archivo de salida.

## 🆕 Nuevas Funcionalidades vs Versión Anterior

### ✅ Mejoras Implementadas:
- **Procesamiento completo**: Flujos marítimos y terrestres
- **Matching inteligente**: Usando rapidfuzz para similitud de nombres
- **Validación exhaustiva**: Detección de errores en todos los niveles
- **Exportación mejorada**: Tablas Excel formateadas automáticamente
- **Interfaz web**: Integración completa con Streamlit
- **Cálculos financieros**: Floating stock y FCF pipe
- **Documentación**: Indicadores de error detallados

### 🔄 Compatibilidad:
- Mantiene funciones de la versión anterior para compatibilidad
- Estructura de archivos de entrada sin cambios
- API similar para integración existente

## 🐛 Troubleshooting

### Errores Comunes:

1. **"Archivo no encontrado"**
   - Verifica que todos los archivos estén en el directorio correcto
   - Revisa nombres de archivos (case-sensitive)

2. **"Combination not found"**
   - Actualiza las bases de datos de referencia
   - Verifica formato de códigos postales y nombres de ciudades

3. **"Memory Error"**
   - Procesa archivos más pequeños
   - Aumenta memoria disponible

## 📞 Soporte

Para reportar bugs o solicitar funcionalidades, contacta al equipo de desarrollo.

---
*Versión 2.0 - Julio 2025*

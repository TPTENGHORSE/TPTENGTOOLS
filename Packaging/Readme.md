# 📦 Empower 3D App

Aplicación interactiva en Streamlit para calcular cuántas cajas caben en un contenedor marítimo, considerando rotaciones y saturaciones de volumen y peso. Incluye visualización 3D.

## 🖥️ ¿Qué hace?

- Permite seleccionar el tipo de contenedor (20ft o 40HC).
- Introducir dimensiones y peso de la caja.
- Calcula el número máximo de cajas que caben dentro.
- Muestra saturación de volumen y peso del contenedor.
- Visualiza la distribución 3D con cajas apiladas.

---

## 🚀 Cómo ejecutar la app en tu ordenador

> 📝 Requisitos: Tener Python instalado (3.9 o superior).

### 1. Clona el repositorio o descarga los archivos

- Si usas OneDrive, asegúrate de copiar todos estos archivos en una misma carpeta:

### 2. Abre una terminal (CMD o PowerShell)

Ubícate dentro de la carpeta del proyecto:
Container3D/
├── app.py
├── logo.png
├── requirements.txt

```bash
cd ruta\a\la\carpeta\Container3D

### 3. Crea un entorno virtual
python -m venv .venv

### 4. Activa el entorno virtual
En Windows:


.venv\Scripts\activate

### 5. Instala los paquetes necesarios

pip install -r requirements.txt

### 6. Ejecuta la app

streamlit run Container3D.py
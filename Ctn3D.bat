@echo off
REM Activar entorno virtual
call .venv\Scripts\activate

REM Ejecutar la app
streamlit run Container3D.py

REM Mantener la ventana abierta al final (opcional)
pause

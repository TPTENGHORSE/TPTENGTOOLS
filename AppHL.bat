@echo off
REM Activar el entorno virtual si existe
IF EXIST .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)
REM Ejecutar la app de Streamlit
echo Iniciando Transport Engineering Tools...
streamlit run app.py
pause

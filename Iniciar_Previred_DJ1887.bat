@echo off
title Conta Marlene - Previred & DJ 1887 SII
chcp 65001 > nul
echo =====================================================================
echo    Iniciando Conta Marlene: Sistema Previred y DJ 1887 SII
echo =====================================================================
echo.
echo Abriendo la aplicacion en su navegador web...
echo.

set PYTHON_CMD="C:\Users\Acer\AppData\Local\Programs\Python\Python311\python.exe"

if exist %PYTHON_CMD% (
    %PYTHON_CMD% -m streamlit run app.py
) else (
    python -m streamlit run app.py
)

pause

@echo off
title Lanzador Proyecto Merge — Etiquetas
echo.
echo ============================================================
echo   ⚡ PROYECTO MERGE: ETIQUETAS DWDM ⚡
echo ============================================================
echo.
echo Verificando dependencias...
pip install customtkinter tkinterdnd2 pdfplumber pandas openpyxl --quiet
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] No se pudieron instalar las librerias necesarias.
    echo Asegurate de tener Python instalado y conexion a internet.
    pause
    exit /b
)

echo Iniciando aplicacion...
python dwdm_app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] La aplicacion se cerro con errores.
    pause
)

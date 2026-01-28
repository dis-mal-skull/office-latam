@echo off
setlocal EnableDelayedExpansion
title Instalador de Office 2019/2021 - Selección Personalizada

:: 1. VERIFICAR PERMISOS DE ADMINISTRADOR
:: Intentamos acceder a una ruta de sistema para verificar admin
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :Admin
) else (
    echo.
    echo ============================================================
    echo    SOLICITANDO PERMISOS DE ADMINISTRADOR...
    echo ============================================================
    echo.
    echo Por favor acepte la solicitud para continuar con la instalacion.
    echo.
    :: Relanzar este mismo script como administrador
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:Admin
:: AQUI YA SOMOS ADMINISTRADOR
cd /d "%~dp0"
cls
echo.
echo ============================================================
echo         INSTALADOR AUTOMATICO DE OFFICE 2019 / 2021
echo                    SELECCIÓN PERSONALIZADA
echo ============================================================
echo.
echo Este script permite seleccionar qué programas instalar.
echo Puede elegir Excel, Word, PowerPoint, Outlook y más.
echo.

:: 2. VERIFICAR PYTHON
python --version > nul 2>&1
if %errorlevel% equ 0 (
    goto :RunPython
)

echo [INFO] Python no detectado. Iniciando instalacion automatica...
echo.

set "PYTHON_URL=https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe"
set "PYTHON_EXE=%TEMP%\python_installer.exe"

echo Descargando Python...
powershell -Command "$ProgressPreference = 'Continue'; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_EXE%'"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] No se pudo descargar Python.
    pause
    exit /b 1
)

echo Instalando Python...
start /w "" "%PYTHON_EXE%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

del "%PYTHON_EXE%" > nul 2>&1
echo [OK] Python instalado.

:RunPython
echo Iniciando instalador...
echo.

:: 3. EJECUTAR SCRIPT DE PYTHON EN LA MISMA VENTANA
:: Intentamos detectar la ruta de python si el PATH no se actualizó
set "PY_CMD=python"
if exist "C:\Program Files\Python311\python.exe" set "PY_CMD=C:\Program Files\Python311\python.exe"

"%PY_CMD%" "install_office.py"

echo.
echo Presione ENTER para salir...
pause > nul

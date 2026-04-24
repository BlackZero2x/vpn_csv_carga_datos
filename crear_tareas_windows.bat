@echo off
REM ============================================================================
REM Script para crear tareas programadas en Windows Task Scheduler
REM Ejecutar como ADMINISTRADOR
REM ============================================================================

setlocal enabledelayedexpansion

echo.
echo ============================================================
echo  CONFIGURADOR DE TAREAS - CSV Sync Anfitrion
echo ============================================================
echo.

REM Variables
set PYTHON_PATH=C:\Users\developer2\AppData\Local\Programs\Python\Python312\python.exe
set SCRIPT_PATH=C:\VPN_MIFIBRA\vpn_csv_sync.py
echo Configuracion:
echo - Horarios: 08:30, 10:30, 12:30, 14:30, 16:30, 18:30, 20:30
echo - Usuario: SYSTEM
echo - Script: %SCRIPT_PATH%
echo - Python: %PYTHON_PATH%
echo.

if not exist "%SCRIPT_PATH%" (
    echo ERROR: Script no encontrado en %SCRIPT_PATH%
    pause
    exit /b 1
)

if not exist "%PYTHON_PATH%" (
    echo ERROR: Python no encontrado en %PYTHON_PATH%
    echo Abre una terminal y ejecuta: where python
    echo Luego actualiza PYTHON_PATH en este archivo .bat
    pause
    exit /b 1
)

echo OK - Validaciones correctas
echo.

set /p CONFIRM="Escribe 'SI' para continuar: "
if /i not "%CONFIRM%"=="SI" (
    echo Cancelado.
    exit /b 0
)

echo.
echo Creando tareas...
echo.

schtasks /create /tn "CSV_Sync_0830" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 08:30 /ru "SYSTEM" /f
if errorlevel 1 (echo ERROR creando tarea 08:30) else (echo OK - Tarea 08:30 creada)

schtasks /create /tn "CSV_Sync_1030" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 10:30 /ru "SYSTEM" /f
if errorlevel 1 (echo ERROR creando tarea 10:30) else (echo OK - Tarea 10:30 creada)

schtasks /create /tn "CSV_Sync_1230" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 12:30 /ru "SYSTEM" /f
if errorlevel 1 (echo ERROR creando tarea 12:30) else (echo OK - Tarea 12:30 creada)

schtasks /create /tn "CSV_Sync_1430" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 14:30 /ru "SYSTEM" /f
if errorlevel 1 (echo ERROR creando tarea 14:30) else (echo OK - Tarea 14:30 creada)

schtasks /create /tn "CSV_Sync_1630" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 16:30 /ru "SYSTEM" /f
if errorlevel 1 (echo ERROR creando tarea 16:30) else (echo OK - Tarea 16:30 creada)

schtasks /create /tn "CSV_Sync_1830" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 18:30 /ru "SYSTEM" /f
if errorlevel 1 (echo ERROR creando tarea 18:30) else (echo OK - Tarea 18:30 creada)

schtasks /create /tn "CSV_Sync_2030" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 20:30 /ru "SYSTEM" /f
if errorlevel 1 (echo ERROR creando tarea 20:30) else (echo OK - Tarea 20:30 creada)

echo.
echo ============================================================
echo Tareas creadas. Verifica en Task Scheduler (taskschd.msc)
echo ============================================================
echo.

pause

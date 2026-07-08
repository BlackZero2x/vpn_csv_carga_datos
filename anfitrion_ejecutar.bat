@echo off
REM ============================================================================
REM Ejecutor del script de sincronizacion CSV -> Google Sheets (ANFITRION)
REM Usado por la tarea programada CSV_Sync (ver crear_tareas_windows.bat)
REM ============================================================================

cd /d C:\proyectos\VPN_MIFIBRA

set PYTHON_PATH=C:\proyectos\.venv\Scripts\python.exe
set SCRIPT=C:\proyectos\VPN_MIFIBRA\vpn_csv_sync.py
set LOG=C:\proyectos\VPN_MIFIBRA\Logs\anfitrion_ejecutar_%DATE:~-4,4%%DATE:~-7,2%%DATE:~0,2%.log

echo [%DATE% %TIME%] Iniciando vpn_csv_sync.py >> "%LOG%" 2>&1
"%PYTHON_PATH%" "%SCRIPT%" >> "%LOG%" 2>&1
echo [%DATE% %TIME%] Fin. Codigo de salida: %ERRORLEVEL% >> "%LOG%" 2>&1

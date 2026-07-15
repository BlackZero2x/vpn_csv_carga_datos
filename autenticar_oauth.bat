@echo off
REM Script para configurar autenticación OAuth con Google Sheets
REM Asegúrate de tener credentials.json en este directorio antes de ejecutar

echo.
echo ============================================================
echo AUTENTICACION OAUTH - Google Sheets
echo ============================================================
echo.

cd /d "%~dp0"

REM Verifica que credentials.json existe
if not exist "credentials.json" (
    echo ERROR: No se encuentra credentials.json
    echo.
    echo Pasos para obtenerlo:
    echo 1. Ve a: https://console.cloud.google.com
    echo 2. Habilita Google Sheets API
    echo 3. Crea credenciales OAuth (tipo: Aplicación de escritorio)
    echo 4. Descarga el archivo JSON
    echo 5. Cópialo aquí y renómbralo a: credentials.json
    echo.
    pause
    exit /b 1
)

REM Ejecuta el script de autenticación
echo Iniciando autenticación...
C:\proyectos\.venv\Scripts\python.exe authenticate_oauth.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================================
    echo OK - AUTENTICACION COMPLETADA
    echo ============================================================
    echo.
    echo El token se ha guardado en: .cache\token.pickle
    echo.
    echo vpn_csv_sync.py ahora usará este token automáticamente.
    echo.
) else (
    echo.
    echo ============================================================
    echo ERROR EN LA AUTENTICACION
    echo ============================================================
    echo.
)

pause

@echo off
REM ============================================================================
REM Script para crear tareas programadas DENTRO DE LA VM (VirtualBox)
REM Ejecutar como ADMINISTRADOR dentro de la maquina virtual
REM ============================================================================

setlocal enabledelayedexpansion

echo.
echo ============================================================
echo  CONFIGURADOR DE TAREAS - VM Descarga CSV
echo ============================================================
echo.

REM Variables - MODIFICA SI INSTALASTE PYTHON EN OTRA RUTA
set PYTHON_PATH=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python314\python.exe
set SCRIPT_PATH=C:\Scripts\vm_descargar_csvs.py
set TASK_USER=%COMPUTERNAME%\%USERNAME%

echo Configuracion:
echo - Horarios: 08:25, 10:25, 12:25, 14:25, 16:25, 18:25, 20:25 (5 min antes que el anfitrion)
echo - Usuario: %TASK_USER%
echo - Script: %SCRIPT_PATH%
echo - Python: %PYTHON_PATH%
echo.

if not exist "%SCRIPT_PATH%" (
    echo ERROR: Script no encontrado en %SCRIPT_PATH%
    echo Copia vm_descargar_csvs.py a C:\Scripts\ dentro de la VM
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

schtasks /create /tn "VM_CSV_Descarga_0825" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 08:25 /ru "%COMPUTERNAME%\%USERNAME%" /f
if errorlevel 1 (echo ERROR creando tarea 08:25) else (echo OK - Tarea 08:25 creada)

schtasks /create /tn "VM_CSV_Descarga_1025" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 10:25 /ru "%COMPUTERNAME%\%USERNAME%" /f
if errorlevel 1 (echo ERROR creando tarea 10:25) else (echo OK - Tarea 10:25 creada)

schtasks /create /tn "VM_CSV_Descarga_1225" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 12:25 /ru "%COMPUTERNAME%\%USERNAME%" /f
if errorlevel 1 (echo ERROR creando tarea 12:25) else (echo OK - Tarea 12:25 creada)

schtasks /create /tn "VM_CSV_Descarga_1425" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 14:25 /ru "%COMPUTERNAME%\%USERNAME%" /f
if errorlevel 1 (echo ERROR creando tarea 14:25) else (echo OK - Tarea 14:25 creada)

schtasks /create /tn "VM_CSV_Descarga_1625" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 16:25 /ru "%COMPUTERNAME%\%USERNAME%" /f
if errorlevel 1 (echo ERROR creando tarea 16:25) else (echo OK - Tarea 16:25 creada)

schtasks /create /tn "VM_CSV_Descarga_1825" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 18:25 /ru "%COMPUTERNAME%\%USERNAME%" /f
if errorlevel 1 (echo ERROR creando tarea 18:25) else (echo OK - Tarea 18:25 creada)

schtasks /create /tn "VM_CSV_Descarga_2025" /tr "^"%PYTHON_PATH%^" ^"%SCRIPT_PATH%^"" /sc daily /st 20:25 /ru "%COMPUTERNAME%\%USERNAME%" /f
if errorlevel 1 (echo ERROR creando tarea 20:25) else (echo OK - Tarea 20:25 creada)

echo.
echo ============================================================
echo Tareas creadas en la VM. Cadena completa:
echo   08:25 VM  ->  08:30 Anfitrion
echo   10:25 VM  ->  10:30 Anfitrion
echo   12:25 VM  ->  12:30 Anfitrion
echo   14:25 VM  ->  14:30 Anfitrion
echo   16:25 VM  ->  16:30 Anfitrion
echo   18:25 VM  ->  18:30 Anfitrion
echo   20:25 VM  ->  20:30 Anfitrion
echo ============================================================
echo.

pause

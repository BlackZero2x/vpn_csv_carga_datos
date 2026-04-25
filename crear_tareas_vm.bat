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

echo Configuracion:
echo - Horarios: 08:25, 10:25, 12:25, 14:25, 16:25, 18:25, 20:25
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
echo Creando tareas via PowerShell (requiere elevation)...
echo.

PowerShell -NonInteractive -Command ^
  "$python = '%PYTHON_PATH%'; $script = '%SCRIPT_PATH%';" ^
  "$accion = New-ScheduledTaskAction -Execute $python -Argument $script;" ^
  "$config = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew;" ^
  "foreach ($hora in @('08:25','10:25','12:25','14:25','16:25','18:25','20:25')) {" ^
  "  $nombre = 'VM_CSV_Descarga_' + $hora.Replace(':','');" ^
  "  $trigger = New-ScheduledTaskTrigger -Daily -At $hora;" ^
  "  Register-ScheduledTask -TaskName $nombre -Action $accion -Trigger $trigger -Settings $config -RunLevel Highest -User 'SYSTEM' -Force | Out-Null;" ^
  "  Write-Host \"OK - $nombre\";" ^
  "}"

echo.
echo ============================================================
echo Tareas creadas en la VM. Cadena completa:
echo   08:25 VM  -^>  08:30 Anfitrion
echo   10:25 VM  -^>  10:30 Anfitrion
echo   12:25 VM  -^>  12:30 Anfitrion
echo   14:25 VM  -^>  14:30 Anfitrion
echo   16:25 VM  -^>  16:30 Anfitrion
echo   18:25 VM  -^>  18:30 Anfitrion
echo   20:25 VM  -^>  20:30 Anfitrion
echo ============================================================
echo.
echo Para verificar: PowerShell -Command "Get-ScheduledTask -TaskName 'VM_CSV_Descarga_*'"
echo.

pause

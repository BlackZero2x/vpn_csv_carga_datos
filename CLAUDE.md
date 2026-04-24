# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Propósito del proyecto

Aplicación de escritorio Windows que automatiza la sincronización de archivos CSV desde un servidor corporativo (accesible por VPN) hacia Google Sheets. El flujo completo es: conectar VPN → descargar CSVs via ruta UNC → desconectar VPN → subir datos a Google Sheets. Se ejecuta mediante tareas programadas de Windows (8:00, 12:00 y 16:00).

## Dependencias Python

No hay `requirements.txt`. Instalar manualmente:

```bash
pip install pandas gspread google-auth-oauthlib google-auth-httplib2
```

## Comandos principales

```bash
# Ejecutar la sincronización completa
python vpn_csv_sync.py

# Ejecutar diagnóstico del entorno (verifica Python, rutas, Google Sheets, VPN, etc.)
python diagnostico_vpn_sync.py

# Crear las tareas programadas en Windows (requiere ejecutar como Administrador)
# Hacer clic derecho sobre el .bat → Ejecutar como administrador
crear_tareas_windows.bat
```

## Arquitectura

El proyecto tiene estructura plana con 3 archivos funcionales:

- **`vpn_csv_sync.py`** — Script principal (~502 líneas). La función `sincronizar_completo()` (línea 383) orquesta los 4 pasos del flujo. La configuración está centralizada en el dict `CONFIG` (líneas 25–55) que el usuario debe editar directamente.
- **`diagnostico_vpn_sync.py`** — Herramienta de diagnóstico (~351 líneas). Ejecuta 7 verificaciones independientes del entorno y genera un reporte final.
- **`crear_tareas_windows.bat`** — Crea 3 tareas en el Programador de tareas de Windows que ejecutan `vpn_csv_sync.py` a las 08:00, 12:00 y 16:00.

### Módulos internos de `vpn_csv_sync.py`

| Función | Líneas | Responsabilidad |
|---|---|---|
| `conectar_vpn()` | 121–166 | Conecta VPN usando `rasdial.exe` con reintentos |
| `desconectar_vpn()` | 168–200 | Desconecta VPN; se llama también en caso de error |
| `descargar_csvs()` | 206–257 | Copia CSVs desde ruta UNC local con `shutil.copy2()` |
| `conectar_google_sheets()` | 259–291 | Autenticación con cuenta de servicio (JSON) |
| `cargar_csv_a_sheets()` | 293–360 | Sube un CSV a una hoja; borra y reescribe en modo overwrite |
| `cargar_todos_csvs()` | 362–380 | Itera sobre todos los CSVs del `CONFIG` |
| `sincronizar_completo()` | 383–480 | Orquestador principal; desconecta VPN en `finally` |

### Configuración (`CONFIG` en vpn_csv_sync.py:25-55)

Los valores que el usuario debe configurar son:
- `vpn_name`: nombre de la conexión VPN en Windows
- `vpn_user` / `vpn_password`: credenciales VPN (almacenadas en texto plano)
- `csv_network_path`: ruta UNC del servidor (ej. `\\SERVIDOR\shared\folder`)
- `google_credentials_file`: ruta al JSON de cuenta de servicio de Google Cloud
- `spreadsheet_id`: ID del Google Sheets de destino
- `csv_files`: dict con nombre de archivo → nombre de hoja de destino

## Rutas esperadas en producción

```
C:\Data\Scripts\        → copias de los scripts .py y .bat
C:\Data\CSV_Sync\       → CSVs descargados temporalmente
C:\Data\Logs\           → logs diarios (vpn_sync_YYYYMMDD.log)
C:\Credentials\         → google_credentials.json (cuenta de servicio)
```

## Integración con VPN

Usa `rasdial.exe` (herramienta nativa de Windows). Es compatible con conexiones VPN tipo RAS/GlobalProtect configuradas en el sistema. La contraseña se pasa como argumento de línea de comandos — esto es una limitación conocida documentada en `GUIA_INSTALACION.md`.

## Google Sheets

Autenticación via cuenta de servicio (no OAuth interactivo). El archivo JSON de credenciales debe obtenerse desde Google Cloud Console. Si una hoja del spreadsheet no existe, se crea automáticamente. Cada sincronización **borra y reescribe** el contenido (modo overwrite). Maneja automáticamente codificaciones UTF-8 y latin-1.

## Logging

Los logs se escriben en consola y en archivo (`C:\Data\Logs\vpn_sync_YYYYMMDD.log`). El sistema de logging se inicializa en las líneas 57–93 de `vpn_csv_sync.py`.

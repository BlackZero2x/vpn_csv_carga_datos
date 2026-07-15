# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Propósito del proyecto

Aplicación de escritorio Windows que automatiza la sincronización de archivos CSV desde un servidor corporativo (accesible por VPN) hacia Google Sheets. El flujo completo es: conectar VPN → descargar CSVs via ruta UNC → desconectar VPN → subir datos a Google Sheets. Se ejecuta mediante tareas programadas de Windows (10:05, 12:30, 14:30, 16:30, 18:30 y 20:30).

## Dependencias Python

Usa el entorno virtual compartido en `C:\proyectos\.venv\`. Las dependencias están en `C:\proyectos\requirements.txt`.

```bash
# Instalar si fuera necesario (desde C:\proyectos\)
pip install pandas gspread google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

**Librerías clave**:
- `gspread`: Cliente para Google Sheets API
- `google-auth-oauthlib`: OAuth 2.0 flow
- `google-api-python-client`: Para operaciones de Sheets API (formato, fórmulas)

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
- **`crear_tareas_windows.bat`** — Crea 6 tareas en el Programador de tareas de Windows que ejecutan `vpn_csv_sync.py` a las 10:05, 12:30, 14:30, 16:30, 18:30 y 20:30.

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

## Rutas del proyecto

```
C:\proyectos\VPN_MIFIBRA\          → scripts .py y .bat
C:\proyectos\VPN_MIFIBRA\.env      → credenciales VPN y configuración sensible
C:\proyectos\VPN_MIFIBRA\CSV_Sync\ → CSVs descargados temporalmente
C:\proyectos\VPN_MIFIBRA\logs\     → logs diarios (vpn_sync_YYYYMMDD.log)
C:\proyectos\shared\credentials\   → google_credentials.json (cuenta de servicio)
```

Las credenciales Google se leen desde `.env` (variable `GOOGLE_CREDENTIALS_FILE`). Las credenciales VPN (`vpn_user`, `vpn_password`) también deben migrarse al `.env` para no quedar en texto plano dentro del `CONFIG` de `vpn_csv_sync.py`.

## Integración con VPN

Usa `rasdial.exe` (herramienta nativa de Windows). Es compatible con conexiones VPN tipo RAS/GlobalProtect configuradas en el sistema. La contraseña se pasa como argumento de línea de comandos — esto es una limitación conocida documentada en `GUIA_INSTALACION.md`.

## Google Sheets

**Autenticación con OAuth (Recomendado)**

El script intenta usar OAuth primero. Para configurarlo:
1. Descarga `credentials.json` desde [Google Cloud Console](https://console.cloud.google.com)
2. Colócalo en `C:\proyectos\VPN_MIFIBRA\`
3. Ejecuta: `python authenticate_oauth.py`
4. Autoriza en el navegador que se abre

El token se guarda en `.cache/token.pickle` y se reutiliza automáticamente en futuras ejecuciones.

**Fallback a Cuenta de Servicio**

Si no hay token OAuth válido, el script intenta usar credenciales de cuenta de servicio desde `GOOGLE_CREDENTIALS_FILE` (`.env`). Esto requiere compartir los spreadsheets con la cuenta de servicio.

Ver `OAUTH_SETUP.md` para instrucciones detalladas.

**Características Generales**

Si una hoja del spreadsheet no existe, se crea automáticamente. Cada sincronización **borra y reescribe** el contenido (modo overwrite). Maneja automáticamente codificaciones UTF-8 y latin-1.

## Logging

Los logs se escriben en consola y en archivo (`C:\proyectos\VPN_MIFIBRA\logs\vpn_sync_YYYYMMDD.log`). El sistema de logging se inicializa en las líneas 57–93 de `vpn_csv_sync.py`.

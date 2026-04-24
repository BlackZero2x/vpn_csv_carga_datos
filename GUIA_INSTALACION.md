# Guía de Instalación y Configuración
## VPN CSV Sync - Automatización de Descarga a Google Sheets

---

## 📋 Índice
1. [Requisitos previos](#requisitos-previos)
2. [Paso 1: Descargar credenciales de Google](#paso-1-descargar-credenciales-de-google)
3. [Paso 2: Configurar el script Python](#paso-2-configurar-el-script-python)
4. [Paso 3: Crear tareas programadas](#paso-3-crear-tareas-programadas)
5. [Paso 4: Pruebas](#paso-4-pruebas)
6. [Solución de problemas](#solución-de-problemas)

---

## Requisitos previos

### Software a instalar

```bash
# 1. Python 3.9+ (si no está instalado)
# Descarga desde: https://www.python.org/downloads/
# ⚠️ MARCA: "Add Python to PATH" durante la instalación

# 2. Librerías Python necesarias
pip install pandas gspread google-auth-oauthlib google-auth-httplib2
```

### Archivos necesarios

- `vpn_csv_sync.py` - Script principal
- `crear_tareas_windows.bat` - Automatización de tareas

---

## Paso 1: Descargar credenciales de Google

### 1.1 Crear proyecto en Google Cloud Console

1. Abre [Google Cloud Console](https://console.cloud.google.com/)
2. Crea un nuevo proyecto:
   - Click en el selector de proyecto (arriba a la izquierda)
   - Click en "NEW PROJECT"
   - Nombre: `VPN_CSV_Sync` (o el que prefieras)
   - Click en "CREATE"

### 1.2 Habilitar API de Google Sheets

1. En el buscador de APIs, busca "Google Sheets API"
2. Click en el resultado
3. Click en "ENABLE"

### 1.3 Crear cuenta de servicio

1. Ve a "Credentials" (lado izquierdo)
2. Click en "Create Credentials" → "Service Account"
3. Completa:
   - Service account name: `vpn-csv-sync-bot`
   - Deja las demás opciones por defecto
   - Click "CREATE AND CONTINUE"
4. Click "CONTINUE" en el siguiente paso (sin llenar "Grant roles")
5. Click "DONE"

### 1.4 Descargar archivo JSON

1. En la lista de "Service Accounts", click en el correo de la cuenta creada
2. Ve a la pestaña "KEYS"
3. Click en "Add Key" → "Create new key"
4. Selecciona "JSON"
5. Click "CREATE" - se descargará automáticamente un archivo JSON

**Guarda este archivo en:** `C:\Credentials\google_credentials.json`

### 1.5 Compartir el Google Sheet

1. Abre el Google Sheet donde quieres recibir los datos
2. Click en "Share" (arriba a la derecha)
3. Copia el email de la service account (está en el JSON descargado, campo "client_email")
4. Pega el email y dale permiso de "Editor"
5. Copia el ID del Sheet desde la URL: 
   ```
   https://docs.google.com/spreadsheets/d/[ESTE_ES_EL_ID]/edit
   ```

---

## Paso 2: Configurar el script Python

### 2.1 Crear carpetas necesarias

```cmd
mkdir C:\Data\CSV_Sync
mkdir C:\Data\Logs
mkdir C:\Data\Scripts
mkdir C:\Credentials
```

### 2.2 Copiar y editar el script

1. Coloca `vpn_csv_sync.py` en `C:\Data\Scripts\`
2. Abre el archivo con un editor (Notepad++ o VS Code)
3. Busca la sección `CONFIG` (línea ~45) y actualiza:

```python
CONFIG = {
    # VPN
    'vpn_name': 'Palo Alto',  # Nombre exacto de tu conexión VPN
    'vpn_user': 'tu_usuario_vpn',  # Tu usuario VPN
    'vpn_password': 'tu_contraseña_vpn',  # Tu contraseña VPN
    
    # Rutas
    'csv_network_path': r'\\192.168.1.100\shared\folder',  # Ruta UNC del servidor
    'csv_local_folder': r'C:\Data\CSV_Sync',
    'logs_folder': r'C:\Data\Logs',
    
    # Google
    'google_credentials_file': r'C:\Credentials\google_credentials.json',
    'spreadsheet_id': 'PASTE_YOUR_SHEET_ID_HERE',
    
    # Nombres de archivos CSV y hojas destino
    'csv_files': {
        'archivo1.csv': 'Hoja1',
        'archivo2.csv': 'Hoja2',
        'archivo3.csv': 'Hoja3',
    },
}
```

### 2.3 Encontrar el nombre exacto de la conexión VPN

En Windows:
1. Abre "Settings" → "Network & Internet" → "VPN"
2. Busca la conexión de Palo Alto - el nombre exacto es lo que necesitas

O en terminal:
```cmd
rasdial
```
Verás todas las conexiones disponibles. Copia el nombre exacto de Palo Alto.

### 2.4 Encontrar la ruta del servidor de archivos

Pídele al cliente:
- El **nombre del servidor** o **IP**
- La **ruta compartida** donde están los CSVs

Formato UNC: `\\SERVIDOR\carpeta\subcarpeta`

Ejemplo: `\\192.168.1.50\data_exports`

**Prueba que funciona:**
1. Conecta manualmente a la VPN
2. Abre "File Explorer"
3. En la barra de direcciones, pega la ruta: `\\192.168.1.50\data_exports`
4. Si accedes, esa es la ruta correcta

---

## Paso 3: Crear tareas programadas

### Opción A: Automático (Recomendado)

1. Haz clic derecho en `crear_tareas_windows.bat`
2. Selecciona "Run as administrator"
3. Escribe `SI` cuando se pida confirmación
4. Las 3 tareas se crearán automáticamente

### Opción B: Manual (si Opción A falla)

1. Abre "Task Scheduler" (Programador de Tareas)
   - Press `Win + R`, escribe `taskschd.msc`, Enter
2. En el panel derecho, click en "Create Basic Task"
3. Nombre: `VPN_CSV_Sync_08AM`
4. Disparador: Diario, 08:00
5. Acción: Programa
   - Programa: `C:\Users\TuUsuario\AppData\Local\Programs\Python\Python311\python.exe`
   - Argumentos: `C:\Data\Scripts\vpn_csv_sync.py`
6. Click "OK"
7. **Repite** para 12:00 y 16:00 con nombres diferentes

### Verificar que funcione

1. Abre Task Scheduler
2. Busca "VPN_CSV_Sync"
3. Click derecho en una tarea
4. "Run" - debería ejecutarse inmediatamente
5. Revisa los logs en `C:\Data\Logs\vpn_sync_YYYYMMDD.log`

---

## Paso 4: Pruebas

### 4.1 Prueba manual (antes de automatizar)

```cmd
# Abre terminal como ADMINISTRADOR
cd C:\Data\Scripts
python vpn_csv_sync.py
```

Debería:
1. Conectar a VPN (espera ~8 segundos)
2. Descargar CSVs (verás archivos en `C:\Data\CSV_Sync`)
3. Desconectar VPN (tu internet vuelve)
4. Subir a Google Sheets
5. Mostrar un reporte final

**Esperado en consola:**
```
2024-04-15 10:30:45 - INFO - ✓ VPN conectada exitosamente
2024-04-15 10:30:53 - INFO - ✓ Descargado: archivo1.csv (125.5 KB)
2024-04-15 10:31:02 - INFO - ✓ VPN desconectada exitosamente
2024-04-15 10:31:15 - INFO - ✓ Completado: Hoja1 (1250 registros)
2024-04-15 10:31:22 - INFO - ✓ SINCRONIZACIÓN COMPLETADA EXITOSAMENTE
```

### 4.2 Revisar logs

```cmd
notepad C:\Data\Logs\vpn_sync_20240415.log
```

Búscalos por:
- `✓` = Éxito
- `✗` = Error
- `⚠` = Advertencia

### 4.3 Verificar Google Sheets

- Abre el Google Sheet compartido
- Verifica que los datos aparecen en las hojas correctas
- Los datos deben estar frescos (misma fecha/hora que la ejecución)

---

## Solución de problemas

### "VPN no se conecta"

**Problema:** `Error: rasdial.exe no reconoce las credenciales`

**Soluciones:**
1. Verifica que el nombre VPN sea exacto (sensible a mayúsculas)
2. Prueba conectar manualmente para validar credenciales
3. Si la credencial es correcta pero no conecta, la VPN podría requerir MFA
   - En ese caso, conecta **manualmente** antes de ejecutar el script

```python
# Comentariza estas líneas en vpn_csv_sync.py si usas MFA
# conectar_vpn()
# desconectar_vpn()
```

### "No se encuentran los archivos CSV"

**Problema:** `Archivo no encontrado: \\SERVIDOR\shared\archivo1.csv`

**Soluciones:**
1. Verifica la ruta UNC manualmente (conectado a VPN):
   ```cmd
   net use Z: \\SERVIDOR\shared
   dir Z:
   ```
2. Revisa los nombres exactos de los archivos (mayúsculas, espacios)
3. Asegúrate de que los nombres en `csv_files` coincidan exactamente

### "Error de autenticación Google Sheets"

**Problema:** `API Error: Invalid authentication`

**Soluciones:**
1. Verifica que el JSON está en `C:\Credentials\google_credentials.json`
2. Comparte el Google Sheet con el email de la service account
3. Descarga un nuevo archivo JSON desde Google Cloud (la clave anterior puede haber expirado)

### "Los datos no se cargan en Google Sheets"

**Problema:** El script corre pero Google Sheets no se actualiza

**Soluciones:**
1. Verifica el `spreadsheet_id` es correcto (copia de la URL)
2. Revisa que el Sheet tenga las hojas exactas: "Hoja1", "Hoja2", etc.
3. Chequea el log para ver si hay error en Unicode

```
# En vpn_csv_sync.py (línea ~308), cambia encoding:
df = pd.read_csv(ruta_completa, encoding='latin-1')  # En lugar de utf-8
```

### "Windows ejecuta la tarea pero no hace nada"

**Problema:** Task Scheduler ejecuta la tarea pero no hay logs

**Soluciones:**
1. Abre la tarea en Task Scheduler, click "Run" manualmente
2. Verifica `C:\Data\Logs\` - debería haber archivos .log
3. Revisa en Task Scheduler → Historial de la tarea

Si no hay logs, quiere decir que Python no se está ejecutando:
- Abre una terminal como admin y corre:
  ```cmd
  python C:\Data\Scripts\vpn_csv_sync.py
  ```
- Si falla, copia el error completo en Google para debuggear

---

## 📅 Horarios personalizados

Para cambiar los horarios (actualmente 08:00, 12:00, 16:00):

### Opción 1: Editar en Task Scheduler

1. Abre Task Scheduler
2. Click derecho en la tarea
3. Click en "Properties"
4. Pestaña "Triggers"
5. Edita la hora

### Opción 2: Desde terminal

```cmd
# Cambiar horario de VPN_CSV_Sync_08AM a las 09:00
schtasks /change /tn "VPN_CSV_Sync_08AM" /st 09:00

# Listar todas las tareas VPN
schtasks /query /tn "VPN_CSV_Sync*"
```

---

## 🔒 Seguridad - Guardar contraseña VPN de forma segura

⚠️ **NO guardes la contraseña en texto plano en el script**

Opciones más seguras:

### Opción 1: Variables de entorno Windows

```cmd
# Crear variables de entorno
setx VPN_USER "tu_usuario"
setx VPN_PASSWORD "tu_contraseña"
```

Luego en el script:
```python
CONFIG = {
    'vpn_user': os.getenv('VPN_USER'),
    'vpn_password': os.getenv('VPN_PASSWORD'),
    # ...
}
```

### Opción 2: Archivo de configuración encriptado

Crea `config.json`:
```json
{
    "vpn_user": "usuario",
    "vpn_password": "contraseña"
}
```

Y en el script:
```python
import json
with open('C:\\Data\\config.json', 'r') as f:
    config_sensible = json.load(f)
    CONFIG['vpn_user'] = config_sensible['vpn_user']
```

---

## 📞 Contacto y soporte

Si tienes problemas:
1. Revisa primero los **logs** en `C:\Data\Logs\`
2. Busca el error exacto en Google
3. Documenta:
   - El comando exacto que ejecutaste
   - El mensaje de error completo
   - El archivo log relevante

---

**¡Listo! Tu sistema está configurado. A partir del primer día, la sincronización se ejecutará automáticamente.**

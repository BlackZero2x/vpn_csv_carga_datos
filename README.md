# VPN CSV → Google Sheets — Sincronización Automática

Automatización que descarga un CSV desde una red corporativa (accesible únicamente desde una VM con VPN) y lo sube a Google Sheets, ejecutándose de forma programada de lunes a sábado.

---

## Arquitectura general

El problema central es que la red corporativa y la salida a internet no coexisten en el mismo equipo:

```
┌─────────────────────────────────────────────────────┐
│  PC ANFITRIÓN                                       │
│                                                     │
│   ┌──────────────────────┐                          │
│   │  VM (VirtualBox)     │  Red corporativa (VPN)  │
│   │                      │ ─────────────────────►  │ \\172.16.10.240\...
│   │  vm_descargar_csvs.py│                          │
│   │                      │  Sin acceso a internet   │
│   └──────────┬───────────┘                          │
│              │ Carpeta compartida VirtualBox         │
│              │ \\VBOXSVR\vpncompartido               │
│              ▼                                       │
│   C:\Users\developer2\Documents\vpncompartido\       │
│              │                                       │
│   ┌──────────┴───────────┐                          │
│   │  vpn_csv_sync.py     │  Internet disponible     │
│   │  (script anfitrión)  │ ─────────────────────►  │ Google Sheets API
│   └──────────────────────┘                          │
└─────────────────────────────────────────────────────┘
```

**Flujo de datos:**

1. La VM (con VPN activa) copia el CSV desde `\\172.16.10.240\8.compartidos\Agencias\Auren` hacia `\\VBOXSVR\vpncompartido`
2. El anfitrión lee ese CSV desde `C:\Users\developer2\Documents\vpncompartido`
3. El anfitrión sube los datos a Google Sheets (columnas A:Y) y escribe fórmulas en Z:AE

---

## Archivos del proyecto

| Archivo | Dónde ejecutar | Descripción |
|---|---|---|
| `vm_descargar_csvs.py` | Dentro de la VM | Copia el CSV desde la red corporativa a la carpeta compartida |
| `vpn_csv_sync.py` | PC anfitrión | Lee el CSV de la carpeta compartida y lo sube a Google Sheets |
| `crear_tareas_vm.bat` | Dentro de la VM (como admin) | Crea las 7 tareas programadas en la VM |
| `crear_tareas_windows.bat` | PC anfitrión (como admin) | Crea las 7 tareas programadas en el anfitrión |
| `csv_mapping.json` | PC anfitrión | Define qué CSV se sube a qué hoja del Spreadsheet |
| `diagnostico_vpn_sync.py` | PC anfitrión | Verifica el entorno: rutas, credenciales, conectividad |
| `.env` | PC anfitrión | Variables de entorno con rutas y credenciales (**no se sube a git**) |
| `.env.example` | — | Plantilla del `.env` sin valores reales |

---

## Horarios de ejecución

Las tareas se ejecutan de **lunes a sábado**. Los domingos ambos scripts salen sin hacer nada.

| Hora | Quién | Qué hace |
|---|---|---|
| 08:25 | VM | Descarga el CSV desde la red corporativa |
| 08:30 | Anfitrión | Sube el CSV a Google Sheets |
| 10:25 | VM | Descarga el CSV |
| 10:30 | Anfitrión | Sube a Sheets |
| 12:25 | VM | Descarga el CSV |
| 12:30 | Anfitrión | Sube a Sheets |
| 14:25 | VM | Descarga el CSV |
| 14:30 | Anfitrión | Sube a Sheets |
| 16:25 | VM | Descarga el CSV |
| 16:30 | Anfitrión | Sube a Sheets |
| 18:25 | VM | Descarga el CSV |
| 18:30 | Anfitrión | Sube a Sheets |
| 20:25 | VM | Descarga el CSV |
| 20:30 | Anfitrión | Sube a Sheets |

---

## Qué se sube a Google Sheets

El archivo `BD_Ventas_Auren_Por_Hora.csv` se carga en la hoja `Altas_MIFIBRA_Abr26` con el siguiente comportamiento:

- **Columnas A:Y** — se borran y reescriben con los 25 campos del CSV en cada ejecución
- **Columnas Z:AE** — se escriben fórmulas automáticamente en cada fila de datos
- **Formato** — se copia el formato de la fila 2 (plantilla) hacia todas las filas de datos

### Las 25 columnas del CSV (A → Y)

`FILIAL`, `numContrato`, `servicio`, `numDocIdentidad`, `codAbonado`, `estadoFichaContrato`, `paqueteInicialInternet`, `fechaInscripcionFicha`, `fechaInstInternet`, `INSTALADA`, `usuarioIngreso`, `canal_atencion`, `vendedor`, `motivoDesaprobacion`, `motivoDesaprobado`, `motivoObservado`, `MOTIVO DE OBSERVACION`, `MOTIVO DE ANULACION`, `ult_actualizacion`, `AÑO_INGRESO`, `MES_INGRESO`, `DIA_INGRESO`, `HORA_INGRESO`, `PORTA`, `CATEGORIA`

### Fórmulas automáticas (columnas Z → AE)

| Col | Fórmula |
|---|---|
| Z | `1` (contador) |
| AA | `=BUSCARV(AE{fila},'cf mi fibra'!$B$5:$C$12,2,0)` |
| AB | `=SI.ERROR(BUSCARV(D{fila},'FORM MIFIBRA'!A:AH,34,0),"")` |
| AC | `=BUSCARV(AB{fila},matriz!A:I,9,0)` |
| AD | `=EXTRAE(G{fila},ENCONTRAR(" ",G{fila},ENCONTRAR(" ",G{fila})+1)+1,LARGO(G{fila}))` |
| AE | `=SI(C{fila}="internet+servicio ott",CONCATENAR(C{fila}," ",AD{fila}),G{fila})` |

---

## Requisitos

### PC anfitrión

```bash
pip install pandas gspread google-auth-oauthlib google-auth-httplib2 google-api-python-client
```

### VM (solo stdlib de Python — sin dependencias externas)

```bash
# No requiere pip install — usa únicamente módulos estándar:
# os, sys, json, logging, datetime, pathlib, shutil
```

---

## Instalación

### 1. Credenciales de Google Cloud

1. Entra a [Google Cloud Console](https://console.cloud.google.com/) y crea un proyecto
2. Habilita la **Google Sheets API**
3. Crea una **Cuenta de servicio** y descarga su clave JSON
4. Guarda el JSON en `C:\Credentials\` (fuera del repositorio)
5. Abre el Google Spreadsheet, haz clic en **Compartir** y agrega el `client_email` del JSON con rol **Editor**

### 2. Configurar el anfitrión

```bash
# Clonar el repositorio
git clone https://github.com/BlackZero2x/vpn_csv_carga_datos.git
cd vpn_csv_carga_datos

# Crear el archivo de configuración
copy .env.example .env
```

Editar `.env` con los valores reales:

```env
CSV_LOCAL_FOLDER=C:\Users\TU_USUARIO\Documents\vpncompartido
LOGS_FOLDER=C:\VPN_MIFIBRA\Logs
GOOGLE_CREDENTIALS_FILE=C:\Credentials\nombre-del-archivo.json
SPREADSHEET_ID=ID_del_spreadsheet_de_destino
```

### 3. Configurar la carpeta compartida en VirtualBox

En VirtualBox → Configuración de la VM → Carpetas compartidas:

| Campo | Valor |
|---|---|
| Ruta en el anfitrión | `C:\Users\developer2\Documents\vpncompartido` |
| Nombre de la carpeta | `vpncompartido` |
| Montaje automático | Activado |
| Solo lectura | Desactivado |

Desde dentro de la VM, la carpeta queda accesible como `\\VBOXSVR\vpncompartido`.

### 4. Crear tareas programadas

**En el anfitrión** (ejecutar como Administrador):
```cmd
crear_tareas_windows.bat
```

**Dentro de la VM** (ejecutar como Administrador):
```cmd
crear_tareas_vm.bat
```

Cada bat crea 7 tareas en el Programador de tareas de Windows usando la cuenta `SYSTEM`.

---

## Prueba manual

**Anfitrión** — verificar que el CSV esté en la carpeta compartida y subirlo:
```cmd
python vpn_csv_sync.py
```

**VM** — copiar el CSV desde la red corporativa:
```cmd
python vm_descargar_csvs.py
```

**Diagnóstico del entorno** (solo anfitrión):
```cmd
python diagnostico_vpn_sync.py
```

Salida esperada del anfitrión:
```
2026-04-24 08:30:01 - INFO - INICIANDO SINCRONIZACIÓN
2026-04-24 08:30:02 - INFO - ✓ Conectado: Avance MIFIBRA
2026-04-24 08:30:04 - INFO - Datos leídos: 4823 filas × 25 columnas
2026-04-24 08:30:06 - INFO - ✓ Fórmulas Z:AE escritas en 4823 filas
2026-04-24 08:30:08 - INFO - ✓ Formato copiado de fila 2 a filas 3:4824
2026-04-24 08:30:08 - INFO - ✓ SINCRONIZACIÓN COMPLETADA EXITOSAMENTE
```

---

## Solución de problemas

### El anfitrión no llega a Google (error SSL / proxy)

El `.env` ya incluye la solución. Verificar que estas líneas estén presentes:
```env
HTTPS_PROXY=
HTTP_PROXY=
NO_PROXY=googleapis.com,google.com,accounts.google.com
```

El script limpia esas variables del entorno antes de llamar a la API para evitar que el proxy corporativo intercepte las conexiones HTTPS a Google.

### La VM no encuentra la carpeta compartida (`\\VBOXSVR\vpncompartido`)

- Verificar que las **Guest Additions** de VirtualBox estén instaladas en la VM
- Verificar que la carpeta compartida esté configurada con **montaje automático**
- Dentro de la VM probar: `dir \\VBOXSVR\vpncompartido`

### La VM no encuentra la red corporativa (`\\172.16.10.240\...`)

- La VPN debe estar activa dentro de la VM antes de ejecutar el script
- Probar desde la VM: `ping 172.16.10.240`
- Si no responde, verificar que el adaptador de red de la VM esté en modo **Bridged** o **NAT con reenvío de puertos VPN**

### Las tareas programadas no se crean (error de SID)

Los `.bat` usan `/ru "SYSTEM"` para evitar problemas de mapeo de SID entre cuentas de usuario. Si aun así falla, abrir el `.bat` como Administrador (clic derecho → Ejecutar como administrador).

### Python no se encuentra al crear las tareas

El `PYTHON_PATH` en `crear_tareas_windows.bat` está fijo como:
```
C:\Users\developer2\AppData\Local\Programs\Python\Python312\python.exe
```
Si Python está instalado en otra ruta, editar esa línea antes de ejecutar el bat. Para encontrar la ruta correcta:
```cmd
where python
```

---

## Seguridad

- El archivo `.env` (con SPREADSHEET_ID y rutas) está en `.gitignore` y **nunca se sube al repositorio**
- El JSON de la cuenta de servicio de Google se guarda en `C:\Credentials\` fuera del repositorio
- `csv_mapping.json` solo contiene nombres de archivo y hoja, sin datos sensibles

---

## Estructura de logs

Los logs se guardan en `C:\VPN_MIFIBRA\Logs\` con el formato `sync_YYYYMMDD.log`.
Los logs de la VM se guardan en `\\VBOXSVR\vpncompartido\Logs\vm_descarga_YYYYMMDD.log`, visibles también desde el anfitrión.

```
✓  Éxito
✗  Error
⚠  Advertencia
```

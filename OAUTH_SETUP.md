# Configuración de OAuth para Google Sheets

El script `vpn_csv_sync.py` ahora soporta **autenticación OAuth**, lo que te permite usar tus propias credenciales de Google sin necesidad de crear una cuenta de servicio en Google Cloud.

## Ventajas de OAuth

✅ No necesitas compartir los spreadsheets con una cuenta de servicio  
✅ Puedes usar tu propia cuenta de Google  
✅ El token se guarda localmente y se reutiliza automáticamente  
✅ Mayor seguridad: sin credenciales JSON expuestas  

## Pasos para Configurar

### 1. Descargar Credenciales OAuth desde Google Cloud

1. Ve a [Google Cloud Console](https://console.cloud.google.com)
2. Selecciona o crea un proyecto
3. Habilita **Google Sheets API**:
   - Panel lateral → APIs y servicios → Biblioteca
   - Busca "Google Sheets API" y haz clic en "Habilitar"
4. Crea credenciales OAuth 2.0:
   - APIs y servicios → Credenciales
   - Crear credenciales → Selecciona "ID de cliente de OAuth 2.0"
   - Tipo de aplicación: **Aplicación de escritorio**
   - Haz clic en "Crear"
5. Descarga el archivo JSON y renómbralo a `credentials.json`

### 2. Coloca `credentials.json` en el Directorio del Proyecto

```
C:\proyectos\VPN_MIFIBRA\credentials.json
```

### 3. Ejecuta el Script de Autenticación

Abre PowerShell y ejecuta:

```powershell
cd "C:\proyectos\VPN_MIFIBRA"
C:\proyectos\.venv\Scripts\python.exe authenticate_oauth.py
```

Esto abrirá una ventana del navegador donde podrás autorizar la aplicación.

### 4. Verifica que el Token se Guardó

Deberías ver:
- La carpeta `.cache` se crea automáticamente
- Archivo `.cache/token.pickle` se genera

### 5. Listo

`vpn_csv_sync.py` ahora usará automáticamente OAuth. No necesitas hacer nada más.

## Fallback a Cuenta de Servicio

Si no tienes un token OAuth válido, el script intentará usar credenciales de cuenta de servicio (si existen en `.env`).

El orden de autenticación es:
1. **OAuth token** (`.cache/token.pickle`) ← Preferido
2. **Cuenta de servicio** (archivo JSON en `.env`) ← Fallback

## Renovación del Token

El token OAuth expira después de cierto tiempo. Para renovarlo:

```powershell
C:\proyectos\.venv\Scripts\python.exe authenticate_oauth.py
```

O simplemente ejecuta `vpn_csv_sync.py` manualmente; si el token está expirado, te pedirá que lo renueves.

## Troubleshooting

### Error: "credentials.json no encontrado"

Asegúrate de que el archivo esté en:
```
C:\proyectos\VPN_MIFIBRA\credentials.json
```

### Error: "Token OAuth expirado"

Ejecuta nuevamente `authenticate_oauth.py` para renovar el token.

### Error: "Sin OAuth token ni credenciales de cuenta de servicio"

- No hay token OAuth en `.cache/token.pickle`, Y
- No hay archivo de cuenta de servicio en `.env`

**Solución**: Coloca `credentials.json` y ejecuta `authenticate_oauth.py`

## Variables de Entorno (.env)

Ya no es **obligatorio** tener `GOOGLE_CREDENTIALS_FILE` en `.env`. Puedes comentarlo:

```bash
# GOOGLE_CREDENTIALS_FILE=C:\proyectos\shared\credentials\...  (opcional)
```

El script usa OAuth primero, y si no está disponible, intenta usar la cuenta de servicio.

---

**Nota**: Los tokens OAuth se guardan localmente en `.cache/token.pickle`. Asegúrate de que esta carpeta no se comparta o publique.

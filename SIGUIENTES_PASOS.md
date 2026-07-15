# Siguientes Pasos - Autenticación OAuth

## Estado Actual

✅ **vpn_csv_sync.py** — Actualizado a OAuth  
✅ **Base Fija** — Funciona con cuenta de servicio  
❌ **FORMSMIFIBRA** — Falla porque cuenta de servicio no tiene permisos  

## Solución: OAuth (Opción Recomendada)

### Paso 1: Obtener `credentials.json` de Google Cloud

1. Ve a: https://console.cloud.google.com
2. Selecciona tu proyecto (o crea uno)
3. En el menú lateral → **APIs y servicios** → **Biblioteca**
4. Busca **"Google Sheets API"** → Haz clic en "Habilitar"
5. Vuelve a: **APIs y servicios** → **Credenciales**
6. Haz clic en **"+ Crear credenciales"**
7. Selecciona: **OAuth 2.0 (ID de cliente)**
   - Tipo de aplicación: **Aplicación de escritorio**
   - Haz clic en **"Crear"**
8. En el listado, busca tu aplicación y haz clic en el icono de descargar (JSON)
9. Guarda el archivo como: `credentials.json`

### Paso 2: Coloca `credentials.json` en el Directorio del Proyecto

```
C:\proyectos\VPN_MIFIBRA\credentials.json
```

Copia el archivo que descargaste en la carpeta anterior.

### Paso 3: Ejecuta el Script de Autenticación

**Opción A: Usar el .bat (recomendado en Windows)**
```powershell
cd "C:\proyectos\VPN_MIFIBRA"
.\autenticar_oauth.bat
```

**Opción B: Ejecutar directamente con Python**
```powershell
cd "C:\proyectos\VPN_MIFIBRA"
C:\proyectos\.venv\Scripts\python.exe authenticate_oauth.py
```

### Paso 4: Autoriza en el Navegador

1. Se abrirá automáticamente una ventana del navegador
2. Verás un diálogo de Google pidiendo autorización
3. Selecciona tu cuenta de Google
4. Haz clic en **"Permitir"** para dar acceso a la aplicación
5. Deberías ver un mensaje: **"La instalación se realizó correctamente"**
6. Cierra el navegador

### Paso 5: Verifica que el Token se Guardó

Debería aparecer:
```
C:\proyectos\VPN_MIFIBRA\.cache\token.pickle
```

Si ves este archivo, **¡la autenticación fue exitosa!**

## Listo ✅

El script `vpn_csv_sync.py` ahora usará OAuth automáticamente:

- ✅ Base Fija se cargará
- ✅ FORMSMIFIBRA se cargará
- ✅ Ambos spreadsheets accesibles sin compartir manualmente
- ✅ Token se auto-renova si expira

## Prueba Rápida

```powershell
cd "C:\proyectos\VPN_MIFIBRA"
C:\proyectos\.venv\Scripts\python.exe vpn_csv_sync.py --forzar
```

Deberías ver:
```
[INFO] ✓ Conectado: Base Fija
[INFO] ✓ Conectado al libro diario: Auren_MiFibra (respuestas)
[INFO] ✓ SINCRONIZACIÓN COMPLETADA EXITOSAMENTE
```

## En Caso de Error

### Error: "credentials.json no encontrado"
- Verifica que `credentials.json` esté en: `C:\proyectos\VPN_MIFIBRA\`
- El nombre debe ser exactamente `credentials.json` (minúsculas)

### Error: "Token OAuth expirado"
- Vuelve a ejecutar: `.\autenticar_oauth.bat`
- El token se renova automáticamente

### El script sigue usando cuenta de servicio
- Si hay `GOOGLE_CREDENTIALS_FILE` en `.env`, el script lo usará de fallback
- Verifica que `credentials.json` sea válido

## Preguntas Frecuentes

**P: ¿Mi contraseña se guarda en el archivo?**
A: No. Solo se guarda un token que permite acceso solo a Google Sheets. No es tu contraseña.

**P: ¿Puedo usar cualquier cuenta de Google?**
A: Sí, cualquier cuenta funciona. No necesita ser la misma de la cuenta de servicio.

**P: ¿Qué pasa si pierdo el token?**
A: Solo ejecuta `autenticar_oauth.bat` de nuevo. Es un proceso de 30 segundos.

**P: ¿Puedo usar OAuth en la máquina VM?**
A: No es necesario. Usa OAuth solo en el anfitrión (PC que carga a Sheets).

## Documentación Adicional

- **Guía detallada:** `OAUTH_SETUP.md`
- **Opciones de migración:** `MIGRACION_A_OAUTH.md`
- **Cambios técnicos:** `CAMBIOS_OAUTH.md`

---

**¡Ahora estás listo para ejecutar el script con OAuth!**

Si tienes dudas o errores, revisa `OAUTH_SETUP.md` para más detalles.

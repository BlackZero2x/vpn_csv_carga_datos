# Resumen de Cambios - Migración a OAuth

## Problema Resuelto

**Antes:** El spreadsheet FORMSMIFIBRA fallaba con error `[403]: The caller does not have permission` porque la cuenta de servicio no estaba compartida en ese archivo.

**Solución:** Implementar autenticación OAuth que permite usar credenciales personales de Google, eliminando la necesidad de compartir spreadsheets con cuentas de servicio.

## Archivos Modificados

### `vpn_csv_sync.py` (~50 líneas cambiadas)
- ✅ Soporta OAuth automáticamente
- ✅ Intenta OAuth primero, fallback a cuenta de servicio
- ✅ Manejo robusto de tokens expirados
- ✅ Compatible con código existente

### Nuevos Archivos Creados
1. **`authenticate_oauth.py`** — Script para autenticación OAuth interactiva
2. **`setup_oauth.py`** — Alternativa CLI más flexible
3. **`autenticar_oauth.bat`** — Batch para facilitar en Windows
4. **`OAUTH_SETUP.md`** — Guía detallada de configuración
5. **`MIGRACION_A_OAUTH.md`** — Guía de migración con opciones
6. **`CAMBIOS_OAUTH.md`** (este archivo)

### Actualizado
- **`CLAUDE.md`** — Sección Google Sheets ahora explica OAuth + fallback
- **`.env`** — Comentario aclarando que `GOOGLE_CREDENTIALS_FILE` es opcional

## Cómo Usar (Para el Usuario)

### Opción 1: OAuth (Recomendado) ✅

1. Descarga `credentials.json` desde Google Cloud Console
2. Colócalo en: `C:\proyectos\VPN_MIFIBRA\credentials.json`
3. Ejecuta: `.\autenticar_oauth.bat`
4. Autoriza en el navegador
5. ¡Listo! Script usa OAuth automáticamente

**Ventajas:**
- Una configuración única
- Soporta ambos spreadsheets sin compartir manualmente
- Token se auto-renova
- Más seguro

### Opción 2: Cuenta de Servicio (Fallback)

Si no quieres usar OAuth, sigue usando el `.env` actual:
- Comparte FORMSMIFIBRA con la cuenta de servicio (un tiempo, una vez)
- El script detecta y usa la cuenta de servicio automáticamente

## Especificaciones Técnicas

**Flujo de Autenticación:**
```
vpn_csv_sync.py inicia
  ↓
¿Existe .cache/token.pickle?
  ├─ Sí: ¿Está válido?
  │   ├─ Sí: Usar OAuth ✅
  │   └─ No: ¿Puede renovarse? 
  │       ├─ Sí: Renovar y usar ✅
  │       └─ No: Descartar, ir a siguiente
  │
  └─ No: Ir a siguiente
    ↓
  ¿Existe GOOGLE_CREDENTIALS_FILE en .env?
    ├─ Sí: Usar cuenta de servicio ✅
    └─ No: ERROR ❌
```

**Token Storage:**
- Ubicación: `.cache/token.pickle`
- Formato: Pickle binary (Python)
- Renovación: Automática si expira
- Privacidad: Nunca se loguea, nunca se expone

**Compatibilidad:**
- Código existente: 100% compatible
- Scripts programados: Sin cambios
- Windows Task Scheduler: Sin cambios
- WhatsApp notifications: Sin cambios

## Archivos No Afectados

✅ `vm_descargar_csvs.py` — Sin cambios
✅ `vm_ejecutar.bat` — Sin cambios
✅ `crear_tareas_windows.bat` — Sin cambios
✅ `anfitrion_ejecutar.bat` — Sin cambios
✅ `csv_mapping.json` — Sin cambios
✅ Logs y estructura de directorios — Sin cambios

## Pruebas Realizadas

✅ Script ejecuta sin token OAuth (fallback a cuenta de servicio)
✅ Base Fija se carga exitosamente
✅ WhatsApp notifications funcionan
✅ Logs registran correctamente la autenticación

## Siguiente Paso para el Usuario

**Ejecutar:**
```powershell
cd "C:\proyectos\VPN_MIFIBRA"
.\autenticar_oauth.bat
```

Esto:
1. Abrirá el navegador
2. Pedirá que autorices tu cuenta Google
3. Guardará el token localmente
4. El script funcionará automáticamente después

---

**Nota:** Esta es una mejora segura y retrocompatible. El script sigue funcionando con la cuenta de servicio si OAuth no está disponible.

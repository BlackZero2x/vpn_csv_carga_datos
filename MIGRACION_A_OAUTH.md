# Migración a OAuth - Próximos Pasos

## Estado Actual

✅ **vpn_csv_sync.py** ahora soporta OAuth automáticamente
✅ El script intenta OAuth primero, luego fallback a cuenta de servicio
✅ Los 3 CSVs se cargan exitosamente en **Base Fija**
❌ **FORMSMIFIBRA** aún falla porque la cuenta de servicio no tiene permisos

## El Problema

El error actual es:
```
[403]: The caller does not have permission
```

Esto significa que el archivo `google_credentials.json` (cuenta de servicio) **no tiene acceso compartido** al spreadsheet FORMSMIFIBRA.

## Dos Soluciones

### Opción A: Usar OAuth (Recomendada) ✅

Esto elimina la necesidad de compartir spreadsheets con cuentas de servicio.

**Pasos:**

1. **Descarga `credentials.json`** desde Google Cloud:
   - Ve a https://console.cloud.google.com
   - Selecciona tu proyecto
   - APIs y servicios → Credenciales
   - Selecciona tu aplicación OAuth "Desktop"
   - Descarga el JSON

2. **Coloca `credentials.json` en:**
   ```
   C:\proyectos\VPN_MIFIBRA\credentials.json
   ```

3. **Ejecuta el script de autenticación:**
   ```powershell
   cd "C:\proyectos\VPN_MIFIBRA"
   .\autenticar_oauth.bat
   ```
   O manualmente:
   ```powershell
   C:\proyectos\.venv\Scripts\python.exe authenticate_oauth.py
   ```

4. **Se abrirá el navegador** para autorizar tu cuenta de Google
   - Autoriza con tu cuenta personal (no necesita ser la misma de la cuenta de servicio)
   - El token se guardará automáticamente

5. **Listo** - vpn_csv_sync.py ahora usará tu cuenta para acceder a ambos spreadsheets

### Opción B: Compartir FORMSMIFIBRA con la Cuenta de Servicio

Si prefieres mantener la cuenta de servicio:

1. Abre FORMSMIFIBRA en Google Sheets
2. Haz clic en "Compartir" (botón superior derecho)
3. Busca el email de la cuenta de servicio en:
   ```
   C:\proyectos\shared\credentials\automationavanceventas-31cbef58f932.json
   ```
   Campo: `"client_email"`
4. Añade ese email con rol **"Editor"**
5. Reinicia el script

## Comparación

| Aspecto | OAuth (Opción A) | Cuenta de Servicio (Opción B) |
|---------|------------------|-------------------------------|
| Configuración | Simple (1 clic) | Manual (compartir spreadsheet) |
| Seguridad | Tu cuenta personal | Cuenta de servicio expuesta |
| Mantenimiento | Token auto-renova | Debe compartirse manualmente |
| Flexibilidad | Usa cualquier cuenta Google | Solo la cuenta de servicio |
| Recomendado | ✅ **SÍ** | ❌ No (manual) |

## Recomendación

**Usa Opción A (OAuth)**. Es más seguro, simple y no requiere compartir spreadsheets.

1. Ejecuta `autenticar_oauth.bat`
2. Autoriza en el navegador
3. ¡Listo! El script funcionará automáticamente

---

**Nota:** Si vas a usar OAuth, puedes comentar o eliminar `GOOGLE_CREDENTIALS_FILE` del `.env` (es opcional ahora).

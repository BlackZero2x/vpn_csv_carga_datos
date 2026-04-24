"""
Script de Diagnóstico - VPN CSV Sync
Verifica configuración, rutas y conectividad antes de ejecutar
"""

import os
import sys
import subprocess
import json
from pathlib import Path
from datetime import datetime

# Colores para consola
class Colors:
    OK = '\033[92m'      # Verde
    ERROR = '\033[91m'   # Rojo
    WARNING = '\033[93m' # Amarillo
    INFO = '\033[94m'    # Azul
    RESET = '\033[0m'

def print_ok(msg):
    print(f"{Colors.OK}✓ {msg}{Colors.RESET}")

def print_error(msg):
    print(f"{Colors.ERROR}✗ {msg}{Colors.RESET}")

def print_warning(msg):
    print(f"{Colors.WARNING}⚠ {msg}{Colors.RESET}")

def print_info(msg):
    print(f"{Colors.INFO}ℹ {msg}{Colors.RESET}")

# ============================================================================
# DIAGNOSTICOS
# ============================================================================

def diagnosticar_python():
    """Verifica instalación de Python y librerías"""
    print("\n" + "="*60)
    print("1. DIAGNÓSTICO PYTHON")
    print("="*60)
    
    # Python version
    print_info(f"Python {sys.version}")
    print_info(f"Ejecutable: {sys.executable}")
    
    if sys.version_info < (3, 9):
        print_error("Python debe ser 3.9 o superior")
        return False
    else:
        print_ok("Versión Python compatible")
    
    # Librerías requeridas
    librerias = {
        'pandas': 'pandas',
        'gspread': 'gspread',
        'google.oauth2': 'google-auth',
        'google.auth.transport': 'google-auth-httplib2',
    }
    
    print_info("\nVerificando librerías...")
    faltantes = []
    
    for modulo, nombre_pip in librerias.items():
        try:
            __import__(modulo)
            print_ok(f"{nombre_pip}")
        except ImportError:
            print_error(f"{nombre_pip} (FALTA INSTALAR)")
            faltantes.append(nombre_pip)
    
    if faltantes:
        print_warning(f"\nFalta instalar: pip install {' '.join(faltantes)}")
        return False
    else:
        print_ok("Todas las librerías están instaladas")
        return True

def diagnosticar_rutas():
    """Verifica existencia de carpetas y archivos necesarios"""
    print("\n" + "="*60)
    print("2. DIAGNÓSTICO DE RUTAS Y ARCHIVOS")
    print("="*60)
    
    rutas_a_verificar = {
        'Script principal': r'C:\Data\Scripts\vpn_csv_sync.py',
        'Carpeta CSVs': r'C:\Data\CSV_Sync',
        'Carpeta Logs': r'C:\Data\Logs',
        'Credenciales Google': r'C:\Credentials\google_credentials.json',
    }
    
    todas_ok = True
    
    for nombre, ruta in rutas_a_verificar.items():
        existe = os.path.exists(ruta)
        
        if existe:
            if os.path.isfile(ruta):
                tamaño = os.path.getsize(ruta) / 1024  # KB
                print_ok(f"{nombre}: {ruta} ({tamaño:.1f} KB)")
            else:
                print_ok(f"{nombre}: {ruta}")
        else:
            print_error(f"{nombre} NO EXISTE: {ruta}")
            todas_ok = False
    
    if not todas_ok:
        print_warning("\nCrea las carpetas faltantes con:")
        print("  mkdir C:\\Data\\Scripts")
        print("  mkdir C:\\Data\\CSV_Sync")
        print("  mkdir C:\\Data\\Logs")
        print("  mkdir C:\\Credentials")
    
    return todas_ok

def diagnosticar_google_sheets():
    """Verifica configuración de Google Sheets"""
    print("\n" + "="*60)
    print("3. DIAGNÓSTICO GOOGLE SHEETS")
    print("="*60)
    
    credenciales_path = r'C:\Credentials\google_credentials.json'
    
    if not os.path.exists(credenciales_path):
        print_error(f"Archivo de credenciales no encontrado: {credenciales_path}")
        print_warning("Descarga google_credentials.json desde Google Cloud Console")
        return False
    
    try:
        with open(credenciales_path, 'r', encoding='utf-8') as f:
            creds = json.load(f)
        
        print_ok(f"Credenciales válidas (JSON bien formado)")
        
        # Verifica campos esenciales
        campos_requeridos = ['type', 'project_id', 'private_key_id', 'client_email']
        
        for campo in campos_requeridos:
            if campo in creds:
                valor = creds[campo]
                # Oculta información sensible
                if len(valor) > 20:
                    valor = valor[:20] + "..."
                print_ok(f"  {campo}: {valor}")
            else:
                print_error(f"  {campo}: FALTA")
                return False
        
        return True
        
    except json.JSONDecodeError:
        print_error("Archivo de credenciales JSON inválido")
        return False
    except Exception as e:
        print_error(f"Error verificando credenciales: {e}")
        return False

def diagnosticar_vpn():
    """Verifica conexiones VPN disponibles"""
    print("\n" + "="*60)
    print("4. DIAGNÓSTICO VPN")
    print("="*60)
    
    try:
        resultado = subprocess.run(['rasdial'], capture_output=True, text=True)
        
        lineas = resultado.stdout.split('\n')
        conexiones = [l.strip() for l in lineas if l.strip() and not l.startswith('entrada')]
        
        if not conexiones:
            print_warning("No hay conexiones VPN configuradas")
            print_info("Crea una conexión VPN en Settings > Network > VPN")
            return False
        
        print_info("Conexiones VPN disponibles:")
        for conexion in conexiones:
            if 'palo alto' in conexion.lower() or 'globalprotect' in conexion.lower():
                print_ok(f"  {conexion} ← Probablemente esta es")
            else:
                print_info(f"  {conexion}")
        
        return True
        
    except Exception as e:
        print_error(f"Error leyendo conexiones VPN: {e}")
        return False

def diagnosticar_red():
    """Verifica conectividad de red"""
    print("\n" + "="*60)
    print("5. DIAGNÓSTICO DE CONECTIVIDAD")
    print("="*60)
    
    # Verifica internet
    print_info("Verificando conexión a Internet...")
    try:
        resultado = subprocess.run(
            ['ping', '-n', '1', '-w', '5000', '8.8.8.8'],
            capture_output=True,
            timeout=7
        )
        
        if resultado.returncode == 0:
            print_ok("Internet disponible")
        else:
            print_warning("No hay conexión a Internet (pero puede ser normal si esperas conexión VPN)")
    except Exception as e:
        print_warning(f"No se puede verificar Internet: {e}")
    
    # Verifica DNS
    print_info("Verificando resolución DNS...")
    try:
        resultado = subprocess.run(
            ['ping', '-n', '1', 'google.com'],
            capture_output=True,
            timeout=7
        )
        
        if resultado.returncode == 0:
            print_ok("DNS funcionando")
        else:
            print_warning("DNS no responde (pero puede ser normal si usas VPN)")
    except Exception as e:
        print_warning(f"No se puede verificar DNS: {e}")
    
    return True

def diagnosticar_task_scheduler():
    """Verifica si las tareas ya existen"""
    print("\n" + "="*60)
    print("6. DIAGNÓSTICO TASK SCHEDULER")
    print("="*60)
    
    try:
        resultado = subprocess.run(
            ['schtasks', '/query', '/tn', 'VPN_CSV_Sync*', '/fo', 'list'],
            capture_output=True,
            text=True
        )
        
        if 'VPN_CSV_Sync' in resultado.stdout:
            print_ok("Tareas VPN_CSV_Sync encontradas")
            
            for linea in resultado.stdout.split('\n'):
                if 'TaskName' in linea:
                    print_info(f"  {linea.strip()}")
        else:
            print_warning("No hay tareas VPN_CSV_Sync creadas aún")
            print_info("Ejecuta crear_tareas_windows.bat para crearlas")
        
        return True
        
    except Exception as e:
        print_error(f"Error accediendo a Task Scheduler: {e}")
        return False

def diagnosticar_configuracion():
    """Verifica el archivo de configuración del script"""
    print("\n" + "="*60)
    print("7. DIAGNÓSTICO DE CONFIGURACIÓN")
    print("="*60)
    
    script_path = r'C:\Data\Scripts\vpn_csv_sync.py'
    
    if not os.path.exists(script_path):
        print_error(f"Script no encontrado: {script_path}")
        return False
    
    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            contenido = f.read()
        
        # Verifica si tiene valores de ejemplo sin configurar
        if 'TU_USUARIO_VPN' in contenido:
            print_error("Script no configurado: 'TU_USUARIO_VPN' aún tiene valor de placeholder")
            return False
        
        if 'TU_CONTRASEÑA_VPN' in contenido:
            print_error("Script no configurado: 'TU_CONTRASEÑA_VPN' aún tiene valor de placeholder")
            return False
        
        if 'TU_SPREADSHEET_ID' in contenido:
            print_error("Script no configurado: 'TU_SPREADSHEET_ID' aún tiene valor de placeholder")
            return False
        
        print_ok("Script configurado correctamente")
        return True
        
    except Exception as e:
        print_error(f"Error leyendo script: {e}")
        return False

# ============================================================================
# REPORTE FINAL
# ============================================================================

def generar_reporte(resultados):
    """Genera reporte final"""
    print("\n" + "="*60)
    print("REPORTE FINAL")
    print("="*60)
    
    total = len(resultados)
    exitosos = sum(1 for v in resultados.values() if v)
    
    print(f"\nResultados: {exitosos}/{total} verificaciones OK")
    
    if exitosos == total:
        print_ok("\n¡TODO LISTO! Puedes ejecutar vpn_csv_sync.py")
    elif exitosos >= total - 1:
        print_warning("\nHay pequeños problemas a resolver")
    else:
        print_error("\nResuelve los errores antes de continuar")
    
    print("\n" + "="*60)
    print("PRÓXIMOS PASOS:")
    print("="*60)
    
    if exitosos == total:
        print("1. Conecta manualmente a la VPN")
        print("2. Ejecuta: python C:\\Data\\Scripts\\vpn_csv_sync.py")
        print("3. Verifica que funciona correctamente")
        print("4. Ejecuta: crear_tareas_windows.bat (como ADMINISTRADOR)")
        print("5. ¡Listo! Las tareas se ejecutarán automáticamente")
    else:
        print("Sigue las instrucciones de error anterior y repite el diagnóstico")

# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    print("\n" + "="*60)
    print("DIAGNÓSTICO - VPN CSV SYNC")
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    resultados = {
        'Python': diagnosticar_python(),
        'Rutas': diagnosticar_rutas(),
        'Google Sheets': diagnosticar_google_sheets(),
        'VPN': diagnosticar_vpn(),
        'Red': diagnosticar_red(),
        'Task Scheduler': diagnosticar_task_scheduler(),
        'Configuración': diagnosticar_configuracion(),
    }
    
    generar_reporte(resultados)
    
    print("\n")
    input("Presiona Enter para salir...")

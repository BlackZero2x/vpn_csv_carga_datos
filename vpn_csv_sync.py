"""
Script del ANFITRIÓN: lee CSVs desde carpeta compartida con la VM y los sube a Google Sheets.
Sube únicamente las 25 columnas del CSV al rango A:Y, sin tocar columnas a la derecha.
"""

import os
import sys
import json
import logging
import pandas as pd
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ============================================================================
# CARGA DE CONFIGURACIÓN
# ============================================================================

def _cargar_env(ruta_env: Path) -> None:
    try:
        with open(ruta_env, encoding='utf-8') as f:
            for linea in f:
                linea = linea.strip()
                if not linea or linea.startswith('#') or '=' not in linea:
                    continue
                clave, _, valor = linea.partition('=')
                os.environ.setdefault(clave.strip(), valor.strip())
    except FileNotFoundError:
        print(f"ERROR: No se encontró .env en {ruta_env}")
        sys.exit(1)

def _cargar_csv_mapping(ruta_json: Path) -> tuple[Dict[str, str], Dict[str, str]]:
    try:
        with open(ruta_json, encoding='utf-8') as f:
            datos = json.load(f)
        return datos['csv_files'], datos['csv_files_diario']
    except FileNotFoundError:
        print(f"ERROR: No se encontró csv_mapping.json en {ruta_json}")
        sys.exit(1)
    except (KeyError, json.JSONDecodeError) as e:
        print(f"ERROR: csv_mapping.json mal formado: {e}")
        sys.exit(1)

_BASE_DIR = Path(__file__).parent
_cargar_env(_BASE_DIR / '.env')
_csv_files, _csv_files_diario = _cargar_csv_mapping(_BASE_DIR / 'csv_mapping.json')

# Desactiva proxy para llamadas a Google (evita error SSL con proxies corporativos HTTP)
for _var in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy'):
    os.environ.pop(_var, None)
os.environ['NO_PROXY'] = 'googleapis.com,google.com,accounts.google.com'

CONFIG = {
    'csv_local_folder':        os.environ['CSV_LOCAL_FOLDER'],
    'logs_folder':             os.environ['LOGS_FOLDER'],
    'google_credentials_file': os.environ['GOOGLE_CREDENTIALS_FILE'],
    'spreadsheet_id':          os.environ['SPREADSHEET_ID'],
    'spreadsheet_id_diario':   os.environ.get('SPREADSHEET_ID_DIARIO', ''),
    'csv_files':               _csv_files,
    'csv_files_diario':        _csv_files_diario,
    'max_retries':             int(os.environ.get('MAX_RETRIES', 2)),
    'retry_delay':             int(os.environ.get('RETRY_DELAY', 5)),
}

# 25 columnas del CSV — define el rango exacto que se sobreescribe en Sheets
CSV_COLUMNAS = [
    'FILIAL', 'numContrato', 'servicio', 'numDocIdentidad', 'codAbonado',
    'estadoFichaContrato', 'paqueteInicialInternet', 'fechaInscripcionFicha',
    'fechaInstInternet', 'INSTALADA', 'usuarioIngreso', 'canal_atencion',
    'vendedor', 'motivoDesaprobacion', 'motivoDesaprobado', 'motivoObservado',
    'MOTIVO DE OBSERVACION', 'MOTIVO DE ANULACION', 'ult_actualizacion',
    'AÑO_INGRESO', 'MES_INGRESO', 'DIA_INGRESO', 'HORA_INGRESO', 'PORTA', 'CATEGORIA'
]
RANGO_COLUMNAS = f"A:Y"  # 25 columnas = A hasta Y

# ============================================================================
# LOGGING
# ============================================================================

def setup_logger() -> logging.Logger:
    Path(CONFIG['logs_folder']).mkdir(parents=True, exist_ok=True)
    log_file = Path(CONFIG['logs_folder']) / f"sync_{datetime.now().strftime('%Y%m%d')}.log"

    logger = logging.getLogger('CSVSync')
    logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger

logger = setup_logger()

# ============================================================================
# GOOGLE SHEETS
# ============================================================================

def _sheets_service():
    """Construye el cliente de Sheets API v4 con la cuenta de servicio."""
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(CONFIG['google_credentials_file'], scopes=scopes)
    return build('sheets', 'v4', credentials=creds, cache_discovery=False)

def copiar_formato_fila2(sheet_id: int, ultima_fila: int) -> None:
    """Copia el formato de la fila 2 (todas las columnas) hacia las filas 3..ultima_fila."""
    if ultima_fila < 3:
        return  # Nada que copiar si solo hay una fila de datos

    service = _sheets_service()
    body = {
        'requests': [{
            'copyPaste': {
                'source': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,   # fila 2 (0-indexed)
                    'endRowIndex': 2,
                    'startColumnIndex': 0,
                    'endColumnIndex': 1000,
                },
                'destination': {
                    'sheetId': sheet_id,
                    'startRowIndex': 2,   # fila 3 (0-indexed)
                    'endRowIndex': ultima_fila,
                    'startColumnIndex': 0,
                    'endColumnIndex': 1000,
                },
                'pasteType': 'PASTE_FORMAT',
                'pasteOrientation': 'NORMAL',
            }
        }]
    }
    service.spreadsheets().batchUpdate(
        spreadsheetId=CONFIG['spreadsheet_id'],
        body=body
    ).execute()
    logger.info(f"✓ Formato copiado de fila 2 a filas 3:{ultima_fila}")

def conectar_google_sheets() -> Tuple[bool, gspread.Spreadsheet | None]:
    try:
        if not os.path.exists(CONFIG['google_credentials_file']):
            logger.error(f"✗ Credenciales no encontradas: {CONFIG['google_credentials_file']}")
            return False, None

        logger.info("Conectando a Google Sheets...")
        gc = gspread.service_account(filename=CONFIG['google_credentials_file'])
        sh = gc.open_by_key(CONFIG['spreadsheet_id'])
        logger.info(f"✓ Conectado: {sh.title}")
        return True, sh

    except gspread.exceptions.SpreadsheetNotFound:
        logger.error(f"✗ Spreadsheet no encontrado (ID: {CONFIG['spreadsheet_id']})")
        return False, None
    except gspread.exceptions.APIError as e:
        logger.error(f"✗ Error de API Google: {e}")
        return False, None
    except Exception as e:
        logger.error(f"✗ Error conectando a Google Sheets: {e}")
        return False, None

def cargar_csv_a_sheets(archivo_csv: str, nombre_hoja: str, sh: gspread.Spreadsheet) -> bool:
    """
    Sube el CSV a la hoja indicada, sobreescribiendo solo el rango A:Y.
    Las columnas a la derecha de Y no se tocan.
    """
    ruta = Path(CONFIG['csv_local_folder']) / archivo_csv

    if not ruta.exists():
        logger.error(f"✗ Archivo no encontrado en carpeta compartida: {ruta}")
        return False

    # Lee el CSV con fallback de encoding
    df = None
    for enc in ('utf-8', 'latin-1'):
        try:
            df = pd.read_csv(ruta, encoding=enc, dtype=str)
            break
        except UnicodeDecodeError:
            continue

    if df is None or df.empty:
        logger.error(f"✗ No se pudo leer o está vacío: {archivo_csv}")
        return False

    # Verifica que las columnas esperadas estén presentes
    columnas_faltantes = [c for c in CSV_COLUMNAS if c not in df.columns]
    if columnas_faltantes:
        logger.warning(f"⚠ Columnas no encontradas en el CSV: {columnas_faltantes}")

    # Selecciona solo las 25 columnas definidas (en el orden correcto)
    columnas_presentes = [c for c in CSV_COLUMNAS if c in df.columns]
    df = df[columnas_presentes]

    logger.info(f"Datos leídos: {len(df)} filas × {len(df.columns)} columnas")

    # Obtiene o crea la hoja
    try:
        ws = sh.worksheet(nombre_hoja)
        logger.info(f"Usando hoja existente: {nombre_hoja}")
    except gspread.exceptions.WorksheetNotFound:
        logger.info(f"Creando hoja nueva: {nombre_hoja}")
        ws = sh.add_worksheet(title=nombre_hoja, rows=len(df) + 1, cols=len(CSV_COLUMNAS))

    # Limpia SOLO el rango A:Y (no toca columnas a la derecha)
    num_filas_actuales = len(ws.get_all_values())
    if num_filas_actuales > 0:
        ws.batch_clear([f"A1:Y{num_filas_actuales}"])
        logger.info(f"Limpiado rango A1:Y{num_filas_actuales}")

    # Prepara y sube los datos
    valores = [df.columns.tolist()] + df.fillna('').values.tolist()
    ws.update(f"A1:Y{len(valores)}", valores, raw=False)

    # Extiende las fórmulas de Z:AE hasta la última fila de datos
    # Fila 1 = encabezados, datos empiezan en fila 2
    primera_fila_datos = 2
    ultima_fila = len(df) + 1  # +1 por la fila de encabezados

    logger.info(f"Escribiendo fórmulas Z:AE en filas {primera_fila_datos}:{ultima_fila}...")

    formulas_por_fila = []
    for fila in range(primera_fila_datos, ultima_fila + 1):
        formulas_por_fila.append([
            1,                                                                                          # Z: contador
            f"=BUSCARV(AE{fila},'cf mi fibra'!$B$5:$C$12,2,0)",                                       # AA
            f'=SI.ERROR(BUSCARV(D{fila},\'FORM MIFIBRA\'!A:AH,34,0),"")',                             # AB
            f"=BUSCARV(AB{fila},matriz!A:I,9,0)",                                                      # AC
            f'=EXTRAE(G{fila},ENCONTRAR(" ",G{fila},ENCONTRAR(" ",G{fila})+1)+1,LARGO(G{fila}))',      # AD
            f'=SI(C{fila}="internet+servicio ott",CONCATENAR(C{fila}," ",AD{fila}),G{fila})',          # AE
        ])

    ws.update(f"Z{primera_fila_datos}:AE{ultima_fila}", formulas_por_fila, raw=False)
    logger.info(f"✓ Fórmulas Z:AE escritas en {len(formulas_por_fila)} filas")

    # Copia formato de la fila 2 (plantilla) hacia todas las filas de datos
    copiar_formato_fila2(ws.id, ultima_fila)

    logger.info(f"✓ Completado: {nombre_hoja} ({len(df)} registros, {len(df.columns)} columnas)")
    return True

# ============================================================================
# FLUJO PRINCIPAL
# ============================================================================

def _es_primera_ejecucion_del_dia() -> bool:
    """
    Devuelve True si hoy no hay ninguna entrada de éxito para el libro diario
    en el log del día. Permite que BD_Ventas_AUREN.csv se cargue solo una vez.
    """
    log_file = Path(CONFIG['logs_folder']) / f"sync_{datetime.now().strftime('%Y%m%d')}.log"
    if not log_file.exists():
        return True
    marcador = 'DIARIO COMPLETADO EXITOSAMENTE'
    with open(log_file, encoding='utf-8', errors='ignore') as f:
        return marcador not in f.read()


def sincronizar_diario(sh_diario: gspread.Spreadsheet) -> Dict[str, bool]:
    """
    Carga al libro diario:
    - BD_Ventas_AUREN.csv → 'Reportes_ventas_dia'   (solo 1ª ejecución del día)
    - BD_Ventas_Auren_Por_Hora.csv → 'Reporte_ventas_hora'  (siempre)
    """
    resultados: Dict[str, bool] = {}
    primera = _es_primera_ejecucion_del_dia()

    for archivo_csv, nombre_hoja in CONFIG['csv_files_diario'].items():
        es_csv_dia = archivo_csv == 'BD_Ventas_AUREN.csv'

        if es_csv_dia and not primera:
            logger.info(f"Saltando {archivo_csv} → {nombre_hoja} (ya se cargó hoy)")
            continue

        logger.info(f"\nProcesando [diario]: {archivo_csv} → {nombre_hoja}")
        resultados[archivo_csv] = cargar_csv_a_sheets(archivo_csv, nombre_hoja, sh_diario)

    return resultados


def sincronizar_completo() -> bool:
    inicio = datetime.now()
    logger.info("=" * 60)
    logger.info(f"INICIANDO SINCRONIZACIÓN: {inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Carpeta compartida: {CONFIG['csv_local_folder']}")
    logger.info("=" * 60)

    # Conectar a Google Sheets principal
    sheets_ok, sh = conectar_google_sheets()
    if not sheets_ok:
        logger.error("✗ No se pudo conectar a Google Sheets principal. Abortando.")
        return False

    # --- Libro principal: BD_Ventas_Auren_Por_Hora → Altas_MIFIBRA_* (cada ejecución) ---
    resultados: Dict[str, bool] = {}
    for archivo_csv, nombre_hoja in CONFIG['csv_files'].items():
        logger.info(f"\nProcesando [principal]: {archivo_csv} → {nombre_hoja}")
        resultados[archivo_csv] = cargar_csv_a_sheets(archivo_csv, nombre_hoja, sh)

    # --- Libro diario: ambos CSVs con su propia lógica de frecuencia ---
    resultados_diario: Dict[str, bool] = {}
    if not CONFIG['spreadsheet_id_diario']:
        logger.warning("⚠ SPREADSHEET_ID_DIARIO no configurado en .env — se omite el libro diario")
    else:
        logger.info(f"\nConectando a Google Sheets diario (ID: {CONFIG['spreadsheet_id_diario'][:12]}…)")
        gc = gspread.service_account(filename=CONFIG['google_credentials_file'])
        try:
            sh_diario = gc.open_by_key(CONFIG['spreadsheet_id_diario'])
            logger.info(f"✓ Conectado al libro diario: {sh_diario.title}")
            resultados_diario = sincronizar_diario(sh_diario)
            if resultados_diario and all(resultados_diario.values()):
                logger.info("✓ DIARIO COMPLETADO EXITOSAMENTE")
        except gspread.exceptions.SpreadsheetNotFound:
            logger.error(f"✗ Spreadsheet diario no encontrado (ID: {CONFIG['spreadsheet_id_diario']})")
        except Exception as e:
            logger.error(f"✗ Error en libro diario: {e}")

    # Reporte final
    tiempo_total = datetime.now() - inicio
    todos = {**resultados, **{f"[diario] {k}": v for k, v in resultados_diario.items()}}
    exitosos = sum(1 for v in todos.values() if v)
    total = len(todos)

    logger.info("\n" + "=" * 60)
    logger.info("REPORTE FINAL")
    logger.info("=" * 60)
    for archivo, ok in todos.items():
        logger.info(f"  {'✓' if ok else '✗'} {archivo}")
    logger.info(f"Resultado: {exitosos}/{total} exitosos | Tiempo: {tiempo_total}")

    if exitosos == total:
        logger.info("✓ SINCRONIZACIÓN COMPLETADA EXITOSAMENTE")
    else:
        logger.warning("⚠ SINCRONIZACIÓN INCOMPLETA — REVISAR LOGS")
    logger.info("=" * 60 + "\n")

    return exitosos == total

# ============================================================================
# ENTRADA
# ============================================================================

if __name__ == "__main__":
    # 0=lunes, 6=domingo — se omite el domingo
    if datetime.now().weekday() == 6:
        print("Domingo — no se ejecuta.")
        sys.exit(0)

    campos_vacios = [k for k in ('csv_local_folder', 'spreadsheet_id', 'google_credentials_file') if not CONFIG[k]]
    if campos_vacios:
        logger.error(f"✗ CONFIGURACIÓN INCOMPLETA en .env: {', '.join(campos_vacios)}")
        sys.exit(1)

    if not CONFIG['spreadsheet_id_diario']:
        logger.warning("⚠ SPREADSHEET_ID_DIARIO vacío en .env — el libro diario no se cargará")

    if not CONFIG['csv_files']:
        logger.error("✗ csv_mapping.json no tiene entradas en 'csv_files'")
        sys.exit(1)

    exito = sincronizar_completo()
    sys.exit(0 if exito else 1)

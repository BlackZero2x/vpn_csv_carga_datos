"""
Script del ANFITRIÓN: lee CSVs desde carpeta compartida con la VM y los sube a Google Sheets.
Sube únicamente las 25 columnas del CSV al rango A:Y, sin tocar columnas a la derecha.
"""

import os
import sys
import json
import logging
import pandas as pd
import traceback
import pickle
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple
import gspread
from google.oauth2.credentials import Credentials as OAuth2Credentials
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow

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
    'google_credentials_file': os.environ.get('GOOGLE_CREDENTIALS_FILE', ''),
    'basefija_id':             os.environ['BASEFIJA_ID'],
    'formsmifibra_id':         os.environ.get('FORMSMIFIBRA_ID', ''),
    'csv_files':               _csv_files,
    'csv_files_diario':        _csv_files_diario,
    'max_retries':             int(os.environ.get('MAX_RETRIES', 2)),
    'retry_delay':             int(os.environ.get('RETRY_DELAY', 5)),
}

_CACHE_DIR = Path(__file__).parent / '.cache'
_TOKEN_PICKLE = _CACHE_DIR / 'token.pickle'
_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

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

# CSVs que solo se cargan en la primera ejecución del día (frecuencia diaria)
CSV_SOLO_DIARIO = {'BD_Ventas_AUREN.csv', 'BD_Cobranza_AUREN.csv'}

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

def _get_oauth_credentials():
    """Obtiene credenciales OAuth válidas, o retorna None si no están disponibles.
    Si el token existe pero está expirado, intenta renovarlo; si no puede, lo descarta."""
    creds = None
    if _TOKEN_PICKLE.exists():
        try:
            with open(_TOKEN_PICKLE, 'rb') as token:
                creds = pickle.load(token)

            if creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception:
                    logger.debug("No se pudo renovar el token OAuth")
                    creds = None
            elif not creds.valid:
                logger.debug("Token OAuth inválido o expirado")
                creds = None
        except Exception as e:
            logger.debug(f"Error cargando token OAuth: {e}")
            creds = None

    return creds

def _sheets_service():
    """Construye cliente de Sheets API v4. Intenta OAuth primero, luego cuenta de servicio."""
    creds = _get_oauth_credentials()
    if creds:
        logger.debug("Usando autenticación OAuth")
        return build('sheets', 'v4', credentials=creds, cache_discovery=False)

    if CONFIG['google_credentials_file'] and os.path.exists(CONFIG['google_credentials_file']):
        logger.debug("Fallback a autenticación con cuenta de servicio")
        creds = Credentials.from_service_account_file(CONFIG['google_credentials_file'], scopes=_SCOPES)
        return build('sheets', 'v4', credentials=creds, cache_discovery=False)

    raise Exception("Sin OAuth token ni credenciales de cuenta de servicio disponibles")

def copiar_formato_fila2(sheet_id: int, ultima_fila: int, spreadsheet_id: str | None = None) -> None:
    """Copia el formato de la fila 2 (todas las columnas) hacia las filas 3..ultima_fila."""
    if ultima_fila < 3:
        return

    sid = spreadsheet_id or CONFIG['basefija_id']
    service = _sheets_service()

    # Quita filtros activos antes de copyPaste: la API rechaza la operación
    # si hay filas filtradas/ocultas en el rango destino.
    quitar_filtro = {'clearBasicFilter': {'sheetId': sheet_id}}
    copiar_formato = {
        'copyPaste': {
            'source': {
                'sheetId': sheet_id,
                'startRowIndex': 1,
                'endRowIndex': 2,
                'startColumnIndex': 0,
                'endColumnIndex': 1000,
            },
            'destination': {
                'sheetId': sheet_id,
                'startRowIndex': 2,
                'endRowIndex': ultima_fila,
                'startColumnIndex': 0,
                'endColumnIndex': 1000,
            },
            'pasteType': 'PASTE_FORMAT',
            'pasteOrientation': 'NORMAL',
        }
    }

    try:
        service.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={'requests': [quitar_filtro, copiar_formato]},
        ).execute()
        logger.info(f"✓ Formato copiado de fila 2 a filas 3:{ultima_fila}")
    except Exception as e:
        logger.warning(f"⚠ No se pudo copiar formato (hoja puede tener filtros protegidos): {e}")

def conectar_google_sheets() -> Tuple[bool, gspread.Spreadsheet | None]:
    try:
        logger.info("Conectando a Google Sheets...")
        creds = _get_oauth_credentials()
        if creds:
            logger.debug("Usando OAuth")
            gc = gspread.Client(auth=creds)
        elif CONFIG['google_credentials_file'] and os.path.exists(CONFIG['google_credentials_file']):
            logger.debug("Fallback a cuenta de servicio")
            gc = gspread.service_account(filename=CONFIG['google_credentials_file'])
        else:
            raise Exception("Sin OAuth token ni credenciales de cuenta de servicio disponibles")

        sh = gc.open_by_key(CONFIG['basefija_id'])
        logger.info(f"✓ Conectado: {sh.title}")
        return True, sh

    except gspread.exceptions.SpreadsheetNotFound:
        logger.error(f"✗ Spreadsheet no encontrado (ID: {CONFIG['basefija_id']})")
        return False, None
    except gspread.exceptions.APIError as e:
        logger.error(f"✗ Error de API Google: {e}")
        return False, None
    except Exception as e:
        logger.error(f"✗ Error conectando a Google Sheets: {e}")
        logger.debug(f"Traza: {traceback.format_exc()}")
        return False, None

def cargar_csv_a_sheets(archivo_csv: str, nombre_hoja: str, sh: gspread.Spreadsheet,
                         con_formulas: bool = True, filtrar_columnas: bool = True) -> bool:
    """
    Sube el CSV a la hoja indicada.
    filtrar_columnas=True  → aplica CSV_COLUMNAS (25 cols, rango A:Y) — usado para BD_Ventas_Auren_Por_Hora.csv
    filtrar_columnas=False → sube todas las columnas del CSV tal cual — usado para BD_Ventas_AUREN.csv
    con_formulas=False     → omite fórmulas Z:AE y copia de formato — usado en FORMSMIFIBRA
    """
    ruta = Path(CONFIG['csv_local_folder']) / archivo_csv

    if not ruta.exists():
        logger.error(f"✗ Archivo no encontrado en carpeta compartida: {ruta}")
        return False

    # Verifica antigüedad del CSV en la carpeta compartida.
    # Si lleva más de 8 horas sin modificarse en un día laboral, es muy probable
    # que la VM no haya podido acceder a la red y esté copiando datos viejos.
    mod_ts = ruta.stat().st_mtime
    antiguedad_min = (datetime.now().timestamp() - mod_ts) / 60
    if antiguedad_min > 480:  # 8 horas
        logger.warning(
            f"⚠ DATOS POSIBLEMENTE VIEJOS: {archivo_csv} no se actualizó hace "
            f"{antiguedad_min / 60:.1f} horas (última modificación: "
            f"{datetime.fromtimestamp(mod_ts).strftime('%Y-%m-%d %H:%M')}). "
            "Es posible que la VM no haya podido conectarse a la red corporativa."
        )
    else:
        logger.info(f"CSV actualizado hace {antiguedad_min:.0f} min — OK")

    # Lee el CSV con fallback de encoding
    df = None
    for enc in ('utf-8', 'latin-1'):
        try:
            df = pd.read_csv(ruta, encoding=enc, dtype=str)
            break
        except UnicodeDecodeError:
            continue
        except pd.errors.EmptyDataError:
            logger.error(f"✗ Archivo vacío (0 bytes o sin columnas): {archivo_csv}")
            return False

    if df is None or df.empty:
        logger.error(f"✗ No se pudo leer o está vacío: {archivo_csv}")
        return False

    if filtrar_columnas:
        # Verifica y selecciona solo las 25 columnas definidas (en el orden correcto)
        columnas_faltantes = [c for c in CSV_COLUMNAS if c not in df.columns]
        if columnas_faltantes:
            logger.warning(f"⚠ Columnas no encontradas en el CSV: {columnas_faltantes}")
        df = df[[c for c in CSV_COLUMNAS if c in df.columns]]

    num_cols = len(df.columns)
    logger.info(f"Datos leídos: {len(df)} filas × {num_cols} columnas")

    # Columna final del rango (A=1, Y=25, etc.)
    col_final = chr(ord('A') + num_cols - 1) if num_cols <= 26 else 'Z'

    # Obtiene o crea la hoja
    try:
        ws = sh.worksheet(nombre_hoja)
        logger.info(f"Usando hoja existente: {nombre_hoja}")
    except gspread.exceptions.WorksheetNotFound:
        logger.info(f"Creando hoja nueva: {nombre_hoja}")
        ws = sh.add_worksheet(title=nombre_hoja, rows=len(df) + 1, cols=max(num_cols, 26))

    # Limpia el rango de datos (solo las columnas que maneja este CSV)
    num_filas_actuales = len(ws.get_all_values())
    if num_filas_actuales > 0:
        ws.batch_clear([f"A1:{col_final}{num_filas_actuales}"])
        logger.info(f"Limpiado rango A1:{col_final}{num_filas_actuales}")

    # Prepara y sube los datos
    valores = [df.columns.tolist()] + df.fillna('').values.tolist()
    ws.update(values=valores, range_name=f"A1:{col_final}{len(valores)}", raw=False)

    if con_formulas:
        primera_fila_datos = 2
        ultima_fila = len(df) + 1  # +1 por la fila de encabezados

        logger.info(f"Escribiendo fórmulas Z:AE en filas {primera_fila_datos}:{ultima_fila}...")

        formulas_por_fila = []
        for fila in range(primera_fila_datos, ultima_fila + 1):
            formulas_por_fila.append([
                1,                                                                                          # Z: contador
                f"=BUSCARV(G{fila},'CF MI FIBRA'!$B$5:$C$12,2,0)",                                        # AA
                f'=SI.ERROR(BUSCARV(D{fila},\'FORM MIFIBRA\'!A:AH,34,0),"")',                             # AB
                f"=BUSCARV(AB{fila},matriz!A:I,9,0)",                                                      # AC
                f'=EXTRAE(G{fila},ENCONTRAR(" ",G{fila},ENCONTRAR(" ",G{fila})+1)+1,LARGO(G{fila}))',      # AD
                f'=SI(C{fila}="internet+servicio ott",CONCATENAR(C{fila}," ",AD{fila}),G{fila})',          # AE
            ])

        ws.update(values=formulas_por_fila, range_name=f"Z{primera_fila_datos}:AE{ultima_fila}", raw=False)
        logger.info(f"✓ Fórmulas Z:AE escritas en {len(formulas_por_fila)} filas")

        # Copia formato de la fila 2 (plantilla) hacia todas las filas de datos
        copiar_formato_fila2(ws.id, ultima_fila, spreadsheet_id=sh.id)

    logger.info(f"✓ Completado: {nombre_hoja} ({len(df)} registros, {len(df.columns)} columnas)")
    return True

# ============================================================================
# NOTIFICACIÓN WHATSAPP
# ============================================================================

_GRUPO_BACK_MIFIBRA = "120363425538176534@g.us"
_STATE_FILE = _BASE_DIR / 'last_sync_state.json'
_CSV_POR_HORA = 'BD_Ventas_Auren_Por_Hora.csv'


def _leer_max_ult_actualizacion() -> str | None:
    """Devuelve el valor máximo de 'ult_actualizacion' del CSV Por_Hora, o None si no se puede leer.
    El campo viene en formato d/MM/YYYY HH:MM:SS (sin cero en el día), por eso se parsea a datetime
    antes de comparar — col.max() sobre strings daría un máximo lexicográfico incorrecto."""
    ruta = Path(CONFIG['csv_local_folder']) / _CSV_POR_HORA
    if not ruta.exists():
        return None
    for enc in ('utf-8', 'latin-1'):
        try:
            df = pd.read_csv(ruta, encoding=enc, dtype=str, usecols=['ult_actualizacion'])
            col = df['ult_actualizacion'].dropna()
            if col.empty:
                return None
            parsed = pd.to_datetime(col, dayfirst=True, errors='coerce').dropna()
            if parsed.empty:
                return None
            return parsed.max().strftime('%d/%m/%Y %H:%M:%S')
        except Exception:
            continue
    return None


def _leer_estado_anterior() -> str | None:
    try:
        with open(_STATE_FILE, encoding='utf-8') as f:
            return json.load(f).get('ult_actualizacion')
    except Exception:
        return None


def _guardar_estado(ult_actualizacion: str) -> None:
    try:
        with open(_STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'ult_actualizacion': ult_actualizacion}, f)
    except Exception as e:
        logger.debug(f"No se pudo guardar estado: {e}")


def notificar_whatsapp(hay_datos_nuevos: bool, ult_actualizacion: str | None) -> None:
    if not hay_datos_nuevos or not ult_actualizacion:
        logger.debug("Sin datos nuevos — notificación WhatsApp omitida")
        return
    try:
        from wa_client import WhatsAppClient
        wa = WhatsAppClient()
        if not wa.is_ready():
            logger.debug("Servidor WhatsApp no disponible — notificación omitida")
            return
        texto = f"Se cargaron en sus drives la info actualizada hasta las {ult_actualizacion}"
        wa.send_text(_GRUPO_BACK_MIFIBRA, texto)
        logger.info(f"✓ Notificación WhatsApp enviada: {texto}")
    except Exception as e:
        logger.debug(f"Notificación WhatsApp fallida: {e}")


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


def sincronizar_diario(sh_diario: gspread.Spreadsheet, primera: bool) -> Dict[str, bool]:
    """Carga al libro FORMSMIFIBRA. CSVs en CSV_SOLO_DIARIO solo se cargan en la primera ejecución del día."""
    resultados: Dict[str, bool] = {}
    for archivo_csv, nombre_hoja in CONFIG['csv_files_diario'].items():
        if archivo_csv in CSV_SOLO_DIARIO and not primera:
            logger.info(f"Saltando [formsmifibra] {archivo_csv} → {nombre_hoja} (ya se cargó hoy)")
            continue
        es_diario = archivo_csv in CSV_SOLO_DIARIO
        logger.info(f"\nProcesando [formsmifibra]: {archivo_csv} → {nombre_hoja}")
        resultados[archivo_csv] = cargar_csv_a_sheets(
            archivo_csv, nombre_hoja, sh_diario,
            con_formulas=False,
            filtrar_columnas=not es_diario,
        )
    return resultados


def sincronizar_completo(forzar: bool = False) -> bool:
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

    # --- Libro BASEFIJA: Por_Hora siempre; BD_Ventas_AUREN solo primera ejecución del día ---
    resultados: Dict[str, bool] = {}
    primera = True if forzar else _es_primera_ejecucion_del_dia()
    for archivo_csv, nombre_hoja in CONFIG['csv_files'].items():
        if archivo_csv in CSV_SOLO_DIARIO and not primera:
            logger.info(f"Saltando {archivo_csv} → {nombre_hoja} (ya se cargó hoy)")
            continue
        logger.info(f"\nProcesando [principal]: {archivo_csv} → {nombre_hoja}")
        es_diario = archivo_csv in CSV_SOLO_DIARIO
        resultados[archivo_csv] = cargar_csv_a_sheets(
            archivo_csv, nombre_hoja, sh,
            con_formulas=not es_diario,
            filtrar_columnas=not es_diario,
        )

    # --- Libro FORMSMIFIBRA: copia del Por_Hora en cada ejecución ---
    resultados_diario: Dict[str, bool] = {}
    if not CONFIG['formsmifibra_id']:
        logger.warning("⚠ FORMSMIFIBRA_ID vacío en .env — se omite el libro FORMSMIFIBRA")
    else:
        logger.info(f"\nConectando a Google Sheets FORMSMIFIBRA (ID: {CONFIG['formsmifibra_id'][:12]}…)")
        try:
            creds = _get_oauth_credentials()
            if creds:
                logger.debug("Usando OAuth para FORMSMIFIBRA")
                gc = gspread.Client(auth=creds)
            elif CONFIG['google_credentials_file'] and os.path.exists(CONFIG['google_credentials_file']):
                logger.debug("Fallback a cuenta de servicio para FORMSMIFIBRA")
                gc = gspread.service_account(filename=CONFIG['google_credentials_file'])
            else:
                raise Exception("Sin OAuth token ni credenciales de cuenta de servicio disponibles")

            sh_diario = gc.open_by_key(CONFIG['formsmifibra_id'])
            logger.info(f"✓ Conectado al libro diario: {sh_diario.title}")
            resultados_diario = sincronizar_diario(sh_diario, primera)
            if resultados_diario and all(resultados_diario.values()):
                logger.info("✓ DIARIO COMPLETADO EXITOSAMENTE")
        except gspread.exceptions.SpreadsheetNotFound:
            logger.error(f"✗ Spreadsheet diario no encontrado (ID: {CONFIG['formsmifibra_id']})")
        except Exception as e:
            logger.error(f"✗ Error en libro diario: {e}")
            logger.debug(f"Traza completa:\n{traceback.format_exc()}")

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

    if exitosos > 0:
        ult_actual = _leer_max_ult_actualizacion()
        ult_anterior = _leer_estado_anterior()
        hay_datos_nuevos = forzar or (ult_actual is not None and ult_actual != ult_anterior)
        if ult_actual is not None:
            _guardar_estado(ult_actual)
        notificar_whatsapp(hay_datos_nuevos, ult_actual)

    return exitosos == total

# ============================================================================
# ENTRADA
# ============================================================================

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Sincronización CSV → Google Sheets')
    parser.add_argument('--forzar', action='store_true',
                        help='Fuerza la carga de todos los CSVs (ignora control diario y estado de notificación)')
    args = parser.parse_args()

    # 0=lunes, 6=domingo — se omite el domingo (salvo con --forzar)
    if datetime.now().weekday() == 6 and not args.forzar:
        print("Domingo — no se ejecuta. Usa --forzar para ignorar esto.")
        sys.exit(0)

    campos_vacios = [k for k in ('csv_local_folder', 'basefija_id', 'google_credentials_file') if not CONFIG[k]]
    if campos_vacios:
        logger.error(f"✗ CONFIGURACIÓN INCOMPLETA en .env: {', '.join(campos_vacios)}")
        sys.exit(1)

    if not CONFIG['formsmifibra_id']:
        logger.warning("⚠ FORMSMIFIBRA_ID vacío en .env — el libro FORMSMIFIBRA no se cargará")

    if not CONFIG['csv_files']:
        logger.error("✗ csv_mapping.json no tiene entradas en 'csv_files'")
        sys.exit(1)

    if args.forzar:
        logger.info("⚡ MODO FORZADO — se cargan todos los CSVs sin restricción diaria")

    exito = sincronizar_completo(forzar=args.forzar)
    sys.exit(0 if exito else 1)

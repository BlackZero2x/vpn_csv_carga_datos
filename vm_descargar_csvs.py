"""
Script de la VM: copia CSVs desde la red corporativa hacia la carpeta compartida con el anfitrion.
Ejecutar dentro de la maquina virtual (VirtualBox).
Carpeta de destino visible desde el anfitrion en: C:/Users/developer2/Documents/vpncompartido
"""

import os
import sys
import json
import logging
from datetime import datetime
from pathlib import Path
from shutil import copy2

# ============================================================================
# CONFIGURACIÓN (editar aquí si cambian las rutas)
# ============================================================================

CONFIG = {
    # Ruta UNC en la red corporativa (accesible desde la VM por VPN)
    'csv_network_path': r'\\172.16.10.240\8.compartidos\Agencias\Auren',

    # Carpeta compartida VirtualBox (visible en el anfitrión como vpncompartido)
    'carpeta_destino': r'\\VBOXSVR\vpncompartido',

    # Archivos a copiar (deben coincidir con csv_mapping.json del anfitrión)
    'archivos': [
        'BD_Ventas_Auren_Por_Hora.csv',
    ],

    # Logs dentro de la carpeta compartida (también visibles desde el anfitrión)
    'logs_folder': r'\\VBOXSVR\vpncompartido\Logs',
}

# ============================================================================
# LOGGING
# ============================================================================

def setup_logger() -> logging.Logger:
    Path(CONFIG['logs_folder']).mkdir(parents=True, exist_ok=True)
    log_file = Path(CONFIG['logs_folder']) / f"vm_descarga_{datetime.now().strftime('%Y%m%d')}.log"

    logger = logging.getLogger('VMDescarga')
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
# DESCARGA
# ============================================================================

def descargar_csvs() -> bool:
    inicio = datetime.now()
    logger.info("=" * 60)
    logger.info(f"INICIANDO DESCARGA VM: {inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Origen:  {CONFIG['csv_network_path']}")
    logger.info(f"Destino: {CONFIG['carpeta_destino']}")
    logger.info("=" * 60)

    # Verifica acceso a la red corporativa
    if not Path(CONFIG['csv_network_path']).exists():
        logger.error(f"✗ No se puede acceder a la red corporativa: {CONFIG['csv_network_path']}")
        logger.error("  Verifica que la VPN esté activa dentro de la VM.")
        return False

    # Verifica acceso a la carpeta compartida
    if not Path(CONFIG['carpeta_destino']).exists():
        logger.error(f"✗ No se puede acceder a la carpeta compartida: {CONFIG['carpeta_destino']}")
        logger.error("  Verifica la configuración de carpetas compartidas en VirtualBox.")
        return False

    exitos = 0
    for archivo in CONFIG['archivos']:
        origen = Path(CONFIG['csv_network_path']) / archivo
        destino = Path(CONFIG['carpeta_destino']) / archivo

        try:
            if not origen.exists():
                logger.warning(f"✗ No encontrado en red: {archivo}")
                continue

            copy2(origen, destino)

            if destino.exists():
                tamaño_kb = destino.stat().st_size / 1024
                logger.info(f"✓ Copiado: {archivo} ({tamaño_kb:.1f} KB)")
                exitos += 1
            else:
                logger.error(f"✗ Falló la copia: {archivo}")

        except PermissionError:
            logger.error(f"✗ Permiso denegado: {archivo}")
        except Exception as e:
            logger.error(f"✗ Error copiando {archivo}: {e}")

    tiempo_total = datetime.now() - inicio
    logger.info(f"\nDescarga completada: {exitos}/{len(CONFIG['archivos'])} archivos | Tiempo: {tiempo_total}")

    if exitos == len(CONFIG['archivos']):
        logger.info("✓ DESCARGA COMPLETADA EXITOSAMENTE")
    else:
        logger.warning("⚠ DESCARGA INCOMPLETA — REVISAR LOGS")
    logger.info("=" * 60 + "\n")

    return exitos > 0

# ============================================================================
# ENTRADA
# ============================================================================

if __name__ == "__main__":
    # 0=lunes, 6=domingo — se omite el domingo
    if datetime.now().weekday() == 6:
        print("Domingo — no se ejecuta.")
        sys.exit(0)

    exito = descargar_csvs()
    sys.exit(0 if exito else 1)

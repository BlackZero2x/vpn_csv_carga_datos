#!/usr/bin/env python3
"""
Script para autenticarse con Google usando OAuth.
Genera un token que se reutiliza en futuras ejecuciones.

Uso:
    python authenticate_oauth.py

Esto abrirá un navegador para que autorices la aplicación.
El token se guardará en .cache/token.pickle
"""

import os
import sys
import pickle
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow

_BASE_DIR = Path(__file__).parent
_CACHE_DIR = _BASE_DIR / '.cache'
_TOKEN_PICKLE = _CACHE_DIR / 'token.pickle'
_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Credenciales OAuth2 — se debe descargar desde Google Cloud Console
# https://developers.google.com/sheets/api/quickstart/python
# Descarga el archivo credentials.json de tu Google Cloud Project y renómbralo aquí
_CREDENTIALS_JSON = _BASE_DIR / 'credentials.json'

def authenticate():
    """Realiza autenticación OAuth interactiva y guarda el token."""
    if not _CREDENTIALS_JSON.exists():
        print(f"❌ ERROR: No se encontró {_CREDENTIALS_JSON}")
        print("\nPasos para obtener credenciales OAuth:")
        print("1. Ve a https://console.cloud.google.com")
        print("2. Crea un nuevo proyecto o selecciona uno existente")
        print("3. Habilita Google Sheets API")
        print("4. Crea credenciales OAuth 2.0 (tipo: Aplicación de escritorio)")
        print("5. Descarga el archivo JSON y guárdalo como credentials.json en este directorio")
        sys.exit(1)

    _CACHE_DIR.mkdir(parents=True, exist_ok=True)

    print("🔐 Iniciando autenticación OAuth con Google...")
    print("Se abrirá una ventana del navegador. Por favor autoriza la aplicación.\n")

    flow = InstalledAppFlow.from_client_secrets_file(
        _CREDENTIALS_JSON,
        _SCOPES,
        redirect_uri='http://localhost:8080'
    )
    creds = flow.run_local_server(port=8080, open_browser=True)

    # Guarda el token para futuras ejecuciones
    with open(_TOKEN_PICKLE, 'wb') as token:
        pickle.dump(creds, token)

    print(f"\n✅ Autenticación completada!")
    print(f"✓ Token guardado en: {_TOKEN_PICKLE}")
    print("\nAhora puedes ejecutar vpn_csv_sync.py normalmente.")

if __name__ == '__main__':
    authenticate()

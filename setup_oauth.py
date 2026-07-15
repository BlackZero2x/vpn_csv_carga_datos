#!/usr/bin/env python3
"""
Script para configurar OAuth sin interfaz gráfica.
Guarda credenciales OAuth desde credenciales.json de Google Cloud.
"""

import os
import sys
import json
import pickle
import webbrowser
from pathlib import Path
from urllib.parse import urlencode, parse_qs
from urllib.request import urlopen
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

_BASE_DIR = Path(__file__).parent
_CACHE_DIR = _BASE_DIR / '.cache'
_TOKEN_PICKLE = _CACHE_DIR / 'token.pickle'
_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def setup_oauth():
    """Configura OAuth automáticamente."""
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)

    # Busca credentials.json en varios lugares
    possible_paths = [
        _BASE_DIR / 'credentials.json',
        Path.home() / 'credentials.json',
        Path.home() / 'Downloads' / 'credentials.json',
    ]

    credentials_json = None
    for path in possible_paths:
        if path.exists():
            credentials_json = path
            break

    if not credentials_json:
        print("❌ No se encontró credentials.json")
        print("\nOpciones:")
        print("1. Coloca credentials.json en:")
        for p in possible_paths:
            print(f"   {p}")
        print("\n2. O proporciona la ruta explícitamente:")
        print("   python setup_oauth.py /ruta/a/credentials.json")
        sys.exit(1)

    if len(sys.argv) > 1:
        credentials_json = Path(sys.argv[1])
        if not credentials_json.exists():
            print(f"❌ No existe: {credentials_json}")
            sys.exit(1)

    print(f"📋 Usando: {credentials_json}")
    print("🔐 Iniciando autenticación OAuth...\n")

    try:
        flow = InstalledAppFlow.from_client_secrets_file(
            credentials_json,
            _SCOPES
        )
        creds = flow.run_local_server(port=8080, open_browser=True)

        with open(_TOKEN_PICKLE, 'wb') as token:
            pickle.dump(creds, token)

        print(f"\n✅ Autenticación completada!")
        print(f"✓ Token guardado en: {_TOKEN_PICKLE}")
        print("\nvpn_csv_sync.py ahora puede ejecutarse automáticamente.")

    except Exception as e:
        print(f"\n❌ Error durante la autenticación: {e}")
        sys.exit(1)

if __name__ == '__main__':
    setup_oauth()

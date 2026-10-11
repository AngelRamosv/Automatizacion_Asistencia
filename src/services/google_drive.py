"""Servicio de autenticacion y descargas robustas desde Google Drive / Google Sheets."""
import time
import requests
from google.oauth2.service_account import Credentials
import google.auth.transport.requests


def get_drive_headers(creds_path, scopes):
    """Obtiene los encabezados HTTP con Bearer Token usando la Service Account."""
    try:
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        auth_req = google.auth.transport.requests.Request()
        creds.refresh(auth_req)
        return {"Authorization": f"Bearer {creds.token}"}
    except Exception as e:
        print(f"Error cargando credenciales de Google: {e}")
        return {}


def download_with_retry(url, headers, output_path, max_retries=3):
    """
    Descarga un archivo xlsx con reintentos para mitigar límites de tasa (HTTP 429)
    y respuestas HTML disfrazadas de Google Sheets.
    """
    for i in range(max_retries):
        try:
            r = requests.get(url, headers=headers)
            if r.status_code == 200 and not r.content.strip().startswith(b'<!DOCTYPE html'):
                with open(output_path, "wb") as f:
                    f.write(r.content)
                return True
            print(f"    [!] Intento {i+1} fallido para {output_path} (HTTP {r.status_code}, Posible bloqueo de Google). Reintentando en 10s...")
            time.sleep(10)
        except Exception as e:
            print(f"    [!] Intento {i+1} error de red: {e}. Reintentando en 10s...")
            time.sleep(10)
    return False


def download_sheet_by_gid(sheet_id, gid, headers, output_path, max_retries=3):
    """Descarga una pestaña específica de Google Sheets en formato XLSX dado su GID."""
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={gid}"
    return download_with_retry(url, headers, output_path, max_retries=max_retries)

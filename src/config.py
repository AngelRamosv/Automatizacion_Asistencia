"""Configuracion centralizada, constantes y rutas del sistema."""
import os

# --- TELEGRAM ---
TELEGRAM_TOKEN = "8730285872:AAHyIXLfOHjYEQRZgEX4S_ONGdaO_G4d5DU"
TELEGRAM_CHAT_ID = "-5016797911"

# --- BASE DE DATOS (CyberHub) ---
DB_CONFIG = {
    "host": "192.168.50.33",
    "user": "cyberhub",
    "password": "masterC1berHUb#",
    "database": "cyber_ideas_hub"
}

# --- GOOGLE DRIVE / SHEETS ---
SHEET_ID = "1HK7mp9l39vXPgyMisrDZVnY3jc-BaDzf43fMrjPUzQU"
GID_PRINCIPAL = "595741460"
GID_VTAS = "187319239"
GID_CYBER = "0"
GID_WINCALL = "99266008"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# --- RUTAS DE DIRECTORIOS Y ARCHIVOS ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE_DIR = os.environ.get("WORKSPACE_DIR", BASE_DIR)
SEMANA_DIR = os.path.join(WORKSPACE_DIR, "semana")
TEMP_DIR = os.path.join(WORKSPACE_DIR, "temp")
CREDS_PATH = os.path.join(WORKSPACE_DIR, "credentials.json")

# Archivos estaticos/temporales
INPUT_PRINCIPAL = os.path.join(WORKSPACE_DIR, "NOMINA POR SUPERVISOR IZZI ATN.xlsx")
INPUT_VTAS = os.path.join(WORKSPACE_DIR, "IZZI VTAS CONCENTRADO.xlsx")
OUTPUT_PRINCIPAL = os.path.join(SEMANA_DIR, "NOMINA POR SUPERVISOR IZZI ATN.xlsx")
OUTPUT_VTAS = os.path.join(SEMANA_DIR, "IZZI VTAS CONCENTRADO.xlsx")
OUTPUT_CYBER_CONC = os.path.join(SEMANA_DIR, "CYBERIDEAS CONCENTRADO.xlsx")
OUTPUT_WINCALL_CONC = os.path.join(SEMANA_DIR, "WINCALL CONCENTRADO.xlsx")
OUTPUT_ATN_RAW = os.path.join(TEMP_DIR, "ATN_RAW.xlsx")

# --- CATALOGOS Y REGLAS DE NEGOCIO ---
MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}

DAY_INDICES = [8, 10, 12, 14, 16, 18, 20]  # Columnas I, K, M, O, Q, S, U
DAY_NAMES = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO"]

ATTENDANCE_CODES = ['A', 'D', 'F', 'V', 'FJ', 'INC', 'SUSP']
CODE_HEADERS = [
    'Asistencias (A)',
    'Descansos (D)',
    'Faltas (F)',
    'Vacaciones (V)',
    'Faltas Justificadas (FJ)',
    'Incapacidades (INC)',
    'Suspensiones (SUSP)'
]

ABSENCE_CODES = {'F', 'FI', 'B', 'BAJA', 'INC', 'FJ'}
DATE_COLS_CONCENTRADO = [17, 19, 21, 23, 25, 27, 29]
SUNDAY_CODE_COL = 29
RETARDOS_COL = 32

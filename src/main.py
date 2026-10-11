import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
import datetime
import os
import requests
import json
import pandas as pd
import mysql.connector
from google.oauth2.service_account import Credentials
import google.auth.transport.requests
import glob
import shutil

TELEGRAM_TOKEN = "8730285872:AAHyIXLfOHjYEQRZgEX4S_ONGdaO_G4d5DU"
TELEGRAM_CHAT_ID = "-5016797911"

def send_telegram_reports(summary_text, file_paths):
    """Envia el resumen de asistencia y los archivos a Telegram."""
    print("--- INICIANDO ENVIO A TELEGRAM ---")
    url_msg = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": summary_text,
        "parse_mode": "HTML"
    }
    
    try:
        r = requests.post(url_msg, data=payload)
        res = r.json()
        if res.get("ok"):
            print("Mensaje de resumen enviado correctamente a Telegram.")
        else:
            print(f"Error al enviar resumen a Telegram: {res}")
            return False
    except Exception as e:
        print(f"Error de red al enviar mensaje a Telegram: {e}")
        return False

    url_doc = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"
    for caption, path in file_paths.items():
        if os.path.exists(path):
            print(f"Enviando '{os.path.basename(path)}'...")
            try:
                with open(path, "rb") as f:
                    files = {"document": f}
                    data = {"chat_id": TELEGRAM_CHAT_ID, "caption": caption}
                    r = requests.post(url_doc, data=data, files=files)
                    res = r.json()
                    if res.get("ok"):
                        print(f"Archivo '{os.path.basename(path)}' enviado con éxito.")
                    else:
                        print(f"Error de Telegram al enviar archivo '{os.path.basename(path)}': {res}")
            except Exception as e:
                print(f"Error al enviar archivo '{os.path.basename(path)}': {e}")
        else:
            print(f"Advertencia: El archivo '{path}' no existe y no se enviará.")
            
    print("--- ENVIO A TELEGRAM FINALIZADO ---")
    return True


def conectar_bd():
    return mysql.connector.connect(
        host='192.168.50.33',
        user='cyberhub',
        password='masterC1berHUb#',
        database='cyber_ideas_hub'
    )

def extraer_empleados(conn):
    query = "SELECT personal_id AS 'ID', personal_id AS 'Clave trabajador', personal_apaterno AS 'A. paterno', personal_amaterno AS 'A. materno', personal_nombre AS 'Nombre (s)', personal_sueldo_diario AS 'Salario diario', personal_fecha_nacimiento AS 'Fecha de nacimiento', personal_ingreso_sa AS 'Fecha de alta', personal_imss AS 'IMSS', personal_curp AS 'CURP', personal_rfc AS 'RFC', personal_celular AS 'Teléfono 1', personal_razonSocial AS 'Razon Social' FROM personal WHERE personal_estatus = 1"
    return pd.read_sql(query, conn)

def extraer_percepciones_deducciones(conn):
    query = "SELECT c.personal_id, c.tipo, c.p_d, c.valor FROM config_percep_deduc c WHERE c.estatus = 'INDEFINIDO'"
    return pd.read_sql(query, conn)

def generar_trabajadores(df_empleados, output_path):
    columnas_trabajadores = [
        'Clave trabajador', 'A. paterno', 'A. materno', 'Nombre (s)', 'Salario diario', 'Sexo', 
        'Fecha de nacimiento', 'Fecha de alta', 'IMSS', 'Lugar de Nacimiento', 'CURP', 'RFC', 
        'Calle', 'Colonia', 'Población', 'Código postal', 'Código postal Fiscal', 'Entidad federativa', 
        'Teléfono 1', 'Teléfono 2', 'Correo electrónico', 'Correo electrónico 2', 'Nivel de estudios', 
        'Profesión', 'Estado civil', 'Tipo de sangre', 'Nombre de Emergencia', 'Telefono de Emergencia', 
        'Forma de pago', 'Banco operador', 'Localidad / sucursal', 'Control de banco', 'Cuenta depósito', 
        'Tipo de empleado', 'Tipo de jornada', 'Turno', 'Unidad médica familiar', 'Descuenta pensión alimenticia', 
        'Registro patronal', 'Clasificación', 'Departamento', 'Puesto', 'Tipo de contrato', 'Clave de Régimen de contrato', 
        'País', 'Tipo ingreso', 'Base cotización', 'Clave de tipo contrato V12', 'Clave de tipo jornada', 
        'Clave de entidad federativa V12', 'Ruta plantilla', 'Clave Régimen', 'Clave de banco', 'Tabla de vacaciones', 
        'Tabla de sal. Diario', 'ModalidaTrabajo', 'SDI capturado', 'Dias contrato'
    ]
    df_trab = pd.DataFrame(columns=columnas_trabajadores)
    for col in df_empleados.columns:
        if col in df_trab.columns:
            df_trab[col] = df_empleados[col]
            
    fecha_limite = pd.Timestamp.now() - pd.Timedelta(days=7)
    df_trab['Fecha de alta'] = pd.to_datetime(df_trab['Fecha de alta'], errors='coerce')
    df_trab_nuevos = df_trab[df_trab['Fecha de alta'] >= fecha_limite]
    if len(df_trab_nuevos) > 0:
        print(f"\n¡ATENCIÓN! Se detectaron {len(df_trab_nuevos)} agentes nuevos contratados en los últimos 7 días.")
    else:
        print("\nNo hubo contrataciones nuevas en los últimos 7 días.")
        
    df_trab = df_trab_nuevos
    df_trab.to_excel(output_path, index=False)
    print(f"Trabajadores.xlsx generado: {output_path}")

def generar_movimientos(df_empleados, df_per_ded, expected_dates_dmy, output_path_cyber, output_path_wincall, path_cyber_conc, path_wincall_conc):
    """
    Genera DOS reportes de Percepciones (Prima Dominical P19), uno por empresa.
    Se guia EXCLUSIVAMENTE de los concentrados de ambas empresas.
    Solo incluye agentes con A o R en el dia domingo.
    Calcula: sueldo_diario * 0.25
    """
    import re as _re
    import openpyxl as _opx
    
    fecha_domingo = expected_dates_dmy[6] if len(expected_dates_dmy) > 6 else ""
    cols = ['Clave trabajador', 'Per/Ded', 'Num Per/Ded', 'Num credito', 'Aplica Destaj', 'Aplicacion',
            'Fecha inicio', 'Fecha fin', 'Monto o formula']
    
    def extraer_nois_con_domingo(path):
        """Lee el concentrado y devuelve los NOIs de agentes que asistieron (A o R) el domingo."""
        nois_domingo = set()
        if not os.path.exists(path):
            print(f"  ADVERTENCIA Percepciones: No se encontro {path}")
            return nois_domingo
        
        wb_c = _opx.load_workbook(path, data_only=True)
        ws_c = wb_c.active
        
        # La columna del domingo siempre es AC (29) en los concentrados, al igual que hicimos con las faltas.
        # Evitamos leer la fila 1 porque Google Sheets corrompe el formato de las fechas.
        sunday_code_col = 29
        
        for r in range(2, ws_c.max_row + 1):
            noi_val = ws_c.cell(row=r, column=3).value
            if noi_val is None:
                continue
            try:
                noi_code = int(float(str(noi_val).strip()))
            except (ValueError, TypeError):
                continue
            if noi_code <= 0:
                continue
            code = str(ws_c.cell(row=r, column=sunday_code_col).value or "").strip().upper()
            if code in ('A', 'R'):
                nois_domingo.add(noi_code)
        
        wb_c.close()
        return nois_domingo
    
    def build_movimientos(nois_set):
        movimientos = []
        for noi_code in nois_set:
            clave = str(noi_code)
            empleado = df_empleados[df_empleados['Clave trabajador'].astype(str) == clave]
            if not empleado.empty:
                sueldo_diario = pd.to_numeric(empleado['Salario diario'].values[0], errors='coerce')
                if pd.notnull(sueldo_diario):
                    monto_prima = round(sueldo_diario * 0.25, 2)
                    movimientos.append({
                        'Clave trabajador': clave, 'Per/Ded': 'P', 'Num Per/Ded': 19,
                        'Num credito': 0, 'Aplica Destaj': 'N', 'Aplicacion': 4,
                        'Fecha inicio': fecha_domingo, 'Fecha fin': fecha_domingo,
                        'Monto o formula': monto_prima
                    })
        df = pd.DataFrame(movimientos)
        if not df.empty:
            return df[cols].sort_values('Clave trabajador').reset_index(drop=True)
        return pd.DataFrame(columns=cols)
    
    nois_cyber = extraer_nois_con_domingo(path_cyber_conc)
    nois_wincall = extraer_nois_con_domingo(path_wincall_conc)
    
    print(f"  -> {len(nois_cyber)} agentes CyberIdeas con asistencia el domingo")
    print(f"  -> {len(nois_wincall)} agentes WinCall con asistencia el domingo")
    
    df_cyber = build_movimientos(nois_cyber)
    df_wincall = build_movimientos(nois_wincall)
    
    with pd.ExcelWriter(output_path_cyber, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_cyber.to_excel(writer, index=False, sheet_name="Prima_Dominical")
    print(f"Percepciones_CyberIdeas.xlsx generado: {output_path_cyber}")
    
    with pd.ExcelWriter(output_path_wincall, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_wincall.to_excel(writer, index=False, sheet_name="Prima_Dominical")
    print(f"Percepciones_WinCall.xlsx generado: {output_path_wincall}")

def extract_bono_from_concentrado(path_concentrado, col_name, expected_dates_dmy):
    """
    Genera el reporte de Percepciones Asistencia P15 desde el concentrado.
    Busca la columna col_name ('BONO' o 'BONO AUTORIZADO').
    Todos los agentes con OK en esa columna reciben P15 con monto fijo 115.4.
    """
    fecha_fin_semana = expected_dates_dmy[6] if len(expected_dates_dmy) > 6 else expected_dates_dmy[-1]
    
    cols = ["Clave trabajador", "Per/Ded", "Num Per/Ded", "Num credito", "Aplica Destaj", "Aplicacion",
            "Fecha inicio", "Fecha fin", "Monto o formula"]
    
    import os, openpyxl
    import pandas as pd
    
    if not os.path.exists(path_concentrado):
        return pd.DataFrame(columns=cols)

    wb = openpyxl.load_workbook(path_concentrado, data_only=True)
    ws = wb.active

    bono_col = None
    for r_h in [1, 2]:
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r_h, column=c).value or "").strip().upper()
            if val == col_name.upper():
                bono_col = c
                break
        if bono_col:
            break

    if bono_col is None:
        wb.close()
        return pd.DataFrame(columns=cols)

    movimientos = []
    for r in range(2, ws.max_row + 1):
        noi_val = ws.cell(row=r, column=3).value
        if noi_val is None:
            continue
        try:
            noi_code = int(float(str(noi_val).strip()))
        except (ValueError, TypeError):
            continue
        if noi_code <= 0:
            continue

        bono_val = str(ws.cell(row=r, column=bono_col).value or "").strip().upper()
        if bono_val == "OK":
            movimientos.append({
                "Clave trabajador": str(noi_code), "Per/Ded": "P", "Num Per/Ded": 15,
                "Num credito": 0, "Aplica Destaj": "N", "Aplicacion": 4,
                "Fecha inicio": fecha_fin_semana, "Fecha fin": fecha_fin_semana, "Monto o formula": 115.4
            })
            movimientos.append({
                "Clave trabajador": str(noi_code), "Per/Ded": "P", "Num Per/Ded": 16,
                "Num credito": 0, "Aplica Destaj": "N", "Aplicacion": 4,
                "Fecha inicio": fecha_fin_semana, "Fecha fin": fecha_fin_semana, "Monto o formula": 115.4
            })

    wb.close()
    
    df_premios = pd.DataFrame(movimientos)
    if not df_premios.empty:
        df_premios = df_premios[cols].sort_values("Clave trabajador").reset_index(drop=True)
    else:
        df_premios = pd.DataFrame(columns=cols)

    return df_premios

def download_with_retry(url, headers, output_path, max_retries=3):
    import time
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

def process_attendance_live():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    workspace = r"C:\Users\Agents\Desktop\Automatizacion_Asistencia"
    pruebas_dir = os.path.join(workspace, "semana")
    os.makedirs(pruebas_dir, exist_ok=True)
    
    creds_path = os.path.join(workspace, "credentials.json")
    
    input_principal = os.path.join(workspace, "NOMINA POR SUPERVISOR IZZI ATN.xlsx")
    input_vtas = os.path.join(workspace, "IZZI VTAS CONCENTRADO.xlsx")
    
    output_principal = os.path.join(pruebas_dir, "NOMINA POR SUPERVISOR IZZI ATN.xlsx")
    output_vtas = os.path.join(pruebas_dir, "IZZI VTAS CONCENTRADO.xlsx")
    
    # Archivos temporales (solo lectura interna, NO se envian ni guardan en la carpeta final)
    temp_dir = os.path.join(workspace, "temp")
    os.makedirs(temp_dir, exist_ok=True)
    
    # El usuario solicitó guardar los concentrados en la carpeta semana para validación (10 reportes en total)
    output_cyber_conc = os.path.join(pruebas_dir, "CYBERIDEAS CONCENTRADO.xlsx")
    output_wincall_conc = os.path.join(pruebas_dir, "WINCALL CONCENTRADO.xlsx")
    
    output_atn_raw = os.path.join(temp_dir, "ATN_RAW.xlsx")  # ATN original con columna BONO
    
    print("--- INICIANDO AUTOMATIZACION DE ASISTENCIA EN VIVO ---")
    
    sheet_id = "1HK7mp9l39vXPgyMisrDZVnY3jc-BaDzf43fMrjPUzQU"
    gid_principal = "595741460"
    gid_vtas = "187319239"
    gid_cyber = "0"
    gid_wincall = "99266008"
    try:
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
        auth_req = google.auth.transport.requests.Request()
        creds.refresh(auth_req)
        headers = {"Authorization": f"Bearer {creds.token}"}
    except Exception as e:
        print(f"Error cargando credenciales: {e}")
        headers = {}
    

    # ================= 1. DESCARGAR NOMINA POR SUPERVISOR IZZI ATN (Y EXTRAER FECHAS) =================
    print("\nDescargando NOMINA POR SUPERVISOR IZZI ATN.xlsx directamente con GID para extraer fechas de la semana...")
    url_principal = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={gid_principal}"
    if download_with_retry(url_principal, headers, output_principal):
        import shutil
        shutil.copy(output_principal, output_atn_raw)
        print("Descarga de nomina principal completa.")
    else:
        print("Error crítico: No se pudo descargar nómina principal después de varios intentos.")
        return
        
    print("Extrayendo fechas inteligentemente desde NOMINA POR SUPERVISOR IZZI ATN...")
    import re
    wb_temp = openpyxl.load_workbook(output_atn_raw, data_only=True)
    ws_temp = wb_temp.active
    
    val_lunes = ws_temp.cell(row=1, column=9).value # Columna I (lunes)
    
    if isinstance(val_lunes, (datetime.datetime, datetime.date)):
        monday_dt = val_lunes
    else:
        v_str = str(val_lunes).strip()
        if re.match(r'^\d{2}/\d{2}/\d{4}', v_str):
            v_str = v_str[:10]
        try:
            monday_dt = datetime.datetime.strptime(v_str, "%d/%m/%Y")
        except:
            monday_dt = datetime.datetime.strptime(v_str, "%m/%d/%Y")
            
    expected_dates_dmy = []
    for i in range(7):
        day_dt = monday_dt + datetime.timedelta(days=i)
        expected_dates_dmy.append(day_dt.strftime("%d/%m/%Y"))
    wb_temp.close()
    
    print(f"Fechas leídas del Drive (semana actual de trabajo): {expected_dates_dmy}")
    reported_monday = monday_dt
    meses = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 
             7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"}
    sunday_dt = monday_dt + datetime.timedelta(days=6)
    if monday_dt.month == sunday_dt.month:
        semana_str = f"semana_{monday_dt.day}_al_{sunday_dt.day}_{meses[sunday_dt.month]}_{sunday_dt.year}"
    else:
        semana_str = f"semana_{monday_dt.day}_{meses[monday_dt.month]}_al_{sunday_dt.day}_{meses[sunday_dt.month]}_{sunday_dt.year}"

    # Nuevos paths (Plan B)
    if 'pruebas_dir' in locals():
        out_dir = pruebas_dir
    else:
        out_dir = prueba_dir

    output_incidencias_cyber = os.path.join(out_dir, f"Incidencias_CyberIdeas_{semana_str}.xlsx")
    output_incidencias_wincall = os.path.join(out_dir, f"Incidencias_WinCall_{semana_str}.xlsx")
    output_premios_cyber    = os.path.join(out_dir, f"Percepciones_Bono_asistencia_CyberIdeas_{semana_str}.xlsx")
    output_premios_wincall  = os.path.join(out_dir, f"Percepciones_Bono_asistencia_WinCall_{semana_str}.xlsx")
    
    # Maintain old paths for reference before ExcelWriter refactoring
    output_faltas_cyber = output_incidencias_cyber
    output_faltas_wincall = output_incidencias_wincall
    output_percepciones_cyber = output_incidencias_cyber
    output_percepciones_wincall = output_incidencias_wincall
    output_cyber_conc = os.path.join(out_dir, "CYBERIDEAS CONCENTRADO.xlsx")
    output_wincall_conc = os.path.join(out_dir, "WINCALL CONCENTRADO.xlsx")

    
    # ================= 1. GENERATING IZZI VTAS CONCENTRADO =================
    print("\nDescargando y generando IZZI VTAS CONCENTRADO.xlsx directamente con GID...")
    url_vtas = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={gid_vtas}"
    if not download_with_retry(url_vtas, headers, output_vtas):
        print("Error crítico al descargar pestaña de ventas.")
        return
    
    try:
        # Auto-fit and force static dates in IZZI VTAS CONCENTRADO to avoid empty cells or '###'
        print("Escribiendo y formateando fechas estáticas en IZZI VTAS CONCENTRADO.xlsx...")
        wb_vtas = openpyxl.load_workbook(output_vtas)
        ws_vtas = wb_vtas.active
        date_cols = ['I', 'K', 'M', 'O', 'Q', 'S', 'U']
        for idx, col in enumerate(date_cols):
            col_date = reported_monday + datetime.timedelta(days=idx)
            cell = ws_vtas[f"{col}1"]
            cell.value = col_date
            cell.number_format = 'dd/mm/yyyy'
            ws_vtas.column_dimensions[col].width = 12.0
        wb_vtas.save(output_vtas)
        wb_vtas.close()
        print("IZZI VTAS CONCENTRADO.xlsx creado con éxito con estilos y colores originales.")
    except Exception as e:
        print(f"Error procesando pestaña de ventas: {e}")
        return
        
    
    print("\nDescargando pestañas de Concentrados (Cyber y Wincall)...")
    url_cyber = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={gid_cyber}"
    url_wincall = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={gid_wincall}"
    import time
    time.sleep(2) # Pausa extra antes de los concentrados
    if not download_with_retry(url_cyber, headers, output_cyber_conc):
        print("Error crítico al descargar CYBERIDEAS CONCENTRADO")
    time.sleep(2)
    if not download_with_retry(url_wincall, headers, output_wincall_conc):
        print("Error crítico al descargar WINCALL CONCENTRADO")
    print("Descarga de Concentrados completa.")
        
    print("Cargando NOMINA POR SUPERVISOR IZZI ATN.xlsx para procesamiento de asistencia...")
    wb_principal = openpyxl.load_workbook(output_principal)
    ws = wb_principal["NOMINA POR SUPERVISOR IZZI ATN"]
    ws.views.sheetView[0].showGridLines = True
    
    max_row = ws.max_row
    print(f"Filas detectadas en la hoja original: {max_row}")
    
    # Extract and Validate Dates in Row 1
    day_indices = [8, 10, 12, 14, 16, 18, 20] # Columns I, K, M, O, Q, S, U
    day_names = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO", "DOMINGO"]
    
    print(f"Fechas correctas esperadas para la semana: {expected_dates_dmy}")

    meses = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 
             7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"}
    sunday_dt = reported_monday + datetime.timedelta(days=6)
    if reported_monday.month == sunday_dt.month:
        semana_str = f"semana_{reported_monday.day}_al_{sunday_dt.day}_{meses[sunday_dt.month]}_{sunday_dt.year}"
    else:
        semana_str = f"semana_{reported_monday.day}_{meses[reported_monday.month]}_al_{sunday_dt.day}_{meses[sunday_dt.month]}_{sunday_dt.year}"

    # Nuevos paths (Plan B)
    if 'pruebas_dir' in locals():
        out_dir = pruebas_dir
    else:
        out_dir = prueba_dir

    output_incidencias_cyber = os.path.join(out_dir, f"Incidencias_CyberIdeas_{semana_str}.xlsx")
    output_incidencias_wincall = os.path.join(out_dir, f"Incidencias_WinCall_{semana_str}.xlsx")
    output_premios_cyber    = os.path.join(out_dir, f"Percepciones_Bono_asistencia_CyberIdeas_{semana_str}.xlsx")
    output_premios_wincall  = os.path.join(out_dir, f"Percepciones_Bono_asistencia_WinCall_{semana_str}.xlsx")
    
    # Maintain old paths for reference before ExcelWriter refactoring
    output_faltas_cyber = output_incidencias_cyber
    output_faltas_wincall = output_incidencias_wincall
    output_percepciones_cyber = output_incidencias_cyber
    output_percepciones_wincall = output_incidencias_wincall
    output_cyber_conc = os.path.join(out_dir, "CYBERIDEAS CONCENTRADO.xlsx")
    output_wincall_conc = os.path.join(out_dir, "WINCALL CONCENTRADO.xlsx")

    
    corrections_made = []
    for idx, col_idx in enumerate(day_indices):
        cell = ws.cell(row=1, column=col_idx+1)
        old_val = str(cell.value or "").strip()
        new_val = expected_dates_dmy[idx]
        if old_val != new_val:
            warning_msg = f"{day_names[idx]} (Columna {col_idx+1}): Se leyó '{old_val}', se corrigió a '{new_val}'"
            print(f"!ADVERTENCIA! {warning_msg}")
            corrections_made.append(warning_msg)
            cell.value = new_val

    # Parse agents for Faltas sheet
    agents_records = []
    for r in range(3, max_row + 1):
        noi_val = ws.cell(row=r, column=4).value # Column D is NOI
        if noi_val is None or str(noi_val).strip() == "" or str(noi_val).strip().lower() == "noi":
            continue
            
        try:
            noi_code = int(float(str(noi_val).strip()))
        except ValueError:
            continue
            
        supervisor = str(ws.cell(row=r, column=2).value or "").strip() # Column B
        puesto = str(ws.cell(row=r, column=5).value or "").strip() # Column E
        jornada = str(ws.cell(row=r, column=6).value or "").strip() # Column F
        nombre = str(ws.cell(row=r, column=7).value or "").strip() # Column G
        apellido = str(ws.cell(row=r, column=8).value or "").strip() # Column H
        
        day_codes = []
        for col_idx in day_indices:
            code = str(ws.cell(row=r, column=col_idx+1).value or "").strip().upper()
            day_codes.append(code if code else "")
            
        agents_records.append({
            'row_index_1based': r,
            'noi_code': noi_code,
            'supervisor': supervisor,
            'puesto': puesto,
            'jornada': jornada,
            'nombre': nombre,
            'apellido': apellido,
            'day_codes': day_codes
        })
        
    print(f"Total de agentes identificados para procesamiento: {len(agents_records)}")
    
    # ================= 3. GENERATING FALTAS =================
    print("\nIniciando extracción de concentrados para filtrar Faltas...")
    
    absence_codes = {'F', 'FI', 'B', 'BAJA', 'INC', 'FJ'}
    
    import re as _re
    
    def extract_faltas_from_concentrado(path):
        """
        Lee el concentrado directamente fila por fila.
        NOI en columna C (3). Las fechas están en columnas impares (17,19,21...).
        Los códigos de asistencia en las columnas pares inmediatamente siguientes (18,20,22...).
        Genera UNA fila por cada día con incidencia válida. Ignora D y blancos. Siempre Tipo falta = 2.
        """
        rows = []
        if not os.path.exists(path):
            print(f"  ADVERTENCIA: No se encontró {path}")
            return rows
        
        wb_c = openpyxl.load_workbook(path, data_only=True)
        ws_c = wb_c.active
        
        # Las fechas siempre están en las columnas: 17, 19, 21, 23, 25, 27, 29
        # Y los códigos de asistencia PRINCIPALES están en esas MISMAS columnas.
        # Las columnas pares (18, 20, 22...) son para retardos, no faltas.
        date_cols = [17, 19, 21, 23, 25, 27, 29]
        
        for r in range(2, ws_c.max_row + 1):
            noi_val = ws_c.cell(row=r, column=3).value  # NOI en col C
            if noi_val is None:
                continue
            try:
                noi_code = int(float(str(noi_val).strip()))
            except (ValueError, TypeError):
                continue
            if noi_code <= 0:
                continue
            
            for i, code_col in enumerate(date_cols):
                # Usar la fecha correcta ya calculada matemáticamente
                date_str = expected_dates_dmy[i] if i < len(expected_dates_dmy) else ""
                if not date_str:
                    continue
                
                # Leer el código de la MISMA columna donde iría la fecha (Q, S, U, W, Y, AA, AC)
                code_raw = str(ws_c.cell(row=r, column=code_col).value or "").strip().upper()
                
                # Ignorar D, blancos y no reconocidos
                if code_raw in ('D', '', 'NONE', 'A', 'R', 'V'):
                    continue
                if code_raw not in absence_codes:
                    continue
                
                rows.append({
                    'Clave trab': noi_code,
                    'Tipo falta': 2,
                    '% falta': 100,
                    'Certificado': 0,
                    '% pagado': 0,
                    'Fecha inic': date_str,
                    'Fecha fin': date_str
                })
            # --- REGLA RETARDOS: si el agente tiene exactamente 3 retardos, se agrega 1 falta extra ---
            retardos_val = ws_c.cell(row=r, column=32).value
            try:
                retardos = int(float(str(retardos_val).strip())) if retardos_val is not None else 0
            except (ValueError, TypeError):
                retardos = 0
            
            if retardos == 3:
                # Fechas que ya tienen falta registrada para este agente en esta iteración de fila
                fechas_con_falta = {row['Fecha inic'] for row in rows if row['Clave trab'] == noi_code}
                # Buscar primer día libre (sin falta) entre los 7 días de la semana
                fecha_extra = None
                for fecha_libre in expected_dates_dmy:
                    if fecha_libre not in fechas_con_falta:
                        fecha_extra = fecha_libre
                        break
                if fecha_extra:
                    rows.append({
                        'Clave trab': noi_code,
                        'Tipo falta': 2,
                        '% falta': 100,
                        'Certificado': 0,
                        '% pagado': 0,
                        'Fecha inic': fecha_extra,
                        'Fecha fin': fecha_extra
                    })
        
        wb_c.close()
        return rows
    
    print("  Leyendo faltas del CYBERIDEAS CONCENTRADO...")
    faltas_cyber_rows = extract_faltas_from_concentrado(output_cyber_conc)
    print(f"  -> {len(faltas_cyber_rows)} incidencias en CyberIdeas")
    
    print("  Leyendo faltas del WINCALL CONCENTRADO...")
    faltas_wincall_rows = extract_faltas_from_concentrado(output_wincall_conc)
    print(f"  -> {len(faltas_wincall_rows)} incidencias en WinCall")
    
    df_faltas_cyber = pd.DataFrame(faltas_cyber_rows)
    df_faltas_wincall = pd.DataFrame(faltas_wincall_rows)
    df_faltas_all = pd.concat([df_faltas_cyber, df_faltas_wincall], ignore_index=True)
    
    empty_cols = ['Clave trab', 'Tipo falta', '% falta', 'Certificado', '% pagado', 'Fecha inic', 'Fecha fin']
    # Se eliminó el filtro de valid_nois para que aparezcan todas las faltas de ambos concentrados (incluyendo VTAS)
    if df_faltas_cyber.empty:
        df_faltas_cyber = pd.DataFrame(columns=empty_cols)
    if df_faltas_wincall.empty:
        df_faltas_wincall = pd.DataFrame(columns=empty_cols)

    def save_and_style_faltas(df_to_save, out_path):
        df_to_save.to_excel(out_path, index=False, sheet_name="Faltas")
        wb_f = openpyxl.load_workbook(out_path)
        ws_f = wb_f["Faltas"]
        ws_f.views.sheetView[0].showGridLines = True
        for col in ws_f.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row > 1 and col_letter in ['F', 'G'] and cell.value is not None:
                    cell.alignment = Alignment(horizontal="center")
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws_f.column_dimensions[col_letter].width = max(max_len + 3, 14)
        wb_f.save(out_path)
        wb_f.close()

    print(f"Guardando archivos de Faltas en: {pruebas_dir}")
    save_and_style_faltas(df_faltas_cyber, output_faltas_cyber)
    save_and_style_faltas(df_faltas_wincall, output_faltas_wincall)
    
    # ================= 4. ADD COUNT COLUMNS & FORMULAS IN PRINCIPAL SHEET =================
    codes = ['A', 'D', 'F', 'V', 'FJ', 'INC', 'SUSP']
    code_headers = [
        'Asistencias (A)',
        'Descansos (D)',
        'Faltas (F)',
        'Vacaciones (V)',
        'Faltas Justificadas (FJ)',
        'Incapacidades (INC)',
        'Suspensiones (SUSP)'
    ]
    
    ref_header_cell = ws.cell(row=2, column=5) # 'EMPLEADO POR NOMBRE'
    ref_font = ref_header_cell.font
    ref_fill = ref_header_cell.fill
    ref_border = ref_header_cell.border
    
    start_col = 22 # Column V
    
    # Write count headers in Row 2
    for idx, header_text in enumerate(code_headers):
        col = start_col + idx
        cell = ws.cell(row=2, column=col, value=header_text)
        if ref_font:
            cell.font = copy(ref_font)
        if ref_fill:
            cell.fill = copy(ref_fill)
        if ref_border:
            cell.border = copy(ref_border)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    # Write Row 1 background/borders for new columns
    for idx in range(len(code_headers)):
        col = start_col + idx
        cell_r1 = ws.cell(row=1, column=col, value="")
        if ref_fill:
            cell_r1.fill = copy(ref_fill)
        if ref_border:
            cell_r1.border = copy(ref_border)
            
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Write COUNTIF formulas for each agent
    for agent in agents_records:
        r = agent['row_index_1based']
        for idx, code in enumerate(codes):
            col = start_col + idx
            col_letter_start = "H"
            col_letter_end = "T"
            
            if code == 'A':
                formula = f'=COUNTIF({col_letter_start}{r}:{col_letter_end}{r}, "A") + COUNTIF({col_letter_start}{r}:{col_letter_end}{r}, "R")'
            else:
                formula = f'=COUNTIF({col_letter_start}{r}:{col_letter_end}{r}, "{code}")'
                
            cell = ws.cell(row=r, column=col, value=formula)
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            
            # Subtle coloring for agent count cells
            code_count = agent['day_codes'].count(code) + (agent['day_codes'].count('R') if code == 'A' else 0)
            if code_count > 0:
                if code in ['F', 'SUSP']:
                    cell.fill = PatternFill(start_color="FCE4D6", fill_type="solid")
                    cell.font = Font(name="Calibri", size=11, bold=True)
                elif code == 'A':
                    cell.fill = PatternFill(start_color="E2EFDA", fill_type="solid")
                elif code == 'V':
                    cell.fill = PatternFill(start_color="FFF2CC", fill_type="solid")
                elif code == 'INC':
                    cell.fill = PatternFill(start_color="E8DAEF", fill_type="solid")
                    
    # Write bottom Total Row
    total_row = max_row + 1
    ws.cell(row=total_row, column=6, value="TOTAL GENERAL").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=total_row, column=6).alignment = Alignment(horizontal="right", vertical="center")
    
    for idx, code in enumerate(codes):
        col = start_col + idx
        col_letter = get_column_letter(col)
        cell = ws.cell(row=total_row, column=col, value=f"=SUM({col_letter}3:{col_letter}{total_row-1})")
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))
        cell.fill = PatternFill(start_color="D9E1F2", fill_type="solid")
        
    # Auto-fit the new count columns
    for col_idx in range(start_col, start_col + len(codes)):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for r in range(2, total_row + 1):
            val = ws.cell(row=r, column=col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)
        
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[total_row].height = 22
    
    wb_principal.save(output_principal)
    wb_principal.close()
    print("NOMINA POR SUPERVISOR IZZI ATN.xlsx actualizado con éxito.")
    
    # ------------------ INYECTANDO REPORTE DE PERCEPCIONES ------------------
    print("\nIniciando extraccion de CyberHub para Percepciones (Prima Dominical)...")
    try:
        conn = conectar_bd()
        df_empleados = extraer_empleados(conn)
        df_per_ded = extraer_percepciones_deducciones(conn)
        conn.close()
        
        generar_movimientos(df_empleados, df_per_ded, expected_dates_dmy, output_percepciones_cyber, output_percepciones_wincall, output_cyber_conc, output_wincall_conc)
    except Exception as e:
        print(f"Error generando reporte de Percepciones: {e}")
    # ---------------------------------------------------------------

    # ------------------ INYECTANDO REPORTE DE PREMIOS DE ASISTENCIA P15 ------------------
    print("\nGenerando Premios de Asistencia (P15 y P16, BONO OK) desde concentrados...")
    try:
        meses = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"}
        start_d = datetime.datetime.strptime(expected_dates_dmy[0], "%d/%m/%Y")
        end_d = datetime.datetime.strptime(expected_dates_dmy[6], "%d/%m/%Y")
        if start_d.month == end_d.month:
            semana_loc = f"semana_{start_d.day}_al_{end_d.day}_{meses[end_d.month]}_{end_d.year}"
        else:
            semana_loc = f"semana_{start_d.day}_{meses[start_d.month]}_al_{end_d.day}_{meses[end_d.month]}_{end_d.year}"

        output_premios_cyber   = os.path.join(pruebas_dir, f"Percepciones_Bono_asistencia_CyberIdeas_{semana_loc}.xlsx")
        output_premios_wincall = os.path.join(pruebas_dir, f"Percepciones_Bono_asistencia_WinCall_{semana_loc}.xlsx")
        
        df_bono_cyber = extract_bono_from_concentrado(output_cyber_conc, "BONO", expected_dates_dmy)
        df_bono_cyber.to_excel(output_premios_cyber, index=False)
        print(f"  -> CyberIdeas: {len(df_bono_cyber)} agentes. Guardado en {output_premios_cyber}")
        
        df_bono_wincall = extract_bono_from_concentrado(output_wincall_conc, "BONO AUTORIZADO", expected_dates_dmy)
        df_bono_wincall.to_excel(output_premios_wincall, index=False)
        print(f"  -> WinCall: {len(df_bono_wincall)} agentes. Guardado en {output_premios_wincall}")
    except Exception as e:
        print(f"Error generando reporte de Premios: {e}")
    # ---------------------------------------------------------------

    # ================= TELEGRAM SENDING =================
    start_date = expected_dates_dmy[0]
    end_date = expected_dates_dmy[6]
    current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    corrections_html = ""
    if corrections_made:
        corrections_html = "⚠️ <b>Correcciones de fecha aplicadas en el reporte:</b>\n"
        for corr in corrections_made:
            corrections_html += f"• {corr}\n"
        corrections_html += "\n"
        
    summary_text = (
        f"<b>📊 REPORTES DE ASISTENCIA Y OPERACIONES</b>\n"
        f"📅 Periodo: <code>{start_date}</code> al <code>{end_date}</code>\n\n"
        f"<b>Resumen de Procesamiento:</b>\n"
        f"👥 Agentes totales procesados: <code>{len(agents_records)}</code>\n"
        f"❌ Faltas totales registradas: <code>{df_faltas_all.shape[0]}</code>\n\n"
        f"{corrections_html}"
        f"Se adjuntan los siguientes reportes para revisión:\n"
        f"1️⃣ <b>NOMINA POR SUPERVISOR IZZI ATN</b>\n"
        f"2️⃣ <b>IZZI VTAS CONCENTRADO</b>\n"
        f"3️⃣ <b>CYBERIDEAS CONCENTRADO (Drive)</b>\n"
        f"4️⃣ <b>WINCALL CONCENTRADO (Drive)</b>\n"
        f"5️⃣ <b>Incidencias CyberIdeas (Faltas + Prima Dominical)</b>\n"
        f"6️⃣ <b>Incidencias WinCall (Faltas + Prima Dominical)</b>\n"
        f"7️⃣ <b>Percepciones Bono Asistencia - CyberIdeas (P15 y P16)</b>\n"
        f"8️⃣ <b>Percepciones Bono Asistencia - WinCall (P15 y P16)</b>\n\n"
        f"<i>🤖 Sistema automatizado de asistencia - {current_time}</i>"
    )
    
    file_paths = {
        "📊 Reporte NOMINA POR SUPERVISOR IZZI ATN": output_principal,
        "📈 Concentrado IZZI VTAS CONCENTRADO": output_vtas,
        "📂 CYBERIDEAS CONCENTRADO (Drive)": output_cyber_conc,
        "📂 WINCALL CONCENTRADO (Drive)": output_wincall_conc,
        "📁 Incidencias CyberIdeas (Faltas + P19)": output_incidencias_cyber,
        "📁 Incidencias WinCall (Faltas + P19)": output_incidencias_wincall,
        "🏆 Percepciones Bono Asistencia CyberIdeas (P15)": output_premios_cyber,
        "🏆 Percepciones Bono Asistencia WinCall (P15)": output_premios_wincall
    }
    
    send_telegram_reports(summary_text, file_paths)
    print('\n*** REPORTES ENVIADOS A TELEGRAM EXITOSAMENTE ***\n')
    
    import glob
    def contar_filas_hoja(patron, hoja):
        archivos = sorted(glob.glob(os.path.join(pruebas_dir, patron)), key=os.path.getmtime, reverse=True)
        if not archivos:
            return "N/A"
        try:
            return len(pd.read_excel(archivos[0], sheet_name=hoja))
        except:
            return "N/A"

    prima_c = contar_filas_hoja("Incidencias_CyberIdeas_*.xlsx", "Prima_Dominical")
    prima_w = contar_filas_hoja("Incidencias_WinCall_*.xlsx", "Prima_Dominical")

    def contar_filas_bono(patron):
        archivos = sorted(glob.glob(os.path.join(pruebas_dir, patron)), key=os.path.getmtime, reverse=True)
        if not archivos:
            return "N/A"
        try:
            return len(pd.read_excel(archivos[0]))
        except:
            return "N/A"

    bono_c = contar_filas_bono("Percepciones_Bono_asistencia_CyberIdeas_*.xlsx")
    bono_w = contar_filas_bono("Percepciones_Bono_asistencia_WinCall_*.xlsx")

    print(f"\n{'='*50}")
    print(f"AUTOMATIZACION COMPLETADA CON EXITO - Reportes en: {pruebas_dir}")
    print(f"  Faltas CyberIdeas:  {len(df_faltas_cyber)}")
    print(f"  Faltas WinCall:     {len(df_faltas_wincall)}")
    print(f"  Total Faltas:       {df_faltas_all.shape[0]}")
    print(f"  Prima Dom. Cyber:   {prima_c}")
    print(f"  Prima Dom. WinCall: {prima_w}")
    print(f"  Bono Cyber:         {bono_c}")
    print(f"  Bono WinCall:       {bono_w}")
    print(f"{'='*50}\n")
    
    return len(agents_records), df_faltas_all.shape[0]

if __name__ == "__main__":
    process_attendance_live()

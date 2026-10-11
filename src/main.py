"""Orquestador principal del proceso de automatizacion de reportes de asistencia y nomina."""
import os
import sys

# Asegurar que el directorio raíz del proyecto esté en sys.path para imports absolutos
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import shutil
import time
import datetime
import glob
import pandas as pd

from src.config import (
    SEMANA_DIR, TEMP_DIR, CREDS_PATH, SCOPES,
    SHEET_ID, GID_PRINCIPAL, GID_VTAS, GID_CYBER, GID_WINCALL,
    OUTPUT_PRINCIPAL, OUTPUT_VTAS, OUTPUT_CYBER_CONC, OUTPUT_WINCALL_CONC, OUTPUT_ATN_RAW
)
from src.services.google_drive import get_drive_headers, download_sheet_by_gid
from src.services.telegram import send_telegram_reports
from src.database.client import conectar_bd, extraer_empleados
from src.processors.dates import extraer_fechas_semana
from src.processors.asistencia import formatear_ventas, procesar_nomina_principal
from src.processors.faltas import extraer_faltas_concentrado, guardar_y_estilizar_faltas
from src.processors.prima_dominical import generar_prima_dominical
from src.processors.bonos import extraer_bonos_concentrado


def process_attendance_live():
    """Ejecuta el pipeline completo de generacion y distribucion de reportes de asistencia."""
    print("--- INICIANDO AUTOMATIZACION DE ASISTENCIA EN VIVO ---")
    os.makedirs(SEMANA_DIR, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)
    
    # 1. Autenticacion con Google Drive
    headers = get_drive_headers(CREDS_PATH, SCOPES)
    
    # 2. Descargar nomina principal y extraer fechas de la semana
    print("\nDescargando NOMINA POR SUPERVISOR IZZI ATN.xlsx directamente con GID...")
    if download_sheet_by_gid(SHEET_ID, GID_PRINCIPAL, headers, OUTPUT_PRINCIPAL):
        shutil.copy(OUTPUT_PRINCIPAL, OUTPUT_ATN_RAW)
        print("Descarga de nomina principal completa.")
    else:
        print("Error crítico: No se pudo descargar nómina principal después de varios intentos.")
        return None, None
        
    expected_dates_dmy, reported_monday, semana_str = extraer_fechas_semana(OUTPUT_ATN_RAW)
    
    # Rutas dinamicas de salida segun la semana
    output_incidencias_cyber = os.path.join(SEMANA_DIR, f"Incidencias_CyberIdeas_{semana_str}.xlsx")
    output_incidencias_wincall = os.path.join(SEMANA_DIR, f"Incidencias_WinCall_{semana_str}.xlsx")
    output_premios_cyber = os.path.join(SEMANA_DIR, f"Percepciones_Bono_asistencia_CyberIdeas_{semana_str}.xlsx")
    output_premios_wincall = os.path.join(SEMANA_DIR, f"Percepciones_Bono_asistencia_WinCall_{semana_str}.xlsx")
    
    # 3. Descargar y formatear IZZI VTAS CONCENTRADO
    print("\nDescargando y generando IZZI VTAS CONCENTRADO.xlsx...")
    if not download_sheet_by_gid(SHEET_ID, GID_VTAS, headers, OUTPUT_VTAS):
        print("Error crítico al descargar pestaña de ventas.")
        return None, None
    formatear_ventas(OUTPUT_VTAS, reported_monday)
    
    # 4. Descargar Concentrados Cyber y Wincall
    print("\nDescargando pestañas de Concentrados (Cyber y Wincall)...")
    time.sleep(2)
    if not download_sheet_by_gid(SHEET_ID, GID_CYBER, headers, OUTPUT_CYBER_CONC):
        print("Error crítico al descargar CYBERIDEAS CONCENTRADO")
    time.sleep(2)
    if not download_sheet_by_gid(SHEET_ID, GID_WINCALL, headers, OUTPUT_WINCALL_CONC):
        print("Error crítico al descargar WINCALL CONCENTRADO")
    print("Descarga de Concentrados completa.")
    
    # 5. Procesar asistencia principal (formulas, estilos y totales)
    agents_records, corrections_made = procesar_nomina_principal(OUTPUT_PRINCIPAL, expected_dates_dmy)
    
    # 6. Extraer Faltas e Incidencias
    print("\nIniciando extracción de concentrados para filtrar Faltas...")
    faltas_cyber_rows = extraer_faltas_concentrado(OUTPUT_CYBER_CONC, expected_dates_dmy)
    faltas_wincall_rows = extraer_faltas_concentrado(OUTPUT_WINCALL_CONC, expected_dates_dmy)
    print(f"  -> {len(faltas_cyber_rows)} incidencias en CyberIdeas")
    print(f"  -> {len(faltas_wincall_rows)} incidencias en WinCall")
    
    df_faltas_cyber = pd.DataFrame(faltas_cyber_rows)
    df_faltas_wincall = pd.DataFrame(faltas_wincall_rows)
    df_faltas_all = pd.concat([df_faltas_cyber, df_faltas_wincall], ignore_index=True)
    
    empty_cols = ['Clave trab', 'Tipo falta', '% falta', 'Certificado', '% pagado', 'Fecha inic', 'Fecha fin']
    if df_faltas_cyber.empty:
        df_faltas_cyber = pd.DataFrame(columns=empty_cols)
    if df_faltas_wincall.empty:
        df_faltas_wincall = pd.DataFrame(columns=empty_cols)
        
    print(f"Guardando archivos de Faltas en: {SEMANA_DIR}")
    guardar_y_estilizar_faltas(df_faltas_cyber, output_incidencias_cyber)
    guardar_y_estilizar_faltas(df_faltas_wincall, output_incidencias_wincall)
    
    # 7. Generar Prima Dominical (P19 - 25%) cruzando con BD
    print("\nIniciando extraccion de CyberHub para Percepciones (Prima Dominical)...")
    try:
        conn = conectar_bd()
        df_empleados = extraer_empleados(conn)
        conn.close()
        generar_prima_dominical(df_empleados, expected_dates_dmy, output_incidencias_cyber, output_incidencias_wincall, OUTPUT_CYBER_CONC, OUTPUT_WINCALL_CONC)
    except Exception as e:
        print(f"Error generando reporte de Percepciones: {e}")
        
    # 8. Generar Premios de Asistencia (Bonos P15 y P16)
    print("\nGenerando Premios de Asistencia (P15 y P16, BONO OK) desde concentrados...")
    try:
        df_bono_cyber = extraer_bonos_concentrado(OUTPUT_CYBER_CONC, "BONO", expected_dates_dmy)
        df_bono_cyber.to_excel(output_premios_cyber, index=False)
        print(f"  -> CyberIdeas: {len(df_bono_cyber)} agentes. Guardado en {output_premios_cyber}")
        
        df_bono_wincall = extraer_bonos_concentrado(OUTPUT_WINCALL_CONC, "BONO AUTORIZADO", expected_dates_dmy)
        df_bono_wincall.to_excel(output_premios_wincall, index=False)
        print(f"  -> WinCall: {len(df_bono_wincall)} agentes. Guardado en {output_premios_wincall}")
    except Exception as e:
        print(f"Error generando reporte de Premios: {e}")
        
    # 9. Envio a Telegram
    corrections_html = ""
    if corrections_made:
        corrections_html = "⚠️ <b>Correcciones de fecha aplicadas en el reporte:</b>\n"
        for corr in corrections_made:
            corrections_html += f"• {corr}\n"
        corrections_html += "\n"
        
    start_date = expected_dates_dmy[0]
    end_date = expected_dates_dmy[6]
    current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
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
        "📊 Reporte NOMINA POR SUPERVISOR IZZI ATN": OUTPUT_PRINCIPAL,
        "📈 Concentrado IZZI VTAS CONCENTRADO": OUTPUT_VTAS,
        "📂 CYBERIDEAS CONCENTRADO (Drive)": OUTPUT_CYBER_CONC,
        "📂 WINCALL CONCENTRADO (Drive)": OUTPUT_WINCALL_CONC,
        "📁 Incidencias CyberIdeas (Faltas + P19)": output_incidencias_cyber,
        "📁 Incidencias WinCall (Faltas + P19)": output_incidencias_wincall,
        "🏆 Percepciones Bono Asistencia CyberIdeas (P15)": output_premios_cyber,
        "🏆 Percepciones Bono Asistencia WinCall (P15)": output_premios_wincall
    }
    
    send_telegram_reports(summary_text, file_paths)
    print('\n*** REPORTES ENVIADOS A TELEGRAM EXITOSAMENTE ***\n')
    
    # 10. Metricas de ejecucion en consola
    def contar_filas_hoja(patron, hoja):
        archivos = sorted(glob.glob(os.path.join(SEMANA_DIR, patron)), key=os.path.getmtime, reverse=True)
        if not archivos:
            return "N/A"
        try:
            return len(pd.read_excel(archivos[0], sheet_name=hoja))
        except Exception:
            return "N/A"

    def contar_filas_bono(patron):
        archivos = sorted(glob.glob(os.path.join(SEMANA_DIR, patron)), key=os.path.getmtime, reverse=True)
        if not archivos:
            return "N/A"
        try:
            return len(pd.read_excel(archivos[0]))
        except Exception:
            return "N/A"

    prima_c = contar_filas_hoja("Incidencias_CyberIdeas_*.xlsx", "Prima_Dominical")
    prima_w = contar_filas_hoja("Incidencias_WinCall_*.xlsx", "Prima_Dominical")
    bono_c = contar_filas_bono("Percepciones_Bono_asistencia_CyberIdeas_*.xlsx")
    bono_w = contar_filas_bono("Percepciones_Bono_asistencia_WinCall_*.xlsx")

    print(f"\n{'='*50}")
    print(f"AUTOMATIZACION COMPLETADA CON EXITO - Reportes en: {SEMANA_DIR}")
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

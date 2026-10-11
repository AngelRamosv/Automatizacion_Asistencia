"""Procesador de extraccion y calculo de fechas de la semana."""
import datetime
import re
import openpyxl
from src.config import MESES


def extraer_fechas_semana(archivo_atn_raw):
    """
    Lee la fecha del lunes desde la celda I1 de NOMINA POR SUPERVISOR IZZI ATN,
    construye el vector de 7 fechas (expected_dates_dmy) y genera la etiqueta semana_str.
    """
    print("Extrayendo fechas inteligentemente desde NOMINA POR SUPERVISOR IZZI ATN...")
    wb_temp = openpyxl.load_workbook(archivo_atn_raw, data_only=True)
    ws_temp = wb_temp.active
    
    val_lunes = ws_temp.cell(row=1, column=9).value  # Columna I (lunes)
    
    if isinstance(val_lunes, (datetime.datetime, datetime.date)):
        monday_dt = val_lunes
    else:
        v_str = str(val_lunes).strip()
        if re.match(r'^\d{2}/\d{2}/\d{4}', v_str):
            v_str = v_str[:10]
        try:
            monday_dt = datetime.datetime.strptime(v_str, "%d/%m/%Y")
        except Exception:
            monday_dt = datetime.datetime.strptime(v_str, "%m/%d/%Y")
            
    expected_dates_dmy = []
    for i in range(7):
        day_dt = monday_dt + datetime.timedelta(days=i)
        expected_dates_dmy.append(day_dt.strftime("%d/%m/%Y"))
    wb_temp.close()
    
    print(f"Fechas leídas del Drive (semana actual de trabajo): {expected_dates_dmy}")
    
    sunday_dt = monday_dt + datetime.timedelta(days=6)
    if monday_dt.month == sunday_dt.month:
        semana_str = f"semana_{monday_dt.day}_al_{sunday_dt.day}_{MESES[sunday_dt.month]}_{sunday_dt.year}"
    else:
        semana_str = f"semana_{monday_dt.day}_{MESES[monday_dt.month]}_al_{sunday_dt.day}_{MESES[sunday_dt.month]}_{sunday_dt.year}"
        
    return expected_dates_dmy, monday_dt, semana_str

@echo off
:: ============================================================
:: Tarea programada - Se ejecuta automaticamente cada martes 12:00 PM
:: Registrada en: Programador de Tareas de Windows
:: ============================================================
cd /d "C:\Users\Agents\Desktop\Automatizacion_Asistencia"
echo [%date% %time%] ===== EJECUCION PROGRAMADA (MARTES) ===== >> log_ejecucion.log
python src\main.py >> log_ejecucion.log 2>&1
echo [%date% %time%] Ejecucion finalizada con codigo %errorlevel%. >> log_ejecucion.log
echo ---------------------------------------------------------- >> log_ejecucion.log

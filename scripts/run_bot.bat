@echo off
cd /d "C:\Users\Agents\Desktop\Automatizacion_Asistencia"
echo [%date% %time%] Iniciando ejecucion automatica... >> log_ejecucion.log
python src\main.py >> log_ejecucion.log 2>&1
echo [%date% %time%] Ejecucion terminada con codigo %errorlevel%. >> log_ejecucion.log
echo ---------------------------------------------------- >> log_ejecucion.log

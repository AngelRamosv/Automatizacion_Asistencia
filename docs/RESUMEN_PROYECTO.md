# Resumen de Proyecto: Automatización de Asistencia

## Estado Actual del Proyecto
Hemos desarrollado un script principal llamado `generar_reportes_live.py` que se encarga de automatizar la generación de reportes de asistencia y nómina. Este script realiza lo siguiente:

1. **Extracción de Datos:** Se conecta a Google Sheets para descargar la nómina del supervisor (`NOMINA POR SUPERVISOR IZZI ATN.xlsx`) y el concentrado de ventas (`IZZI VTAS CONCENTRADO.xlsx`).
2. **Cruce y Faltas:** Lee los códigos de asistencia diarios y determina a los agentes que tienen faltas reales (códigos 'F', 'FI', 'B', 'BAJA', 'INC', 'FJ'). Estos registros se consolidan y guardan en el archivo `Faltas.xlsx`.
3. **Cálculo de Percepciones y Prima Dominical:** Se conecta a la base de datos de CyberHub para extraer los empleados y percepciones/deducciones. Evalúa qué agentes asistieron el día domingo ('A' o 'R'). Quienes asistieron reciben una Prima Dominical equivalente al 25% de su Salario Diario. Todos los conceptos de nómina y esta prima se empaquetan en el archivo `Percepciones.xlsx`.
4. **Resumen de Asistencia (Conteo):** En el archivo de Nomina del supervisor, se inyectan columnas al final con el conteo automatizado semanal de: Asistencias (A/R), Faltas (F), Vacaciones (V), Incapacidades (INC), Faltas Justificadas (FJ), etc.
5. **Envío Automático a Telegram:** Finalmente, el bot empaqueta los archivos resultantes y los envía automáticamente al grupo de Telegram configurado con un pequeño resumen en texto.

## Archivos Generados
Actualmente el bot genera y envía los siguientes 4 archivos:
1. `NOMINA POR SUPERVISOR IZZI ATN.xlsx` (Nómina con cruce de asistencias)
2. `IZZI VTAS CONCENTRADO.xlsx` (Ventas con fechas actualizadas)
3. `Faltas.xlsx` (Archivo para Aspel NOI con las faltas de la semana)
4. `Percepciones.xlsx` (Archivo con Prima Dominical y Percepciones/Deducciones)

*Nota: El archivo de 'Trabajadores nuevos' ya no es generado por el script debido a que el director lo descarga manualmente de CyberHub.*

## Reglas de Negocio Claves
- Las Faltas mapeadas a NOI son las letras: F, FI, B, BAJA, INC, FJ.
- Se otorga Prima Dominical a todo agente cuyo registro del domingo sea 'A' (Asistencia) o 'R' (Retardo), respetando el marco legal laboral.
- El monto de la Prima Dominical corresponde al 25% del `Salario diario` extraído de la base de datos de CyberHub. El número de percepción asignado para este bono es el `19`.

# Especificación Funcional: Automatización de Reportes de Asistencia y Nómina

## 1. Propósito
Automatizar la extracción, consolidación, cálculo matemático y distribución de 8 reportes semanales de nómina e incidencias de asistencia para las empresas **CyberIdeas** y **WinCall**, asegurando su entrega puntual vía Telegram y el registro de incidencias en formato compatible con el sistema NOI.

---

## 2. Insumos y Fuentes de Datos

### 2.1 Google Drive (Google Sheets)
- **ID Hoja Maestra:** `1HK7mp9l39vXPgyMisrDZVnY3jc-BaDzf43fMrjPUzQU`
- **Pestañas y GIDs:**
  - `NOMINA POR SUPERVISOR IZZI ATN` (GID: `595741460`): Registro de asistencia diaria de agentes de atención.
  - `IZZI VTAS CONCENTRADO` (GID: `187319239`): Registro de asistencia de ventas.
  - `CYBERIDEAS CONCENTRADO` (GID: `0`): Concentrado operativo de CyberIdeas.
  - `WINCALL CONCENTRADO` (GID: `99266008`): Concentrado operativo de WinCall.

### 2.2 Base de Datos Operativa (MySQL CyberHub)
- **Host:** `192.168.50.33` (cyber_ideas_hub)
- **Tablas:**
  - `personal`: Consulta de empleados activos (`personal_estatus = 1`), salarios diarios y datos fiscales/personales.
  - `config_percep_deduc`: Configuración de percepciones y deducciones indefinidas.

---

## 3. Reglas de Negocio y Lógica de Procesamiento

### 3.1 Detección Dinámica de Fechas
1. La fecha de inicio de semana se lee de la celda **I1** (columna 9) de `NOMINA POR SUPERVISOR IZZI ATN`.
2. Se genera el vector de 7 días: `[Lunes, Martes, Miércoles, Jueves, Viernes, Sábado, Domingo]`.
3. Se genera la etiqueta de semana (ejemplo: `semana_8_al_14_octubre_2026`).

### 3.2 Reporte IZZI VTAS CONCENTRADO
- Se colocan fechas estáticas en las columnas `['I', 'K', 'M', 'O', 'Q', 'S', 'U']` en la fila 1 con formato `dd/mm/yyyy` y ancho de columna 12.0 para evitar errores de renderizado `###`.

### 3.3 Reporte Principal NOMINA POR SUPERVISOR IZZI ATN
- Se corrigen encabezados de fecha en fila 1 si difieren del cálculo matemático.
- Se inyectan columnas de conteo a partir de la columna V (22):
  - `Asistencias (A)`: `=COUNTIF(H{r}:T{r}, "A") + COUNTIF(H{r}:T{r}, "R")`
  - `Descansos (D)`: `=COUNTIF(H{r}:T{r}, "D")`
  - `Faltas (F)`: `=COUNTIF(H{r}:T{r}, "F")`
  - `Vacaciones (V)`: `=COUNTIF(H{r}:T{r}, "V")`
  - `Faltas Justificadas (FJ)`: `=COUNTIF(H{r}:T{r}, "FJ")`
  - `Incapacidades (INC)`: `=COUNTIF(H{r}:T{r}, "INC")`
  - `Suspensiones (SUSP)`: `=COUNTIF(H{r}:T{r}, "SUSP")`
- Fila inferior con `=SUM(...)` para cada código con bordes dobles y fondo `#D9E1F2`.
- Formato condicional de celdas por tipo de código.

### 3.4 Reportes de Incidencias / Faltas
- **Fuente:** `CYBERIDEAS CONCENTRADO` y `WINCALL CONCENTRADO`.
- **Columnas de Asistencia:** 17, 19, 21, 23, 25, 27, 29 (Q, S, U, W, Y, AA, AC).
- **Códigos de Falta considerados:** `{'F', 'FI', 'B', 'BAJA', 'INC', 'FJ'}`.
- Cada incidencia genera un registro:
  - `Tipo falta`: `2`
  - `% falta`: `100`
  - `Certificado`: `0`
  - `% pagado`: `0`
- **Regla Especial de Retardos:** 
  - Si la columna 32 (retardos) es exactamente igual a **3**, se añade **1 falta adicional** asignada al primer día libre (sin falta previa) de la semana.

### 3.5 Prima Dominical (P19)
- **Fuente:** Concentrados de ambas empresas.
- **Criterio:** Agentes con código `'A'` o `'R'` en el día domingo (Columna 29 / AC).
- **Cálculo:** $\text{Monto} = \text{round}(\text{Salario Diario} \times 0.25, 2)$ obtenido de MySQL.
- **Estructura del movimiento NOI:**
  - `Per/Ded`: `'P'`
  - `Num Per/Ded`: `19`
  - `Num credito`: `0`
  - `Aplica Destaj`: `'N'`
  - `Aplicacion`: `4`
  - `Fecha inicio` y `Fecha fin`: Fecha del domingo.
- Se inyecta como hoja `Prima_Dominical` en los archivos de Incidencias correspondientes.

### 3.6 Premios de Asistencia (Bonos P15 y P16)
- **CyberIdeas:** Columna `"BONO"`. Si el valor es `"OK"`, genera:
  - Concepto P15 ($115.40).
  - Concepto P16 ($115.40).
- **WinCall:** Columna `"BONO AUTORIZADO"`. Si el valor es `"OK"`, genera:
  - Concepto P15 ($115.40).
  - Concepto P16 ($115.40).
- Se genera archivo independiente para cada empresa: `Percepciones_Bono_asistencia_{Empresa}_{semana}.xlsx`.

---

## 4. Distribución y Salidas

### 4.1 Archivos Generados (Carpeta `semana/`)
1. `NOMINA POR SUPERVISOR IZZI ATN.xlsx`
2. `IZZI VTAS CONCENTRADO.xlsx`
3. `CYBERIDEAS CONCENTRADO.xlsx`
4. `WINCALL CONCENTRADO.xlsx`
5. `Incidencias_CyberIdeas_{semana}.xlsx` (Hojas: Faltas, Prima_Dominical)
6. `Incidencias_WinCall_{semana}.xlsx` (Hojas: Faltas, Prima_Dominical)
7. `Percepciones_Bono_asistencia_CyberIdeas_{semana}.xlsx`
8. `Percepciones_Bono_asistencia_WinCall_{semana}.xlsx`

### 4.2 Notificación Telegram
- Mensaje formateado en HTML con resumen de total de agentes, total de faltas y correcciones detectadas.
- Envío de los 8 documentos adjuntos.

"""Cliente de conexion y consultas a la base de datos MySQL (CyberHub)."""
import mysql.connector
import pandas as pd
from src.config import DB_CONFIG


def conectar_bd(config=None):
    """Establece conexion con la base de datos MySQL CyberHub."""
    cfg = config or DB_CONFIG
    return mysql.connector.connect(
        host=cfg['host'],
        user=cfg['user'],
        password=cfg['password'],
        database=cfg['database']
    )


def extraer_empleados(conn):
    """Extrae todos los empleados activos (personal_estatus = 1) con sus salarios diarios."""
    query = (
        "SELECT personal_id AS 'ID', personal_id AS 'Clave trabajador', "
        "personal_apaterno AS 'A. paterno', personal_amaterno AS 'A. materno', "
        "personal_nombre AS 'Nombre (s)', personal_sueldo_diario AS 'Salario diario', "
        "personal_fecha_nacimiento AS 'Fecha de nacimiento', personal_ingreso_sa AS 'Fecha de alta', "
        "personal_imss AS 'IMSS', personal_curp AS 'CURP', personal_rfc AS 'RFC', "
        "personal_celular AS 'Teléfono 1', personal_razonSocial AS 'Razon Social' "
        "FROM personal WHERE personal_estatus = 1"
    )
    return pd.read_sql(query, conn)


def extraer_percepciones_deducciones(conn):
    """Extrae configuracion de percepciones y deducciones indefinidas."""
    query = "SELECT c.personal_id, c.tipo, c.p_d, c.valor FROM config_percep_deduc c WHERE c.estatus = 'INDEFINIDO'"
    return pd.read_sql(query, conn)


def generar_trabajadores(df_empleados, output_path):
    """Genera archivo Trabajadores.xlsx para altas recientes (ultimos 7 dias)."""
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

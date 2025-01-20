from common.__init__ import *
from querys import *
from db import inciar_db
# Consulta para obtener tipos_elementos_del_usuario_id de todos los usuarios

def generar_data():
    db_path = inciar_db()
    # Conectar a la base de datos SQLite
    conn = sqlite3.connect(db_path)

    # Obtener datos de las tablas y almacenarlos en DataFrames de pandas
    esd_items_df = pd.read_sql_query("SELECT * FROM esd_items", conn)
    personal_esd_df = pd.read_sql_query("SELECT * FROM personal_esd", conn)
    actividades_df = pd.read_sql_query("SELECT * FROM actividades", conn)
    usuarios_elementos_df = pd.read_sql_query("SELECT * FROM usuarios_elementos", conn)
    evidencias_asignacion_df = pd.read_sql_query("SELECT * FROM evidencias_asignacion", conn)
    registro_actividades_df = pd.read_sql_query("SELECT * FROM registro_actividades", conn)
    actividades_registradas_df = pd.read_sql_query("SELECT * FROM actividades_registradas", conn)

    elementos_usuarios_df = pd.read_sql_query(users_estatus, conn)

    # Obtener datos de tipos_elementos_del_usuario_id y almacenarlos en un DataFrame de pandas
    tipos_elementos_del_usuario_id_df = pd.read_sql_query(all_esd_users, conn)

    # Cerrar la conexión a la base de datos
    conn.close()

    # Lista de DataFrames y sus nombres de tabla correspondientes
    dataframes = [
        (esd_items_df, 'ESD Items'),
        (personal_esd_df, 'Personal ESD'),
        (actividades_df, 'Actividades'),
        (usuarios_elementos_df, 'Usuarios Elementos'),
        (evidencias_asignacion_df, 'Evidencias Asignación'),
        (registro_actividades_df, 'Registro Actividades'),
        (actividades_registradas_df, 'Actividades Registradas'),
        (elementos_usuarios_df, 'Usuarios Elementos Todos'),
        (tipos_elementos_del_usuario_id_df, 'Tipos Elementos del Usuario ID')
    ]

    # Agregar una columna 'tabla_origen' en cada DataFrame
    for df, nombre_tabla in dataframes:
        df['tabla_origen'] = nombre_tabla

    # Combinar todos los DataFrames en uno solo
    df_unificado = pd.concat([df for df, _ in dataframes], ignore_index=True)

    # Limpiar y formatear la columna 'fecha_maestra', si existe
    if 'fecha_maestra' in df_unificado.columns:
        df_unificado['fecha_maestra'] = pd.to_datetime(df_unificado['fecha_maestra'], errors='coerce').dt.strftime('%d-%m-%Y')

    # Nombre del archivo fijo
    file_name = 'db_datos_excel.xlsx'

    reporte_c_path = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\data_new_report'

    # Ruta completa para guardar el archivo (sobreescribir si ya existe)
    save_path = os.path.join(reporte_c_path, file_name)

    # Guardar el DataFrame unificado en el archivo Excel
    guardar_excel(save_path, df_unificado)

def guardar_excel(save_path, df_unificado):
    try:
        # Guardar DataFrame en Excel (sobreescribir si ya existe)
        df_unificado.to_excel(save_path, index=False)

        print(f'El archivo ha sido guardado correctamente en: {save_path}')
    except Exception as e:
        print(f'Ocurrió un error al guardar el archivo: {e}')

if __name__ == "__main__":
    try:
        generar_data()
    except Exception as e:
        print(f'Ocurrió un error al ejecutar el script: {e}')
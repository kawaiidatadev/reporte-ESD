from common import *

def inciar_db():
    db_patch = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\database.sqlite'
    copias = r'\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Data\copias de seguridad'
    """
    Copia el archivo database.sqlite, y lo pega en la direccion de copias de seguridad, cambiandole el nombre a:
    {"ESD_BC_copia de seguridad_"dd-mm-aaaa_hh-mm-ss_usuario_windows}
    """

    # Obtener fecha y hora actual
    fecha_hora_actual = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

    # Obtener nombre del usuario de Windows
    usuario_windows = os.getlogin()

    # Crear el nuevo nombre para la copia de seguridad
    nombre_copia = f"ESD_BC_copia_de_seguridad_{fecha_hora_actual}_{usuario_windows}.sqlite"

    # Ruta completa para la copia de seguridad
    destino_copia = os.path.join(copias, nombre_copia)

    try:
        # Copiar el archivo y renombrarlo
        shutil.copy(db_patch, destino_copia)
        print(f"Copia de seguridad creada: {destino_copia}")
    except Exception as e:
        print(f"Error al copiar la base de datos: {e}")

    return db_patch

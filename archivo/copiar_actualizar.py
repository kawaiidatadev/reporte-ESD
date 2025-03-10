import shutil
import os
import time
import traceback
from datetime import datetime
import win32com.client as win32
import ctypes

def show_error(message):
    """ Muestra un mensaje de error al usuario. """
    ctypes.windll.user32.MessageBoxW(0, message, "Error", 0)

def kill_excel():
    """ Cierra todas las instancias de Excel abiertas. """
    # os.system("taskkill /f /im excel.exe")

def check_permissions(path):
    """ Verifica si el usuario tiene permisos de acceso al archivo o directorio. """
    if not os.path.exists(path):
        show_error(f"La ruta no existe: {path}")
        return False
    if not os.access(path, os.R_OK):
        show_error(f"No tienes permisos de lectura en: {path}")
        return False
    return True

def copy_file(source, destination):
    """ Copia el archivo fuente al destino. """
    try:
        shutil.copy2(source, destination)
        return True
    except Exception as e:
        print(f"Error al copiar el archivo: {e}")
        show_error(f"Error al copiar el archivo: {e}")
        return False

def main():
    try:
        # Cerrar instancias de Excel antes de iniciar
        kill_excel()

        source_path = r"\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\data_new_report\ReporteETL\ReporteESD.xlsx"
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        dest_filename = f"Informe ESD ({datetime.now().strftime('%Y-%m-%d - %H-%M-%S')}).xlsx"
        dest_path = os.path.join(downloads_folder, dest_filename)

        # Verificar permisos de acceso
        if not check_permissions(source_path) or not check_permissions(downloads_folder):
            return

        if not copy_file(source_path, dest_path):
            return

        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False  # Desactiva alertas emergentes
        workbook = excel.Workbooks.Open(dest_path)
        excel.Visible = False  # Mantener Excel en segundo plano

        # Mostrar todas las hojas
        for sheet in workbook.Sheets:
            sheet.Visible = -1

        # Actualizar consultas y tablas dinámicas
        try:
            workbook.RefreshAll()
            time.sleep(5)  # Esperar para evitar conflictos

            for sheet in workbook.Sheets:
                try:
                    for pivot_table in sheet.PivotTables():
                        pivot_table.RefreshTable()
                except Exception as e:
                    print(f"Error al actualizar tabla dinámica en hoja '{sheet.Name}': {e}")
        except Exception as e:
            if "-2147418111" in str(e):  # Ignorar este error específico
                print(f"[IGNORADO] Error al actualizar consultas/tablas dinámicas: {e}")
            else:
                show_error(f"Error al actualizar consultas/tablas dinámicas: {e}")

        # Ocultar todas las hojas excepto "Informe" y "Reporte de mediciones semestral"
        for sheet in workbook.Sheets:
            if sheet.Name not in ["Informe", "Reporte de mediciones semestral"]:
                sheet.Visible = 0

        # Guardar y cerrar
        workbook.Save()
        workbook.Close()
        excel.Quit()

        os.startfile(downloads_folder)  # Abrir carpeta de descargas

    except Exception:
        error_message = f"Error durante el proceso:\n{traceback.format_exc()}"
        show_error(error_message)

    finally:
        if 'excel' in locals():
            try:
                excel.DisplayAlerts = True  # Restaurar alertas
                excel.Quit()
            except Exception as e:
                print(f"Error al cerrar Excel: {e}")

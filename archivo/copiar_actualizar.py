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


def copy_file(source, destination):
    """ Copia el archivo fuente al destino. """
    try:
        shutil.copy2(source, destination)
        return True
    except Exception as e:
        show_error(f"Error al copiar el archivo: {e}")
        return False


def main():
    try:
        source_path = r"\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\data_new_report\ReporteETL\ReporteESD.xlsx"
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        dest_filename = f"Informe ESD ({datetime.now().strftime('%Y-%m-%d - %H-%M-%S')}).xlsx"
        dest_path = os.path.join(downloads_folder, dest_filename)

        if not copy_file(source_path, dest_path):
            return

        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False  # Desactiva alertas emergentes
        workbook = excel.Workbooks.Open(dest_path)
        excel.Visible = False  # Mantener Excel en segundo plano

        # Mostrar todas las hojas
        for sheet in workbook.Sheets:
            sheet.Visible = -1

        # Actualizar consultas y tablas dinámicas mínimo 1 veces
        for _ in range(1):
            workbook.RefreshAll()
            time.sleep(3)  # Esperar para evitar conflictos

            for sheet in workbook.Sheets:
                try:
                    for pivot_table in sheet.PivotTables():
                        pivot_table.RefreshTable()
                except Exception:
                    continue  # Ignorar errores en hojas sin tablas dinámicas

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
        show_error(f"Error durante el proceso:\n{traceback.format_exc()}")
    finally:
        if 'excel' in locals():
            try:
                excel.DisplayAlerts = True  # Restaurar alertas
                excel.Quit()
            except:
                pass



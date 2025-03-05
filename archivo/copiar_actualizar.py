"""
Copia el archivo de:
\\mercury\Mtto_Prod\00_Departamento_Mantenimiento\ESD\Software\Recurses\data_new_report\ReporteETL\ReporteESD.xlsx

A la carpeta de descargas del usuario de windows actual.nombrendolo como
Informe ESD (AAAA-MMMM-DD) - (HH-MM-SS)
LUEGRO LO ABRIRA (SIN MOSTRARLE AL USUARIO)
para mostrar todas las hojas ocultas, luego parasar por cada una de ellas actualizando todo minimo 3 veces.
Luego ocultar todas las hojas menos la de "Informe"
Para luego actualizar todas las tablas dinamiscas que tenga 3 veces cada una.
luego actualizar 3 veces las consultas.
Luego actualizar todo lo restante por actualizar.
y ahora si abrir la carpeta de descargas del usuario señalando el archivo.

en cualquier caso de error, en alguna parte de este proceso, dar un maneje al usuario por medio de msgbx

"""

import shutil
import os
import time
import traceback
from datetime import datetime
import win32com.client as win32
import ctypes


def show_error(message):
    ctypes.windll.user32.MessageBoxW(0, message, "Error", 0)


def copy_file(source, destination):
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

        # Actualizar consultas y tablas dinámicas
        excel.ActiveWorkbook.RefreshAll()
        time.sleep(5)  # Esperar para evitar conflictos

        for sheet in workbook.Sheets:
            try:
                for pivot_table in sheet.PivotTables():
                    pivot_table.RefreshTable()
            except Exception:
                continue  # Ignorar errores en hojas sin tablas dinámicas

        # Ocultar todas las hojas excepto "Informe"
        for sheet in workbook.Sheets:
            if sheet.Name != "Informe":
                sheet.Visible = 0

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


if __name__ == "__main__":
    main()
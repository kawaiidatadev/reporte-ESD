import sys
import os
import time
from concurrent.futures import ProcessPoolExecutor
import psutil  # Manejo de procesos

from common.__init__ import *
from obtener_data import generar_data
from archivo.copiar_actualizar import main


def set_high_priority():
    """ Aumenta la prioridad del proceso de forma segura. """
    try:
        p = psutil.Process(os.getpid())
        p.nice(psutil.HIGH_PRIORITY_CLASS)  # Prioridad alta pero no extrema
        print("Prioridad del proceso aumentada a HIGH.")
    except Exception as e:
        print(f"No se pudo cambiar la prioridad: {e}")


def run_tasks():
    """ Ejecuta las tareas en paralelo. """
    start_time = time.time()

    # Ajustar el número de procesos (dejar un núcleo libre para el sistema)
    num_workers = max(1, os.cpu_count() - 1)
    print(f"Usando {num_workers} trabajadores.")

    # Ejecutar tareas en paralelo
    with ProcessPoolExecutor(max_workers=num_workers) as executor:
        # Lanzar las tareas asíncronamente
        future_data = executor.submit(generar_data)
        future_main = executor.submit(main)

        # Esperar a que ambas tareas terminen
        future_data.result()  # Bloquea hasta que generar_data termine
        future_main.result()  # Bloquea hasta que main termine

    print(f"Tiempo total: {time.time() - start_time:.2f} segundos.")


if __name__ == '__main__':
    try:
        set_high_priority()  # Aumentar prioridad del proceso
        run_tasks()  # Ejecutar tareas
        sys.exit(0)  # Salida exitosa
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)  # Salida con error
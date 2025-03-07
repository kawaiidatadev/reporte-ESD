import sys
from common.__init__ import *
from obtener_data import generar_data
from archivo.copiar_actualizar import main

if __name__ == '__main__':
    try:
        generar_data()  # Copia de seguridad de la base de datos
        main()  # Copia y actualización del archivo
        sys.exit(0)  # Salida exitosa
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)  # Salida con error

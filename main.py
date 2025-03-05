from common.__init__ import *
from obtener_data import generar_data
from archivo.copiar_actualizar import main



if __name__ == '__main__':
    generar_data()  # Copia de seguridad de db
    main()  # Da el archivo
    sys.exit(0)
    sys.exit(1)
    sys.exit()





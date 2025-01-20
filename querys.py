from common import *

# Consulta de todos los usuarios, y sus estatus
users_estatus = """
 SELECT 
        nombre_usuario AS "Nombre Completo", 
        area AS "Área", 
        linea AS "Línea", 
        rol AS "Rol", 
        estatus_usuario AS "Estatus", 
        bata_estatus AS "Bata Estatus",
        bata_polar_estatus AS "Bata Polar Estatus",
        pulsera_estatus AS "Pulsera Estatus",
        talonera_estatus AS "Talonera Estatus"
    FROM 
        personal_esd
    WHERE
        estatus_usuario != '0'
"""


# Consulta todos los usuarios y sus elementos ESD
all_esd_users = """
SELECT
    esd_items.numero_serie, 
    personal_esd.nombre_usuario, 
    personal_esd.estatus_usuario,  -- Agregando el estatus_usuario
    esd_items.tipo_elemento, 
	esd_items.tamaño, -- Agregando el tamaño
    personal_esd.area, 
    personal_esd.linea, 
    esd_items.comentarios, 
    esd_items.fecha_maestra
FROM
    usuarios_elementos
JOIN
    esd_items
ON
    usuarios_elementos.esd_item_id = esd_items.id
JOIN
    personal_esd
ON
    usuarios_elementos.usuario_id = personal_esd.id
WHERE
    LOWER(esd_items.tipo_elemento) != 1
ORDER BY
    personal_esd.nombre_usuario;

"""
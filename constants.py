# -*- coding: utf-8 -*-

TIPO_DOC_DNI_STR = 'DNI'

# Cantidad máxima de filas a procesar. Como la librería no distingue las filas con
# datos de las vacías, si no se define un límite se genera una salida con cientos
# de miles de filas.
INPUT_SPREADSHEET_MAX_ROWS = 2000

REQUIRED_COLUMNS = [
    'codigo_refes',
    'nro_documento',
    'apellido',
    'fecha_resultado',
    'codigo_resultado',
]

OUTPUT_CSV_COLUMNS = [
    'codigo_refes',
    'tipo_documento',
    'nro_documento',
    'apellido',
    'nombre',
    'fecha_nac',
    'sexo',
    'fecha_procesamiento',
    'fecha_resultado',
    'codigo_localidad',
    'codigo_resultado',
    'telefono',
    'tipo_muestra',
    'pais',
    'provincia',
    'laboratorio',
    'partido',
    'calle',
    'altura',
    'tel_particular',
    'tel_laboral',
    'tel_celular',
]
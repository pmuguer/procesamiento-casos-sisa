# -*- coding: utf-8 -*-
import os

from procesar_entrada_excel import generate_columns_dict
from generar_salida_csv import generate_csv_casos, generate_csv_empadronamiento

if __name__ == '__main__':
    input_filename = raw_input('Ingrese la ruta completa del archivo Excel: ')

    basename = os.path.basename(input_filename)
    # Quito la extensión .xlsx
    basename_no_extension = os.path.splitext(basename)[0]
    # Reemplazo espacios por underscores en el nombre del archivo
    underscore_basename = basename_no_extension.replace(' ', '_')
    casos_basename = '{}.csv'.format(underscore_basename)
    empadronamiento_basename = 'empadronamiento_{}'.format(casos_basename)

    output_filename_casos = os.path.join(os.sep, 'tmp', casos_basename)
    output_filename_empadronamiento = os.path.join(os.sep, 'tmp',
                                                   empadronamiento_basename)

    columns_dict = generate_columns_dict(input_filename)
    generate_csv_casos(columns_dict, output_filename_casos)
    generate_csv_empadronamiento(columns_dict, output_filename_empadronamiento)

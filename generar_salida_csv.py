# -*- coding: utf-8 -*-
import csv

from constants import OUTPUT_CSV_COLUMNS, TIPO_DOC_DNI_STR, INPUT_SPREADSHEET_MAX_ROWS


def stringify_cell(cell):
    """Convertir el valor de una celda openpyxl a un string valido para el archivo CSV"""
    if not cell.value:
        return ''
    if cell.is_date:
        return cell.value.strftime('%Y-%m-%d')
    else:
        return unicode(cell.value).encode('utf-8')


def is_empty(row):
    """Chequea si una fila no tiene datos"""
    values = row.values()
    # Si todos los strings son vacíos, la fila no tiene datos
    if not ''.join(values):
        return True
    else:
        return False


def generate_csv_casos(columns_dict, output_filename):
    """Recibe un diccionario nombre col -> column (openpyxl) y genera el CSV con los datos de los casos."""
    with open(output_filename, "w") as output_file:
        csv_dict_writer = csv.DictWriter(output_file, fieldnames=OUTPUT_CSV_COLUMNS)
        csv_dict_writer.writeheader()

        for i in range(INPUT_SPREADSHEET_MAX_ROWS):
            row = {}
            for column_name in OUTPUT_CSV_COLUMNS:
                openpyxl_column = columns_dict.get(column_name)
                if openpyxl_column:
                    openpyxl_cell = openpyxl_column[i]
                    row[column_name] = stringify_cell(openpyxl_cell)

            # Si se encuentra una fila sin datos se considera que ya no hay
            # más filas para procesar
            if is_empty(row):
                break
            if 'tipo_documento' not in row.keys():
                row['tipo_documento'] = TIPO_DOC_DNI_STR
            csv_dict_writer.writerow(row)


def generate_csv_empadronamiento(columns_dict, output_filename):
    """Recibe un diccionario nombre col -> column (openpyxl) y genera el CSV con los datos para empadronar."""
    with open(output_filename, "w") as output_file:
        csv_dict_writer = csv.DictWriter(output_file, fieldnames=['documento_numero', 'sexo'])
        csv_dict_writer.writeheader()

        for i in range(INPUT_SPREADSHEET_MAX_ROWS):
            # Obtengo los datos completos de la fila porque hay filas con el documento
            # vacío que, en caso de que se procese sólo ese dato, se interpretan erróneamente
            # como el fin del archivo Excel
            row = {}
            for column_name in OUTPUT_CSV_COLUMNS:
                openpyxl_column = columns_dict.get(column_name)
                if openpyxl_column:
                    openpyxl_cell = openpyxl_column[i]
                    row[column_name] = stringify_cell(openpyxl_cell)

            # Si se encuentra una fila sin datos se considera que ya no hay
            # más filas para procesar
            if is_empty(row):
                break

            output_row = {}
            documento = columns_dict.get('nro_documento')
            if documento:
                output_row['documento_numero'] = stringify_cell(documento[i])
                output_row['sexo'] = ''

            csv_dict_writer.writerow(output_row)

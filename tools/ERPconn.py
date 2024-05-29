import os
import psycopg2
import openpyxl
from datetime import datetime

def importar_archivos_excel_en_carpeta(carpeta):
    """
    Importa los datos de archivos Excel en una carpeta a una base de datos PostgreSQL.

    Parameters:
        carpeta (str): La ruta de la carpeta que contiene los archivos Excel.

    Returns:
        None
    """

    date = datetime.now()
    dia = date.strftime('%d-%m-%Y')

    # Recorre todos los archivos en la carpeta
    for archivo in os.listdir(carpeta):
        # Verifica si el archivo es un archivo Excel
        if archivo.endswith('.xlsx'):
            # Construye la ruta completa del archivo
            excel_file = os.path.join(carpeta, archivo)

            # Abre el archivo Excel y realiza las operaciones que deseas
            # Abre el Excel workbook y carga la hoja activa en una variable
            workbook = openpyxl.load_workbook(excel_file)
            sheet = workbook.active

            # Crea una lista con los nombres de las columnas en la primera fila del libro de trabajo
            column_names = [column.value for column in sheet[1]]

            # Crea una lista vacía para almacenar los datos
            data = []
            # Itera sobre las filas y agrega los datos a la lista
            for row in sheet.iter_rows(min_row=2, values_only=True):
                data.append(row)

            # Define el esquema y la tabla de PostgreSQL donde se almacenarán los datos
            schema_name = 'test3'
            table_name = 'test_3'

            # Ejecuta el resto del código que maneja la conexión a la base de datos y la inserción de datos
            connection = psycopg2.connect(
                database='postgres',
                user='postgres',
                password='Aa123456',
                host='localhost',
                port='5432'
            )
            cursor = connection.cursor()

            schema_creation_query = f'CREATE SCHEMA IF NOT EXISTS {schema_name}'
            cursor.execute(schema_creation_query)

            table_creation_query = f"""
                CREATE TABLE IF NOT EXISTS {schema_name}.{table_name} (
                    "Documento EIPSA" TEXT PRIMARY KEY,
                    {", ".join([f'"{name}" TEXT' for name in column_names if name != 'Documento EIPSA'])}
                )
            """
            cursor.execute(table_creation_query)

            insert_update_data_query = f"""
                INSERT INTO {schema_name}.{table_name} ({", ".join([f'"{name}"' for name in column_names])})
                VALUES ({", ".join(['%s' for _ in column_names])})
                ON CONFLICT ("Documento EIPSA")
                DO UPDATE SET {", ".join([f'"{name}"=EXCLUDED."{name}"' for name in column_names if name != 'Documento EIPSA'])}
            """
            cursor.executemany(insert_update_data_query, data)

            connection.commit()

            cursor.close()
            connection.close()

            print(f'Importación de {excel_file} completada exitosamente!')
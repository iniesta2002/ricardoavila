# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 08:50:30 2023

@author: AvilaRiM
"""

import time
import pyodbc
#import serial

# Configuración de la conexión a la base de datos
connection_string = (
    r"DRIVER={SQL Server};"
    r"Server=CMXMOA17\l40SQLQAENV,1432;"
    r"Database=CMXBDataCollection;"
)
# Configuración del puerto serial (ajusta el puerto COM según tu configuración)
#serial_port = serial.Serial('COM4', 9600)  

# Array para almacenar los últimos 10 valores de LightsAscii
data_arrayopen = []
data_arrayclose = []

try:
    # Conexión a la base de datos
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    while True:
        cursor.execute(
            f'SELECT TOP 1 Department, Line, Area, EndAttend, idTransaction, ID FROM AndonTabletOutput ORDER BY ID DESC'
        )
        ultimo_dato_consulta1 = cursor.fetchone()

        cursor.execute(
            f'SELECT TOP 1 Department, Line, Area, EndAttend, idTransaction, ID FROM AndonTabletOutput WHERE EndAttend IS NOT NULL ORDER BY EndAttend DESC'
        )
        ultimo_dato_consulta2 = cursor.fetchone()

        if ultimo_dato_consulta1:
            department = ultimo_dato_consulta1.Department
            if department == "Mantenimiento":
                department = "MTO"
            elif department == "Manufactura":
                department = "MFG"
            elif department == "Calidad":
                department = "CAL"
            elif department == "Supply Chain":
                department = "SUP"
            elif department == "Ingenieria":
                department = "ING"
            elif department == "Facilities":
                department = "FAC"
            line = ultimo_dato_consulta1.Line
            if line == "39M+":
                line = "39W"

            area = ultimo_dato_consulta1.Area
            end_attend = ultimo_dato_consulta1.EndAttend
            idt = ultimo_dato_consulta1.ID
            idtt = ultimo_dato_consulta1.idTransaction

        if ultimo_dato_consulta2:
            department2 = ultimo_dato_consulta2.Department
            if department2 == "Mantenimiento":
                department2 = "REL"
            elif department2 == "Manufactura":
                department2 = "REL"
            elif department2 == "Calidad":
                department2 = "REL"
            elif department2 == "Supply Chain":
                department2 = "REL"
            elif department2 == "Ingenieria":
                department2 = "REL"
            elif department2 == "Facilities":
                department2 = "REL"
            line2 = ultimo_dato_consulta2.Line
            if line2 == "39M+":
                line2 = "39W"

            area2 = ultimo_dato_consulta2.Area
            end_attend2 = ultimo_dato_consulta2.EndAttend
            idt2 = ultimo_dato_consulta2.ID
            idtt2 = ultimo_dato_consulta2.idTransaction

            if idtt not in data_arrayopen:
                data_arrayopen.append(idtt)

                resultserial = "".join(map(str, department[:3]))+ "".join(map(str, line[:3])) + str(idt)
                print(f'{resultserial}')
            if idtt2 not in data_arrayclose:
                data_arrayclose.append(idtt2)
                
                resultserial2 = "".join(map(str, department2[:3]))+ "".join(map(str, line2[:3])) + str(idt2)
                print(f'{resultserial2}')

                #AQUI VA LO DE SERIAL
                #serial_port.write(f'{resultserial}\n'.encode())  # Agregar un salto de línea
               
        time.sleep(1)

except Exception as e:
    print("error en la conexión:", e)
finally:
    # Cerrar la conexión a la base de datos
        if cursor:
            cursor.close()
        if connection:
            connection.close()
#Graficos a partir de los dataframes

import openpyxl
import os
import csv
import traceback
from datetime import datetime
from pywinauto import Application, findwindows
import time
import threading
from novawinmng import manejar_novawin, leer_csv_y_crear_dataframe,agregar_csv_a_plantilla_excel, guardar_dataframe_en_ini,generar_nombre_unico,agregar_dataframe_a_excel_sin_borrar,agregar_dataframe_a_nueva_hoja
from pywinauto.keyboard import send_keys
from openpyxl import Workbook
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook
from pandas import ExcelWriter
from openpyxl import load_workbook
from pandas import ExcelWriter
import subprocess
from openpyxl.utils.dataframe import dataframe_to_rows

def agregar_dataframe_a_nueva_hoja(archivo_excel, dataframe, nombre_hoja):
    # Cargar el archivo Excel existente
    book = load_workbook(archivo_excel)
    
    # Usar el modo 'append' para evitar sobrescribir el archivo
    with pd.ExcelWriter(archivo_excel, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        # Pasar el objeto book al escritor
        writer._book = book  # Nota: Usamos '_book' en lugar de 'book'
        
        # Escribir el DataFrame en la nueva hoja
        dataframe.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
def BET_BI(df, ruta_excel, Rango_de_Absorpcion, Rango_de_Desorpcion):

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    try:
        # Convertir Rango_de_Desorpcion a entero y filtrar
        num_filas = int(Rango_de_Desorpcion)
        ultimo_rango_df = df.tail(num_filas)
        print("Últimos elementos filtrados según Rango de Desorción:")
        print(ultimo_rango_df)
    except ValueError:
        print("El valor de Rango_de_Desorpcion no es válido para realizar el filtrado.")
    
    # Filtrar los rangos y mostrar los datos
    filtered_df_a = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.45) & (ultimo_rango_df['Relative Pressure'] <= 0.55)
    ]
    print("Filtrado en rango 0.45-0.55:")
    print(filtered_df_a)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not ultimo_rango_df['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_a = filtered_df_a['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_a}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    filtered_df_b = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.6) & (ultimo_rango_df['Relative Pressure'] <= 0.75)
    ]
    print("Filtrado en rango 0.6-0.75:")
    print(filtered_df_b)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not filtered_df_b['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_b = filtered_df_b['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_b}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    # Dividir los promedios
    div = promedio_volume_stp_a / promedio_volume_stp_b
    print("La división es: " + str(round(div, 2)))  # Convierte el número a cadena y lo redondea
    if (div >= 1):
        return True
    else:
        return False
    # Filtrar los últimos N elementos según el valor de Rango_de_Desorpcion
   
        
def BET_P(df, ruta_excel, Rango_de_Absorpcion, Rango_de_Desorpcion):

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    try:
        # Convertir Rango_de_Desorpcion a entero y filtrar
        num_filas = int(Rango_de_Desorpcion)
        ultimo_rango_df = df.tail(num_filas)
        print("Últimos elementos filtrados según Rango de Desorción:")
        print(ultimo_rango_df)
    except ValueError:
        print("El valor de Rango_de_Desorpcion no es válido para realizar el filtrado.")

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    
     # Filtrar los rangos y mostrar los datos
    filtered_df_a = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.6) & (ultimo_rango_df['Relative Pressure'] <= 0.75)
    ]
    print("Filtrado en rango 0.6-0.75:")
    print(filtered_df_a)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not ultimo_rango_df['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_a = filtered_df_a['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_a}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")
    
    try:
        # Convertir Rango_de_Absorpcion a entero y filtrar
        num_filas = int(Rango_de_Absorpcion)
        
        # Obtener los primeros "num_filas" elementos
        primeros_rango_df = df.head(num_filas)
        print("Primeros elementos filtrados según Rango de Rango_de_Absorpcion:")
        print(primeros_rango_df)
    except ValueError:
        print("El valor de Rango_de_Absorpcion no es válido para realizar el filtrado.")
        return
   
    filtered_df_b = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.6) & (ultimo_rango_df['Relative Pressure'] <= 0.75)
    ]
    print("Filtrado en rango 0.6-0.75:")
    print(filtered_df_b)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not filtered_df_b['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_b = filtered_df_b['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP': {promedio_volume_stp_b}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    # Dividir los promedios
    div = promedio_volume_stp_a / promedio_volume_stp_b
    print("La división es: " + str(round(div, 2)))  # Convierte el número a cadena y lo redondea
    if (div >= 1):
        return True
    else:
        return False
    # Filtrar los últimos N elementos según el valor de Rango_de_Desorpcion
def BET_C(df, ruta_excel, Rango_de_Absorpcion, Rango_de_Desorpcion):


    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    try:
        # Convertir Rango_de_Desorpcion a entero y filtrar
        num_filas = int(Rango_de_Desorpcion)
        ultimo_rango_df = df.tail(num_filas)
        print("Últimos elementos filtrados según Rango de Desorción:")
        print(ultimo_rango_df)
    except ValueError:
        print("El valor de Rango_de_Desorpcion no es válido para realizar el filtrado.")

    # Verificar si las columnas necesarias existen en el DataFrame
    if 'Volume @ STP' not in df.columns or '1 / [ W((P/Po) - 1) ]' not in df.columns:
        print("Las columnas necesarias no están en el DataFrame.")
        return
    
     # Filtrar los rangos y mostrar los datos
    filtered_df_a = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.95) & (ultimo_rango_df['Relative Pressure'] <= 1.0)
    ]
    print("Filtrado en rango 0.95-1.0:")
    print(filtered_df_a)
   
    try:
        # Convertir Rango_de_Absorpcion a entero y filtrar
        num_filas = int(Rango_de_Absorpcion)
        
        # Obtener los primeros "num_filas" elementos
        primeros_rango_df = df.head(num_filas)
        print("Primeros elementos filtrados según Rango de Rango_de_Absorpcion:")
        print(primeros_rango_df)
    except ValueError:
        print("El valor de Rango_de_Absorpcion no es válido para realizar el filtrado.")
        return
   
    filtered_df_b = ultimo_rango_df[
        (ultimo_rango_df['Relative Pressure'] >= 0.95) & (ultimo_rango_df['Relative Pressure'] <= 1.0)
    ]
    print("Filtrado en rango 0.95-1.0:")
    print(filtered_df_b)

    # Calcular el promedio de la columna 'Volume @ STP'
    if not filtered_df_b['Volume @ STP'].isnull().all():  # Verificar si la columna no está vacía
        promedio_volume_stp_ab = filtered_df_b['Volume @ STP'].mean()
        promedio_volume_stp_ab = promedio_volume_stp_ab + filtered_df_a['Volume @ STP'].mean()
        print(f"Promedio de 'Volume @ STP de valores de Absorbcion y Desorbcion es ': {promedio_volume_stp_ab}")
    else:
        print("La columna 'Volume @ STP' está vacía o contiene solo valores nulos.")

    if (promedio_volume_stp_ab >= 1):
        return True
    else:
        return False
    # Filtrar los últimos N elementos según el valor de Rango_de_Desorpcion
def tests_main(archivo_ruta_completa, archivo_csv):
    print("Inicio de graphs_main")
    print(archivo_ruta_completa)
    
    # Crear un DataFrame para almacenar los resultados
    resultados = pd.DataFrame(columns=["Test", "Resultado", "Promedio_A", "Promedio_B", "División"])

    # Leer los datos del archivo CSV
    df = pd.read_csv(archivo_csv)

    # Calcular el cambio relativo en la columna 'Relative Pressure'
    df['delta_pressure'] = df['Relative Pressure'].diff()

    # Identificar índices donde ocurre una disminución significativa
    umbral_disminucion = -0.05
    puntos_disminucion = df[df['delta_pressure'] < umbral_disminucion]

    print("Puntos de disminución:")
    print(puntos_disminucion)

    # Separar los datos en rangos de absorción y desorción según los cambios
    Rango_de_Absorpcion = df[df['delta_pressure'] >= 0]
    Rango_de_Desorpcion = df[df['delta_pressure'] < 0]

    print("Rango de absorción:")
    print(Rango_de_Absorpcion)

    print("Rango de desorción:")
    print(Rango_de_Desorpcion)

    num_absorcion = len(Rango_de_Absorpcion)
    num_desorcion = len(Rango_de_Desorpcion)

    # Ejecución de los tests (se asume que estas funciones están definidas)
    resultado_bi = BET_BI(df, archivo_csv, num_absorcion, num_desorcion)
    if resultado_bi:
         resultados.loc[len(resultados)] = ["BET_BI", "Hay poros cuello de botella", "-", "-", "-"]
    else:
        resultados.loc[len(resultados)] = ["BET_BI", "No hay poros cuello de botella", "-", "-", "-"]

    resultado_p = BET_P(df, archivo_csv, num_absorcion, num_desorcion)
    if resultado_p:
        resultados.loc[len(resultados)] = ["BET_P", "Hay poros planos", "-", "-", "-"]
    else:
        resultados.loc[len(resultados)] = ["BET_P", "No hay poros planos", "-", "-", "-"]

    resultado_c = BET_C(df, archivo_csv, num_absorcion, num_desorcion)
    if resultado_c:
        resultados.loc[len(resultados)] = ["BET_C", "Hay poros cilindricos", "-", "-", "-"]
    else:
        resultados.loc[len(resultados)] = ["BET_C", "No hay poros cilindricos", "-", "-", "-"]

    # Guardar resultados en una nueva hoja Excel (archivo_planilla.xlsx)
    archivo_excel = archivo_csv.replace(".csv", ".xlsx")

    # Crear un nuevo archivo Excel si no existe
    if not os.path.exists(archivo_excel):
        wb = openpyxl.Workbook()
        wb.save(archivo_excel)

    # Abrir el archivo Excel
    wb = openpyxl.load_workbook(archivo_excel)

    # Si la hoja ya existe, eliminarla
    if "Resultados Tests" in wb.sheetnames:
        del wb["Resultados Tests"]

    # Crear una nueva hoja "Resultados Tests" con los datos actualizados
    ws = wb.create_sheet("Resultados Tests")

    # Escribir los resultados en la hoja
    for r in dataframe_to_rows(resultados, index=False, header=True):
        ws.append(r)

    wb.save(archivo_excel)

    # Generar gráficos (ejemplo: histograma de 'Volume @ STP')
    plt.hist(df['Volume @ STP'], bins=20, color='blue', alpha=0.7)
    plt.title("Histograma de Volume @ STP")
    plt.xlabel("Volume @ STP")
    plt.ylabel("Frecuencia")
    grafico_path = os.path.join(os.path.dirname(archivo_ruta_completa), "histograma.png")
    plt.savefig(grafico_path)
    plt.close()

    print("Proceso completado y gráficos generados.")

    # Ejecutar un módulo específico
    result = subprocess.run(["python", "-m", "novarep_ide"], capture_output=True, text=True)
    print("Salida estándar:", result.stdout)
    print("Errores estándar:", result.stderr)
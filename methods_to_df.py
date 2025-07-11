import openpyxl
import os
import csv
import traceback
from datetime import datetime
from pywinauto import Application, findwindows
import pandas as pd
import time
import threading
# manejar_novawin, leer_csv_y_crear_dataframe,agregar_csv_a_plantilla_excel, guardar_dataframe_en_ini,generar_nombre_unico,agregar_dataframe_a_excel_sin_borrar,agregar_dataframe_a_nueva_hoja,close_window_novawin
from pywinauto.keyboard import send_keys
from openpyxl import Workbook
# ejecutor.py
import subprocess
from queue import Queue
import queue
def generar_nombre_unico(base_path, namext):
    # Normalizar las barras a formato Unix (/)
    base_path = base_path.replace("\\", "/")
    
    if not base_path.endswith(namext):
        base_path += namext

    # Extraer nombre base y extensión
    name, ext = os.path.splitext(base_path)
    
    # Agregar fecha y hora actual al nombre base
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name_with_timestamp = f"{name}_{timestamp}"
    base_path = f"{name_with_timestamp}{ext}"
    
    # Asegurarse de que el nombre sea único
    counter = 1
    while os.path.exists(base_path):
        base_path = f"{name_with_timestamp}_{counter}{ext}"
        counter += 1
    
    # Normalizar las barras de regreso a formato Windows (\)
    return base_path.replace("/", "\\")
    
# Función para manejar la exportación de reportes en un hilo
def hilo_exportar_HK(main_window, path_csv, app, queue):
    try:
        # Exportar el reporte y guardar la ruta en la cola
        ruta_csv = exportar_reporte_HK(main_window, path_csv, app)
        queue.put(ruta_csv)  # Almacenar la ruta exportada
    except Exception as e:
        print(f"Error en la exportación: {e}")
        queue.put(None)
def exportar_reporte_HK(main_window, ruta_exportacion, app):
    try:
        print("Buscando componente 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if graph_view_window.exists(timeout=5):
            print("Componente 'TGraphViewWindow' encontrado.")
            graph_view_window.right_click_input()
            time.sleep(1)
        else:
            raise Exception("No se encontró el componente 'TGraphViewWindow'.")
        
        send_keys('t')  # 'Tables'
        time.sleep(0.3)
        send_keys('e')  # 'HK method'
        time.sleep(0.3)
        send_keys('p')  # 'Pore Size Distribution'
        print("Menú 'Pore Size Distribution' seleccionado.")
        time.sleep(1)

        # Volver a hacer clic derecho para acceder al menú de exportación
        graph_view_window.right_click_input()
        time.sleep(0.5)
        send_keys('x')  # 'Pore Size Distribution'
        time.sleep(1)
        csv_dialog = app.window(class_name="#32770")
        csv_dialog.wait("visible ready", timeout=10)
        print("Diálogo de guardado encontrado.")

        ruta_exportacion = generar_nombre_unico(ruta_exportacion, "hk.csv")

        send_keys('%m')
        time.sleep(1)
        send_keys(ruta_exportacion)
        time.sleep(0.5)
        send_keys('%g')  # Alt + G para guardar
        print(" Presionado Alt+G")

        # Esperar posible diálogo de sobrescritura (max 2 seg)
        time.sleep(1.5)
        print(" Intentando confirmar sobrescritura con Alt+S...")
        send_keys('%s')  # Alt + S para confirmar "Sí, sobrescribir"
        print(" Si apareció el diálogo, fue confirmado con Alt+S.")

        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f" Archivo guardado en: {ruta_relativa}")
        return ruta_relativa
       



    except Exception as e:
        print(f"Error durante la exportación: {e}")
        traceback.print_exc()
        return None

def hilo_exportar_DFT(main_window, path_csv, app, queue):
    try:
        ruta_csv = exportar_reporte_DFT(main_window, path_csv, app)
        queue.put(ruta_csv)
    except Exception as e:
        print(f" Error en la exportación: {e}")
        queue.put(None)

def exportar_reporte_DFT(main_window, ruta_exportacion, app):
    try:
        print(" Buscando componente 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        print(" Componente encontrado. Clic derecho para menú.")
        graph_view_window.right_click_input()
        time.sleep(0.5)

        print("Enviando teclas: T  F  P  X")
        send_keys('t')  # Tables
        time.sleep(0.3)
        send_keys('f')  # DFT method
        time.sleep(0.3)
        send_keys('p')  # Pore Size Distribution
        time.sleep(0.3)
        
        # Volver a hacer clic derecho para acceder al menú de exportación
        graph_view_window.right_click_input()
        time.sleep(0.5)
        send_keys('x')  # 'Pore Size Distribution'
        time.sleep(1)
        csv_dialog = app.window(class_name="#32770")
        csv_dialog.wait("visible ready", timeout=10)
        print("Diálogo de guardado encontrado.")

        ruta_exportacion = generar_nombre_unico(ruta_exportacion, "dft.csv")

        send_keys('%m')
        time.sleep(1)
        send_keys(ruta_exportacion)
        time.sleep(0.5)
        send_keys('%g')  # Alt + G para guardar
        print(" Presionado Alt+G")

        # Esperar posible diálogo de sobrescritura (max 2 seg)
        time.sleep(1.5)
        print(" Intentando confirmar sobrescritura con Alt+S...")
        send_keys('%s')  # Alt + S para confirmar "Sí, sobrescribir"
        print(" Si apareció el diálogo, fue confirmado con Alt+S.") 

        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f" Archivo DFT exportado en: {ruta_relativa}")
        return ruta_relativa

    except Exception as e:
        print(f" Error durante la exportación DFT: {e}")
        traceback.print_exc()
        return None
 
def exportar_reporte_BJH_con_teclas( main_window, ruta_exportacion, app,tipo):
    try:
        print(" Buscando componente 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        print(" Componente encontrado. Clic derecho para menú.")
        graph_view_window.right_click_input()
        time.sleep(0.5)

        print("Enviando teclas: T  F  P  X")
        send_keys('t')  # Tables
        time.sleep(0.3)
        send_keys('j')  # BJH Pore Size Distribution
        time.sleep(0.3)
        send_keys(tipo)
           
        ruta_exportacion = generar_nombre_unico(ruta_exportacion, "bjhd.csv")
     

      # Volver a hacer clic derecho para acceder al menú de exportación
        graph_view_window.right_click_input()
        time.sleep(0.5)
        send_keys('x')  # 'Exportar'
        ruta_exportacion = generar_nombre_unico(ruta_exportacion, tipo+"bjh.csv")

        send_keys('%m')
        time.sleep(1)
        send_keys(ruta_exportacion)
        time.sleep(0.5)
        send_keys('%g')  # Alt + G para guardar
        print(" Presionado Alt+G")

        # Esperar posible diálogo de sobrescritura (max 2 seg)
        time.sleep(1.5)
        print(" Intentando confirmar sobrescritura con Alt+S...")
        send_keys('%s')  # Alt + S para confirmar "Sí, sobrescribir"
        print(" Si apareció el diálogo, fue confirmado con Alt+S.") 

        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f" Archivo DFT exportado en: {ruta_relativa}")
        return ruta_relativa

    except Exception as e:
        print(f" Error exportando BJH: {e}")
        traceback.print_exc()
        return None

def hilo_exportar_reporte_fractal_con_teclas(main_window, path_csv, app,queue):
    try:
        # Ejecutar la exportación y guardar el resultado en la cola
        ruta_csv = exportar_reporte_fractal_con_teclas(main_window, ruta_exportacion, app, tipo)
        queue.put(ruta_csv)
    except Exception as e:
        print(f"Error en la exportación: {e}")
        queue.put(None)
def exportar_reporte_fractal_con_teclas(main_window, ruta_exportacion, app,tipo):
    try:
        print(" Buscando 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        # Click derecho para abrir el menú
        graph_view_window.right_click_input()
        time.sleep(0.5)
        print(" Enviando teclas:T C  N  K  H")
        send_keys('t')  # Tables
        time.sleep(0.3)

        send_keys('c')  # Tables
        time.sleep(0.3)

        send_keys(tipo)  # FHH Method Fractal Dimension (Adsorption)
      
        print(" Menú fractal seleccionado.")
        # Volver a hacer clic derecho para acceder al menú de exportación
        graph_view_window.right_click_input()
        time.sleep(0.5)
        send_keys('x')  # 'Exportar'
        time.sleep(1.5)
        csv_dialog = app.window(class_name="#32770")
        csv_dialog.wait("visible", timeout=10)

        ruta_exportacion = generar_nombre_unico(ruta_exportacion, "fractal.csv")

        print(f" Ingresando ruta: {ruta_exportacion}")
        send_keys(ruta_exportacion)
        time.sleep(0.5)

        send_keys('%g')  # Alt + G para guardar
        print(" Alt+G enviado")

        time.sleep(1.5)
        send_keys('%s')  # Alt + S para confirmar si ya existe
        print(" Alt+S enviado (si corresponde)")

        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f" Archivo exportado correctamente: {ruta_relativa}")
        
        return ruta_relativa

    except Exception as e:
        print(f" Error durante la exportación FHH Adsorption: {e}")
        traceback.print_exc()
        return None
        
def hilo_exportar_BET(main_window, path_csv, app,queue):
    try:
        # Aquí va la lógica para exportar el reporte
        ruta_csv=exportar_reporte_BET(main_window, path_csv, app)        
        queue.put(ruta_csv)  # Almacenar la ruta exportada
    except Exception as e:
        print(f"Error en la exportación: {e}")
        queue.put(None)
def exportar_reporte_BET_con_teclas(main_window, ruta_exportacion, app):
    try:
        print(" Buscando 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        # Click derecho para abrir el menú
        graph_view_window.right_click_input()
        time.sleep(0.5)

        print(" Enviando teclas :T B S")
        send_keys('t')  # Tables
        time.sleep(0.3)
        send_keys('b')  # BET
        time.sleep(0.3)
        send_keys('s')  # Single Point Surface Area
        print(" Menú BET   Export seleccionado.")
        # Volver a hacer clic derecho para acceder al menú de exportación
        graph_view_window.right_click_input()
        time.sleep(0.5)
        send_keys('x')  # 'Exportar'
        time.sleep(1.5)
        csv_dialog = app.window(class_name="#32770")
        csv_dialog.wait("visible", timeout=10)

        ruta_exportacion = generar_nombre_unico(ruta_exportacion, "bet.csv")

        print(f" Ingresando ruta: {ruta_exportacion}")
        send_keys(ruta_exportacion)
        time.sleep(0.5)

        send_keys('%g')  # Alt + G para guardar
        print(" Alt+G enviado")

        time.sleep(1.5)
        send_keys('%s')  # Alt + S para confirmar si ya existe
        print(" Alt+S enviado (si corresponde)")

        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f" Archivo exportado correctamente: {ruta_relativa}")
        
        return ruta_relativa

    except Exception as e:
         
        print(f" Error durante la exportación BET: {e}")
        traceback.print_exc()
        return None

def hilo_leer_csv_y_crear_dataframe(ruta_csv, resultado_dict):
    try:
        resultado_dict['dataframe'] = leer_csv_y_crear_dataframe(ruta_csv)
    except Exception as e:
        resultado_dict['error'] = f"Error al leer CSV: {e}"

# Función para agregar el CSV al Excel en un hilo
def hilo_agregar_csv_a_plantilla_excel(ruta_csv, ruta_excel, resultado_dict):
    try:
        agregar_csv_a_plantilla_excel(ruta_csv, ruta_excel)
        resultado_dict['agregado'] = True
    except Exception as e:
        resultado_dict['error'] = f"Error al agregar datos del CSV a Excel: {e}"

# Función para guardar el DataFrame en un archivo INI en un hilo
# Función para guardar el DataFrame en un archivo INI en un hilo
def hilo_guardar_dataframe_en_ini(df, archivo_ini, resultado_dict):
    try:
        guardar_dataframe_en_ini(df, archivo_ini)
        resultado_dict['guardado'] = True
    except Exception as e:
        resultado_dict['error'] = f"Error al guardar INI: {e}"
def preparar_archivo_excel(path_qps, archivo_planilla):
    archivo_planilla = archivo_planilla.replace("/", "\\")
    archivo_planilla = os.path.normpath(archivo_planilla)
    if os.path.exists(archivo_planilla):
        os.remove(archivo_planilla)
        print(f"Archivo '{archivo_planilla}' eliminado.")
    if not os.path.exists(archivo_planilla):
        workbook = Workbook()
        hoja = workbook.active
        hoja["A2"] = "Nombre de la muestra: " + os.path.basename(path_qps)
        workbook.save(archivo_planilla)
        print(f"Archivo Excel creado en: {archivo_planilla}")
    return archivo_planilla

def exportar_y_guardar_fractal(tipo, hoja, path_novawin, path_qps, path_csv, archivo_planilla, resultado_dict):
    queue = Queue()
    app, main_window = manejar_novawin(path_novawin, path_qps)
    hilo = threading.Thread(target=hilo_exportar_reporte_fractal_con_teclas,
                            args=(main_window, path_csv, app, tipo, queue))
    hilo.start()
    hilo.join()
    ruta_csv = queue.get()
    close_window_novawin()
    if ruta_csv:
        hilo_excel = threading.Thread(target=hilo_agregar_csv_to_plantilla_excel,
                                      args=(ruta_csv, archivo_planilla, resultado_dict, hoja))
        hilo_excel.start()
        hilo_excel.join()
    else:
        raise ValueError(f"Exportación fractal '{tipo}' fallida")
def guardar_final(path_csv, resultado_dict):
    df = resultado_dict.get("HK") or list(resultado_dict.values())[0]
    if df is not None:
        hilo_guardar_ini = threading.Thread(target=hilo_guardar_dataframe_en_ini,
                                            args=(df, os.path.join(path_csv, "dataframe.ini"), resultado_dict))
        hilo_guardar_ini.start()
        hilo_guardar_ini.join()
def ejecutar_en_hebra(funcion):
    hebra = threading.Thread(target=funcion)
    hebra.start()
    hebra.join()
    print("Comando ejecutado en hebra.")
    
def df_main(path_qps, path_csv, path_novawin, archivo_planilla):
    try:
        archivo_planilla = preparar_archivo_excel(path_qps, archivo_planilla)
        resultado_dict = {}

        # Bloques de exportación
        exportar_y_guardar('HK', hilo_exportar_HK, path_novawin, path_qps, path_csv, archivo_planilla, resultado_dict)
        exportar_y_guardar('DFT', hilo_exportar_DFT, path_novawin, path_qps, path_csv, archivo_planilla, resultado_dict)
        exportar_y_guardar('BJHD', lambda *args: ejecutar_bjh_en_hilo('d'), path_novawin, path_qps, path_csv, archivo_planilla, resultado_dict)
        exportar_y_guardar('BJHA', lambda *args: ejecutar_bjh_en_hilo('a'), path_novawin, path_qps, path_csv, archivo_planilla, resultado_dict)

        # Fractales
        tipos_fractales = {'n': 'Fractal_N', 'f': 'Fractal_F', 'k': 'Fractal_K', 'h': 'Fractal_H'}
        for tipo, hoja in tipos_fractales.items():
            exportar_y_guardar_fractal(tipo, hoja, path_novawin, path_qps, path_csv, archivo_planilla, resultado_dict)

        # BET
        exportar_y_guardar('BET', hilo_exportar_BET, path_novawin, path_qps, path_csv, archivo_planilla, resultado_dict)

        # Guardar .ini y Excel final
        guardar_final(path_csv, resultado_dict)

        print("Proceso completado exitosamente.")
        ejecutar_en_hebra(ejecutar_ide)

    except Exception as e:
        print(f"Error en df_main: {e}")
        traceback.print_exc()
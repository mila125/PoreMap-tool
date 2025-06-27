import openpyxl
import os
import csv
import traceback
from datetime import datetime
from pywinauto import Application, findwindows

import time
import threading
#from novawinmng import manejar_novawin, leer_csv_y_crear_dataframe,agregar_csv_a_plantilla_excel, guardar_dataframe_en_ini,generar_nombre_unico,agregar_dataframe_a_excel_sin_borrar,agregar_dataframe_a_nueva_hoja,close_window_novawin
from pywinauto.keyboard import send_keys
from openpyxl import Workbook
# ejecutor.py
import subprocess
from queue import Queue
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
       

        print("Archivo exportado exitosamente.")
        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f"Archivo exportado en: {ruta_relativa}")
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
        send_keys('x')  # 'Exportar'

        print(" Exportar a CSV solicitado.")

        # Esperar a que aparezca el diálogo
        time.sleep(1.5)
        csv_dialog = app.window(class_name="#32770")
        csv_dialog.wait("visible ready", timeout=10)

        ruta_exportacion = generar_nombre_unico(ruta_exportacion, "dft.csv")
        print(f" Ingresando ruta: {ruta_exportacion}")
        send_keys(ruta_exportacion)
        time.sleep(0.5)

        send_keys('%g')  # Alt + G para guardar
        print(" Presionado Alt+G")

        # Posible diálogo de sobrescritura
        time.sleep(1.5)
        send_keys('%s')  # Alt + S para confirmar sobrescritura
        print(" Alt+S enviado por si hay confirmación.")

        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f" Archivo DFT exportado en: {ruta_relativa}")
        return ruta_relativa

    except Exception as e:
        print(f" Error durante la exportación DFT: {e}")
        traceback.print_exc()
        return None
def exportar_reporte_BJH_con_teclas( main_window, ruta_exportacion, app,tipo):
    """
    Exporta reporte BJH Pore Size Distribution.
    tipo: 'adsorption' o 'desorption'
    """
    try:
        print(" Buscando 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

            graph_view_window.right_click_input()
            time.sleep(0.5)

            print(" Enviando secuencia de teclas: T  J  A/D  X")
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

            time.sleep(1.5)
            csv_dialog = app.window(class_name="#32770")
            csv_dialog.wait("visible", timeout=10)

            print(" Ingresando ruta y guardando archivo...")
            send_keys(ruta_exportacion)
            time.sleep(0.5)
            send_keys('%g')  # Alt + G para Guardar
            print(" Alt+G enviado.")

            time.sleep(1.5)
            send_keys('%s')  # Alt + S para Sobrescribir si aparece
            print(" Alt+S enviado (si es necesario).")

            ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
            print(f" Archivo exportado: {ruta_relativa}")
        
            graph_view_window.right_click_input()
            time.sleep(0.5)

            print(" Enviando secuencia de teclas: T  J  A/D  X")
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
            print(f" Menú BJH h  Export seleccionado.")
 
            time.sleep(1.5)
            csv_dialog = app.window(class_name="#32770")
            csv_dialog.wait("visible", timeout=10)

            print(" Ingresando ruta y guardando archivo...")
            send_keys(ruta_exportacion)
            time.sleep(0.5)
            send_keys('%g')  # Alt + G para Guardar
            print(" Alt+G enviado.")

            time.sleep(1.5)
            send_keys('%s')  # Alt + S para Sobrescribir si aparece
            print(" Alt+S enviado (si es necesario).")

            ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
            print(f" Archivo exportado: {ruta_relativa}")
            return ruta_relativa

    except Exception as e:
        print(f" Error exportando BJH: {e}")
        traceback.print_exc()
        return None

def hilo_exportar_FFHA(main_window, path_csv, app,queue):
    try:
        # Aquí va la lógica para exportar el reporte
        ruta_csv=exportar_reporte_FFHA(main_window, path_csv, app)
        queue.put(ruta_csv)  # Almacenar la ruta exportada
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

        print(" Enviando teclas: C  N  K  H")
        send_keys('c')  # Tables
        time.sleep(0.3)
        send_keys('n')  # Fractal Dimension Methods
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

        print(" Enviando teclas :C B S")
        send_keys('c')  # Tables
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
def hilo_guardar_dataframe_en_ini(df, archivo_ini, resultado_dict):
    try:
        guardar_dataframe_en_ini(df, archivo_ini)
        resultado_dict['guardado'] = True
    except Exception as e:
        resultado_dict['error'] = f"Error al guardar INI: {e}"
        
def df_main(path_qps, path_csv, path_novawin,archivo_planilla):
    queue = Queue()

    resultado_dict = {}

    archivo_planilla = archivo_planilla.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
    # Normalizar la ruta del archivo
    archivo_planilla = os.path.normpath(archivo_planilla)

    print("Inicio de df_main")
    
    # Imprimir la ruta completa
    print(archivo_planilla)
    # Si el archivo ya existe, eliminarlo
    if os.path.exists(archivo_planilla):
      os.remove(archivo_planilla)
      print(f"Archivo '{archivo_planilla}' eliminado.")
 
    # Crear el archivo Excel si no existe
    if not os.path.exists(archivo_planilla):
     workbook = Workbook()
     hoja = workbook.active
     file_name = os.path.basename(path_qps)
     hoja["A2"] = "Nombre de la muestra: " + file_name  # Agregar el nombre del archivo en A2
     workbook.save(archivo_planilla)
     print(f"Archivo Excel creado en: {archivo_planilla}")
    else:
     print(f"El archivo ya existe en: {archivo_planilla}")
    
    try:
        # Inicializar y manejar NovaWin
        app, main_window = manejar_novawin(path_novawin, path_qps)

        hilo_exportacion_HK = threading.Thread(target=hilo_exportar_HK, args=(main_window, path_csv, app, queue))
        hilo_exportacion_HK.start()

        # Esperar a que el hilo termine antes de proceder
        hilo_exportacion_HK.join()

        # Recuperar la ruta del archivo exportado
        ruta_csv_HK = queue.get() 
        if ruta_csv_HK is None:
           raise ValueError("La exportación no devolvió una ruta válida. Verifica la función exportar_reporte_HK.")
        close_window_novawin()

        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(ruta_csv_HK)

       # Cerrar la ventana de NovaWin
        close_window_novawin()

        print(dataframe)


        # Crear hilos para cada tarea
        hilo_leer_csv_HK = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv_HK, resultado_dict))
        hilo_leer_csv_HK.start()
        hilo_leer_csv_HK.join()
        agregar_dataframe_a_nueva_hoja(archivo_planilla, "HK", dataframe)
        
        # Inicializar y manejar NovaWin nuevamente
        app, main_window = manejar_novawin(path_novawin, path_qps)

        # Crear un hilo para la exportación (ya no es necesario exportar de nuevo)
        hilo_exportacion_DFT = threading.Thread(target=hilo_exportar_DFT, args=(main_window, path_csv, app,queue))
        hilo_exportacion_DFT.start()

        # Esperar a que el hilo termine antes de proceder
        hilo_exportacion_DFT.join()
        # Recuperar la ruta del archivo exportado
        ruta_csv_DFT = queue.get() 
        if ruta_csv_DFT is None:
           raise ValueError("La exportación no devolvió una ruta válida. Verifica la función exportar_reporte_DFT.")
        close_window_novawin()

        # Crear DataFrame y guardar
        
        dataframe = leer_csv_y_crear_dataframe(ruta_csv_DFT)
        print(dataframe)
        # Crear hilos para cada tarea
        hilo_leer_csv_DFT = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv_DFT, resultado_dict))
        hilo_leer_csv_DFT.start()
        hilo_leer_csv_DFT.join()
        agregar_dataframe_a_nueva_hoja(archivo_planilla, "DFT", dataframe)
        
        # Inicializar y manejar NovaWin nuevamente
        app, main_window = manejar_novawin(path_novawin, path_qps)
       
        # Crear un hilo para la exportación (ya no es necesario exportar de nuevo)
        hilo_exportacion_BJHD = threading.Thread(target=hilo_exportar_BJHD, args=(main_window, path_csv, app,queue))
        hilo_exportacion_BJHD.start()

        # Esperar a que el hilo termine antes de proceder
        hilo_exportacion_BJHD.join()
        # Recuperar la ruta del archivo exportado
        ruta_csv_BJHD = queue.get() 
        if ruta_csv_BJHD is None:
           raise ValueError("La exportación no devolvió una ruta válida. Verifica la función exportar_reporte_BJHD.")
        close_window_novawin()
       
        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(ruta_csv_BJHD)
        print(dataframe)
        # Crear hilos para cada tarea
        hilo_leer_csv_BJHD = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv_BJHD, resultado_dict))
        hilo_leer_csv_BJHD.start()
        hilo_leer_csv_BJHD.join()
        agregar_dataframe_a_nueva_hoja(archivo_planilla, "BJHD", dataframe)
        
        # Inicializar y manejar NovaWin nuevamente
        app, main_window = manejar_novawin(path_novawin, path_qps)
       
        # Crear un hilo para la exportación (ya no es necesario exportar de nuevo)
        hilo_exportacion_BJHA = threading.Thread(target=hilo_exportar_BJHA, args=(main_window, path_csv, app,queue))
        hilo_exportacion_BJHA.start()

        # Esperar a que el hilo termine antes de proceder
        hilo_exportacion_BJHA.join()

        ruta_csv_BJHA = queue.get() 
        if ruta_csv_BJHA is None:
           raise ValueError("La exportación no devolvió una ruta válida. Verifica la función exportar_reporte_BJHA.")
        close_window_novawin()
        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(ruta_csv_BJHA)
        print(dataframe)
        # Crear hilos para cada tarea
        hilo_leer_csv_BJHA = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv_BJHA, resultado_dict))
        hilo_leer_csv_BJHA.start()
        hilo_leer_csv_BJHA.join()
        agregar_dataframe_a_nueva_hoja(archivo_planilla, "BJHA", dataframe)
       
        #guardar_dataframe_en_ini(dataframe, path_csv+"dataframe.ini")
        
          # Inicializar y manejar NovaWin nuevamente
        app, main_window = manejar_novawin(path_novawin, path_qps)
    
        # Crear un hilo para la exportación (ya no es necesario exportar de nuevo)
        hilo_exportacion_FFHA = threading.Thread(target=hilo_exportar_FFHA, args=(main_window, path_csv, app,queue))
        hilo_exportacion_FFHA.start()

        # Esperar a que el hilo termine antes de proceder
        hilo_exportacion_FFHA.join()

        ruta_csv_FFHA = queue.get() 
        if ruta_csv_FFHA is None:
           raise ValueError("La exportación no devolvió una ruta válida. Verifica la función exportar_reporte_FFHA.")
        close_window_novawin()
        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(ruta_csv_FFHA)
        print(dataframe)
        # Crear hilos para cada tarea
        hilo_leer_csv_FFHA = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv_FFHA, resultado_dict))
        hilo_leer_csv_FFHA.start()
        hilo_leer_csv_FFHA.join()
        agregar_dataframe_a_nueva_hoja(archivo_planilla, "FFHA", dataframe)
        #guardar_dataframe_en_ini(dataframe, path_csv+"dataframe.ini")
         
        # Inicializar y manejar NovaWin nuevamente
        app, main_window = manejar_novawin(path_novawin, path_qps)
        
        # Crear un hilo para la exportación (ya no es necesario exportar de nuevo)
        hilo_exportacion_NKA = threading.Thread(target=hilo_exportar_NKA, args=(main_window, path_csv, app,queue))
        hilo_exportacion_NKA.start()

        # Esperar a que el hilo termine antes de proceder
        hilo_exportacion_NKA.join()
        ruta_csv_NKA = queue.get() 
        if ruta_csv_NKA is None:
           raise ValueError("La exportación no devolvió una ruta válida. Verifica la función exportar_reporte_NKA.")
        close_window_novawin()
        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(ruta_csv_NKA)
        print(dataframe)
        # Crear hilos para cada tarea
        hilo_leer_csv_NKA = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv_NKA, resultado_dict))
        hilo_leer_csv_NKA.start()
        hilo_leer_csv_NKA.join()
        agregar_dataframe_a_nueva_hoja(archivo_planilla, "NKA", dataframe)
        
        
        # Inicializar y manejar NovaWin nuevamente
        app, main_window = manejar_novawin(path_novawin, path_qps)
    
        # Crear un hilo para la exportación (ya no es necesario exportar de nuevo)
        hilo_exportacion_BET = threading.Thread(target=hilo_exportar_BET, args=(main_window, path_csv, app,queue))
        hilo_exportacion_BET.start()

        # Esperar a que el hilo termine antes de proceder
        hilo_exportacion_BET.join()
        ruta_csv_BET = queue.get() 
        if ruta_csv_BET is None:
           raise ValueError("La exportación no devolvió una ruta válida. Verifica la función exportar_reporte_BET.")
        close_window_novawin()
        # Crear DataFrame y guardar
        dataframe = leer_csv_y_crear_dataframe(ruta_csv_BET)
        
        print(dataframe)
        # Crear hilos para cada tarea
        hilo_leer_csv_BET = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv_BET, resultado_dict))
        hilo_leer_csv_BET.start()
        hilo_leer_csv_BET.join()
        agregar_dataframe_a_nueva_hoja(archivo_planilla, "BET", dataframe)
        
        #guardar_dataframe_en_ini(dataframe, path_csv+"dataframe.ini")      

        print("Proceso completado exitosamente.")
        
        # Crear hilos para cada tarea
        #hilo_leer_csv = threading.Thread(target=hilo_leer_csv_y_crear_dataframe, args=(ruta_csv, resultado_dict))
        hilo_agregar_excel = threading.Thread(target=hilo_agregar_csv_a_plantilla_excel, args=(ruta_csv, path_csv, resultado_dict))
        hilo_guardar_ini = threading.Thread(target=hilo_guardar_dataframe_en_ini, args=(resultado_dict.get('dataframe', None), path_csv + "dataframe.ini", resultado_dict))

        # Iniciar hilos
        hilo_leer_csv.start()
        hilo_agregar_excel.start()
        hilo_guardar_ini.start()

        # Esperar a que todos los hilos terminen
        hilo_leer_csv.join()
        hilo_agregar_excel.join()
        hilo_guardar_ini.join()

        # Verificar errores o resultados en el diccionario
        if 'error' in resultado_dict:
            print(f"Error: {resultado_dict['error']}")
        else:
            print("Todas las tareas completadas exitosamente.")

        # Continuar con otras tareas si es necesario
        print("Proceso completado exitosamente.")
        # Ejecutar un módulo específico
        # Crear y ejecutar la hebra
        hebra = threading.Thread(target=ejecutar_ide)
        hebra.start()

        print("El comando se está ejecutando en una hebra separada.")
        hebra.join()
    except Exception as e:
        print(f"Error en df_main: {e}")
        traceback.print_exc()
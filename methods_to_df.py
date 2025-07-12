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
from pywinauto.findwindows import find_elements
from pywinauto.application import Application
import pyautogui
def generar_nombre_unico(base_path, namext):
    import os
    from datetime import datetime

    # Normalizar las barras a formato Unix (/)
    base_path = base_path.replace("\\", "/")

    # Asegurarse de que termine en barra
    if not base_path.endswith("/"):
        base_path += "/"

    # Concatenar el nombre base
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

    # Convertir de vuelta a formato Windows con barras invertidas
    return base_path.replace("/", "\\")

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
        print("Esperando aparición de diálogo...")
        for i in range(5):
            dialogs = find_elements(class_name="#32770", process=app.process)
            visibles = [d for d in dialogs if d.visible]
            print(f"[{i}] Diálogos visibles encontrados:", visibles)
            if visibles:
                break
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

def exportar_reporte_DFT(main_window, ruta_exportacion, app):
    try:
        print(" Buscando componente 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        print(" Componente encontrado. Clic derecho para menú.")
        graph_view_window.right_click_input()
        time.sleep(0.5)

        print(" Enviando teclas: T F P")
        send_keys('t')
        time.sleep(0.3)
        send_keys('f')  # DFT
        time.sleep(0.3)
        send_keys('p')  # Pore size

        time.sleep(1)

        # Volver a hacer clic derecho para acceder al menú de exportación
        graph_view_window.right_click_input()
        time.sleep(0.5)
        send_keys('x')  # 'Pore Size Distribution'
        time.sleep(1)
        print("Esperando aparición de diálogo...")
        for i in range(5):
            dialogs = find_elements(class_name="#32770", process=app.process)
            visibles = [d for d in dialogs if d.visible]
            print(f"[{i}] Diálogos visibles encontrados:", visibles)
            if visibles:
                break
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
        print(f" Archivo guardado en: {ruta_relativa}")
        return ruta_relativa


        # Cerrar la ventana correctamente antes de retornar
        print("Cerrando ventana TGraphViewWindow...")
        graph_view_window.close()
        time.sleep(1)
    except Exception as e:
        print(f" Error en exportación DFT: {e}")
        traceback.print_exc()
        return None

def actualizar_label_estado(texto):
    global label_estado  # Necesario si vamos a crear o modificar la variable

    try:
        # Si label_estado no existe o fue destruido, se crea nuevamente
        if 'label_estado' not in globals() or not label_estado.winfo_exists():
            # Asegúrate de que 'ventana' está definido globalmente o pásalo como argumento
            label_estado = Label(ventana, text=texto, fg="blue")
            label_estado.grid(row=10, column=0, pady=10)
            print("[INFO] Label recreado.")
        else:
            label_estado.config(text=texto)
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo actualizar o crear label_estado: {e}")
        
def exportar_reporte_BJH_con_teclas(main_window, ruta_exportacion, app):
    try:
        import configparser
        import os
        import time
        import traceback
        from pywinauto.keyboard import send_keys
        from pywinauto.findwindows import find_elements
        from pywinauto.controls.hwndwrapper import HwndWrapper

        rutas_relativas = []

        # --- Limpiar config.ini ---
        ruta_ini = "config.ini"
        config = configparser.ConfigParser()
        config.read(ruta_ini)

        try:
            if label_estado.winfo_exists():
                label_estado.config(text="Iniciando exportación BJH...")
        except:
            print("Label no disponible para actualizar estado.")

        if 'SeccionDeseada' in config and 'label3' in config['SeccionDeseada']:
            config.remove_option('SeccionDeseada', 'label3')
            with open(ruta_ini, 'w') as configfile:
                config.write(configfile)

        # --- Exportar función básica (asume que vista ya está seleccionada) ---
        def exportar(tipo_letra):
            print(f" Exportando BJH tipo {tipo_letra.upper()}...")
            tlistboxes = find_elements(class_name="TListBox", top_level_only=False,
                                       enabled_only=False, visible_only=False,
                                       parent=main_window.element_info)

            if not tlistboxes:
                raise Exception(" No se encontraron elementos TListBox.")

            ultimo_tlistbox = tlistboxes[-1]
            wrapper = HwndWrapper(ultimo_tlistbox.handle)
            wrapper.click_input(button='right')
            time.sleep(0.5)

            send_keys('x')  # 'Pore Size Distribution'
            ruta_completa = generar_nombre_unico(ruta_exportacion, tipo_letra + "bjh.csv")

            send_keys('%m')
            time.sleep(1)
            send_keys(ruta_completa)
            time.sleep(0.5)
            send_keys('%g')
            print(" Presionado Alt+G")

            time.sleep(1.5)
            print(" Intentando confirmar sobrescritura con Alt+S...")
            send_keys('%s')
            print(" Si apareció el diálogo, fue confirmado con Alt+S.")

            ruta_relativa = os.path.relpath(ruta_completa, start=os.getcwd())
            print(f" Archivo BJH exportado en: {ruta_relativa}")
            rutas_relativas.append(ruta_relativa)

        # --- Preparar ventana principal ---
        graph_view_window = main_window.child_window(class_name="QGraph")
        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        graph_view_window.set_focus()
        graph_view_window.click_input()

        # --- Exportar Adsorción ---
        graph_view_window.right_click_input()
        time.sleep(0.5)
        print(" Seleccionando BJH Adsorción...")
        send_keys('t')
        time.sleep(0.5)
        send_keys('j')
        time.sleep(0.5)
        send_keys('a')  # Adsorción
        time.sleep(1.5)
        exportar("a")
        time.sleep(0.5)
        # Reemplaza este bloque por uno más robusto:
           # --- Preparar ventana principal ---
        graph_view_window = main_window.child_window(class_name="QGraph")
        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'QGraph'.")
        graph_view_window.set_focus()
        graph_view_window.click_input()  # asegura foco
        time.sleep(0.5)
        graph_view_window.right_click_input()
        time.sleep(1)  # <--- aumenta tiempo de espera
        print(" Seleccionando BJH Desorción...")
        send_keys('t')
        time.sleep(1)
        send_keys('j')
        time.sleep(1)
        send_keys('d')  # Desorción
        time.sleep(2)  # espera más antes de exportar
        exportar("d")

        return rutas_relativas

    except Exception as e:
        print(f" Error en exportación BJH: {e}")
        traceback.print_exc()
        try:
            if label_estado.winfo_exists():
                label_estado.config(text=f"Error en exportación BJH: {e}")
        except:
            print("No se pudo actualizar el label de estado.")
        return None
def buscar_graph_view_window(main_window):
    try:
        print(" Intentando encontrar 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")
        if graph_view_window.exists(timeout=3):
            print(" 'TGraphViewWindow' encontrado.")
            return graph_view_window
        else:
            raise Exception("No se encontró 'TGraphViewWindow'")
    except:
        print(" 'TGraphViewWindow' no encontrado, buscando hijo con clase que comience con 'QGraph'...")
        for child in main_window.children():
            class_name = child.element_info.class_name
            if class_name and class_name.startswith("QGraph"):
                print(f" Componente alternativo encontrado: {class_name}")
                return child
        raise Exception(" No se encontró ni 'TGraphViewWindow' ni un hijo con clase 'QGraph'")
def exportar_reporte_fractal_con_teclas(main_window, ruta_exportacion, app, tipo):
    import configparser
    import os
    import time
    import traceback
    from pywinauto.keyboard import send_keys
    from pywinauto.findwindows import find_elements
    from pywinauto.controls.hwndwrapper import HwndWrapper
    try:
        # --- Preparar ventana principal ---
        graph_view_window = main_window.child_window(class_name="QGraph")
        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        graph_view_window.click_input()

        # Click derecho para abrir el menú principal
        graph_view_window.set_focus()
        graph_view_window.right_click_input()
        time.sleep(0.5)

        print(" Enviando teclas:T C N K H")
        send_keys('t')
        time.sleep(0.3)
        send_keys('c')
        time.sleep(0.3)
        send_keys(tipo)  # 'n' o 'f'
        time.sleep(0.3)
     

        print(" Menú fractal seleccionado. Abriendo exportación...")

 
        time.sleep(0.5)
        print(f" Exportando FRACTAL tipo {tipo.upper()}...")
        tlistboxes = find_elements(class_name="TListBox", top_level_only=False,
            enabled_only=False, visible_only=False,
            parent=main_window.element_info)

        if not tlistboxes:
             raise Exception(" No se encontraron elementos TListBox.")

        ultimo_tlistbox = tlistboxes[-1]
        wrapper = HwndWrapper(ultimo_tlistbox.handle)
        wrapper.click_input(button='right')
        time.sleep(0.5)

        send_keys('x')  # 'Pore Size Distribution'
        ruta_completa = generar_nombre_unico(ruta_exportacion, tipo + "fractal.csv")

        send_keys('%m')
        time.sleep(1)
        send_keys(ruta_completa)
        time.sleep(0.5)
        send_keys('%g')
        print(" Presionado Alt+G")

        time.sleep(1.5)
        print(" Intentando confirmar sobrescritura con Alt+S...")
        send_keys('%s')
        print(" Si apareció el diálogo, fue confirmado con Alt+S.")

        ruta_relativa = os.path.relpath(ruta_completa, start=os.getcwd())
        print(f" Archivo exportado correctamente: {ruta_relativa}")
        return ruta_relativa

    except Exception as e:
        print(f" Error durante la exportación FHH Adsorption: {e}")
        traceback.print_exc()
        return None
def exportar_reporte_BET_con_teclas(main_window, ruta_exportacion, app):
    try:
        print(" Buscando 'TGraphViewWindow'...")
        graph_view_window = main_window.child_window(class_name="TGraphViewWindow")

        if not graph_view_window.exists(timeout=5):
            raise Exception(" No se encontró 'TGraphViewWindow'.")

        graph_view_window.right_click_input()
        time.sleep(0.5)

        print(" Enviando teclas :T B S")
        send_keys('t')  # Tables
        time.sleep(0.3)
        send_keys('b')  # BET
        time.sleep(0.3)
        send_keys('s')  # Surface area
        graph_view_window.set_focus()
        graph_view_window.right_click_input()
        time.sleep(0.5)
        send_keys('x')
        time.sleep(1.5)

        dialogs = find_elements(class_name="#32770", process=app.process)
        visibles = [d for d in dialogs if d.visible]
        if not visibles:
            raise Exception("No se encontró diálogo exportar BET.")
        csv_dialog = app.window(handle=visibles[-1].handle)
        csv_dialog.wait("visible ready", timeout=10)

        ruta_exportacion = generar_nombre_unico(ruta_exportacion, "bet.csv")
        print(f" Ingresando ruta: {ruta_exportacion}")
        send_keys(ruta_exportacion)
        time.sleep(0.5)
        send_keys('%g')
        print(" Alt+G enviado")
        time.sleep(1.5)
        send_keys('%s')
        print(" Alt+S enviado")

        ruta_relativa = os.path.relpath(ruta_exportacion, start=os.getcwd())
        print(f"BET exportado: {ruta_relativa}")
        return ruta_relativa

    except Exception as e:
        print(f"Error exportando BET: {e}")
        traceback.print_exc()
        return None

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
    

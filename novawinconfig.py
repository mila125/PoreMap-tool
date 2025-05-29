from tkinter import Tk, Label, Button, Entry, filedialog
import configparser
import threading
from methods_to_df import df_main
from graphs import graphs_main
from datetime import datetime
from tests import tests_main
from rectangles import draw_nested_rectangles
from tkinter import ttk
import pandas as pd
from tkinter import *
from rangos_dft import rangos_dft_main
import os

# Archivo de configuración
config_file = "config.ini"
def ejecutar_modulo_rectangles(funcion):

      
    ventana.quit()
    ventana.destroy()
    funcion(ruta_excel)
      
       

# Función para ejecutar el módulo de tests
def ejecutar_modulo_rangos_dft_main(funcion):

        try:
            # Obtener la ruta del archivo
            ruta_csv = entry_csv.get()
            # Corregir las barras
            ruta_csv = ruta_csv.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
            # Corregir las barras
            ruta_excel = entry_excel.get()
            ruta_excel = ruta_excel.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
            
            # Validar los valores y ejecutar la función

            print(f"Ejecutando {funcion.__name__} con rutas:")
            print(f"CSV: {ruta_csv}")
            print(f"EXCEL: {ruta_excel}")
            ventana.quit()
            ventana.destroy()
            funcion(ruta_csv, ruta_excel)
      
            try:
              label_estado.config(text="Se ha cargado el archivo")
            except TclError:
              print("El widget 'label_estado' ya no está disponible.")
        except ValueError:       
            try:
              label_estado.config(text="Error: Seleccione valores válidos en las ComboBox.")
            except TclError:
              print("El widget 'label_estado' ya no está disponible.")

# Función para ejecutar el módulo de tests
def ejecutar_modulo_tests(funcion):
    # Obtener los valores seleccionados en las ComboBox
    try:
        ruta_csv = entry_csv.get()
    
        ruta_excel = entry_excel.get()
        # Corregir las barras
    
        ruta_csv = ruta_csv.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
    
        ruta_excel = ruta_excel.replace("/", "\\")  # Reemplazar barras normales por barras invertidas

        print(f"Ejecutando {funcion.__name__} con rutas:")
        print(f"CSV: {ruta_csv}")
        print(f"EXCEL: {ruta_excel}")
        ventana.quit()
        ventana.destroy()
        funcion(ruta_csv,ruta_excel)
        try:
          label_estado.config(text="Se ha cargado el archivo")
        except TclError:
          print("El widget 'label_estado' ya no está disponible.")
    except ValueError:       
        try:
          label_estado.config(text="Error: Seleccione valores válidos en las ComboBox.")
        except TclError:
          print("El widget 'label_estado' ya no está disponible.")
# Función para cargar el archivo de Excel y crear el DataFrame
def cargar_archivo():

    ruta_csv = entry_csv.get()
    # Construir la ruta del archivo correctamente
    ruta_excel =  entry_excel.get()
    if ruta_csv:
        entry_csv.delete(0, END)
        entry_csv.insert(0, ruta_csv)
        # Leer el archivo de Excel
        df = pd.read_excel(ruta_excel)
        # Llenar las ComboBox con los valores de la columna "Relative Pressure"
        relative_pressure_values = df["Relative Pressure"].tolist()
        combo_absorcion["values"] = relative_pressure_values
        combo_desorcion["values"] = relative_pressure_values
        label_estado.config(text="Archivo cargado correctamente")

def ejecutar_modulo_grafico(funcion):
    """Ejecutar una función pasando las rutas como argumentos."""
    
    ruta_csv = entry_csv.get()
    
    ruta_excel = entry_excel.get()
    # Corregir las barras
    
    ruta_csv = ruta_csv.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
    
    ruta_excel = ruta_excel.replace("/", "\\")  # Reemplazar barras normales por barras invertidas

    print(f"Ejecutando {funcion.__name__} con rutas:")
    print(f"CSV: {ruta_csv}")
    print(f"EXCEL: {ruta_excel}")
    ventana.quit()
    ventana.destroy()
    funcion(ruta_excel)

# Función genérica para ejecutar módulos con rutas
def ejecutar_modulo(funcion):
    """Ejecutar una función pasando las rutas como argumentos."""
    ruta_qps = entry_qps.get()

    ruta_novawin = entry_novawin.get()
    
    ruta_csv = entry_csv.get()
    
    ruta_excel = entry_excel.get()
    # Corregir las barras
    
    ruta_qps = ruta_qps.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
    
    ruta_novawin = ruta_novawin.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
    
    ruta_csv = ruta_csv.replace("/", "\\")  # Reemplazar barras normales por barras invertidas
    
    ruta_excel = ruta_excel.replace("/", "\\")  # Reemplazar barras normales por barras invertidas


    print(f"Ejecutando {funcion.__name__} con rutas:")
    print(f"QPS: {ruta_qps}, CSV: {ruta_csv}, NovaWin: {ruta_novawin} , EXCEL: {ruta_excel}")
    ventana.quit()
    ventana.destroy()
    funcion(ruta_qps, ruta_csv, ruta_novawin,ruta_excel)

# Función genérica para seleccionar rutas
def seleccionar_ruta(entry, is_file=True):
    """Seleccionar archivo o carpeta y actualizar el Entry correspondiente."""
    try:
        ruta = (
            filedialog.askopenfilename(title="Seleccionar archivo") if is_file
            else filedialog.askdirectory(title="Seleccionar directorio")
        )
        if ruta:
            entry.delete(0, "end")
            entry.insert(0, ruta)
    except Exception as e:
        print(f"Error al seleccionar la ruta: {e}")

def seleccionar_ruta_async(entry, is_file=True):
    """Abrir el diálogo de selección en un hilo separado."""
    threading.Thread(target=seleccionar_ruta, args=(entry, is_file), daemon=True).start()

# Funciones para cargar y guardar configuración
def cargar_configuracion():
    config = configparser.ConfigParser()
    config.read(config_file)
    if "Rutas" in config:
        entry_qps.insert(0, config["Rutas"].get("ruta_qps", ""))
        entry_novawin.insert(0, config["Rutas"].get("ruta_novawin", ""))
        entry_excel.insert(0, config["Rutas"].get("ruta_excel", ""))
        # --> Eliminada la línea de entry_reportes

def guardar_configuracion():
    config = configparser.ConfigParser()
    config["Rutas"] = {
        "ruta_qps": entry_qps.get(),
        "ruta_novawin": entry_novawin.get(),
        "ruta_excel": entry_excel.get()
        # --> Eliminado ruta_reportes
    }
    with open(config_file, "w") as configfile:
        config.write(configfile)
    print("Rutas guardadas en config.ini")
# Crear ventana principal
ventana = Tk()
ventana.title("Selector de Rutas")
ventana.geometry("1000x800")
ventana.resizable(False, False)

# Botón para seleccionar archivo
def seleccionar_archivo():
    archivo_seleccionado = filedialog.askopenfilename(filetypes=[("Archivos XLSX", "*.xlsx")])
    if archivo_seleccionado:
        entry_excel.delete(0, "end")
        entry_excel.insert(0, archivo_seleccionado)


# Etiqueta
Label(ventana, text="Ruta de archivos .qps:").grid(row=0, column=0, padx=5, pady=5, sticky="w")

# Campo de entrada
entry_qps = Entry(ventana, width=50)
entry_qps.grid(row=1, column=0, padx=5, pady=5)

# Botón
Button(ventana, text="Seleccionar Archivo", command=lambda: seleccionar_ruta(entry_qps)).grid(row=1, column=1, padx=5, pady=5)

Label(ventana, text="Ruta de NovaWin:").grid(row=6, column=0, padx=5, pady=5, sticky="w")
entry_novawin = Entry(ventana, width=50)
entry_novawin.grid(row=7, column=0, padx=10, pady=10)
Button(ventana, text="Seleccionar Archivo", command=lambda: seleccionar_ruta(entry_novawin)).grid(row=7, column=1, padx=5, pady=5)

# Etiqueta y campo de entrada para archivo Excel
Label(ventana, text="Ruta del archivo Excel:").grid(row=13, column=0, padx=5, pady=5, sticky="w")
entry_excel = Entry(ventana, width=50)
entry_excel.grid(row=14, column=0, padx=5, pady=5)
# Botón para seleccionar archivo
boton_xlsx = Button(ventana, text="Seleccionar archivo", command=lambda: seleccionar_ruta(entry_excel))
boton_xlsx.grid(row=14, column=1, pady=5)

# Label para el estado
label_estado = Label(ventana, text="", fg="blue", font=("Arial", 10))
label_estado.grid(row=16, column=1, pady=5)

# Botones para guardar configuración y ejecutar funciones
Button(ventana, text="Guardar Configuración", command=guardar_configuracion).grid(row=11, column=0, pady=5)

Button(ventana, text="Ejecutar Metodos", command=lambda: ejecutar_modulo(df_main)).grid(row=12, column=0, pady=5)

cargar_configuracion()

# Iniciar ventana
ventana.mainloop()
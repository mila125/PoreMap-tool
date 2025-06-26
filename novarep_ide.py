from tkinter import Tk, Label, Button, Entry, filedialog, Frame, Scrollbar, VERTICAL, HORIZONTAL, RIGHT, Y, BOTTOM, X, BOTH  
from tkinter import ttk
import configparser
import pandas as pd
import threading
from graphs import graphs_main
from tests import tests_main
from rangos_dft import rangos_dft_main
from rectangles import draw_nested_rectangles
from config_ini import main
from cesarofrac import visualizar_poro_fractal
from porespy_view import visualizar_poro_con_porespy
# Archivo de configuración
config_file = "config.ini"

# Crear ventana principal
ventana = Tk()
ventana.title("Selector de Archivo Excel")
ventana.geometry("1000x800")
ventana.resizable(True, True)  # Permitir que se redimensione

# Etiqueta y campo de entrada para archivo Excel
Label(ventana, text="Ruta del archivo Excel:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_excel = Entry(ventana, width=50)
entry_excel.grid(row=0, column=1, padx=10, pady=10)

# Función para seleccionar archivo Excel
def seleccionar_archivo():
    archivo_seleccionado = filedialog.askopenfilename(filetypes=[("Archivos XLSX", "*.xlsx")])
    if archivo_seleccionado:
        entry_excel.delete(0, "end")
        entry_excel.insert(0, archivo_seleccionado)

# Selección de archivo (ya está en row=0)
Button(ventana, text="Seleccionar archivo", command=seleccionar_archivo).grid(row=0, column=2, padx=10, pady=10)


# Estado
label_estado = Label(ventana, text="", fg="blue", font=("Arial", 10))
label_estado.grid(row=1, column=1, columnspan=2, pady=5, sticky="w")

# Combobox para seleccionar hoja
Label(ventana, text="Selecciona hoja:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
combo_hojas = ttk.Combobox(ventana, state="readonly")
combo_hojas.grid(row=3, column=1, padx=10, pady=10, sticky="w")

# Ocultar el combo y la tabla inicialmente
combo_hojas.grid_forget()
frame_tabla = Frame(ventana, width=960, height=600)
frame_tabla.grid(row=3, column=0, columnspan=3, padx=10, pady=10)
frame_tabla.grid_forget()  # Inicialmente oculto

scrollbar_y = Scrollbar(frame_tabla, orient=VERTICAL)
scrollbar_y.pack(side=RIGHT, fill=Y)
scrollbar_x = Scrollbar(frame_tabla, orient=HORIZONTAL)
scrollbar_x.pack(side=BOTTOM, fill=X)

tree = ttk.Treeview(frame_tabla, yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
tree.pack(expand=True, fill=BOTH)

scrollbar_y.config(command=tree.yview)
scrollbar_x.config(command=tree.xview)

# Función para mostrar DataFrame
def mostrar_dataframe(df):
    # Hacer visibles el combo y el treeview
    combo_hojas.grid(row=2, column=1, padx=10, pady=10, sticky="w")
    frame_tabla.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

    tree.delete(*tree.get_children())
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    
    # Establecer las columnas con un ancho fijo total
    ancho_fijo = 1000
    column_width = ancho_fijo // len(df.columns)  # Dividir el ancho entre el número de columnas
    
    # Establecer el ancho de cada columna
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=column_width)
    
    # Insertar los datos en el Treeview
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

def cargar_archivo(ruta_excel):
    if not ruta_excel:
        label_estado.config(text="Por favor, proporciona la ruta del archivo Excel")
        return
    try:
        # Leer solo nombres de hojas
        hojas = pd.ExcelFile(ruta_excel).sheet_names
        combo_hojas["values"] = hojas
        if hojas:
            combo_hojas.current(0)
            cargar_hoja(ruta_excel, hojas[0])
        label_estado.config(text="Archivo cargado correctamente")
    except Exception as e:
        label_estado.config(text=f"Error: {e}")

# Función para ejecutar módulos
def ejecutar_modulo(funcion):
    ruta_excel = entry_excel.get()
    hoja_seleccionada = combo_hojas.get()
    if not ruta_excel:
        label_estado.config(text="Selecciona un archivo Excel")
        return
    print(f"Ejecutando {funcion.__name__} con Excel: {ruta_excel} y hoja: {hoja_seleccionada}")
    ventana.quit()
    ventana.destroy()
    funcion(ruta_excel, hoja_seleccionada)  # PASA LA HOJA

# Visualizar Excel (va en row=2 debajo del selector de archivo)
Button(ventana, text="Visualizar Excel", command=lambda: cargar_archivo(entry_excel.get())).grid(row=2, column=0, columnspan=3, pady=10)
# Botones de acciones (bajan a row=5)
Button(ventana, text="Visualizar histograma", command=lambda: ejecutar_modulo(graphs_main)).grid(row=5, column=0, pady=10)
Button(ventana, text="Tests de poros (BET)", command=lambda: ejecutar_modulo(tests_main)).grid(row=5, column=1, pady=10)
Button(ventana, text="Clasificar poros (DFT)", command=lambda: ejecutar_modulo(rangos_dft_main)).grid(row=5, column=2, pady=10)
Button(
    ventana,
    text="Visualizar poro fractal",
    command=lambda: ejecutar_modulo(
        lambda ruta, hoja: visualizar_poro_fractal(ruta, hoja)
    )
).grid(row=7, column=1, pady=10)
Button(ventana, text="Configuracion de inicio", command=lambda: ejecutar_modulo(main)).grid(row=7, column=2, pady=10)
Button(
    ventana,
    text="Visualizar poro (porespy)",
    command=lambda: ejecutar_modulo(lambda ruta, hoja: visualizar_poro_con_porespy(ruta, hoja))
).grid(row=8, column=1, pady=10)

# Cargar configuración inicial
def cargar_configuracion():
    config = configparser.ConfigParser()
    config.read(config_file)
    if "Rutas" in config:
        entry_excel.insert(0, config["Rutas"].get("ruta_excel", ""))

# Guardar configuración al cerrar
def guardar_configuracion():
    config = configparser.ConfigParser()
    config["Rutas"] = {"ruta_excel": entry_excel.get()}
    with open(config_file, "w") as configfile:
        config.write(configfile)

def cargar_hoja(ruta_excel, hoja):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=hoja)
        mostrar_dataframe(df)
        label_estado.config(text=f"Mostrando hoja: {hoja}")
    except Exception as e:
        label_estado.config(text=f"Error al cargar hoja: {e}")

def on_hoja_seleccionada(event):
    ruta_excel = entry_excel.get()
    hoja_seleccionada = combo_hojas.get()
    if ruta_excel and hoja_seleccionada:
        cargar_hoja(ruta_excel, hoja_seleccionada)

combo_hojas.bind("<<ComboboxSelected>>", on_hoja_seleccionada)

ventana.protocol("WM_DELETE_WINDOW", lambda: [guardar_configuracion(), ventana.destroy()])

cargar_configuracion()

# Iniciar la ventana
ventana.mainloop()
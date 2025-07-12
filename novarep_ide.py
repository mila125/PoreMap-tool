from tkinter import Tk, Label, Button, Entry, filedialog, Frame, Scrollbar, VERTICAL, HORIZONTAL, RIGHT, Y, BOTTOM, X, BOTH  
from tkinter import ttk
import configparser
import pandas as pd
from config_ini import main as config_main
import subprocess

config_file = "config.ini"

ventana = Tk()
ventana.title("Selector de Archivo Excel")
ventana.geometry("600x600")
ventana.resizable(False, False)
# --- Estado ---
label_estado = Label(ventana, text="", fg="blue", font=("Arial", 10))
label_estado.grid(row=2, column=1, columnspan=2, pady=5, sticky="w")
# --- Ruta archivo Excel ---
Label(ventana, text="Ruta del archivo Excel:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_excel = Entry(ventana, width=50)
entry_excel.grid(row=0, column=1, padx=10, pady=10)
def abrir_config_ini():
    ruta_carpeta = entry_carpeta.get()
    if not ruta_carpeta:
        label_estado.config(text="Por favor, selecciona una carpeta primero")
        return
    try:
     config_main(ruta_carpeta)
     print("Módulo configuración abierto correctamente")
     ventana.destroy()
    except Exception as e:
     print(f"Error al abrir módulo: {e}")
     try:
        label_estado.config(text=f"Error: {e}")
     except:
        pass
def seleccionar_archivo():

    ventana.geometry("1000x600")        # Cambia el tamaño
    ventana.update_idletasks()          # Fuerza la actualización visual
    ventana.resizable(False, False)     # Opcional: evitar que el usuario redimensione
    archivo = filedialog.askopenfilename(filetypes=[("Archivos XLSX", "*.xlsx")])
    if archivo:
        
        entry_excel.delete(0, 'end')
        entry_excel.insert(0, archivo)

Button(ventana, text="Seleccionar archivo", command=seleccionar_archivo).grid(row=0, column=2, padx=10, pady=10)

# --- Ruta carpeta ---
Label(ventana, text="Ruta de carpeta:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
entry_carpeta = Entry(ventana, width=50)
entry_carpeta.grid(row=1, column=1, padx=10, pady=10)

def seleccionar_carpeta():
    carpeta = filedialog.askdirectory(title="Selecciona carpeta")
    if carpeta:
        entry_carpeta.delete(0, 'end')
        entry_carpeta.insert(0, carpeta)

Button(ventana, text="Seleccionar carpeta", command=seleccionar_carpeta).grid(row=1, column=2, padx=10, pady=10)



# --- Selección de hoja ---
Label(ventana, text="Selecciona hoja:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
combo_hojas = ttk.Combobox(ventana, state="readonly")
combo_hojas.grid(row=4, column=1, padx=10, pady=10, sticky="w")
combo_hojas.grid_forget()

# --- Frame para tabla ---
frame_tabla = Frame(ventana, width=960, height=600)
frame_tabla.grid(row=4, column=0, columnspan=3, padx=10, pady=10)
frame_tabla.grid_forget()

scrollbar_y = Scrollbar(frame_tabla, orient=VERTICAL)
scrollbar_y.pack(side=RIGHT, fill=Y)
scrollbar_x = Scrollbar(frame_tabla, orient=HORIZONTAL)
scrollbar_x.pack(side=BOTTOM, fill=X)

tree = ttk.Treeview(frame_tabla, yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
tree.pack(expand=True, fill=BOTH)

scrollbar_y.config(command=tree.yview)
scrollbar_x.config(command=tree.xview)

def mostrar_dataframe(df):
    combo_hojas.grid(row=3, column=1, padx=10, pady=10, sticky="w")
    frame_tabla.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    tree.delete(*tree.get_children())
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"

    ancho_fijo = 1000
    column_width = ancho_fijo // len(df.columns)

    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=column_width)

    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

def cargar_archivo():
    ruta_excel = entry_excel.get()
    if not ruta_excel:
        label_estado.config(text="Por favor, proporciona la ruta del archivo Excel")
        return
    try:
        hojas = pd.ExcelFile(ruta_excel).sheet_names
        combo_hojas["values"] = hojas
        if hojas:
            combo_hojas.current(0)
            cargar_hoja(ruta_excel, hojas[0])
        label_estado.config(text="Archivo cargado correctamente")
    except Exception as e:
        label_estado.config(text=f"Error: {e}")

def cargar_hoja(ruta_excel, hoja):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=hoja)
        mostrar_dataframe(df)
        label_estado.config(text=f"Mostrando hoja: {hoja}")
    except Exception as e:
        label_estado.config(text=f"Error al cargar hoja: {e}")

def on_hoja_seleccionada(event):
    ruta_excel = entry_excel.get()
    hoja = combo_hojas.get()
    if ruta_excel and hoja:
        cargar_hoja(ruta_excel, hoja)

combo_hojas.bind("<<ComboboxSelected>>", on_hoja_seleccionada)

# Botón para cargar Excel y mostrar hojas
Button(ventana, text="Visualizar Excel", command=cargar_archivo).grid(row=5, column=0, columnspan=3, pady=10)

Button(ventana, text="Ir a Configuración", command=lambda: abrir_config_ini()).grid(row=6, column=0, pady=20)
Button(ventana, text="Acción 2", command=lambda: label_estado.config(text="Acción 2 ejecutada")).grid(row=6, column=1, pady=10)
Button(ventana, text="Acción 3", command=lambda: label_estado.config(text="Acción 3 ejecutada")).grid(row=6, column=2, pady=10)

# Configuración para cerrar y guardar si usas configparser
def cargar_configuracion():
    config = configparser.ConfigParser()
    config.read(config_file)
    if "Rutas" in config:
        entry_excel.insert(0, config["Rutas"].get("ruta_excel", ""))
        entry_carpeta.insert(0, config["Rutas"].get("ruta_carpeta", ""))

def guardar_configuracion():
    config = configparser.ConfigParser()
    config["Rutas"] = {
        "ruta_excel": entry_excel.get(),
        "ruta_carpeta": entry_carpeta.get()
    }
    with open(config_file, "w") as configfile:
        config.write(configfile)

ventana.protocol("WM_DELETE_WINDOW", lambda: [guardar_configuracion(), ventana.destroy()])

cargar_configuracion()

ventana.mainloop()

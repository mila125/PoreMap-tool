from tkinter import Tk, Label, Button, Entry, filedialog, Frame, Scrollbar, VERTICAL, HORIZONTAL, RIGHT, Y, BOTTOM, X, BOTH  
from tkinter import ttk
import configparser
import pandas as pd
import threading
from graphs import graphs_main
from novawinmng import manejar_novawin

def main(ruta_excel,hoja):
    config_file = "startconfig.ini"
    print("Ejecutando configuración con:")
    print("Excel:", ruta_excel)

    # Crear ventana principal
    ventana = Tk()
    ventana.title("Selector de Archivo Excel")
    ventana.geometry("1000x800")
    ventana.resizable(True, True)  # Permitir que se redimensione

    # === Entrada: Planilla Excel ===
    Label(ventana, text="Ruta del archivo Excel:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    entry_excel = Entry(ventana, width=70)
    entry_excel.grid(row=0, column=1, padx=10, pady=10)
    Button(ventana, text="Seleccionar Excel", command=lambda: seleccionar_archivo(entry_excel, [("Archivos Excel", "*.xlsx")])).grid(row=0, column=2)

    # === Entrada: Archivo .qps ===
    Label(ventana, text="Ruta del archivo .qps:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    entry_qps = Entry(ventana, width=70)
    entry_qps.grid(row=1, column=1, padx=10, pady=10)
    Button(ventana, text="Seleccionar .qps", command=lambda: seleccionar_archivo(entry_qps, [("Archivos QPS", "*.qps")])).grid(row=1, column=2)

    # === Entrada: Ruta NovaWin ===
    Label(ventana, text="Ruta del ejecutable NovaWin:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
    entry_novawin = Entry(ventana, width=70)
    entry_novawin.grid(row=2, column=1, padx=10, pady=10)
    Button(ventana, text="Seleccionar NovaWin", command=lambda: seleccionar_archivo(entry_novawin, [("Ejecutables", "*.exe")])).grid(row=2, column=2)

    # === Botón: Manejar NovaWin ===
    Button(
        ventana,
        text="Manejar NovaWin",
        command=lambda: manejar_novawin(entry_novawin.get(), entry_qps.get())
    ).grid(row=4, column=0, pady=20)

    def cargar_configuracion():
        config = configparser.ConfigParser()
        config.read(config_file)
        if "Rutas" in config:
            entry_excel.insert(0, config["Rutas"].get("ruta_excel", ""))
            entry_qps.insert(0, config["Rutas"].get("ruta_qps", ""))
            entry_novawin.insert(0, config["Rutas"].get("ruta_novawin", ""))

    def guardar_configuracion():
        config = configparser.ConfigParser()
        config["Rutas"] = {
            "ruta_excel": entry_excel.get(),
            "ruta_qps": entry_qps.get(),
            "ruta_novawin": entry_novawin.get()
        }
        with open(config_file, "w") as configfile:
            config.write(configfile)
    Button(
     ventana,
     text="Guardar Configuración",
     command=guardar_configuracion
    ).grid(row=4, column=1, pady=20)
    
    ventana.protocol("WM_DELETE_WINDOW", lambda: [guardar_configuracion(), ventana.destroy()])
    cargar_configuracion()
    ventana.mainloop()


if __name__ == "__main__":
    main()
# === Función para seleccionar archivo y colocar en entrada ===
def seleccionar_archivo(entry_widget, filetypes):
    archivo = filedialog.askopenfilename(filetypes=filetypes)
    if archivo:
        entry_widget.delete(0, "end")
        entry_widget.insert(0, archivo)
# === Botón: Guardar Configuración ===

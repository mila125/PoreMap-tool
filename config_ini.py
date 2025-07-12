import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import configparser
import threading
import subprocess

config_file = "startconfig.ini"

def safe(entry_qps,entry_csv,entry_novawin,entry_excel):

    from novawinmng import manejar_novawin
    manejar_novawin(entry_qps,entry_csv,entry_novawin,entry_excel)
def seleccionar_archivo(entry_widget, tipos_archivo):
    archivo = filedialog.askopenfilename(filetypes=tipos_archivo)
    if archivo:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, archivo)

def cargar_configuracion(entry_excel, entry_qps, entry_novawin):
    config = configparser.ConfigParser()
    config.read(config_file)
    if "Rutas" in config:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, config["Rutas"].get("ruta_excel", ""))
        entry_qps.delete(0, tk.END)
        entry_qps.insert(0, config["Rutas"].get("ruta_qps", ""))
        entry_novawin.delete(0, tk.END)
        entry_novawin.insert(0, config["Rutas"].get("ruta_novawin", ""))

# --- CORREGIR GUARDAR ---
def guardar_configuracion(ruta_excel, ruta_qps, ruta_novawin):
    config = configparser.ConfigParser()
    config["Rutas"] = {
        "ruta_excel": ruta_excel,
        "ruta_qps": ruta_qps,
        "ruta_novawin": ruta_novawin
    }
    with open(config_file, "w") as configfile:
        config.write(configfile)
    messagebox.showinfo("Configuración guardada", "La configuración se ha guardado correctamente.")
def main(entry_excel):
    ventana = tk.Tk()
    ventana.title("Configuración Inicial")
    ventana.geometry("600x180")
    ventana.resizable(False, False)

    ttk.Label(ventana, text="Ruta archivo Excel:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
    entry_excel = ttk.Entry(ventana, width=50)
    entry_excel.grid(row=0, column=1, padx=10, pady=10)


    # En config_ini.py
    def seleccionar_carpeta(entry_carpeta):
        carpeta = filedialog.askdirectory(title="Selecciona carpeta")
        if carpeta:
            entry_carpeta.delete(0, 'end')
            entry_carpeta.insert(0, carpeta)

    ttk.Button(ventana, text="Seleccionar carpeta", command=seleccionar_carpeta).grid(row=0, column=2, padx=10, pady=10)
  
    ttk.Label(ventana, text="Ruta archivo .qps:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    entry_qps = ttk.Entry(ventana, width=50)
    entry_qps.grid(row=1, column=1, padx=10, pady=10)
    ttk.Button(ventana, text="Seleccionar .qps", command=lambda: seleccionar_archivo(entry_qps, [("Archivos QPS", "*.qps")])).grid(row=1, column=2, padx=5, pady=10)

    ttk.Label(ventana, text="Ruta ejecutable NovaWin:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
    entry_novawin = ttk.Entry(ventana, width=50)
    entry_novawin.grid(row=2, column=1, padx=10, pady=10)
    ttk.Button(ventana, text="Seleccionar NovaWin", command=lambda: seleccionar_archivo(entry_novawin, [("Ejecutables", "*.exe")])).grid(row=2, column=2, padx=5, pady=10)

    ttk.Button(ventana, text="Manejar NovaWin",
               command=lambda: threading.Thread(target=safe,
                                               args=( entry_qps.get(), entry_excel.get(), entry_novawin.get(),entry_excel.get())).start()
              ).grid(row=3, column=0, pady=20, padx=10)

    ttk.Button(ventana, text="Guardar Configuración",
               command=lambda: guardar_configuracion(
                   entry_excel.get(),
                   entry_qps.get(),
                   entry_novawin.get()
               )
    ).grid(row=3, column=1, pady=20)

    ttk.Button(ventana, text="Volver a inicio",
               command=lambda: [subprocess.Popen(["python", "-m", "novarep_ide"]), ventana.destroy()]
              ).grid(row=3, column=2, pady=20, padx=10)

    ventana.protocol("WM_DELETE_WINDOW", lambda: [
        guardar_configuracion(entry_excel.get(), entry_qps.get(), entry_novawin.get()),
        ventana.destroy()
    ])

    cargar_configuracion(entry_excel, entry_qps, entry_novawin)

    ventana.mainloop()
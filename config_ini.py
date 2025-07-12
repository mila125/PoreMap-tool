import tkinter as tk
from tkinter import filedialog, ttk

def seleccionar_carpeta(entry_widget):
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta destino")
    if carpeta:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, carpeta)

ventana = tk.Tk()
ventana.title("Test Selector Carpeta")
ventana.geometry("600x100")

ttk.Label(ventana, text="Carpeta destino:").grid(row=0, column=0, padx=10, pady=10)
entry = ttk.Entry(ventana, width=50)
entry.grid(row=0, column=1, padx=10, pady=10)
ttk.Button(ventana, text="Seleccionar carpeta", command=lambda: seleccionar_carpeta(entry)).grid(row=0, column=2, padx=10, pady=10)

ventana.mainloop()


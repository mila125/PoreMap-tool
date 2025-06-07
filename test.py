from pywinauto import Application
import time

# Iniciar NovaWin (modifica la ruta)
app = Application(backend="uia").start(r"C:\Quantachrome Instruments\NovaWin\NovaWin.exe")

# Esperar a que la ventana principal esté lista
main_win = app.window(title_re=".*NovaWin.*")  # Usa el título real de la ventana

# Espera a que cargue completamente
main_win.wait('ready', timeout=15)

# Listar controles disponibles (esto es clave para saber qué puedes automatizar)
main_win.print_control_identifiers()

# === EJEMPLO DE ACCIONES EN SEGUNDO PLANO ===

# Escribir en un campo (si existe)
# main_win.child_window(title="Nombre del Campo", control_type="Edit").set_text("12345")

# Hacer clic en botón (si existe)
# main_win.child_window(title="Aceptar", control_type="Button").click()

# Seleccionar una pestaña, menú, etc.
# main_win.child_window(title="Archivo", control_type="MenuItem").select()

# NOTA: estas acciones no requieren que la ventana esté al frente

# Puedes minimizar la ventana
main_win.minimize()

# Continuar usando el PC mientras NovaWin es controlado automáticamente
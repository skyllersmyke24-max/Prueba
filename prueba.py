
import os
import time
from pywinauto import Application

base_path = os.path.dirname(os.path.abspath(__file__))

# Leer contraseña
password_file = os.path.join(base_path, "contrasena.txt")
if not os.path.exists(password_file):
    raise FileNotFoundError(f"No se encontró el archivo: {password_file}")

with open(password_file, "r", encoding="utf-8") as f:
    password = f.read().strip()

# Ejecutar acceso directo o EXE
shortcut_path = os.path.join(base_path, "Actualiza RSIRAT.lnk")
if not os.path.exists(shortcut_path):
    # Si no existe el .lnk, usar ruta directa al EXE
    shortcut_path = r"C:\Program Files\SIRAT\SIRAT.exe"

if not os.path.exists(shortcut_path):
    raise FileNotFoundError(f"No se encontró el acceso directo ni el EXE: {shortcut_path}")

os.startfile(shortcut_path)

# Esperar a que la ventana aparezca
time.sleep(8)

# Conectar a la ventana
app = Application(backend="uia").connect(title="SIRAT")
dlg = app.window(title="SIRAT")

# Escribir contraseña
password_edit = dlg.child_window(auto_id="1005", control_type="Edit")
password_edit.type_keys(password, with_spaces=True)

# Click en Aceptar
dlg.child_window(title="Aceptar", control_type="Button").click()

print("✅ Login completado.")


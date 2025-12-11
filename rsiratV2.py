import os
import sys
import time
import pandas as pd
from io import StringIO
import pyautogui
from PIL import ImageGrab
import pytesseract

from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys

# ======================================================
# 0. CONFIGURAR SALIDA A ARCHIVO DE TEXTO
# ======================================================
log_output = StringIO()

class DualWriter:
    """Escribe en consola y en archivo simultáneamente."""
    def __init__(self, console, log_file):
        self.console = console
        self.log_file = log_file
    
    def write(self, message):
        self.console.write(message)
        self.log_file.write(message)
        self.console.flush()
        self.log_file.flush()
    
    def flush(self):
        self.console.flush()
        self.log_file.flush()

sys.stdout = DualWriter(sys.__stdout__, log_output)

# ======================================================
# 1. RUTAS BASE Y UTILIDADES
# ======================================================
if getattr(sys, "frozen", False):
    # Cuando está compilado con PyInstaller
    base_path = os.path.dirname(sys.executable)
else:
    # Cuando se ejecuta como .py
    base_path = os.path.dirname(os.path.abspath(__file__))


def find_external_file(name: str) -> str:
    """Busca un archivo primero en base_path y luego en el cwd."""
    candidates = [
        os.path.join(base_path, name),
        os.path.join(os.getcwd(), name),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    # Devuelve la primera ruta para mostrar en el error
    return candidates[0]


def wait_and_connect_window(title="SIRAT", timeout=30, poll=0.5):
    """Espera una ventana por título o regex y se conecta con backend UIA."""
    desktop = Desktop(backend="uia")
    end_time = time.time() + timeout
    while time.time() < end_time:
        # 1) Título exacto
        try:
            win = desktop.window(title=title)
            if win.exists(timeout=1):
                handle = win.handle
                app = Application(backend="uia").connect(handle=handle)
                return app, app.window(handle=handle)
        except Exception:
            pass
        # 2) Regex
        try:
            win = desktop.window(title_re=f".*{title}.*")
            if win.exists(timeout=1):
                handle = win.handle
                app = Application(backend="uia").connect(handle=handle)
                return app, app.window(handle=handle)
        except Exception:
            pass

        time.sleep(poll)

    raise ElementNotFoundError({"title": title, "backend": "uia", "timeout": timeout})


def normalizar_nombre(nombre):
    """Normaliza nombres para comparación robusta."""
    return " ".join(str(nombre).strip().upper().split())


# ======================================================
# 2. LECTURA DE CONTRASEÑA Y EXCEL
# ======================================================
password_file = find_external_file("contrasena.txt")
if not os.path.exists(password_file):
    raise FileNotFoundError(
        f"No se encontró el archivo de contraseña. Se buscó en: {password_file}"
    )

with open(password_file, "r", encoding="utf-8") as f:
    password = f.read().strip()

excel_file = find_external_file("expedientes.xlsx")
if not os.path.exists(excel_file):
    raise FileNotFoundError(f"No se encontró el archivo Excel: {excel_file}")

expedientes = pd.read_excel(excel_file, engine="openpyxl")

for col_necesaria in ["DEPENDENCIA", "EXPEDIENTE", "EJECUTOR", "TIPO DE MEDIDA"]:
    if col_necesaria not in expedientes.columns:
        raise ValueError(f"El Excel no contiene la columna obligatoria '{col_necesaria}'")

# Columna RESULTADO si no existe
if "RESULTADO" not in expedientes.columns:
    expedientes["RESULTADO"] = ""

# Mapear dependencia
def map_dependencia(val):
    val_str = str(val).zfill(4)
    if val_str.endswith("21") or val_str.endswith("021") or val_str.endswith("0021"):
        return "0021 I.R. Lima - PRICO"
    elif val_str.endswith("23") or val_str.endswith("023") or val_str.endswith("0023"):
        return "0023 I.R. Lima - MEPECO"
    else:
        return "DESCONOCIDO"


expedientes["DEPENDENCIA_COMPLETA"] = expedientes["DEPENDENCIA"].apply(map_dependencia)

orden_prioridad = {"0021 I.R. Lima - PRICO": 1, "0023 I.R. Lima - MEPECO": 2}
expedientes.sort_values(
    by="DEPENDENCIA_COMPLETA",
    key=lambda x: x.map(orden_prioridad),
    inplace=True,
)

# La dependencia que usaremos para loguear (la de mayor prioridad)
dependencia_login = expedientes.iloc[0]["DEPENDENCIA_COMPLETA"]
print(f"Dependencia usada en login: {dependencia_login}")

# ======================================================
# 3. ABRIR RSIRAT Y LOGUEARSE
# ======================================================
shortcut_path = os.path.join(base_path, "Actualiza RSIRAT.lnk")
if not os.path.exists(shortcut_path):
    shortcut_path = r"C:\Program Files\SIRAT\SIRAT.exe"

if not os.path.exists(shortcut_path):
    raise FileNotFoundError(
        f"No se encontró el acceso directo ni el EXE de SIRAT: {shortcut_path}"
    )

os.startfile(shortcut_path)

try:
    app, dlg = wait_and_connect_window("SIRAT", timeout=30)
except ElementNotFoundError as e:
    all_windows = Desktop(backend="uia").windows()
    titles = [w.window_text() for w in all_windows]
    raise RuntimeError(
        f"No se encontró la ventana de login 'SIRAT'. Ventanas detectadas: {titles}"
    ) from e

# Campos de login
dependencia_edit = dlg.child_window(auto_id="1001", control_type="Edit")
password_edit = dlg.child_window(auto_id="1005", control_type="Edit")

dependencia_edit.type_keys(dependencia_login, with_spaces=True)
password_edit.type_keys(password, with_spaces=True)

dlg.child_window(title="Aceptar", control_type="Button").click()
print("✔ Login completado exitosamente.")

time.sleep(3)

# ======================================================
# 4. PASOS CON PYAUTOGUI - USANDO COORDENADAS
# ======================================================
pyautogui.FAILSAFE = True

# Conectarse a SIRAT (ya logueado)
app = Application(backend="uia").connect(title_re=".*SIRAT.*")
win = app.window(title_re=".*SIRAT.*")
win.set_focus()
time.sleep(1)

# Datos de Inspect para "Cobranza Coactiva"
left = 151
top = 228
width = 226
height = 16

x = left + width // 2
y = top + height // 2

print(f"Haciendo clic en coordenadas ({x}, {y})")
pyautogui.moveTo(x, y, duration=0.3)
pyautogui.click()
time.sleep(1)

# ======================================================
# 5. GUARDAR CAMBIOS EN EL EXCEL Y EXPORTAR LOG
# ======================================================
expedientes.to_excel(excel_file, index=False)
print("\n✔ Proceso finalizado. Excel actualizado.")

# Guardar el log de salida en un archivo de texto
log_dir = os.path.dirname(excel_file)
log_file_path = os.path.join(log_dir, "proceso_log.txt")

sys.stdout = sys.__stdout__  # Restaurar salida estándar
log_content = log_output.getvalue()

with open(log_file_path, "w", encoding="utf-8") as log_file:
    log_file.write(log_content)

print(f"✔ Log del proceso guardado en: {log_file_path}")

import os
import sys
import time
import pandas as pd
import atexit
import datetime
import traceback

from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys

# ======================================================
# 1. RUTAS BASE Y UTILIDADES
# ======================================================

if getattr(sys, "frozen", False):
    # Cuando está compilado con PyInstaller
    base_path = os.path.dirname(sys.executable)
else:
    # Cuando se ejecuta como .py
    base_path = os.path.dirname(os.path.abspath(__file__))

# Duplicar salida terminal a txt en la misma carpeta
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
log_file_path = os.path.join(base_path, f"terminal_output_{timestamp}.txt")
_orig_stdout = sys.stdout
_orig_stderr = sys.stderr

class _Tee:
    def __init__(self, filepath, mode="a", encoding="utf-8"):
        self.terminal = _orig_stdout
        try:
            self.log = open(filepath, mode, encoding=encoding)
        except Exception:
            self.log = None
    def write(self, message):
        try:
            if self.terminal:
                self.terminal.write(message)
        except Exception:
            pass
        try:
            if self.log:
                self.log.write(message)
        except Exception:
            pass
    def flush(self):
        try:
            if self.terminal:
                self.terminal.flush()
        except Exception:
            pass
        try:
            if self.log:
                self.log.flush()
        except Exception:
            pass

def _close_log():
    try:
        sys.stdout.flush()
    except Exception:
        pass
    try:
        sys.stderr.flush()
    except Exception:
        pass
    try:
        if hasattr(sys.stdout, "log") and sys.stdout.log:
            sys.stdout.log.close()
    except Exception:
        pass

try:
    _tee = _Tee(log_file_path)
    sys.stdout = _tee
    sys.stderr = _tee
except Exception:
    pass
atexit.register(_close_log)

def _excepthook(exc_type, exc_value, exc_tb):
    traceback.print_exception(exc_type, exc_value, exc_tb)
    try:
        _close_log()
    except Exception:
        pass
sys.excepthook = _excepthook


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


def wait_and_connect_window(title="SIRAT", timeout=90, poll=0.5):
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
    app, dlg = wait_and_connect_window("SIRAT", timeout=90)
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
print(" Login completado.")

time.sleep(3)

# ======================================================
# 4. FUNCIONES DE NAVEGACIÓN DENTRO DE RSIRAT
# ======================================================
def abrir_menu_principal():
    """Devuelve la ventana de 'Menú de Opciones' principal y su árbol, con reintentos."""
    last_error = None
    for intento in range(3):
        try:
            menu = app.window(title_re=".*Menú de Opciones.*")
            menu.wait("visible", timeout=90)
            menu.set_focus()
            tree = menu.child_window(control_type="Tree")
            return menu, tree
        except Exception as e:
            last_error = e
            print(f"[RETRY abrir_menu_principal] intento {intento+1} falló: {e}")
            time.sleep(3 * (intento+1))
    print("[ERROR] No se pudo abrir el menú principal tras 3 intentos.")
    if last_error:
        raise last_error


def abrir_seleccion_expediente():
    """
    Desde el menú principal:
    Cobranza Coactiva -> Exp. Cob. Coactiva - Individual
    y espera la ventana 'Selección de Expediente Coactivo'. Con reintentos.
    """
    last_error = None
    for intento in range(3):
        try:
            _, tree = abrir_menu_principal()
            cobranza = tree.child_window(
                title_re=r"^Cobranza Coactiva$",
                control_type="TreeItem",
            ).wrapper_object()
            cobranza.wait("enabled visible", timeout=15)
            try:
                cobranza.expand()
            except Exception:
                cobranza.double_click_input()
            exp_indiv = tree.child_window(
                title_re=r"^Exp\. Cob\. Coactiva - Individual$",
                control_type="TreeItem",
            ).wrapper_object()
            exp_indiv.wait("enabled visible", timeout=15)
            exp_indiv.double_click_input()
            sel = app.window(title_re=".*Selección de Expediente Coactivo.*")
            sel.wait("visible", timeout=90)
            sel.set_focus()
            return sel
        except Exception as e:
            last_error = e
            print(f"[RETRY abrir_seleccion_expediente] intento {intento+1} falló: {e}")
            time.sleep(3 * (intento+1))
    print("[ERROR] No se pudo abrir la selección de expediente tras 3 intentos.")
    if last_error:
        raise last_error


def leer_y_validar_ejecutor(sel_ventana, i):
    """
    En la ventana 'Selección de Expediente Coactivo':
    - Escribe el numero de EXPEDIENTE.
    - Verifica ejecutor de la columna Grupo.
    - Actualiza expedientes['RESULTADO'] según el caso.
    - Devuelve:
        - 'coincide' si puede continuar con el flujo de embargo.
        - 'sin_ejecutor' o 'diferente' en otros casos.
    """
    expediente_actual = str(expedientes.loc[i, "EXPEDIENTE"]).strip()
    ejecutor_excel = expedientes.loc[i, "EJECUTOR"]

    # Edit "Número": tomamos el primer Edit
    numero_edit = sel_ventana.child_window(control_type="Edit", found_index=0).wrapper_object()
    numero_edit.set_focus()
    numero_edit.select()  # seleccionar por si había texto previo
    numero_edit.type_keys("^a{BACKSPACE}")  # limpiar
    numero_edit.type_keys(expediente_actual, with_spaces=False)

    # Lanzar la búsqueda (ENTER suele actualizar la grilla)
    send_keys("{ENTER}")
    time.sleep(1.5)

    # Intentar obtener la tabla/grilla
    try:
        tabla = sel_ventana.child_window(control_type="Table")
    except ElementNotFoundError:
        # Puede que la grilla sea List o DataGrid: ajustar si es necesario
        tabla = sel_ventana.child_window(control_type="DataGrid")

    filas = tabla.children()

    if not filas:
        expedientes.loc[i, "RESULTADO"] = "Sin ejecutor designado"
        print(f"{expediente_actual}: sin filas en la grilla (sin ejecutor).")
        # Cerrar solo esta ventana de selección
        sel_ventana.child_window(title="Cerrar", control_type="Button").click_input()
        return "sin_ejecutor"

    fila0 = filas[0]
    textos = fila0.texts()  # textos de Número, Tipo, Saldos, Grupo, etc.

    ejecutor_sirat = textos[-1].strip() if textos else ""

    if ejecutor_sirat == "" or ejecutor_sirat.isspace():
        expedientes.loc[i, "RESULTADO"] = "Sin ejecutor designado"
        sel_ventana.child_window(title="Cerrar", control_type="Button").click_input()
        print(f"{expediente_actual}: sin ejecutor en columna Grupo.")
        return "sin_ejecutor"

    if normalizar_nombre(ejecutor_sirat) != normalizar_nombre(ejecutor_excel):
        expedientes.loc[i, "RESULTADO"] = "Ejecutor diferente"
        sel_ventana.child_window(title="Cerrar", control_type="Button").click_input()
        print(
            f"{expediente_actual}: ejecutor diferente. SIRAT='{ejecutor_sirat}', "
            f"EXCEL='{ejecutor_excel}'."
        )
        return "diferente"

    # Coinciden
    expedientes.loc[i, "RESULTADO"] = "Ejecutor coincide"

    boton_aceptar = sel_ventana.child_window(title="Aceptar", control_type="Button")
    boton_aceptar.click_input()
    print(f"{expediente_actual}: ejecutor coincide, se hizo clic en Aceptar.")
    return "coincide"


def abrir_menu_individual():
    """
    Luego de aceptar en 'Selección de Expediente Coactivo', queda
    un Menú de Opciones para 'Exp. Cob. Coactiva - Individual'.
    Devolvemos esa ventana y su árbol.
    """
    menu = app.window(title_re=".*Menú de Opciones.*")
    menu.wait("visible", timeout=90)
    menu.set_focus()
    tree = menu.child_window(control_type="Tree")
    return menu, tree


def abrir_proceso_embargo_y_trabar():
    """Desde el Menú de 'Exp. Cob. Coactiva - Individual', abre Proceso de Embargo -> Trabar Embargo."""
    _, tree = abrir_menu_individual()

    proc_emb = tree.child_window(
        title_re=r"^Proceso de Embargo$",
        control_type="TreeItem",
    ).wrapper_object()
    proc_emb.wait("enabled visible", timeout=15)
    try:
        proc_emb.expand()
    except Exception:
        proc_emb.double_click_input()

    trabar = tree.child_window(
        title_re=r"^Trabar Embargo$",
        control_type="TreeItem",
    ).wrapper_object()
    trabar.wait("enabled visible", timeout=15)
    trabar.double_click_input()
    print(" Abierto menú 'Trabar Embargo'.")


def abrir_opcion_por_tipo_medida(i):
    """
    Dentro de 'Trabar Embargo', decide según TIPO DE MEDIDA:
    IEI -> Trabar Intervención en Información
    DSE -> Trabar Depósito sin Extracción
    Solo abre la pantalla correspondiente.
    """
    _, tree = abrir_menu_individual()

    tipo_medida = str(expedientes.loc[i, "TIPO DE MEDIDA"]).strip().upper()

    if "IEI" in tipo_medida:
        clave = "IEI"
        patron = r"^Trabar Intervención en Información$"
    elif "DSE" in tipo_medida:
        clave = "DSE"
        patron = r"^Trabar Depósito sin Extracción$"
    else:
        expedientes.loc[i, "RESULTADO"] = "Tipo de medida no soportado"
        print(
            f"{expedientes.loc[i, 'EXPEDIENTE']}: tipo de medida no IEI/DSE ({tipo_medida})."
        )
        return

    opcion = tree.child_window(
        title_re=patron,
        control_type="TreeItem",
    ).wrapper_object()
    opcion.wait("enabled visible", timeout=15)
    opcion.double_click_input()

    print(
        f"{expedientes.loc[i, 'EXPEDIENTE']}: TIPO_DE_MEDIDA={tipo_medida} -> opción '{clave}' abierta."
    )

    # -----------------------------------------------------------------
    # TODO: A partir de aquí se abre un formulario específico (IEI o DSE).
    # Aquí deberías:
    #   - Llenar campos necesarios.
    #   - Guardar / Aceptar.
    #   - Cerrar la ventana y volver al menú individual.
    # -----------------------------------------------------------------


# ======================================================
# 5. BUCLE PRINCIPAL POR EXPEDIENTES
# ======================================================
for i, fila in expedientes.iterrows():
    expediente_actual = str(fila["EXPEDIENTE"]).strip()

    if not expediente_actual or expediente_actual.lower() == "nan":
        continue

    dep_completa = fila["DEPENDENCIA_COMPLETA"]
    if dep_completa == "DESCONOCIDO":
        expedientes.loc[i, "RESULTADO"] = "Dependencia desconocida"
        print(f"{expediente_actual}: dependencia desconocida, se omite.")
        continue

    print("\n==============================")
    print(f"Procesando expediente: {expediente_actual}")
    print("==============================")

    # 1) Abrir Selección de Expediente
    sel_vent = abrir_seleccion_expediente()

    # 2) Buscar expediente y validar ejecutor
    estado = leer_y_validar_ejecutor(sel_vent, i)

    if estado != "coincide":
        # Ya se cerró la ventana de selección dentro de la función
        # y RESULTADO quedó seteado, pasamos al siguiente expediente
        continue

    # 3) Proceso de Embargo -> Trabar Embargo
    abrir_proceso_embargo_y_trabar()

    # 4) IEI o DSE según TIPO DE MEDIDA
    abrir_opcion_por_tipo_medida(i)

    # IMPORTANTE:
    # Aquí podría ser necesario cerrar la pantalla que se abrió (IEI/DSE)
    # y luego cerrar el menú de 'Exp. Cob. Coactiva - Individual' (botón Cerrar)
    # para regresar al Menú de Opciones principal antes del siguiente loop.
    # Eso depende de cómo se comporten exactamente esas pantallas.


# ======================================================
# 6. GUARDAR CAMBIOS EN EL EXCEL
# ======================================================
expedientes.to_excel(excel_file, index=False)
print("\n Proceso finalizado. Excel actualizado con la columna 'RESULTADO'.")
print(f"\n[LOG] La salida del terminal se guardó en: {log_file_path}")

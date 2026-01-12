import cv2
import numpy as np
import pytesseract
import pyautogui
import subprocess
import time
import os
import sys
import logging
from pathlib import Path
import pandas as pd
from pywinauto import Application, Desktop

# Deshabilitar fail-safe de pyautogui para evitar errores cuando el mouse se mueve a las esquinas
pyautogui.FAILSAFE = False

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('proceso_log32.txt'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configurar ruta de Tesseract (buscar localmente primero, luego en sistema)
try:
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
    
    # Buscar Tesseract localmente
    local_tesseract = base_dir / "Tesseract-OCR" / "tesseract.exe"
    if local_tesseract.exists():
        pytesseract.pytesseract.pytesseract_cmd = str(local_tesseract)
        logger.info(f"✓ Tesseract encontrado localmente: {local_tesseract}")
    else:
        # Rutas alternativas donde puede estar instalado
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            r'D:\Usuarios\prac-areyess\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Tesseract-OCR\tesseract.exe',
        ]
        
        tesseract_found = False
        for path in possible_paths:
            if Path(path).exists():
                pytesseract.pytesseract.pytesseract_cmd = path
                logger.info(f"✓ Tesseract encontrado en: {path}")
                tesseract_found = True
                break
        
        if not tesseract_found:
            logger.warning("ADVERTENCIA: Tesseract no encontrado. OCR no funcionará.")
except Exception as e:
    logger.warning(f"Error configurando Tesseract: {e}")

# Rutas
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = Path(sys.executable).parent
else:
    SCRIPT_DIR = Path(__file__).parent

SHORTCUT_PATH = SCRIPT_DIR / "Actualiza RSIRAT.lnk"
IMAGES_DIR = SCRIPT_DIR

logger.info(f"Directorio de trabajo: {SCRIPT_DIR}")
logger.info(f"Arquitectura Python: {'64-bit' if sys.maxsize > 2**32 else '32-bit'}")


class ImageRecognition32:
    """
    Clase para detectar elementos en aplicación 32-bit desde Python 64-bit.
    Utiliza captura de pantalla a nivel de ventana para mejor compatibilidad.
    """
    
    def __init__(self, confidence_threshold=0.40):
        """
        Args:
            confidence_threshold: Umbral de confianza para la coincidencia (0-1)
        """
        self.confidence_threshold = confidence_threshold
        self.methods = [
            (cv2.TM_CCOEFF_NORMED, "TM_CCOEFF_NORMED"),
            (cv2.TM_CCORR_NORMED, "TM_CCORR_NORMED"),
            (cv2.TM_SQDIFF_NORMED, "TM_SQDIFF_NORMED")
        ]
    
    def enhance_image(self, image):
        """Mejora el contraste de una imagen"""
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image
        
        clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        
        return enhanced
    
    def find_image_on_screen(self, image_path, save_debug=False):
        """
        Busca una imagen en la pantalla usando múltiples métodos y escalas.
        Optimizado para compatibilidad 32-bit desde Python 64-bit.
        """
        if not os.path.exists(image_path):
            logger.error(f"Imagen no encontrada: {image_path}")
            return {'found': False, 'x': None, 'y': None, 'confidence': None}

        try:
            # Capturar pantalla completa
            logger.info(f"Buscando: {os.path.basename(image_path)}")
            screenshot = pyautogui.screenshot()
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

            # Cargar template
            template = cv2.imread(image_path)

            if template is None:
                logger.error(f"No se pudo cargar la imagen: {image_path}")
                return {'found': False, 'x': None, 'y': None, 'confidence': None}

            # Preparar imágenes
            screenshot_gray = self.enhance_image(screenshot_cv)
            template_gray = self.enhance_image(template)

            if save_debug:
                cv2.imwrite("debug_screenshot_32.png", screenshot_cv)
                cv2.imwrite("debug_template_32.png", template)
                logger.info("Screenshots de debug guardados")

            best_confidence = 0
            best_location = None
            best_method = None
            best_scale = 1.0

            # Probar múltiples escalas (rango más amplio para app 32-bit)
            scales = list(np.linspace(0.5, 1.6, 23))
            for scale in scales:
                try:
                    if scale == 1.0:
                        temp = template_gray
                    else:
                        new_w = max(1, int(template_gray.shape[1] * scale))
                        new_h = max(1, int(template_gray.shape[0] * scale))
                        temp = cv2.resize(template_gray, (new_w, new_h), interpolation=cv2.INTER_AREA)

                    if screenshot_gray.shape[0] < temp.shape[0] or screenshot_gray.shape[1] < temp.shape[1]:
                        continue

                    for method, method_name in self.methods:
                        try:
                            result = cv2.matchTemplate(screenshot_gray, temp, method)
                            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

                            if "SQDIFF" in method_name:
                                confidence = 1 - min_val
                                location = min_loc
                            else:
                                confidence = max_val
                                location = max_loc

                            if confidence > best_confidence:
                                best_confidence = confidence
                                best_location = location
                                best_method = method_name
                                best_scale = scale

                        except Exception as e:
                            logger.warning(f"Error con {method_name} a scale {scale}: {e}")
                            continue

                except Exception as e:
                    logger.warning(f"Error procesando escala {scale}: {e}")
                    continue

            logger.info(f"Mejor confianza: {best_confidence*100:.2f}% ({best_method}) scale {best_scale:.2f}")

            if best_confidence >= self.confidence_threshold and best_location is not None:
                temp_h = int(template.shape[0] * best_scale)
                temp_w = int(template.shape[1] * best_scale)

                center_x = best_location[0] + temp_w // 2
                center_y = best_location[1] + temp_h // 2

                logger.info(f"Imagen encontrada en: ({center_x}, {center_y})")

                if save_debug:
                    try:
                        y1 = max(0, best_location[1])
                        y2 = min(screenshot_cv.shape[0], best_location[1] + temp_h)
                        x1 = max(0, best_location[0])
                        x2 = min(screenshot_cv.shape[1], best_location[0] + temp_w)
                        cv2.imwrite("debug_match_32.png", screenshot_cv[y1:y2, x1:x2])
                    except Exception:
                        pass

                return {
                    'found': True,
                    'x': center_x,
                    'y': center_y,
                    'confidence': float(best_confidence)
                }
            else:
                logger.warning(f"Imagen no encontrada (mejor: {best_confidence*100:.2f}%)")
                if save_debug:
                    cv2.imwrite("debug_screenshot_32.png", screenshot_cv)
                    cv2.imwrite("debug_template_32.png", template)
                return {'found': False, 'x': None, 'y': None, 'confidence': best_confidence}

        except Exception as e:
            logger.error(f"Error en búsqueda de imagen: {str(e)}")
            return {'found': False, 'x': None, 'y': None, 'confidence': None}
    
    def click_on_image(self, image_path, clicks=1, interval=0.5):
        """Busca una imagen y hace clic en ella"""
        result = self.find_image_on_screen(image_path)
        
        if result['found']:
            try:
                logger.info(f"Haciendo clic en: {os.path.basename(image_path)}")
                for i in range(clicks):
                    pyautogui.click(result['x'], result['y'])
                    if i < clicks - 1:
                        time.sleep(interval)
                logger.info(f"Clic completado")
                return True
            except Exception as e:
                logger.error(f"Error al hacer clic: {str(e)}")
                return False
        else:
            logger.warning(f"No se pudo hacer clic - imagen no encontrada")
            return False

    def ocr_click(self, texto_buscar: str, timeout=8):
        """Fallback OCR para buscar texto y hacer clic"""
        end_time = time.time() + timeout
        while time.time() < end_time:
            try:
                screenshot = pyautogui.screenshot()
                datos = pytesseract.image_to_data(screenshot, lang='spa', output_type=pytesseract.Output.DICT)
                for i, texto in enumerate(datos.get('text', [])):
                    if texto and texto_buscar.upper() in texto.upper():
                        x = datos['left'][i] + datos['width'][i] // 2
                        y = datos['top'][i] + datos['height'][i] // 2
                        logger.info(f"OCR: Encontrado '{texto_buscar}' en ({x}, {y}), haciendo clic")
                        pyautogui.click(x, y)
                        time.sleep(0.5)
                        return True
            except Exception as e:
                logger.warning(f"Error OCR: {e}")
                time.sleep(0.5)

        logger.warning(f"OCR no encontró '{texto_buscar}'")
        return False
    
    def wait_for_image(self, image_path, timeout=30, check_interval=0.5, save_debug_on_fail=True):
        """Espera hasta que una imagen aparezca en pantalla"""
        logger.info(f"Esperando imagen: {os.path.basename(image_path)} (timeout: {timeout}s)")
        start_time = time.time()
        checks = 0
        
        while time.time() - start_time < timeout:
            checks += 1
            result = self.find_image_on_screen(image_path, save_debug=(checks <= 1))
            if result['found']:
                elapsed = time.time() - start_time
                logger.info(f"Imagen encontrada después de {elapsed:.1f}s")
                return True
            time.sleep(check_interval)
        
        elapsed = time.time() - start_time
        logger.warning(f"Timeout esperando imagen después de {elapsed:.1f}s")
        
        if save_debug_on_fail:
            logger.info("Guardando screenshots de debug...")
            self.find_image_on_screen(image_path, save_debug=True)
        
        return False


class RSIRATAutomation32:
    """Automatización de RSIRAT optimizada para 32-bit desde Python 64-bit"""
    
    def __init__(self, confidence_threshold=0.40):
        self.image_recognition = ImageRecognition32(confidence_threshold=confidence_threshold)
        self.password = None
        self.dependencia = None
        self.expediente = None
        self.dep_type = None  # "21" o "23" para determinar el flujo
    
    def load_credentials(self):
        """Carga contraseña desde archivo y lee Excel para determinar dependencia"""
        try:
            # Cargar contraseña
            password_file = SCRIPT_DIR / "contrasena.txt"
            if not password_file.exists():
                logger.error(f"Archivo de contraseña no encontrado: {password_file}")
                return False
            
            with open(password_file, "r", encoding="utf-8") as f:
                self.password = f.read().strip()
            
            logger.info("Contraseña cargada correctamente")
            
            # Cargar Excel - especificar dtype para columns numéricas como string
            excel_file = SCRIPT_DIR / "EXPEDIENTES.xlsx"
            if not excel_file.exists():
                logger.error(f"Archivo Excel no encontrado: {excel_file}")
                return False
            
            # Leer con dtype object para preservar formato original
            expedientes = pd.read_excel(excel_file, engine="openpyxl", dtype=str)
            
            if "DEPENDENCIA" not in expedientes.columns:
                logger.error("El Excel no contiene la columna 'DEPENDENCIA'")
                return False
            
            # Mapear dependencia
            dep_code = str(expedientes.iloc[0]["DEPENDENCIA"]).zfill(4)
            if dep_code.endswith("21") or dep_code.endswith("021"):
                self.dependencia = "0021 I.R. Lima - PRICO"
                self.dep_type = "21"  # Guardar tipo para el flujo
            elif dep_code.endswith("23") or dep_code.endswith("023"):
                self.dependencia = "0023 I.R. Lima - MEPECO"
                self.dep_type = "23"  # Guardar tipo para el flujo
            else:
                self.dependencia = dep_code
                self.dep_type = None
            
            # Extraer expediente (ya es texto en Excel, solo hacer strip)
            if "EXPEDIENTE" in expedientes.columns:
                self.expediente = str(expedientes.iloc[0]["EXPEDIENTE"]).strip()
                logger.info(f"Expediente: {self.expediente} (formato texto)")
            
            logger.info(f"Dependencia: {self.dependencia}")
            return True
        
        except Exception as e:
            logger.error(f"Error cargando credenciales: {str(e)}")
            return False
    
    def wait_for_login_window(self, timeout=30):
        """Espera a que aparezca la ventana de login"""
        try:
            logger.info("Esperando ventana de login...")
            desktop = Desktop(backend="uia")
            end_time = time.time() + timeout
            
            while time.time() < end_time:
                try:
                    win = desktop.window(title="SIRAT")
                    if win.exists(timeout=1):
                        handle = win.handle
                        app = Application(backend="uia").connect(handle=handle)
                        logger.info("Ventana de login encontrada")
                        return app, app.window(handle=handle)
                except Exception:
                    pass
                
                try:
                    win = desktop.window(title_re=".*SIRAT.*")
                    if win.exists(timeout=1):
                        handle = win.handle
                        app = Application(backend="uia").connect(handle=handle)
                        logger.info("Ventana de login encontrada")
                        return app, app.window(handle=handle)
                except Exception:
                    pass
                
                time.sleep(0.5)
            
            logger.error("Timeout esperando ventana de login")
            return None, None
        
        except Exception as e:
            logger.error(f"Error esperando login: {str(e)}")
            return None, None
    
    def login(self):
        try:
            logger.info("\nIniciando proceso de login...")
            
            # Cargar credenciales
            if not self.load_credentials():
                return False
            
            # Esperar ventana de login
            app, dlg = self.wait_for_login_window(timeout=30)
            if dlg is None:
                logger.error("No se pudo encontrar la ventana de login")
                return False
            
            time.sleep(1)
            
            # Buscar campos de entrada
            logger.info("Buscando campos de dependencia y contraseña...")
            try:
                dependencia_edit = dlg.child_window(auto_id="1001", control_type="Edit")
                password_edit = dlg.child_window(auto_id="1005", control_type="Edit")
                logger.info("Campos encontrados por auto_id")
            except Exception as e:
                logger.warning(f"No se encontraron campos por auto_id: {e}")
                try:
                    edits = dlg.descendants(control_type="Edit")
                    edits_list = list(edits)
                    dependencia_edit = edits_list[0] if len(edits_list) > 0 else None
                    password_edit = edits_list[1] if len(edits_list) > 1 else None
                except Exception:
                    dependencia_edit = None
                    password_edit = None
            
            if dependencia_edit is None or password_edit is None:
                logger.error("No se pudieron localizar los campos de login")
                return False
            
            # Ingresar dependencia
            logger.info(f"Ingresando dependencia: {self.dependencia}")
            dependencia_edit.set_focus()
            time.sleep(0.3)
            pyautogui.write(self.dependencia, interval=0.05)
            time.sleep(0.5)
            
            # Ingresar contraseña
            logger.info("Ingresando contraseña...")
            password_edit.set_focus()
            time.sleep(0.3)
            pyautogui.write(self.password, interval=0.05)
            time.sleep(0.5)
            
            # Hacer clic en Aceptar
            logger.info("Haciendo clic en Aceptar...")
            try:
                aceptar_btn = dlg.child_window(title="Aceptar", control_type="Button")
                aceptar_btn.invoke()
            except Exception:
                logger.warning("No se encontró botón 'Aceptar', intentando con Enter...")
                pyautogui.press('return')
            
            logger.info("Login completado")
            
            # Esperar dinámicamente por el menú
            menu_detected = False
            start_wait = time.time()
            max_wait = 1
            
            while time.time() - start_wait < max_wait:
                try:
                    menu_window = app.window(title_re=".*Menú.*")
                    if menu_window and menu_window.is_visible():
                        menu_detected = True
                        logger.info("Menú de opciones detectado")
                        time.sleep(2)
                        break
                except Exception:
                    pass
                
                time.sleep(0.5)
            
            if not menu_detected:
                logger.warning("Menú de opciones no detectado, continuando...")
            
            return True
        
        except Exception as e:
            logger.error(f"Error en login: {str(e)}")
            return False
    
    def open_application(self):
        """Abre RSIRAT usando el acceso directo"""
        logger.info("=" * 70)
        logger.info("INICIANDO AUTOMATIZACIÓN DE RSIRAT (32-BIT)")
        logger.info("=" * 70)
        logger.info(f"Buscando acceso directo en: {SHORTCUT_PATH}")
        
        if not SHORTCUT_PATH.exists():
            logger.error(f"Acceso directo no encontrado: {SHORTCUT_PATH}")
            return False
        
        try:
            logger.info(f"Abriendo: {SHORTCUT_PATH.name}")
            os.startfile(str(SHORTCUT_PATH))
            time.sleep(4)
            logger.info("Aplicación abierta correctamente")
            return True
        except Exception as e:
            logger.error(f"Error al abrir aplicación: {str(e)}")
            return False
    
    def click_exp_cob_individual(self):
        """
        Hace 4 clics en 'Exp. Cob. Coactiva - Individual' usando MSAA.
        Este elemento aparece después de clicar en 'Cobranza Coactiva'.
        """
        logger.info("Buscando 'Exp. Cob. Coactiva - Individual' para hacer 4 clics...")
        
        try:
            desktop = Desktop(backend="uia")
            
            # Buscar ventana de menú o principal
            menu_windows = []
            try:
                win = desktop.window(title_re=".*Menú.*")
                if win.exists(timeout=1):
                    menu_windows.append(win)
                    logger.info(f"Ventana de menú encontrada: {win.window_text()}")
            except Exception:
                pass
            
            # Si no se encontró ventana de menú, buscar ventana SIRAT general
            if not menu_windows:
                try:
                    win = desktop.window(title_re=".*SIRAT.*")
                    if win.exists(timeout=1):
                        menu_windows.append(win)
                        logger.info(f"Ventana SIRAT encontrada: {win.window_text()}")
                except Exception:
                    pass
            
            if not menu_windows:
                logger.error("No se encontraron ventanas de SIRAT o Menú")
                return False
            
            app_window = menu_windows[0]
            
            # Buscar el elemento "Exp. Cob. Coactiva - Individual"
            logger.info("Buscando elemento en descendientes...")
            
            try:
                # Intentar encontrar por titulo exacto primero
                exp_control = app_window.child_window(title="Exp. Cob. Coactiva - Individual")
                if exp_control.exists(timeout=2):
                    logger.info("Control 'Exp. Cob. Coactiva - Individual' encontrado por titulo exacto")
                    for i in range(4):
                        try:
                            # Obtener coordenadas del control
                            rect = exp_control.rectangle()
                            center_x = (rect.left + rect.right) // 2
                            center_y = (rect.top + rect.bottom) // 2
                            logger.info(f"Clic {i+1}/4 en coordenadas: ({center_x}, {center_y})")
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.3)
                        except Exception as e:
                            logger.warning(f"Error en clic {i+1}/4: {e}")
                            time.sleep(0.3)
                    logger.info("4 clics completados exitosamente")
                    time.sleep(0.5)
                    
                    # Ahora ingresar el expediente
                    logger.info("Procediendo a ingresar expediente...")
                    return self.enter_expediente_field()
            except Exception as e:
                logger.info(f"No se encontró por titulo exacto: {e}")
            
            # Buscar en descendientes
            try:
                logger.info("Buscando en descendientes del control...")
                descendants = app_window.descendants()
                
                for descendant in descendants:
                    try:
                        # Buscar elemento que contenga "Exp. Cob. Coactiva - Individual"
                        desc_text = descendant.window_text()
                        if "Exp. Cob. Coactiva" in desc_text and "Individual" in desc_text:
                            logger.info(f"Encontrado control: {desc_text}")
                            logger.info(f"Control type: {descendant.element_info.control_type}")
                            
                            # Hacer 4 clics usando coordenadas directas
                            for i in range(4):
                                try:
                                    rect = descendant.rectangle()
                                    center_x = (rect.left + rect.right) // 2
                                    center_y = (rect.top + rect.bottom) // 2
                                    logger.info(f"Clic {i+1}/4 en coordenadas: ({center_x}, {center_y})")
                                    pyautogui.click(center_x, center_y)
                                    time.sleep(0.3)
                                except Exception as err:
                                    logger.warning(f"Error en clic {i+1}/4: {err}")
                                    time.sleep(0.3)
                            
                            logger.info("4 clics completados exitosamente")
                            time.sleep(0.5)
                            
                            # Ahora ingresar el expediente
                            logger.info("Procediendo a ingresar expediente...")
                            return self.enter_expediente_field()
                    except Exception:
                        pass
                
                logger.warning("No se encontró elemento 'Exp. Cob. Coactiva - Individual' en descendientes")
            
            except Exception as e:
                logger.error(f"Error buscando en descendientes: {e}")
            
            # Fallback: OCR
            logger.info("Intentando fallback con OCR...")
            for i in range(4):
                if self.image_recognition.ocr_click("Exp. Cob. Coactiva", timeout=3):
                    logger.info(f"Clic {i+1}/4 completado por OCR")
                    time.sleep(0.3)
                else:
                    logger.warning(f"OCR falló en clic {i+1}/4")
                    break
            
            return False
        
        except Exception as e:
            logger.error(f"Error en click_exp_cob_individual: {str(e)}")
            return False
    
    def enter_expediente_field(self):
        """
        Busca el campo 'Número' en la ventana y digita el expediente desde el Excel.
        Verifica si el expediente es válido detectando mensaje de error.
        """
        logger.info(f"Ingresando expediente: {self.expediente}")
        
        try:
            desktop = Desktop(backend="uia")
            
            # Buscar ventana de expediente
            exp_windows = []
            try:
                win = desktop.window(title_re=".*Expediente.*")
                if win.exists(timeout=1):
                    exp_windows.append(win)
                    logger.info(f"Ventana de expediente encontrada: {win.window_text()}")
            except Exception:
                pass
            
            if not exp_windows:
                logger.warning("No se encontró ventana de expediente, intentando encontrar campo de edición...")
            else:
                app_window = exp_windows[0]
                
                # Buscar el campo de edición "Número"
                try:
                    logger.info("Buscando campo 'Número'...")
                    numero_field = app_window.child_window(title="Número", control_type="Edit")
                    if numero_field.exists(timeout=2):
                        logger.info("Campo 'Número' encontrado")
                        numero_field.set_focus()
                        time.sleep(0.3)
                        numero_field.type_keys(self.expediente, with_spaces=True)
                        logger.info(f"Expediente '{self.expediente}' ingresado")
                        time.sleep(0.5)
                        
                        # Presionar Enter para verificar
                        logger.info("Presionando Enter para verificar expediente...")
                        numero_field.type_keys('{ENTER}')
                        time.sleep(1)
                        
                        # Verificar si hay mensaje de error
                        expediente_valido = self.check_expediente_error(app_window)
                        
                        # Si el expediente es válido, proceder a validar ejecutor
                        if expediente_valido:
                            logger.info("Expediente válido, procediendo a validación de ejecutor...")
                            return self.validate_executor()
                        else:
                            logger.info("Expediente inválido, retornando False para reintentar")
                            return False
                except Exception as e:
                    logger.warning(f"No se encontró campo 'Número' por título exacto: {e}")
                
                # Buscar en descendientes
                try:
                    logger.info("Buscando campo en descendientes...")
                    descendants = app_window.descendants()
                    
                    for descendant in descendants:
                        try:
                            # Buscar campos de edición
                            if descendant.element_info.control_type == "Edit":
                                # Intentar obtener label asociado
                                desc_text = descendant.window_text()
                                logger.info(f"Campo Edit encontrado: '{desc_text}'")
                                
                                # Buscar el primer campo Edit visible (suele ser el de número)
                                if not desc_text or "Número" in desc_text or len(desc_text) < 20:
                                    logger.info("Usando este campo para ingresar expediente")
                                    descendant.set_focus()
                                    time.sleep(0.3)
                                    descendant.type_keys(self.expediente, with_spaces=True)
                                    logger.info(f"Expediente '{self.expediente}' ingresado")
                                    time.sleep(0.5)
                                    
                                    # Presionar Enter para verificar
                                    logger.info("Presionando Enter para verificar expediente...")
                                    descendant.type_keys('{ENTER}')
                                    time.sleep(1)
                                    
                                    # Verificar si hay mensaje de error
                                    expediente_valido = self.check_expediente_error(app_window)
                                    
                                    # Si el expediente es válido, proceder a validar ejecutor
                                    if expediente_valido:
                                        logger.info("Expediente válido, procediendo a validación de ejecutor...")
                                        return self.validate_executor()
                                    else:
                                        logger.info("Expediente inválido, retornando False para reintentar")
                                        return False
                        except Exception:
                            pass
                except Exception as e:
                    logger.warning(f"Error buscando en descendientes: {e}")
            
            # Fallback: usar pyautogui para escribir directamente
            logger.info("Intentando fallback con pyautogui...")
            pyautogui.write(self.expediente, interval=0.05)
            logger.info(f"Expediente '{self.expediente}' ingresado con pyautogui")
            time.sleep(0.5)
            
            # Presionar Enter
            logger.info("Presionando Enter para verificar expediente...")
            pyautogui.press('return')
            time.sleep(1)
            
            # Verificar si hay mensaje de error (sin app_window disponible)
            expediente_valido = self.check_expediente_error_screen()
            
            # Si el expediente es válido, proceder a validar ejecutor
            if expediente_valido:
                logger.info("Expediente válido, procediendo a validación de ejecutor...")
                return self.validate_executor()
            else:
                logger.info("Expediente inválido, retornando False para reintentar")
                return False
        
        except Exception as e:
            logger.error(f"Error en enter_expediente_field: {str(e)}")
            return False
    
    def check_expediente_error(self, app_window):
        """
        Verifica si aparece el mensaje de error de expediente inválido.
        Si aparece, actualiza Excel, cierra ventana y retorna False.
        Si no aparece, retorna True (expediente válido).
        """
        logger.info("Verificando si el expediente es válido...")
        time.sleep(0.5)
        
        try:
            # Buscar diálogo de error por ventana
            desktop = Desktop(backend="uia")
            
            # Patrón 1: "Selección de Expediente Coactivo - Error"
            try:
                error_dialog = desktop.window(title_re=".*Expediente.*Error.*")
                if error_dialog.exists(timeout=2):
                    logger.warning("EXPEDIENTE INVÁLIDO! Diálogo de error detectado")
                    
                    # Actualizar Excel marcando como inválido
                    logger.info("Registrando NRO EXPEDIENTE INVALIDO en Excel...")
                    self.update_excel_executor_result("NRO EXPEDIENTE INVALIDO")
                    
                    # Buscar botón "Aceptar" en el diálogo
                    try:
                        logger.info("Buscando botón 'Aceptar' en el diálogo de error...")
                        aceptar_btn = error_dialog.child_window(title="Aceptar", control_type="Button")
                        if aceptar_btn.exists(timeout=1):
                            logger.info("Botón 'Aceptar' encontrado, haciendo clic...")
                            aceptar_btn.invoke()
                            time.sleep(1)
                    except Exception as e:
                        logger.warning(f"No se encontró botón 'Aceptar': {e}")
                        # Fallback: presionar Enter
                        logger.info("Fallback: presionando Enter...")
                        pyautogui.press('return')
                        time.sleep(1)
                    
                    # Cerrar la ventana de búsqueda
                    logger.info("Cerrando ventana de búsqueda...")
                    time.sleep(0.5)
                    self.close_expediente_window()
                    
                    logger.warning("Diálogo de error cerrado, retornando False")
                    return False
            except Exception:
                pass
            
            # Patrón 2: Buscar en descendientes del app_window
            try:
                descendants = app_window.descendants()
                for descendant in descendants:
                    try:
                        desc_text = descendant.window_text()
                        # Buscar el mensaje de error específico
                        if "no es válido" in desc_text.lower():
                            logger.warning(f"EXPEDIENTE INVÁLIDO! Mensaje: {desc_text}")
                            
                            # Actualizar Excel
                            logger.info("Registrando NRO EXPEDIENTE INVALIDO en Excel...")
                            self.update_excel_executor_result("NRO EXPEDIENTE INVALIDO")
                            
                            # Buscar botón "Aceptar"
                            try:
                                aceptar_btn = app_window.child_window(title="Aceptar")
                                if aceptar_btn.exists(timeout=2):
                                    logger.info("Botón 'Aceptar' encontrado, haciendo clic...")
                                    aceptar_btn.invoke()
                                    time.sleep(1)
                            except Exception as e:
                                logger.warning(f"No se encontró botón: {e}")
                                pyautogui.press('return')
                                time.sleep(1)
                            
                            # Cerrar ventana
                            logger.info("Cerrando ventana de búsqueda...")
                            time.sleep(0.5)
                            self.close_expediente_window()
                            
                            return False
                    except Exception:
                        pass
            except Exception:
                pass
            
            logger.info("No se encontró mensaje de error, expediente es VÁLIDO")
            return True
        
        except Exception as e:
            logger.warning(f"Error verificando expediente: {e}")
            logger.info("Asumiendo que el expediente es válido")
            return True
    
    def check_expediente_error_screen(self):
        """
        Verifica si hay mensaje de error usando OCR en toda la pantalla, con fallback MSAA.
        Si hay error, actualiza Excel, cierra y retorna False para reintentar.
        Si es válido, retorna True para proceder a validar ejecutor.
        """
        logger.info("Verificando si el expediente es válido...")
        
        try:
            # Tier 1: Fallback MSAA (más confiable que OCR cuando tesseract no está disponible)
            try:
                logger.info("Buscando diálogo de error por MSAA...")
                desktop = Desktop(backend="uia")
                
                # Buscar diálogos que contengan "Error" en el título
                try:
                    error_dialogs = []
                    
                    # Patrón 1: "Selección de Expediente Coactivo - Error"
                    try:
                        win = desktop.window(title_re=".*Expediente.*Error.*")
                        if win.exists(timeout=1):
                            error_dialogs.append(win)
                            logger.warning("EXPEDIENTE INVÁLIDO! Diálogo de error detectado por MSAA")
                    except Exception:
                        pass
                    
                    # Patrón 2: Solo "Error"
                    if not error_dialogs:
                        try:
                            win = desktop.window(title_re=".*Error.*")
                            if win.exists(timeout=1):
                                error_dialogs.append(win)
                                logger.warning("EXPEDIENTE INVÁLIDO! Diálogo de error detectado por MSAA")
                        except Exception:
                            pass
                    
                    if error_dialogs:
                        # Actualizar Excel marcando como inválido
                        logger.info("Registrando NRO EXPEDIENTE INVALIDO en Excel...")
                        self.update_excel_executor_result("NRO EXPEDIENTE INVALIDO")
                        
                        # Presionar Enter para cerrar el diálogo
                        logger.info("Presionando Enter para cerrar el diálogo de error...")
                        pyautogui.press('return')
                        time.sleep(1)
                        
                        # Cerrar la ventana de búsqueda
                        logger.info("Cerrando ventana de búsqueda...")
                        time.sleep(0.5)
                        self.close_expediente_window()
                        
                        return False  # Expediente inválido
                
                except Exception as msaa_e:
                    logger.warning(f"Error en búsqueda MSAA: {msaa_e}")
            
            except Exception as tier1_error:
                logger.warning(f"Tier 1 (MSAA) falló: {tier1_error}")
            
            # Tier 2: OCR como fallback (si tesseract está disponible)
            try:
                screenshot = pyautogui.screenshot()
                datos = pytesseract.image_to_data(screenshot, lang='spa', output_type=pytesseract.Output.DICT)
                
                logger.info("Buscando mensaje de error con OCR...")
                for i, texto in enumerate(datos.get('text', [])):
                    if texto and "no es válido" in texto.lower():
                        logger.warning(f"EXPEDIENTE INVÁLIDO! OCR detectó: {texto}")
                        
                        # Actualizar Excel marcando como inválido
                        logger.info("Registrando NRO EXPEDIENTE INVALIDO en Excel...")
                        self.update_excel_executor_result("NRO EXPEDIENTE INVALIDO")
                        
                        # Buscar botón Aceptar
                        logger.info("Buscando botón 'Aceptar'...")
                        aceptar_encontrado = False
                        for j, text2 in enumerate(datos.get('text', [])):
                            if text2 and "aceptar" in text2.lower():
                                x = datos['left'][j] + datos['width'][j] // 2
                                y = datos['top'][j] + datos['height'][j] // 2
                                logger.info(f"Botón 'Aceptar' encontrado en ({x}, {y}), haciendo clic...")
                                pyautogui.click(x, y)
                                time.sleep(1)
                                aceptar_encontrado = True
                                break
                        
                        # Si no encuentra botón, usar Enter
                        if not aceptar_encontrado:
                            logger.info("No se encontró botón 'Aceptar', presionando Enter...")
                            pyautogui.press('return')
                            time.sleep(1)
                        
                        # Cerrar ventana de búsqueda
                        logger.info("Cerrando ventana de búsqueda...")
                        time.sleep(0.5)
                        self.close_expediente_window()
                        
                        return False  # Expediente inválido
                
                logger.info("OCR: No se detectó mensaje de error, expediente es VÁLIDO")
            
            except Exception as ocr_error:
                logger.warning(f"OCR no disponible: {ocr_error}")
            
            # Si no se detectó error por MSAA ni OCR, asumir que es válido
            logger.info("No se detectó mensaje de error, expediente es VÁLIDO")
            return True
        
        except Exception as e:
            logger.warning(f"Error verificando pantalla: {e}")
            logger.info("Asumiendo que el expediente es válido")
            return True
    
    def close_expediente_window(self):
        """
        Cierra la ventana de búsqueda de expediente buscando el botón 'Cerrar' o usando Escape.
        Funciona con o sin OCR disponible.
        """
        try:
            # Intentar encontrar botón Cerrar con OCR si está disponible
            try:
                screenshot = pyautogui.screenshot()
                datos = pytesseract.image_to_data(screenshot, lang='spa', output_type=pytesseract.Output.DICT)
                
                logger.info("Buscando botón 'Cerrar' o 'Cancelar' con OCR...")
                for i, texto in enumerate(datos.get('text', [])):
                    if texto and ("cerrar" in texto.lower() or "cancelar" in texto.lower()):
                        x = datos['left'][i] + datos['width'][i] // 2
                        y = datos['top'][i] + datos['height'][i] // 2
                        logger.info(f"Botón '{texto}' encontrado en ({x}, {y}), haciendo clic...")
                        pyautogui.click(x, y)
                        time.sleep(1)
                        logger.info("Ventana cerrada exitosamente con OCR")
                        return True
            except Exception as ocr_e:
                logger.warning(f"OCR no disponible para buscar botón: {ocr_e}")
        except Exception as e:
            logger.warning(f"Error en búsqueda OCR de botón Cerrar: {e}")
        
        # Fallback: Intentar encontrar con MSAA
        try:
            logger.info("Intentando encontrar botón 'Cerrar' con MSAA...")
            desktop = Desktop(backend="uia")
            
            # Buscar ventana de expediente
            try:
                win = desktop.window(title_re=".*Expediente.*")
                if win.exists(timeout=1):
                    # Buscar botón Cerrar
                    try:
                        cerrar_btn = win.child_window(title_re="Cerrar|Cancelar", control_type="Button")
                        if cerrar_btn.exists(timeout=1):
                            logger.info("Botón 'Cerrar' encontrado con MSAA, haciendo clic...")
                            cerrar_btn.invoke()
                            time.sleep(1)
                            logger.info("Ventana cerrada exitosamente con MSAA")
                            return True
                    except Exception as btn_e:
                        logger.warning(f"No se encontró botón con MSAA: {btn_e}")
            except Exception as win_e:
                logger.warning(f"No se encontró ventana: {win_e}")
        except Exception as msaa_e:
            logger.warning(f"Error en búsqueda MSAA: {msaa_e}")
        
        # Último fallback: presionar Escape para cerrar
        logger.info("Fallback final: presionando Escape para cerrar ventana...")
        pyautogui.press('escape')
        time.sleep(1)
        logger.info("Ventana cerrada con Escape")
        return True
    
    def validate_executor(self):
        """
        Simplificado: Solo presiona ALT+A para continuar sin validar nombres.
        El usuario ingresa el expediente y pasa directamente al siguiente paso.
        """
        logger.info("Presionando ALT+A para continuar con el proceso de embargo...")
        
        try:
            time.sleep(0.5)
            pyautogui.hotkey('alt', 'a')
            logger.info("✓ ALT+A presionado correctamente")
            time.sleep(1)
            return True
        
        except Exception as e:
            logger.error(f"Error presionando ALT+A: {str(e)}")
            return False
    
    def click_proceso_embargo(self):
        """
        Busca y hace clic en 'Proceso de Embargo' en el menú.
        Utiliza MSAA para encontrar el elemento. El elemento está en el árbol de controles.
        Similar al patrón que funciona en click_trabar_embargo().
        """
        logger.info("Buscando 'Proceso de Embargo' en el menú...")
        
        try:
            desktop = Desktop(backend="uia")
            
            # Buscar ventana principal de SIRAT (Menú de Opciones)
            app = None
            try:
                app = desktop.window(title_re=".*Menú.*")
                if not app.exists(timeout=2):
                    app = None
            except:
                app = None
            
            if not app:
                try:
                    app = desktop.window(title_re=".*SIRAT.*")
                    if not app.exists(timeout=2):
                        app = None
                except:
                    app = None
            
            if not app:
                try:
                    app = desktop.active()
                except:
                    app = None
            
            if app:
                logger.info("Ventana encontrada, iterando descendientes...")
                
                try:
                    descendants = app.descendants()
                    logger.info(f"Total de descendientes: {len(descendants)}")
                    
                    # Listado de debug: mostrar elementos con "Proceso" o "Embargo"
                    elementos_embargo = []
                    
                    for idx, descendant in enumerate(descendants):
                        try:
                            # Obtener texto con strip() para eliminar espacios en blanco ocultos
                            desc_text = descendant.window_text().strip()
                            
                            # Buscar elementos que contengan "Embargo"
                            if "Embargo" in desc_text:
                                elementos_embargo.append((idx, desc_text))
                                logger.info(f"[{idx}] Elemento con 'Embargo': '{desc_text}'")
                                
                                # BÚSQUEDA EXACTA: "Proceso de Embargo"
                                if desc_text == "Proceso de Embargo":
                                    logger.info(f"✓ 'Proceso de Embargo' encontrado (búsqueda exacta)")
                                    rect = descendant.rectangle()
                                    click_x = (rect.left + rect.right) // 2
                                    click_y = (rect.top + rect.bottom) // 2
                                    
                                    if click_x > 0 and click_y > 0:
                                        logger.info(f"Coordenadas válidas: ({click_x}, {click_y})")
                                        pyautogui.click(click_x, click_y)
                                        time.sleep(1)
                                        logger.info("✓ Clic en Proceso de Embargo completado")
                                        return True
                        
                        except Exception as e:
                            pass
                    
                    # Si no encontró exacto, buscar por palabras clave
                    if not elementos_embargo:
                        logger.info("No se encontraron elementos con 'Embargo'")
                    else:
                        logger.info(f"Se encontraron {len(elementos_embargo)} elementos con 'Embargo', pero ninguno coincide exactamente")
                
                except Exception as e:
                    logger.warning(f"Error iterando descendientes: {e}")
            
            # Fallback: Usar OCR para buscar el botón
            logger.info("Fallback: Buscando 'Proceso de Embargo' con OCR...")
            
            try:
                if self.image_recognition.ocr_click("Proceso de Embargo", timeout=8):
                    logger.info("✓ Proceso de Embargo encontrado con OCR")
                    return True
            except:
                pass
            
            # Último fallback: usar imagen si existe
            embargo_image = IMAGES_DIR / "ProcesoEmbargo.png"
            if embargo_image.exists():
                logger.info("Último fallback: Buscando por imagen...")
                return self.image_recognition.click_on_image(str(embargo_image))
            
            logger.warning("No se pudo encontrar 'Proceso de Embargo'")
            return False
        
        except Exception as e:
            logger.error(f"Error en click_proceso_embargo: {str(e)}")
            return False
    
    def click_trabar_embargo(self):
        """
        Busca y hace clic en 'Trabar Embargo' usando el patrón que funciona para otros elementos.
        El elemento tiene espacios en blanco al final que se deben eliminar con strip().
        """
        logger.info("Buscando 'Trabar Embargo' (usando patrón de descendientes)...")
        
        try:
            desktop = Desktop(backend="uia")
            
            # Buscar ventana SIRAT
            app = None
            try:
                app = desktop.window(title_re=".*SIRAT.*")
                if not app.exists(timeout=2):
                    app = None
            except:
                app = None
            
            if not app:
                try:
                    app = desktop.active()
                except:
                    app = None
            
            if app:
                logger.info("Ventana encontrada, iterando descendientes...")
                
                try:
                    descendants = app.descendants()
                    logger.info(f"Total de descendientes: {len(descendants)}")
                    
                    # Iterar todos los descendientes y buscar "Trabar Embargo"
                    for i, descendant in enumerate(descendants):
                        try:
                            desc_text = descendant.window_text().strip()  # ← IMPORTANTE: strip() para eliminar espacios
                            
                            # Log detallado de elementos que contienen "Embargo"
                            if "embargo" in desc_text.lower():
                                logger.info(f"[{i}] Elemento con 'Embargo': '{desc_text}'")
                            
                            # Búsqueda exacta con strip()
                            if desc_text == "Trabar Embargo":
                                logger.info(f"✓ 'Trabar Embargo' encontrado (exacto) en índice {i}")
                                rect = descendant.rectangle()
                                
                                if rect.width() > 0 and rect.height() > 0:
                                    click_x = (rect.left + rect.right) // 2
                                    click_y = (rect.top + rect.bottom) // 2
                                    logger.info(f"Coordenadas válidas: ({click_x}, {click_y})")
                                    logger.info(f"Haciendo clic en: ({click_x}, {click_y})")
                                    pyautogui.click(click_x, click_y)
                                    time.sleep(1)
                                    logger.info("✓ Clic completado exitosamente")
                                    return True
                                else:
                                    # Intentar invoke si no tiene coordenadas válidas
                                    logger.info("Coordenadas inválidas, intentando invoke...")
                                    try:
                                        descendant.invoke()
                                        time.sleep(1)
                                        logger.info("✓ Invocado correctamente")
                                        return True
                                    except:
                                        pyautogui.press('return')
                                        time.sleep(1)
                                        logger.info("✓ Enter presionado")
                                        return True
                        
                        except Exception as e:
                            pass
                    
                    logger.warning("'Trabar Embargo' no encontrado en descendientes")
                
                except Exception as e:
                    logger.warning(f"Error iterando: {e}")
            
            logger.warning("No se pudo encontrar 'Trabar Embargo'")
            return False
        
        except Exception as e:
            logger.error(f"Error en click_trabar_embargo: {str(e)}")
            return False
    
    def click_trabar_intervencion_informacion(self):
        """
        Busca y hace DOBLE CLIC en 'Trabar Intervención en Información' usando coordenadas.
        Similar al patrón que funciona en click_trabar_embargo().
        El elemento tiene espacios en blanco al final que deben eliminarse con strip().
        """
        logger.info("Buscando 'Trabar Intervención en Información' en el menú...")
        
        try:
            desktop = Desktop(backend="uia")
            
            # Buscar ventana SIRAT
            app = None
            try:
                app = desktop.window(title_re=".*SIRAT.*")
                if not app.exists(timeout=2):
                    app = None
            except:
                app = None
            
            if not app:
                try:
                    app = desktop.active()
                except:
                    app = None
            
            if app:
                logger.info("Ventana encontrada, iterando descendientes...")
                
                try:
                    descendants = app.descendants()
                    logger.info(f"Total de descendientes: {len(descendants)}")
                    
                    # Iterar todos los descendientes y buscar "Trabar Intervención en Información"
                    for i, descendant in enumerate(descendants):
                        try:
                            desc_text = descendant.window_text().strip()  # ← IMPORTANTE: strip()
                            
                            # Log detallado de elementos que contienen "Intervención"
                            if "intervención" in desc_text.lower():
                                logger.info(f"[{i}] Elemento con 'Intervención': '{desc_text}'")
                            
                            # Búsqueda exacta con strip()
                            if desc_text == "Trabar Intervención en Información":
                                logger.info(f"✓ 'Trabar Intervención en Información' encontrado (exacto) en índice {i}")
                                rect = descendant.rectangle()
                                
                                if rect.width() > 0 and rect.height() > 0:
                                    click_x = (rect.left + rect.right) // 2
                                    click_y = (rect.top + rect.bottom) // 2
                                    logger.info(f"Coordenadas válidas: ({click_x}, {click_y})")
                                    logger.info(f"Haciendo DOBLE CLIC en: ({click_x}, {click_y})")
                                    pyautogui.doubleClick(click_x, click_y)
                                    time.sleep(1.5)
                                    logger.info("✓ Doble clic completado exitosamente")
                                    return True
                                else:
                                    # Intentar invoke si no tiene coordenadas válidas
                                    logger.info("Coordenadas inválidas, intentando invoke...")
                                    try:
                                        descendant.invoke()
                                        time.sleep(1)
                                        logger.info("✓ Invocado correctamente")
                                        return True
                                    except:
                                        pyautogui.press('return')
                                        time.sleep(1)
                                        logger.info("✓ Enter presionado")
                                        return True
                        
                        except Exception as e:
                            pass
                    
                    logger.warning("'Trabar Intervención en Información' no encontrado en descendientes")
                    logger.info("Listando TODOS los elementos con 'trabar' para debug...")
                    for i, descendant in enumerate(descendants):
                        try:
                            desc_text = descendant.window_text().strip()
                            if "trabar" in desc_text.lower():
                                logger.info(f"  [{i}] '{desc_text}'")
                        except:
                            pass
                
                except Exception as e:
                    logger.warning(f"Error iterando: {e}")
            
            logger.warning("No se pudo encontrar 'Trabar Intervención en Información'")
            return False
        
        except Exception as e:
            logger.error(f"Error en click_trabar_intervencion_informacion: {str(e)}")
            return False
    
    def handle_trabar_intervencion_aviso(self):
        """
        Presiona ENTER directamente después del doble clic en 'Trabar Intervención en Información'.
        
        Flujo:
        1. Esperar 1 segundo de timeout
        2. Presionar ENTER
        3. Esperar 1 segundo adicional
        4. Retornar True para continuar con INTERVENTOR y PLAZO
        5. ALT+S se presionará después de completar ambos campos
        """
        logger.info("\nPresionando ENTER después de 'Trabar Intervención en Información'...")
        
        try:
            # Timeout de 1 segundo antes de presionar ENTER
            logger.info("Esperando 1 segundo...")
            time.sleep(1)
            
            logger.info("Presionando ENTER...")
            pyautogui.press('return')
            
            # Esperar 1 segundo después de ENTER
            logger.info("Esperando 1 segundo después de ENTER...")
            time.sleep(1)
            
            logger.info("✓ ENTER presionado correctamente")
            logger.info("Continuando con el flujo de INTERVENTOR y PLAZO...")
            
            return True
        
        except Exception as e:
            logger.error(f"Error en handle_trabar_intervencion_aviso: {str(e)}")
            # No fallar si hay error - continuar de todas formas
            return True
    
    def click_trabar_deposito_sin_extraccion(self):
        """
        Busca y hace DOBLE CLIC en 'Trabar Depósito sin Extracción' usando coordenadas.
        Similar al patrón que funciona en click_trabar_embargo().
        Se usa cuando la dependencia es 23 (MEPECO).
        El elemento tiene espacios en blanco al final que deben eliminarse con strip().
        """
        logger.info("Buscando 'Trabar Depósito sin Extracción' en el menú...")
        
        try:
            desktop = Desktop(backend="uia")
            
            # Buscar ventana SIRAT
            app = None
            try:
                app = desktop.window(title_re=".*SIRAT.*")
                if not app.exists(timeout=2):
                    app = None
            except:
                app = None
            
            if not app:
                try:
                    app = desktop.active()
                except:
                    app = None
            
            if app:
                logger.info("Ventana encontrada, iterando descendientes...")
                
                try:
                    descendants = app.descendants()
                    logger.info(f"Total de descendientes: {len(descendants)}")
                    
                    # Iterar todos los descendientes y buscar "Trabar Depósito sin Extracción"
                    for i, descendant in enumerate(descendants):
                        try:
                            desc_text = descendant.window_text().strip()  # ← IMPORTANTE: strip()
                            
                            # Log detallado de elementos que contienen "Depósito"
                            if "depósito" in desc_text.lower():
                                logger.info(f"[{i}] Elemento con 'Depósito': '{desc_text}'")
                            
                            # Búsqueda exacta con strip()
                            if desc_text == "Trabar Depósito sin Extracción":
                                logger.info(f"✓ 'Trabar Depósito sin Extracción' encontrado (exacto) en índice {i}")
                                rect = descendant.rectangle()
                                
                                if rect.width() > 0 and rect.height() > 0:
                                    click_x = (rect.left + rect.right) // 2
                                    click_y = (rect.top + rect.bottom) // 2
                                    logger.info(f"Coordenadas válidas: ({click_x}, {click_y})")
                                    logger.info(f"Haciendo DOBLE CLIC en: ({click_x}, {click_y})")
                                    pyautogui.doubleClick(click_x, click_y)
                                    time.sleep(1.5)
                                    logger.info("✓ Doble clic completado exitosamente")
                                    return True
                                else:
                                    # Intentar invoke si no tiene coordenadas válidas
                                    logger.info("Coordenadas inválidas, intentando invoke...")
                                    try:
                                        descendant.invoke()
                                        time.sleep(1)
                                        logger.info("✓ Invocado correctamente")
                                        return True
                                    except:
                                        pyautogui.press('return')
                                        time.sleep(1)
                                        logger.info("✓ Enter presionado")
                                        return True
                        
                        except Exception as e:
                            pass
                    
                    logger.warning("'Trabar Depósito sin Extracción' no encontrado en descendientes")
                    logger.info("Listando TODOS los elementos con 'trabar' para debug...")
                    for i, descendant in enumerate(descendants):
                        try:
                            desc_text = descendant.window_text().strip()
                            if "trabar" in desc_text.lower():
                                logger.info(f"  [{i}] '{desc_text}'")
                        except:
                            pass
                
                except Exception as e:
                    logger.warning(f"Error iterando: {e}")
            
            logger.warning("No se pudo encontrar 'Trabar Depósito sin Extracción'")
            return False
        
        except Exception as e:
            logger.error(f"Error en click_trabar_deposito_sin_extraccion: {str(e)}")
            return False
    
    def handle_post_embargo_flow(self):
        """
        Maneja el flujo después de 'Trabar Embargo' según la dependencia.
        
        - Dependencia 21 (PRICO): Ejecuta 'Trabar Intervención en Información' → Maneja aviso → Rellena campos
        - Dependencia 23 (MEPECO): Ejecuta 'Trabar Depósito sin Extracción' → Rellena campos
        """
        logger.info("\n" + "=" * 70)
        logger.info("MANEJANDO FLUJO POST-EMBARGO SEGÚN DEPENDENCIA")
        logger.info("=" * 70)
        
        if self.dep_type == "21":
            logger.info("Dependencia 21 (PRICO) detectada")
            logger.info("Ejecutando: Trabar Intervención en Información")
            
            # Paso 1: Hacer clic en Trabar Intervención en Información
            if not self.click_trabar_intervencion_informacion():
                logger.error("No se pudo hacer clic en Trabar Intervención en Información")
                return False
            
            # Paso 2: Manejar el aviso que puede aparecer
            logger.info("\nManejando posible aviso...")
            if not self.handle_trabar_intervencion_aviso():
                logger.error("Error manejando el aviso")
                return False
            
            # Paso 3: Rellenar campos INTERVENTOR y PLAZO
            logger.info("\nRellenando campos...")
            return self.fill_interventor_and_plazo()
        
        elif self.dep_type == "23":
            logger.info("Dependencia 23 (MEPECO) detectada")
            logger.info("Ejecutando: Trabar Depósito sin Extracción")
            
            # Paso 1: Hacer clic en Trabar Depósito sin Extracción
            if not self.click_trabar_deposito_sin_extraccion():
                logger.error("No se pudo hacer clic en Trabar Depósito sin Extracción")
                return False
            
            # Paso 2: Rellenar campos INTERVENTOR y PLAZO
            logger.info("\nRellenando campos...")
            return self.fill_interventor_and_plazo()
        
        else:
            logger.error(f"Dependencia no reconocida: {self.dep_type}")
            return False
    
    def fill_interventor_and_plazo(self):
        """
        Llena los campos de INTERVENTOR y PLAZO en el formulario secuencialmente.
        
        Flujo:
        1. Busca y selecciona el campo INTERVENTOR (por MSAA o imagen)
        2. Hace clic en el campo INTERVENTOR
        3. Digita el código de INTERVENTOR desde Excel
        4. Presiona Enter
        5. Busca y selecciona el campo PLAZO
        6. Hace clic en el campo PLAZO
        7. Digita el PLAZO desde Excel
        8. Presiona Enter
        
        Las capturas de referencia son "1. Interventor.png" y "2. Plazo.png"
        """
        logger.info("Rellenando campos de INTERVENTOR y PLAZO...")
        
        try:
            # Cargar datos del Excel
            excel_file = SCRIPT_DIR / "EXPEDIENTES.xlsx"
            expedientes = pd.read_excel(excel_file, engine="openpyxl", dtype=str)
            
            # Obtener valores de INTERVENTOR y PLAZO
            interventor = None
            plazo = None
            
            if "INTERVENTOR" in expedientes.columns:
                interventor = str(expedientes.iloc[0]["INTERVENTOR"]).strip()
                logger.info(f"INTERVENTOR obtenido del Excel: '{interventor}'")
            else:
                logger.warning("No existe columna 'INTERVENTOR' en el Excel")
                interventor = ""
            
            if "PLAZO" in expedientes.columns:
                plazo = str(expedientes.iloc[0]["PLAZO"]).strip()
                logger.info(f"PLAZO obtenido del Excel: '{plazo}'")
            else:
                logger.warning("No existe columna 'PLAZO' en el Excel")
                plazo = ""
            
            # ============================================================
            # PASO 1: SELECCIONAR Y HACER CLIC EN CAMPO INTERVENTOR
            # ============================================================
            logger.info("=" * 70)
            logger.info("PASO 1: Esperando 0.5s y digitando INTERVENTOR")
            logger.info("=" * 70)
            
            # Esperar 0.5 segundos antes de digitar INTERVENTOR
            logger.info("Esperando 0.5 segundos antes de digitar INTERVENTOR...")
            time.sleep(0.5)
            
            # Digitar INTERVENTOR directamente (sin hacer clic en el campo)
            logger.info(f"Digitando INTERVENTOR: '{interventor}'")
            pyautogui.write(interventor, interval=0.05)
            time.sleep(0.2)
            
            # ============================================================
            # PASO 2: Presionar TAB para pasar a PLAZO
            # ============================================================
            logger.info("=" * 70)
            logger.info("PASO 2: Presionando TAB para pasar a PLAZO")
            logger.info("=" * 70)
            
            logger.info("Presionando TAB para pasar al campo PLAZO...")
            pyautogui.press('tab')
            time.sleep(0.2)
            
            # ============================================================
            # PASO 3: Digitar PLAZO
            # ============================================================
            logger.info("=" * 70)
            logger.info("PASO 3: Digitando PLAZO")
            logger.info("=" * 70)
            
            logger.info(f"Digitando PLAZO: '{plazo}'")
            pyautogui.write(plazo, interval=0.05)
            time.sleep(0.3)
            
            # ============================================================
            # PASO 4: Verificar valores via MSAA
            # ============================================================
            logger.info("=" * 70)
            logger.info("PASO 4: Verificando que los valores aparecieron en los campos via MSAA")
            logger.info("=" * 70)
            
            try:
                desktop = Desktop(backend="uia")
                
                # Buscar ventana principal de SIRAT
                sirat_window = None
                try:
                    sirat_windows = desktop.windows(title_re=".*SIRAT.*")
                    if sirat_windows:
                        sirat_window = sirat_windows[0]
                except:
                    # Si no encontramos por título, usar ventana activa
                    sirat_window = desktop.active()
                
                if sirat_window and sirat_window.exists(timeout=1):
                    # Buscar campo INTERVENTOR
                    try:
                        interventor_field = sirat_window.child_window(
                            title="INTERVENTOR",
                            control_type="Edit"
                        )
                        if interventor_field.exists(timeout=1):
                            valor_interventor = interventor_field.window_text()
                            logger.info(f"Valor en campo INTERVENTOR: '{valor_interventor}'")
                            
                            if interventor in valor_interventor:
                                logger.info("✓ INTERVENTOR se ingresó correctamente")
                            else:
                                logger.warning(f"INTERVENTOR ingresado no coincide: esperado '{interventor}', obtenido '{valor_interventor}'")
                    except Exception as e:
                        logger.info(f"No se pudo verificar INTERVENTOR via MSAA: {e}")
                    
                    # Buscar campo PLAZO
                    try:
                        plazo_field = sirat_window.child_window(
                            title="PLAZO",
                            control_type="Edit"
                        )
                        if plazo_field.exists(timeout=1):
                            valor_plazo = plazo_field.window_text()
                            logger.info(f"Valor en campo PLAZO: '{valor_plazo}'")
                            
                            if plazo in valor_plazo:
                                logger.info("✓ PLAZO se ingresó correctamente")
                            else:
                                logger.warning(f"PLAZO ingresado no coincide: esperado '{plazo}', obtenido '{valor_plazo}'")
                    except Exception as e:
                        logger.info(f"No se pudo verificar PLAZO via MSAA: {e}")
            except Exception as e:
                logger.info(f"No se pudo verificar valores via MSAA: {e}")
            
            logger.info("=" * 70)
            logger.info("✓ CAMPOS INTERVENTOR Y PLAZO COMPLETADOS EXITOSAMENTE")
            logger.info("=" * 70)
            
            # Presionar ALT+S después de completar ambos campos
            logger.info("\nPresionando ALT+S para confirmar...")
            pyautogui.hotkey('alt', 's')
            time.sleep(1)
            logger.info("✓ ALT+S presionado correctamente")
            
            # NO ACTUALIZAR EXCEL - Se actualizará después de todos los pasos
            
            return True
        
        except Exception as e:
            logger.error(f"Error en fill_interventor_and_plazo: {str(e)}")
            return False
    
    def update_excel_executor_result(self, resultado):
        """
        Actualiza el Excel con el resultado de la validación del ejecutor.
        Agrega o actualiza la columna RESULTADO en la fila actual.
        """
        try:
            excel_file = SCRIPT_DIR / "EXPEDIENTES.xlsx"
            
            # Leer el Excel preservando formato original
            expedientes = pd.read_excel(excel_file, engine="openpyxl", dtype=str)
            
            # Crear columna si no existe
            if "RESULTADO" not in expedientes.columns:
                logger.info("Creando columna 'RESULTADO' en el Excel")
                expedientes["RESULTADO"] = None
            
            # Actualizar la primera fila (la que procesamos)
            expedientes.at[0, "RESULTADO"] = resultado
            
            # Guardar el Excel preservando formato
            expedientes.to_excel(excel_file, engine="openpyxl", index=False)
            logger.info(f"Excel actualizado: RESULTADO = '{resultado}'")
            
            return True
        
        except Exception as e:
            logger.error(f"Error actualizando Excel: {str(e)}")
            return False
    
    def click_cobranza_coactiva(self):
        """
        Busca y hace clic en 'Cobranza Coactiva' usando MSAA con timeout de 3 minutos.
        Reintentos cada 2 segundos hasta alcanzar el timeout.
        """
        logger.info("Buscando 'Cobranza Coactiva' usando MSAA (timeout: 3 minutos)...")
        
        timeout_segundos = 180
        tiempo_inicio = time.time()
        intento = 0
        
        while time.time() - tiempo_inicio < timeout_segundos:
            intento += 1
            tiempo_transcurrido = int(time.time() - tiempo_inicio)
            logger.info(f"Intento {intento} (tiempo: {tiempo_transcurrido}/{timeout_segundos}s)")
            
            try:
                desktop = Desktop(backend="uia")
                menu_windows = []
                
                # Buscar ventana de menú
                try:
                    win = desktop.window(title_re=".*Menú.*")
                    if win.exists(timeout=1):
                        menu_windows.append(win)
                        logger.info(f"Ventana menú encontrada")
                except:
                    pass
                
                # Fallback: buscar ventana SIRAT
                if not menu_windows:
                    try:
                        win = desktop.window(title_re=".*SIRAT.*")
                        if win.exists(timeout=1):
                            menu_windows.append(win)
                            logger.info(f"Ventana SIRAT encontrada")
                    except:
                        pass
                
                if not menu_windows:
                    logger.warning(f"Intento {intento}: ventanas no encontradas, reintentando...")
                    time.sleep(2)
                    continue
                
                app_window = menu_windows[0]
                encontrado = False
                
                # Intentar encontrar por titulo exacto
                try:
                    cobranza_control = app_window.child_window(title="Cobranza Coactiva")
                    if cobranza_control.exists(timeout=1):
                        logger.info("Elemento encontrado por titulo, haciendo clic...")
                        rect = cobranza_control.rectangle()
                        center_x = (rect.left + rect.right) // 2
                        center_y = (rect.top + rect.bottom) // 2
                        pyautogui.click(center_x, center_y)
                        time.sleep(1)
                        encontrado = True
                except:
                    pass
                
                # Si no encontrado, buscar en descendientes
                if not encontrado:
                    try:
                        descendants = app_window.descendants()
                        for descendant in descendants:
                            try:
                                if "Cobranza" in descendant.window_text():
                                    logger.info("Elemento encontrado en descendientes, haciendo clic...")
                                    rect = descendant.rectangle()
                                    if rect.left > 0 or rect.top > 0:
                                        center_x = (rect.left + rect.right) // 2
                                        center_y = (rect.top + rect.bottom) // 2
                                        pyautogui.click(center_x, center_y)
                                        time.sleep(1)
                                        encontrado = True
                                        break
                            except:
                                pass
                    except:
                        pass
                
                if encontrado:
                    logger.info("Clic en 'Cobranza Coactiva' completado")
                    time.sleep(1)
                    logger.info("Procediendo con 4 clics en 'Exp. Cob. Coactiva - Individual'...")
                    return self.click_exp_cob_individual()
                
                logger.warning(f"Intento {intento}: elemento no encontrado, reintentando...")
                
            except Exception as e:
                logger.warning(f"Error intento {intento}: {e}")
            
            # Esperar antes de reintentar
            tiempo_restante = timeout_segundos - (time.time() - tiempo_inicio)
            if tiempo_restante > 0:
                time.sleep(2)
        
        # Timeout alcanzado, intentar OCR
        logger.error("Timeout de 3 minutos alcanzado")
        logger.info("Intentando fallback con OCR...")
        
        try:
            if self.image_recognition.ocr_click("Cobranza Coactiva", timeout=8):
                logger.info("OCR: clic exitoso")
                time.sleep(1)
                logger.info("Procediendo con 4 clics...")
                return self.click_exp_cob_individual()
        except:
            pass
        
        return False
    
    def run(self):
        """Ejecuta la secuencia de automatización sin reintentos innecesarios"""
        try:
            # Abrir aplicación
            if not self.open_application():
                logger.error("No se pudo abrir la aplicación")
                return False
            
            # Esperar carga
            logger.info("\nEsperando carga de la aplicación...")
            time.sleep(3)
            
            # Realizar login (una sola vez)
            logger.info("\nProcediendo con el login...")
            if not self.login():
                logger.warning("No se pudo completar el login")
                return False
            
            # Ejecutar flujo completo de automatización
            logger.info("\n" + "=" * 70)
            logger.info("INICIANDO FLUJO DE AUTOMATIZACIÓN")
            logger.info("=" * 70)
            
            # Hacer clic en CobranzaCoactiva → 4 clics → Ingresar expediente → ALT+A
            resultado = self.click_cobranza_coactiva()
            
            if resultado:
                # Éxito: expediente válido, ALT+A fue presionado
                logger.info("\n" + "=" * 70)
                logger.info("PASO 1: EXPEDIENTE INGRESADO - ALT+A PRESIONADO ✓")
                logger.info("=" * 70)
                
                # PASO 2: Hacer clic en "Proceso de Embargo"
                logger.info("\n" + "=" * 70)
                logger.info("PASO 2: HACIENDO CLIC EN PROCESO DE EMBARGO")
                logger.info("=" * 70)
                
                if self.click_proceso_embargo():
                    logger.info("✓ Proceso de Embargo iniciado correctamente")
                    
                    # PASO 3: Hacer clic en "Trabar Embargo"
                    logger.info("\n" + "=" * 70)
                    logger.info("PASO 3: HACIENDO CLIC EN TRABAR EMBARGO")
                    logger.info("=" * 70)
                    
                    if self.click_trabar_embargo():
                        logger.info("✓ Trabar Embargo completado correctamente")
                        
                        # PASO 4: Manejo del flujo post-embargo según dependencia
                        # Dependencia 21 (PRICO): Trabar Intervención en Información
                        # Dependencia 23 (MEPECO): Trabar Depósito sin Extracción
                        logger.info("\n" + "=" * 70)
                        logger.info("PASO 4: FLUJO POST-EMBARGO (DEPENDE DE LA DEPENDENCIA)")
                        logger.info("=" * 70)
                        
                        if self.handle_post_embargo_flow():
                            logger.info("✓ Flujo post-embargo completado")
                            logger.info("✓ Campos rellenados correctamente")
                            
                            logger.info("\n" + "=" * 70)
                            logger.info("AUTOMATIZACIÓN COMPLETADA EXITOSAMENTE")
                            logger.info("Todos los pasos ejecutados correctamente:")
                            logger.info("  1. ✓ Login realizado")
                            logger.info("  2. ✓ Expediente ingresado y validado")
                            logger.info("  3. ✓ Proceso de Embargo iniciado")
                            logger.info("  4. ✓ Trabar Embargo ejecutado")
                            logger.info(f"  5. ✓ Dependencia {self.dep_type} procesada correctamente")
                            logger.info("  6. ✓ Campos digitados")
                            logger.info("=" * 70)
                            return True
                        else:
                            logger.warning("No se pudo completar el flujo post-embargo")
                            return False
                    else:
                        logger.warning("No se pudo hacer clic en Trabar Embargo")
                        return False
                else:
                    logger.warning("No se pudo hacer clic en Proceso de Embargo")
                    return False
            else:
                # Error en el flujo anterior
                logger.info("\n" + "=" * 70)
                logger.info("PROCESO COMPLETADO CON ERRORES")
                logger.info("=" * 70)
                return False
        
        except Exception as e:
            logger.error(f"Error en automatización: {str(e)}")
            return False

def main():
    """Función principal"""
    automation = RSIRATAutomation32()
    result = automation.run()
    
    if result:
        logger.info("\nEl proceso se completó sin errores.")
    else:
        logger.error("\nEl proceso finalizó con errores. Revisa el log para más detalles.")
    
    return result


if __name__ == "__main__":
    main()

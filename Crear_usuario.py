import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import logging
from dotenv import load_dotenv
import json
from typing import Dict, Optional, Tuple, List

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/esi_clinic_automation.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuración inicial
CONFIG = {
    'BASE_URL': "https://esiclinic.com/",
    'EXCEL_PATH': "data/clientes/pacientes.xlsx",
    'SCREENSHOT_DIR': "data/screenshots",
    'WAIT_TIMEOUT': 15,
    'SHORT_WAIT': 5,
    'IMPLICIT_WAIT': 3,
    'HEADLESS': False,
    'RETRY_ATTEMPTS': 3,
    'RETRY_DELAY': 2
}

# Cargar configuración adicional desde archivo si existe
try:
    with open('config.json') as config_file:
        CONFIG.update(json.load(config_file))
except FileNotFoundError:
    logger.info("No se encontró archivo config.json, usando configuración por defecto")

# Cargar variables de entorno
load_dotenv("env/.env")

class ESIClinicPageObjects:
    """Clase para mantener todos los selectores de la página"""
    LOGIN = {
        'username_field': (By.ID, "esi_user"),
        'password_field': (By.ID, "esi_pass"),
        'login_button': (By.ID, "bt_acceder")
    }
    
    PATIENT_FORM = {
        'new_patient_button': (By.CSS_SELECTOR, 'button#bt_nuevo, #bt_nuevo, [title*="Añadir nuevo"]'),
        'name_field': (By.ID, "Tnombre"),
        'lastname_field': (By.ID, "Tapellidos"),
        'dni_field': (By.ID, "TCIF"),
        'phone_field': (By.ID, "Tmovil"),
        'email_field': (By.ID, "Temail"),
        'birthdate_field': (By.ID, "Tfechadenacimiento"),
        'save_button': (By.ID, "guardarRegistro")
    }
    
    MODALS = {
        'confirmation_modal': (By.CSS_SELECTOR, ".jconfirm-scrollpane"),
        'confirm_button': (By.CSS_SELECTOR, ".btn-confirm, .jconfirm-buttons button"),
        'close_button': (By.CSS_SELECTOR, ".jconfirm-closeIcon, .close-button"),
        'success_message': (By.CSS_SELECTOR, ".alert-success, .success-message"),
        'error_message': (By.CSS_SELECTOR, ".error-message, .alert-danger")
    }

class ESIClinicAutomator:
    def __init__(self):
        """Inicializa el automatizador con configuración predeterminada"""
        self._setup_driver()
        self.wait = WebDriverWait(self.driver, CONFIG['WAIT_TIMEOUT'])
        self.page = ESIClinicPageObjects()
        self._create_dirs()
        
    def _setup_driver(self):
        """Configura el driver de Chrome con opciones"""
        self.options = webdriver.ChromeOptions()
        
        if CONFIG['HEADLESS']:
            self.options.add_argument("--headless=new")
            self.options.add_argument("--window-size=1920,1080")
        
        self.options.add_argument("--no-sandbox")
        self.options.add_argument("--disable-dev-shm-usage")
        self.options.add_argument("--disable-gpu")
        self.options.add_argument("--start-maximized")
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=self.options)
        self.driver.implicitly_wait(CONFIG['IMPLICIT_WAIT'])
        
    def _create_dirs(self):
        """Crea directorios necesarios si no existen"""
        os.makedirs(CONFIG['SCREENSHOT_DIR'], exist_ok=True)
        os.makedirs(os.path.dirname(CONFIG['EXCEL_PATH']), exist_ok=True)
        
    def close(self):
        """Cierra el navegador y realiza limpieza"""
        if hasattr(self, 'driver'):
            self.driver.quit()
            logger.info("Navegador cerrado correctamente")
            
    def _take_screenshot(self, prefix: str = "error"):
        """Toma un screenshot y guarda con timestamp"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(
            CONFIG['SCREENSHOT_DIR'], 
            f"{prefix}_{timestamp}.png"
        )
        self.driver.save_screenshot(screenshot_path)
        logger.info(f"Screenshot guardado en: {screenshot_path}")
        return screenshot_path
        
    def _retry_on_failure(self, func, *args, **kwargs):
        """Intenta ejecutar una función varias veces antes de fallar"""
        for attempt in range(CONFIG['RETRY_ATTEMPTS']):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                if attempt == CONFIG['RETRY_ATTEMPTS'] - 1:
                    raise
                logger.warning(f"Intento {attempt + 1} fallido, reintentando... Error: {str(e)}")
                time.sleep(CONFIG['RETRY_DELAY'])
                
    def login(self) -> bool:
        """Realiza el login en ESIClinic
        
        Returns:
            bool: True si el login fue exitoso, False en caso contrario
        """
        try:
            logger.info("Accediendo a ESIClinic...")
            self.driver.get(CONFIG['BASE_URL'])
            
            # Ingresar credenciales
            self.wait.until(EC.presence_of_element_located(
                self.page.LOGIN['username_field']
            )).send_keys(os.getenv("USUARIO_ESICLINIC"))
            
            self.driver.find_element(*self.page.LOGIN['password_field']).send_keys(
                os.getenv("PASSWORD_ESICLINIC")
            )
            
            # Click en login
            self.driver.find_element(*self.page.LOGIN['login_button']).click()
            
            # Esperar redirección
            self.wait.until(EC.url_contains("agenda.php"))
            logger.info("Login exitoso")
            return True
            
        except Exception as e:
            logger.error(f"Error en login: {str(e)}")
            self._take_screenshot("login_error")
            return False

    @staticmethod
    def validate_patient_data(patient_data: Dict) -> List[str]:
        """Valida los datos del paciente
        
        Args:
            patient_data: Diccionario con los datos del paciente
            
        Returns:
            Lista de mensajes de error, vacía si no hay errores
        """
        errors = []
        required_fields = {
            'nombre': "Nombre es obligatorio",
            'apellidos': "Apellidos son obligatorios",
            'dni': "DNI/NIF es obligatorio",
            'movil': "Móvil es obligatorio",
            'email': "Email es obligatorio"
        }
        
        # Validar campos obligatorios
        for field, error_msg in required_fields.items():
            if not patient_data.get(field):
                errors.append(error_msg)
        
        # Validar formato de fecha
        if patient_data.get('fecha_nacimiento'):
            try:
                datetime.strptime(patient_data['fecha_nacimiento'], '%d-%m-%Y')
            except ValueError:
                errors.append("Formato de fecha inválido. Use dd-mm-yyyy")
        
        # Validar email
        if patient_data.get('email') and '@' not in patient_data['email']:
            errors.append("Email no válido")
        
        return errors

    @staticmethod
    def check_excel_duplicates(patient_data: Dict) -> Tuple[bool, Optional[str]]:
        """Verifica duplicados en el archivo Excel
        
        Args:
            patient_data: Diccionario con los datos del paciente
            
        Returns:
            Tupla (allow_create, message) donde:
            - allow_create: False si no se debe crear el paciente por duplicado
            - message: Mensaje descriptivo del resultado
        """
        try:
            if not os.path.exists(CONFIG['EXCEL_PATH']):
                logger.warning("Archivo Excel no encontrado, se omitirá verificación de duplicados")
                return True, None
                
            try:
                # Intentar leer con diferentes motores por compatibilidad
                try:
                    df = pd.read_excel(CONFIG['EXCEL_PATH'])
                except:
                    df = pd.read_excel(CONFIG['EXCEL_PATH'], engine='openpyxl')
                
                # Verificar estructura del DataFrame
                required_columns = {'CIF', 'E-Mail'}
                if not required_columns.issubset(df.columns):
                    logger.warning("El archivo Excel no tiene las columnas esperadas")
                    return True, None
                
                # Buscar duplicados
                dni_duplicate = not df[df['CIF'].str.lower() == patient_data['dni'].lower()].empty
                email_duplicate = not df[df['E-Mail'].str.lower() == patient_data['email'].lower()].empty
                
                if dni_duplicate:
                    return False, "Error: Este DNI ya existe en la base de datos"
                if email_duplicate:
                    return True, "Advertencia: Este email ya está registrado"
                    
                return True, None
                
            except Exception as e:
                logger.error(f"Error al leer archivo Excel: {str(e)}")
                return True, None
            
        except Exception as e:
            logger.error(f"No se pudo verificar duplicados: {str(e)}")
            return True, None

    def _handle_modal(self):
        """Maneja los modales que aparecen después de guardar"""
        try:
            modal = WebDriverWait(self.driver, CONFIG['SHORT_WAIT']).until(
                EC.visibility_of_element_located(self.page.MODALS['confirmation_modal']))
            
            logger.info("Modal de confirmación detectado")
            
            # Intentar hacer clic en el botón de confirmación
            try:
                confirm_button = modal.find_element(*self.page.MODALS['confirm_button'])
                confirm_button.click()
                logger.info("Clic en botón de confirmación realizado")
                return
            except:
                pass
                
            # Intentar hacer clic en cerrar
            try:
                close_button = modal.find_element(*self.page.MODALS['close_button'])
                close_button.click()
                logger.info("Clic en botón cerrar (X) realizado")
                return
            except:
                pass
                
            # Fallback: clic fuera del modal
            ActionChains(self.driver).move_by_offset(20, 10).click().perform()
            logger.info("Clic fuera del modal realizado")
            
        except Exception as modal_error:
            logger.debug(f"No se detectó modal de confirmación: {str(modal_error)}")

    def create_patient(self, patient_data: Dict) -> bool:
        """Crea un nuevo paciente en el sistema
        
        Args:
            patient_data: Diccionario con los datos del paciente
            
        Returns:
            bool: True si el paciente fue creado exitosamente
        """
        try:
            logger.info("Abriendo formulario de paciente...")
            
            # Intentar acceso directo primero
            self.driver.get('https://app.esiclinic.com/pacientes.php?autoclose=1&load=')
            time.sleep(3)
            
            # Localizar y hacer clic en el botón "+" para nuevo paciente
            try:
                new_patient_button = self.wait.until(
                    EC.element_to_be_clickable(self.page.PATIENT_FORM['new_patient_button']))
                self.driver.execute_script("arguments[0].click();", new_patient_button)
                logger.info("Botón '+' encontrado y clickeado")
                time.sleep(2)
            except Exception as btn_e:
                logger.warning(f"No se pudo encontrar el botón '+': {str(btn_e)}")
                logger.info("Intentando método alternativo para abrir formulario...")
                self.driver.get('https://app.esiclinic.com/pacientes.php?action=new')
                time.sleep(3)

            # Rellenar el formulario
            fields_map = {
                'nombre': self.page.PATIENT_FORM['name_field'],
                'apellidos': self.page.PATIENT_FORM['lastname_field'],
                'dni': self.page.PATIENT_FORM['dni_field'],
                'movil': self.page.PATIENT_FORM['phone_field'],
                'email': self.page.PATIENT_FORM['email_field'],
                'fecha_nacimiento': self.page.PATIENT_FORM['birthdate_field']
            }

            logger.info("Rellenando formulario...")
            for field, locator in fields_map.items():
                if patient_data.get(field):
                    try:
                        element = self.wait.until(EC.presence_of_element_located(locator))
                        element.clear()
                        element.send_keys(str(patient_data[field]))
                        logger.info(f"Campo {field} completado")
                    except Exception as e:
                        logger.error(f"Error al rellenar el campo {field}: {str(e)}")
                        self._take_screenshot(f"form_field_error_{field}")
                        return False
            
            # Guardar el paciente
            logger.info("Guardando paciente...")
            try:
                save_button = self.wait.until(
                    EC.element_to_be_clickable(self.page.PATIENT_FORM['save_button']))
                
                self._retry_on_failure(
                    lambda: self.driver.execute_script("arguments[0].click();", save_button)
                )
                
                # Manejar posibles modales
                self._handle_modal()
                
                # Verificar éxito
                try:
                    success_msg = self.wait.until(
                        EC.visibility_of_element_located(self.page.MODALS['success_message']))
                    logger.info(f"Paciente creado exitosamente: {success_msg.text}")
                    return True
                except:
                    # Verificar si hay errores visibles
                    try:
                        error_msg = self.driver.find_element(*self.page.MODALS['error_message'])
                        if error_msg.is_displayed():
                            logger.error(f"Error al crear paciente: {error_msg.text}")
                            return False
                    except:
                        logger.info("Paciente creado exitosamente (sin mensaje de confirmación)")
                        return True
                        
            except Exception as e:
                logger.error(f"Error al guardar paciente: {str(e)}")
                self._take_screenshot("save_error")
                return False

        except Exception as e:
            logger.error(f"Error al crear paciente: {str(e)}")
            self._take_screenshot("patient_creation_error")
            return False

def get_patient_data() -> Dict:
    """Obtiene los datos del paciente por terminal
    
    Returns:
        Diccionario con los datos del paciente
    """
    print("\n=== DATOS DEL NUEVO PACIENTE ===")
    
    patient = {
        'nombre': input("Nombre: ").strip(),
        'apellidos': input("Apellidos: ").strip(),
        'dni': input("DNI/NIF: ").strip(),
        'movil': input("Móvil: ").strip(),
        'email': input("Email: ").strip().lower(),
        'fecha_nacimiento': input("Fecha nacimiento (dd-mm-yyyy, opcional): ").strip(),
    }
    
    return patient

def main():
    """Función principal de ejecución del script"""
    print("=== SISTEMA DE CREACIÓN DE PACIENTES ===")
    
    try:
        # Obtener datos del paciente
        patient_data = get_patient_data()
        
        # Validar datos
        automator = ESIClinicAutomator()
        errors = automator.validate_patient_data(patient_data)
        
        if errors:
            print("\n❌ Errores encontrados:")
            print("\n".join(errors))
            return
        
        # Verificar duplicados
        allow_create, duplicate_msg = automator.check_excel_duplicates(patient_data)
        if not allow_create:
            print(f"\n{duplicate_msg}")
            return
        elif duplicate_msg:
            print(f"\n⚠️ {duplicate_msg}")
            confirm = input("¿Desea continuar? (s/n): ").lower()
            if confirm != 's':
                print("Operación cancelada")
                return
        
        # Iniciar proceso automatizado
        logger.info("Iniciando login...")
        if not automator.login():
            print("❌ Error en el login, no se puede proceder")
            return
        
        logger.info("Login exitoso, continuando con la creación del paciente...")
        
        # Mostrar resumen
        print("\nResumen del paciente:")
        for field, value in patient_data.items():
            print(f"{field:20}: {value}")
        
        # Confirmar
        confirm = input("\n¿Confirmar creación? (s/n): ").lower()
        if confirm != 's':
            print("Operación cancelada")
            return
        
        # Crear paciente
        logger.info("Iniciando creación del paciente...")
        if automator.create_patient(patient_data):
            print("\n🎉 Proceso completado con éxito")
        else:
            print("\n❌ Hubo un error en el proceso")
        
    except Exception as e:
        logger.error(f"Error inesperado en el proceso principal: {str(e)}")
    finally:
        if 'automator' in locals():
            automator.close()

if __name__ == "__main__":
    main()
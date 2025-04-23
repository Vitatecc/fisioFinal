from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
import os
import time
import traceback
from dotenv import load_dotenv
from google_sheets import subir_a_google_sheets
# Cargar variables de entorno
load_dotenv("env/.env")
def eliminar_excel_antiguo():
    """Elimina archivos Excel antiguos en la carpeta data"""
    for filename in os.listdir("data"):
        if filename.endswith(('.xls', '.xlsx')) and not filename.startswith('~$'):
            try:
                os.remove(os.path.join("data", filename))
                print(f"Archivo antiguo eliminado: {filename}")
            except Exception as e:
                print(f"No se pudo eliminar el archivo {filename}: {str(e)}")
def convertir_a_xlsx(ruta_archivo):
    """Convierte el archivo a formato .xlsx si es necesario"""
    if ruta_archivo.endswith('.xls'):
        nuevo_nombre = ruta_archivo.replace('.xls', '.xlsx')
        os.rename(ruta_archivo, nuevo_nombre)
        print(f"Archivo convertido a XLSX: {nuevo_nombre}")
        return nuevo_nombre
    return ruta_archivo
def descargar_excel():
    """Descarga el Excel de esiclinic y lo sube a Google Sheets"""
    eliminar_excel_antiguo()
    # Configuración de Chrome
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath("data"),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
    })
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    try:
        # 1. Acceso e inicio de sesión
        driver.get("https://esiclinic.com/")
        print("Accediendo a esiclinic.com...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "esi_user"))
        ).send_keys(os.getenv("USUARIO_ESICLINIC"))
        driver.find_element(By.ID, "esi_pass").send_keys(os.getenv("PASSWORD_ESICLINIC"))
        login_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "bt_acceder"))
        )
        try:
            login_button.click()
        except:
            driver.execute_script("arguments[0].click();", login_button)
        print("Credenciales enviadas, iniciando sesión...")
        # 2. Esperar login exitoso
        WebDriverWait(driver, 15).until(EC.url_contains("agenda.php"))
        print("Login exitoso")
        # 3. Navegar a listado de citas
        driver.get("https://app.esiclinic.com/listadodecitas.php")
        print("Navegación a Listado de Citas completada")
        # 4. Configurar fechas
        fecha_hoy = datetime.now().strftime("%d-%m-%Y")
        fecha_manana = (datetime.now() + timedelta(days=30)).strftime("%d-%m-%Y")
        def set_fecha(field_id, value):
            field = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, field_id))
            )
            field.clear()
            field.send_keys(Keys.CONTROL + 'a')
            field.send_keys(Keys.DELETE)
            field.send_keys(value)
            print(f"Fecha {field_id} actualizada a {value}")
        set_fecha("fecha", fecha_hoy)
        set_fecha("fecha2", fecha_manana)
        # 5. Descargar Excel
        print("Esperando 10 segundos para la carga de datos...")
        time.sleep(10)
        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "bt_excel"))
        )
        download_button.click()
        print("Descargando Excel...")
        # 6. Esperar descarga
        max_espera = 60
        archivo_descargado = None
        for _ in range(max_espera):
            archivos = [f for f in os.listdir("data") if f.endswith('.xls') and not f.startswith('~$')]
            if archivos:
                archivo_descargado = os.path.join("data", archivos[0])
                archivo_descargado = convertir_a_xlsx(archivo_descargado)
                print(f"Archivo descargado: {archivo_descargado}")
                break
            time.sleep(1)
        else:
            print("No se detectó el archivo descargado")
            driver.save_screenshot("error_descarga.png")
            return False
        # 7. Subir a Google Sheets
        if not subir_a_google_sheets(archivo_descargado, os.getenv("NOMBRE_HOJA")):
            return False
        return True
    except Exception as e:
        print(f"Error crítico: {str(e)}")
        print(traceback.format_exc())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        driver.save_screenshot(f"error_{timestamp}.png")
        return False
    finally:
        driver.quit()
        print("Navegador cerrado")
if __name__ == "__main__":
    if descargar_excel():
        print("Proceso completado con éxito")
    else:
        print("Hubo errores durante el proceso")
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import traceback
import pandas as pd
from dotenv import load_dotenv
from google_sheets import subir_a_google_sheets

# Configuración común
load_dotenv("env/.env")
DOWNLOAD_DIR = os.path.abspath("data/clientes")

def eliminar_archivos_antiguos():
    """Elimina archivos Excel antiguos en la carpeta de descarga"""
    for filename in os.listdir(DOWNLOAD_DIR):
        if filename.endswith(('.xls', '.xlsx')) and not filename.startswith('~$'):
            try:
                os.remove(os.path.join(DOWNLOAD_DIR, filename))
                print(f" Archivo antiguo eliminado: {filename}")
            except Exception as e:
                print(f" No se pudo eliminar el archivo {filename}: {str(e)}")

def convertir_a_xlsx(ruta_archivo):
    if ruta_archivo.endswith('.xls'):
        nuevo_nombre = ruta_archivo.replace('.xls', '.xlsx')
        
        df = pd.read_excel(ruta_archivo, engine='xlrd')  # Abres el .xls
        df.to_excel(nuevo_nombre, index=False, engine='openpyxl')  # Guardas como .xlsx
        
        os.remove(ruta_archivo)  # Borras el .xls original si quieres
        
        print(f" Archivo convertido correctamente a XLSX: {nuevo_nombre}")
        return nuevo_nombre
    return ruta_archivo

def configurar_navegador():
    """Configura y retorna una instancia del navegador Chrome"""
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)
    
    eliminar_archivos_antiguos()

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
    })

    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

def descargar_pacientes(driver):
    """Realiza el proceso completo de descarga del archivo de pacientes"""
    try:
        print(" Accediendo a esiclinic.com...")
        driver.get("https://esiclinic.com/")
        
        print(" Enviando credenciales...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "esi_user"))
        ).send_keys(os.getenv("USUARIO_ESICLINIC"))
        
        driver.find_element(By.ID, "esi_pass").send_keys(os.getenv("PASSWORD_ESICLINIC"))
        
        login_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "bt_acceder"))
        )
        login_button.click()
        
        print(" Iniciando sesión...")
        WebDriverWait(driver, 15).until(EC.url_contains("agenda.php"))
        print(" Login exitoso")

        print(" Navegando a pacientes...")
        driver.get("https://app.esiclinic.com/pacientes.php")
        time.sleep(5)  # Espera para que cargue la página

        print(" Descargando Excel...")
        download_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "bt_excel"))
        )
        download_button.click()

        # Esperar descarga
        max_espera = 60
        archivo_descargado = None
        
        for _ in range(max_espera):
            archivos = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.xls') and not f.startswith('~$')]
            if archivos:
                archivo_descargado = os.path.join(DOWNLOAD_DIR, archivos[0])
                archivo_descargado = convertir_a_xlsx(archivo_descargado)
                print(f" Archivo descargado: {archivo_descargado}")
                break
            time.sleep(1)
        else:
            print(" No se detectó el archivo descargado")
            driver.save_screenshot("error_descarga.png")
            return False
        
        # Subir a Google Sheets
        if not subir_a_google_sheets(archivo_descargado, os.getenv("NOMBRE_HOJA_PACIENTES")):
            return False
            
        return True

    except Exception as e:
        print(f" Error durante la descarga: {str(e)}")
        print(traceback.format_exc())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        driver.save_screenshot(f"error_{timestamp}.png")
        return False

def main():
    driver = configurar_navegador()
    try:
        if descargar_pacientes(driver):
            print(" Proceso completado con éxito")
        else:
            print(" Hubo errores durante el proceso")
    finally:
        driver.quit()
        print(" Navegador cerrado")

if __name__ == "__main__":
    main()
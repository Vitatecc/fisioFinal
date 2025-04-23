import os
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from dotenv import load_dotenv
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Cargar variables de entorno
load_dotenv("env/.env")

# Configuración
WHATSAPP_SESSION_DIR = os.path.abspath("./whatsapp_session")
TU_NUMERO_WHATSAPP = "+34643053023"  # Tu número en formato internacional
REGISTRO_ENVIOS = os.path.abspath("./data/registro_envios.txt")  # Archivo de registro

def normalizar_telefono(numero):
    """Normaliza números de teléfono al formato internacional"""
    if pd.isna(numero):
        return None
    numero = str(numero).strip().replace(" ", "").replace("-", "")
    if not numero.startswith("+"):
        if len(numero) == 9 and numero[0] in ['6','7']:
            return "+34" + numero
        return None
    return numero

def registrar_envio(telefono, fecha_cita, DNI):
    """Registra un mensaje enviado para evitar duplicados"""
    os.makedirs(os.path.dirname(REGISTRO_ENVIOS), exist_ok=True)
    print(f"Guardando en el archivo de registro: {REGISTRO_ENVIOS}")
    with open(REGISTRO_ENVIOS, "a", encoding="utf-8") as f:
        f.write(f"{telefono},{fecha_cita},{DNI}\n")
    print(f"Mensaje registrado para {telefono} en la fecha {fecha_cita} y DNI {DNI}")  

def fue_enviado(telefono, fecha_cita, DNI):
    """Verifica si ya se envió el mensaje al mismo DNI y número de teléfono para esa fecha"""
    if not os.path.exists(REGISTRO_ENVIOS):
        return False
    try:
        with open(REGISTRO_ENVIOS, "r", encoding="utf-8") as f:
            for linea in f:
                partes = linea.strip().split(",")
                if len(partes) != 3:
                    continue
                tel_reg, fecha_reg, dni_reg = partes
                if tel_reg == telefono and fecha_reg == fecha_cita and dni_reg == DNI:
                    return True  # Ya fue enviado a este número y DNI para esta fecha
        return False
    except FileNotFoundError:
        return False

def limpiar_registro():
    """Elimina registros de citas pasadas y líneas inválidas"""
    if not os.path.exists(REGISTRO_ENVIOS):
        return
    
    hoy = datetime.now().strftime("%d-%m-%Y")
    lineas_validas = []
    
    with open(REGISTRO_ENVIOS, "r", encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if not linea:  # Saltar líneas vacías
                continue
                
            partes = linea.split(",")
            if len(partes) < 3:  # Saltar líneas mal formateadas
                continue
                
            fecha = partes[1].strip()
            try:
                if fecha >= hoy:  # Mantener solo futuras o de hoy
                    lineas_validas.append(linea + "\n")
            except TypeError:
                continue
    
    with open(REGISTRO_ENVIOS, "w", encoding="utf-8") as f:
        f.writelines(lineas_validas)

def iniciar_whatsapp():
    """Inicia sesión en WhatsApp Web"""
    try:
        options = webdriver.ChromeOptions()
        options.add_argument(f"user-data-dir={WHATSAPP_SESSION_DIR}")
        options.add_argument("--disable-notifications")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--profile-directory=Default")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        print("Abriendo WhatsApp Web...")
        driver.get("https://web.whatsapp.com")
        
        print("Esperando carga de WhatsApp...")
        WebDriverWait(driver, 60).until(
            lambda d: d.find_element(By.XPATH, '//div[@role="textbox"]') or 
                      d.find_element(By.XPATH, '//canvas[@aria-label="Scan me!"]')
        )
        print("WhatsApp cargado correctamente")
        return driver
        
    except Exception as e:
        print(f"Error al iniciar WhatsApp: {str(e)}")
        if 'driver' in locals():
            driver.save_screenshot("error_login.png")
        raise

def enviar_whatsapp(driver, numero, mensaje):
    try:
        print(f"\nIniciando envío a {numero}...")
        
        # Abrir chat directamente (sin mensaje en URL)
        driver.get(f"https://web.whatsapp.com/send?phone={numero}")
        
        # Esperar hasta 60 segundos a que cargue el chat
        input_box = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        
        # Limpiar el cuadro de texto por si acaso
        input_box.clear()
        
        # Enviar el mensaje línea por línea (sin emojis)
        for line in mensaje.split('\n'):
            input_box.send_keys(line)
            input_box.send_keys(Keys.SHIFT + Keys.ENTER)  # Nuevo línea sin enviar
        input_box.send_keys(Keys.RETURN)  # Enviar final
        
        # Pequeña pausa para asegurar el envío
        time.sleep(2)
        
        # Verificación simple de envío
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-icon="msg-time"]'))
            )
            print(f"Envio confirmado a {numero}")
            return True
        except:
            print(f"Envio no confirmado (aunque el script continuó)")
            return False
        
    except Exception as e:
        print(f"Error al enviar a {numero}: {str(e)}")
        driver.save_screenshot(f"error_envio_{numero}.png")
        return False

def procesar_citas():
    """Procesa las citas y envía recordatorios"""
    carpeta_descargas = "data"
    os.makedirs(carpeta_descargas, exist_ok=True)
    limpiar_registro()  # Limpia registros antiguos
    
    try:
        # Buscar archivo Excel más reciente
        archivos = [f for f in os.listdir(carpeta_descargas) 
                   if f.endswith((".xls", ".xlsx")) and not f.startswith("~$")]
        
        if not archivos:
            print("No se encontró ningún archivo Excel en la carpeta 'data'")
            return
        
        archivo_excel = os.path.join(carpeta_descargas, archivos[0])
        print(f"Procesando archivo: {archivo_excel}")
        
        # Leer archivo Excel
        try:
            df = pd.read_excel(archivo_excel, engine="openpyxl")  # Intentamos con openpyxl primero
        except Exception as e:
            try:
                df = pd.read_excel(archivo_excel, engine="xlrd")  # Si falla, intentamos con xlrd
            except Exception as ex:
                print(f"Error al leer el archivo Excel: {ex}")
                return
        
        # Verificar columnas necesarias
        if not all(col in df.columns for col in ["Fecha", "Hora", "Paciente", "DNI"]):
            print("El archivo Excel no tiene las columnas esperadas (Fecha, Hora, Paciente, DNI)")
            return
        
        # Convertir la columna de fecha
        df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
        
        # Filtrar citas de hoy y mañana
        ahora = datetime.now()
        fecha_hoy = ahora.date()
        fecha_manana = (ahora + timedelta(days=1)).date()
        
        citas_recientes = df[
            (df["Fecha"].dt.date == fecha_hoy) | 
            (df["Fecha"].dt.date == fecha_manana)
        ].copy()
        
        if citas_recientes.empty:
            print("No hay citas para hoy o mañana")
            return
        
        print(f"Encontradas {len(citas_recientes)} citas para notificar")
        
        # Iniciar WhatsApp (una sola vez)
        driver = iniciar_whatsapp()
        
        # Procesar cada cita
        for _, fila in citas_recientes.iterrows():
            DNI = fila.get("DNI", "DNI")
            nombre = fila.get("Paciente", "Paciente")
            telefono = fila.get("Móvil") if pd.notna(fila.get("Móvil")) else fila.get("Teléfono")
            telefono = normalizar_telefono(telefono)
            
            if not telefono:
                print(f"No se encontró teléfono para {DNI}")
                continue
            
            fecha_cita = fila["Fecha"].strftime("%d-%m-%Y")
            hora_cita = fila.get("Hora", "00:00")

            # Combinar fecha y hora en un datetime
            try:
                hora_cita_dt = datetime.strptime(str(hora_cita), "%H:%M").time()
            except Exception:
                hora_cita_dt = datetime.strptime("00:00", "%H:%M").time()

            # Combinar la fecha y la hora en un solo objeto datetime
            fecha_hora_cita = datetime.combine(datetime.strptime(fecha_cita, "%d-%m-%Y").date(), hora_cita_dt)

            # Verificar si la cita ya pasó
            if fecha_hora_cita < datetime.now():
                print(f"La cita de {nombre} ya ha pasado ({fecha_hora_cita.strftime('%d-%m-%Y %H:%M')}). No se enviará mensaje.")
                continue

            # Verificar si ya se envió
            if fue_enviado(telefono, fecha_cita, DNI):
                print(f"Mensaje ya enviado a {DNI} para el {fecha_cita}")
                continue
            
            hora_cita = fila.get("Hora", "Desconocida")
            
            if fila["Fecha"].date() == fecha_hoy:
                mensaje = (f" Recordatorio: Hola {nombre}, tienes cita HOY en Vitatec\n"
                         f" Hora: {hora_cita}\n"
                         f"¡Te esperamos!")
            else:
                mensaje = (f" Recordatorio: Hola {nombre}, tienes cita mañana en Vitatec\n"
                         f"Fecha: {fecha_cita}\n"
                         f" Hora: {hora_cita}\n"
                         f"¡Te esperamos!")
            
            enviar_whatsapp(driver, telefono, mensaje)
            registrar_envio(telefono, fecha_cita, DNI)  # Registrar después de enviar
    
    except Exception as e:
        print(f"Error crítico al procesar citas: {str(e)}")
        raise
    finally:
        if 'driver' in locals():
            driver.quit()
            print("Navegador cerrado")

if __name__ == "__main__":
    procesar_citas()

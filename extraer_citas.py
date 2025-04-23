from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import os
import json
import time
from dotenv import load_dotenv
import glob
import re
from collections import defaultdict

# Cargar variables de entorno
load_dotenv("env/.env")

COLORES = {
    'jose': (9, 142, 67),
    'arnau': (108, 14, 33),
    'david': (242, 159, 44)
}

def determinar_facultativo(dia_semana: int, hora: str):
    hora_dt = datetime.strptime(hora, "%H:%M")
    if dia_semana == 0:  # lunes
        if hora_dt < datetime.strptime("14:00", "%H:%M"):
            return 'arnau'
        else:
            return 'david'
    elif dia_semana == 1:  # martes
        return 'arnau'
    elif dia_semana == 2:  # miércoles
        if hora_dt < datetime.strptime("12:30", "%H:%M"):
            return 'david'
        else:
            return 'arnau'
    elif dia_semana == 3:  # jueves
        return 'jose'
    elif dia_semana == 4:  # viernes
        return 'arnau'
    return None

def extraer_rgb(style: str):
    match = re.search(r'rgb\((\d+), (\d+), (\d+)\)', style)
    if match:
        return tuple(map(int, match.groups()))
    return None

def determinar_agenda(color_rgb, fecha: str, hora: str):
    try:
        dia_dt = datetime.strptime(fecha, "%Y-%m-%d")
        dia_semana = dia_dt.weekday()
        facultativo = determinar_facultativo(dia_semana, hora)
        if facultativo is None:
            return "2"

        color_facultativo = COLORES[facultativo]
        if color_rgb == color_facultativo:
            return "1"
        else:
            return "2"
    except Exception as e:
        print(f"⚠️ Error determinando agenda: {e}")
        return "2"

def parsear_fecha(dia_str):
    hoy = datetime.now()
    partes = dia_str.split()
    if len(partes) < 2:
        return dia_str
    dia_mes = partes[1].split('/')
    if len(dia_mes) != 2:
        return dia_str
    dia = int(dia_mes[0])
    mes = int(dia_mes[1])
    año = hoy.year
    if mes < hoy.month:
        año += 1
    return f"{año}-{mes:02d}-{dia:02d}"

def extraer_citas_por_semanas():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        driver.get("https://esiclinic.com/")
        print("➡️ Accediendo a esiclinic.com...")

        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "esi_user"))).send_keys(os.getenv("USUARIO_ESICLINIC"))
        driver.find_element(By.ID, "esi_pass").send_keys(os.getenv("PASSWORD_ESICLINIC"))
        login_button = driver.find_element(By.ID, "bt_acceder")
        driver.execute_script("arguments[0].click();", login_button)
        print("🔑 Credenciales enviadas, iniciando sesión...")

        WebDriverWait(driver, 15).until(EC.url_contains("agenda.php"))
        print("✅ Login exitoso. Accediendo a la agenda...")

        boton_semana = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".fc-agendaWeek-button"))
        )
        boton_semana.click()
        print("📅 Cambiando a vista SEMANAL...")
        time.sleep(5)

        citas_totales = []
        hoy = datetime.now().date()

        for semana in range(2):
            print(f"\n🔍 Extrayendo citas de la SEMANA {semana + 1}...")
            try:
                dias_semana = []
                encabezados_dias = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".fc-day-header"))
                )
                for dia in encabezados_dias:
                    dia_texto = dia.text.split('\n')[0].strip()
                    if dia_texto:
                        dias_semana.append(dia_texto)

                eventos = WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".fc-event-container .fc-event"))
                )
                columnas_dias = driver.find_elements(By.CSS_SELECTOR, "td.fc-day")
                citas_por_hora = defaultdict(int)

                for evento in eventos:
                    try:
                        ubicacion = evento.location
                        tamaño = evento.size
                        centro_x = ubicacion['x'] + (tamaño['width'] / 2)
                        dia_index = 0
                        for i, columna in enumerate(columnas_dias):
                            col_x = columna.location['x']
                            col_width = columna.size['width']
                            if col_x <= centro_x <= (col_x + col_width):
                                dia_index = i
                                break
                        dia_str = dias_semana[dia_index] if dia_index < len(dias_semana) else "Día no encontrado"
                        hora_texto = evento.find_element(By.CSS_SELECTOR, ".fc-time").get_attribute("data-full")
                        hora_inicio, _ = hora_texto.split(" - ")
                        clave_hora = f"{dia_str}_{hora_inicio.strip()}"
                        citas_por_hora[clave_hora] += 1
                    except Exception:
                        continue

                for evento in eventos:
                    try:
                        ubicacion = evento.location
                        tamaño = evento.size
                        centro_x = ubicacion['x'] + (tamaño['width'] / 2)
                        dia_index = 0
                        for i, columna in enumerate(columnas_dias):
                            col_x = columna.location['x']
                            col_width = columna.size['width']
                            if col_x <= centro_x <= (col_x + col_width):
                                dia_index = i
                                break

                        dia_str = dias_semana[dia_index] if dia_index < len(dias_semana) else "Día no encontrado"
                        fecha_iso = parsear_fecha(dia_str)
                        fecha_cita = datetime.strptime(fecha_iso, "%Y-%m-%d").date()
                        if semana == 0 and fecha_cita < hoy:
                            continue

                        hora_texto = evento.find_element(By.CSS_SELECTOR, ".fc-time").get_attribute("data-full")
                        hora_inicio, hora_fin = hora_texto.split(" - ")
                        paciente = evento.find_element(By.CSS_SELECTOR, ".fc-title").text

                        style_attr = evento.get_attribute("style")
                        color_rgb = extraer_rgb(style_attr)
                        agenda_id = determinar_agenda(color_rgb, fecha_iso, hora_inicio.strip())

                        citas_totales.append({
                            "semana": semana + 1,
                            "dia": fecha_iso,
                            "hora_inicio": hora_inicio.strip(),
                            "hora_fin": hora_fin.strip(),
                            "paciente": paciente,
                            "agenda": agenda_id
                        })

                        print(f"   📅 {fecha_iso} ⏰ {hora_inicio}-{hora_fin} - {paciente} (Agenda {agenda_id})")
                    except Exception as e:
                        print(f"⚠️ Error extrayendo cita (ignorada): {str(e)}")

                if semana < 1:
                    boton_siguiente = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".fc-next-button"))
                    )
                    boton_siguiente.click()
                    print("⏭️ Avanzando a la próxima semana...")
                    time.sleep(5)

            except Exception as e:
                print(f"⚠️ Error procesando semana {semana + 1}: {str(e)}")
                continue

        if not os.path.exists("data"):
            os.makedirs("data")

        archivos_json = glob.glob(os.path.join("data", "citas_2_semanas_*.json"))
        for archivo in archivos_json:
            try:
                os.remove(archivo)
                print(f"🗑️ Archivo antiguo eliminado: {archivo}")
            except Exception as e:
                print(f"⚠️ Error al eliminar archivo {archivo}: {str(e)}")

        archivo_json = os.path.join("data", "citas_2_semanas.json")
        with open(archivo_json, "w", encoding="utf-8") as f:
            json.dump(citas_totales, f, indent=2, ensure_ascii=False)
        print(f"\n💾 Todas las citas guardadas en: {archivo_json}")

    except Exception as e:
        print(f"❌ Error crítico: {str(e)}")
        driver.save_screenshot("error_agenda.png")
    finally:
        driver.quit()
        print("🚪 Navegador cerrado")

if __name__ == "__main__":
    extraer_citas_por_semanas()
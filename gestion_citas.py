from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
import os
import time
import pandas as pd
import json
from datetime import datetime, timedelta
from dotenv import load_dotenv
import random
import subprocess
import sys
# Configuración común
load_dotenv("env/.env")
DOWNLOAD_DIR = os.path.abspath("data/clientes")
ARCHIVO_EXCEL = os.path.join(DOWNLOAD_DIR, "pacientes.xlsx")
RUTA_CITAS = os.path.abspath("data/citas_2_semanas.json")
def configurar_navegador():
    """Configura y retorna una instancia del navegador Chrome"""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
    })
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
def login(driver):
    """Realiza el proceso de login en esiclinic"""
    try:
        print("➡️ Accediendo a esiclinic.com...")
        driver.get("https://esiclinic.com/")
        time.sleep(3)
        print("🔑 Enviando credenciales...")
        driver.find_element(By.ID, "esi_user").send_keys(os.getenv("USUARIO_ESICLINIC"))
        driver.find_element(By.ID, "esi_pass").send_keys(os.getenv("PASSWORD_ESICLINIC"))
        driver.find_element(By.ID, "bt_acceder").click()
        time.sleep(5)
        print("📂 Navegando a agenda...")
        driver.get("https://app.esiclinic.com/agenda.php")
        time.sleep(5)
        return True
    except Exception as e:
        print(f"❌ Error durante el login: {str(e)}")
        driver.save_screenshot("error_login.png")
        return False
    
def generar_rango_horas(inicio, fin, intervalo):
    """Genera todas las horas en un rango con intervalos específicos"""
    horas = []
    hora_actual = datetime.strptime(inicio, "%H:%M")
    hora_fin = datetime.strptime(fin, "%H:%M")

    while hora_actual <= hora_fin - timedelta(minutes=intervalo):
        horas.append((hora_actual.strftime("%H:%M"), 1))  # Por defecto agenda 1
        hora_actual += timedelta(minutes=intervalo)

    return horas

def verificar_paciente():
    """Verifica si el paciente existe en el Excel por DNI (CIF) o correo (E-Mail)"""
    try:
        identificador = input("🔍 Ingrese DNI (CIF) o correo (E-Mail) del paciente: ").strip().lower()
        df = pd.read_excel(ARCHIVO_EXCEL, engine='openpyxl')

        # Verificar columnas esperadas
        columnas_esperadas = ['CIF', 'E-Mail']
        for col in columnas_esperadas:
            if col not in df.columns:
                raise ValueError(f"❌ Faltan columnas esperadas en el Excel: {col}")

        df['CIF'] = df['CIF'].astype(str).str.strip().str.lower()
        df['E-Mail'] = df['E-Mail'].astype(str).str.strip().str.lower()

        pacientes = df[(df['CIF'] == identificador) | (df['E-Mail'] == identificador)]

        if not pacientes.empty:
            if identificador in df['CIF'].values:
                paciente = pacientes.iloc[0]
            else:
                if len(pacientes) > 1:
                    print("\n⚠️ Este correo está asociado a múltiples pacientes:")
                    for i, (_, row) in enumerate(pacientes.iterrows(), 1):
                        print(f"{i}. {row['Nombre']} {row['Apellidos']} (DNI: {row['CIF']})")
                    while True:
                        seleccion = input("🔢 Ingrese el número del paciente o el DNI completo: ").strip().lower()
                        if seleccion.isdigit() and 1 <= int(seleccion) <= len(pacientes):
                            paciente = pacientes.iloc[int(seleccion)-1]
                            break
                        elif seleccion in pacientes['CIF'].values:
                            paciente = pacientes[pacientes['CIF'] == seleccion].iloc[0]
                            break
                        else:
                            print("❌ Opción no válida. Intente nuevamente.")
                else:
                    paciente = pacientes.iloc[0]

            nombre_completo = f"{paciente['Nombre']} {paciente['Apellidos']}"
            print(f"\n🎉 ¡Paciente encontrado!\n👤 Nombre: {nombre_completo}\n🆔 DNI: {paciente['CIF']}\n📧 Correo: {paciente['E-Mail']}")
            return nombre_completo
        else:
            print("\n❌ Paciente no encontrado. Verifique:\n- El DNI (CIF) o correo (E-Mail) ingresado\n- Que el paciente esté registrado en el sistema")
            return None

    except Exception as e:
        print(f"\n⚠️ Error al verificar paciente: {str(e)}")
        print("Posibles causas:\n- El archivo pacientes.xlsx no tiene las columnas esperadas (CIF, E-Mail)\n- El archivo está abierto en otro programa\n- El formato del archivo no es compatible")
        return None
    
def lunes_con_agenda_2(fecha_dt):
    base = datetime(2025, 4, 14)
    diferencia_semanas = (fecha_dt.date() - base.date()).days // 7
    return diferencia_semanas % 2 == 0


def seleccionar_cita():
    """Seleccionar cita con manejo automático de agendas secundarias"""
    try:
        with open(RUTA_CITAS, 'r', encoding='utf-8') as f:
            citas = json.load(f)
            
        while True:
            fecha_input = input("📅 Ingrese fecha (DD-MM-YYYY): ").strip()
            try:
                fecha_dt = datetime.strptime(fecha_input, "%d-%m-%Y")
                fecha_iso = fecha_dt.strftime("%Y-%m-%d")
                dia_semana = fecha_dt.weekday()
                
                # Validaciones básicas
                if fecha_dt.weekday() >= 5:
                    print("⚠️ Solo se pueden agendar citas de lunes a viernes")
                    continue
                if fecha_dt.date() < datetime.now().date():
                    print("⚠️ No se pueden agendar citas en fechas pasadas")
                    continue
                    
                # Obtener citas existentes para esta fecha
                citas_fecha = [c for c in citas if c['dia'] == fecha_iso]
                
                # Mostrar citas existentes
                if citas_fecha:
                    print("\n⛔ Citas existentes para este día:")
                    for c in citas_fecha:
                        print(f"⏰ {c['hora_inicio']}-{c['hora_fin']} | {c['paciente']} (Agenda {c.get('agenda', '1')}")
                
                # Definir horarios por día y tipo de agenda
                horarios = {
                    'manana': {
                        'primaria': ("10:15", "14:00"),
                        'secundaria': ("10:15", "14:00") if dia_semana == 0 and lunes_con_agenda_2(fecha_dt) else ("10:30", "13:30") if dia_semana == 3 else None  # Lunes(0): mismo horario | Jueves(3): escalonado
                    },
                    'tarde': {
                        'primaria': ("15:00", "20:15"),
                        'secundaria': ("15:00", "20:15") if dia_semana == 0 and lunes_con_agenda_2(fecha_dt) else ("15:30", "20:00") if dia_semana == 3 else None  # Lunes(0): mismo horario | Jueves(3): escalonado
                    }
                }
                
                # Función para verificar disponibilidad en una agenda específica
                def verificar_disponibilidad(inicio, fin, agenda_num):
                    disponibles = []
                    hora_actual = datetime.strptime(inicio, "%H:%M")
                    hora_fin = datetime.strptime(fin, "%H:%M")
                    
                    while hora_actual <= hora_fin - timedelta(minutes=45):
                        hora_str = hora_actual.strftime("%H:%M")
                        conflicto = False
                        
                        for cita in citas_fecha:
                            # Solo verificar conflictos en la misma agenda
                            if cita.get('agenda', '1') == str(agenda_num):
                                inicio_cita = datetime.strptime(cita['hora_inicio'], "%H:%M")
                                fin_cita = datetime.strptime(cita['hora_fin'], "%H:%M")
                                if inicio_cita <= hora_actual < fin_cita:
                                    conflicto = True
                                    break
                        
                        if not conflicto:
                            disponibles.append((hora_str, agenda_num))
                        
                        hora_actual += timedelta(minutes=45)
                    return disponibles
                
                # Verificar disponibilidad en todas las agendas
                todas_disponibles = []
                
                # MAÑANA - Primero agenda primaria
                print("\n🌅 Horarios de MAÑANA:")
                disp_manana_primaria = verificar_disponibilidad(*horarios['manana']['primaria'], 1)
                print(f"\n🏥 Agenda Principal Mañana ({horarios['manana']['primaria'][0]} - {horarios['manana']['primaria'][1]}):")
                for hora, agenda in generar_rango_horas(*horarios['manana']['primaria'], 45):
                    disp = any(h == hora for h, a in disp_manana_primaria)
                    print(f"{'✅' if disp else '❌'} {hora} - {'Disponible' if disp else 'Ocupado'}")
                    if disp:
                        todas_disponibles.append((hora, 1))

                # MAÑANA - Agenda secundaria si primaria está llena y existe
                if horarios['manana']['secundaria'] and len(disp_manana_primaria) == 0:
                    disp_manana_secundaria = verificar_disponibilidad(*horarios['manana']['secundaria'], 2)
                    print(f"\n🏥 Agenda Secundaria Mañana ({horarios['manana']['secundaria'][0]} - {horarios['manana']['secundaria'][1]}):")
                    for hora, agenda in generar_rango_horas(*horarios['manana']['secundaria'], 45):
                        disp = any(h == hora for h, a in disp_manana_secundaria)
                        print(f"{'✅' if disp else '❌'} {hora} - {'Disponible' if disp else 'Ocupado'}")
                        if disp:
                            todas_disponibles.append((hora, 2))
                
                # TARDE - Primero agenda primaria
                print("\n🌇 Horarios de TARDE:")
                disp_tarde_primaria = verificar_disponibilidad(*horarios['tarde']['primaria'], 1)
                print(f"\n🏥 Agenda Principal Tarde ({horarios['tarde']['primaria'][0]} - {horarios['tarde']['primaria'][1]}):")
                for hora, agenda in generar_rango_horas(*horarios['tarde']['primaria'], 45):
                    disp = any(h == hora for h, a in disp_tarde_primaria)
                    print(f"{'✅' if disp else '❌'} {hora} - {'Disponible' if disp else 'Ocupado'}")
                    if disp:
                        todas_disponibles.append((hora, 1))
                
                # TARDE - Agenda secundaria si primaria está llena y existe
                if horarios['tarde']['secundaria'] and len(disp_tarde_primaria) == 0:
                    disp_tarde_secundaria = verificar_disponibilidad(*horarios['tarde']['secundaria'], 2)
                    print(f"\n🏥 Agenda Secundaria Tarde ({horarios['tarde']['secundaria'][0]} - {horarios['tarde']['secundaria'][1]}):")
                    for hora, agenda in generar_rango_horas(*horarios['tarde']['secundaria'], 45):
                        disp = any(h == hora for h, a in disp_tarde_secundaria)
                        print(f"{'✅' if disp else '❌'} {hora} - {'Disponible' if disp else 'Ocupado'}")
                        if disp:
                            todas_disponibles.append((hora, 2))
                
                if not todas_disponibles:
                    print("\n❌ No hay horarios disponibles para este día")
                    continue
                    
                hora_input = input("\n⏰ Ingrese la hora deseada (HH:MM): ").strip()
                
                # Buscar la hora en las disponibles
                for hora, agenda in todas_disponibles:
                    if hora == hora_input:
                        return fecha_input, hora_input, agenda  # Ahora devolvemos también el número de agenda
                
                print("❌ Hora no disponible. Intente de nuevo.")
                    
            except ValueError:
                print("⚠️ Formato de fecha u hora incorrecto. Use DD-MM-YYYY y HH:MM")
    except Exception as e:
        print(f"❌ Error al leer las citas: {str(e)}")
        return None

def navegar_a_fecha(driver, fecha):
    """Navega en la agenda hasta encontrar la semana que contiene la fecha especificada (excluyendo fines de semana)"""
    try:
        fecha_dt = datetime.strptime(fecha, "%d-%m-%Y")
        
        # Verificar si la fecha es fin de semana (sábado=5, domingo=6)
        if fecha_dt.weekday() >= 5:
            print("❌ No se pueden agendar citas los fines de semana (sábado y domingo)")
            return False
            
        # Resto de la función (sin cambios)
        meses_es = {
            'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
            'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
        }
        max_intentos = 12
        intentos = 0
        
        while intentos < max_intentos:
            try:
                rango_fechas = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".fc-center h2"))
                ).text
                print(f"Texto obtenido de la agenda: '{rango_fechas}'")
                # Procesar diferentes formatos de fecha
                if " — " in rango_fechas:
                    inicio_semana_str, fin_semana_str = rango_fechas.split(" — ")
                    inicio_semana = parsear_fecha_es(inicio_semana_str, meses_es)
                    fin_semana = parsear_fecha_es(fin_semana_str, meses_es)
                elif "de" in rango_fechas:
                    inicio_semana = fin_semana = parsear_fecha_es(rango_fechas, meses_es)
                else:
                    try:
                        inicio_semana = fin_semana = datetime.strptime(rango_fechas, "%d/%m/%Y")
                    except ValueError:
                        print(f"⚠️ Formato de fecha no reconocido: {rango_fechas}")
                        break
                # Verificar si la fecha buscada está en este rango
                if inicio_semana.date() <= fecha_dt.date() <= fin_semana.date():
                    print(f"✅ Semana encontrada: {rango_fechas}")
                    return True
                # Avanzar o retroceder según sea necesario
                if fecha_dt.date() > fin_semana.date():
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".fc-next-button"))).click()
                    print("⏭️ Avanzando a la próxima semana...")
                else:
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".fc-prev-button"))).click()
                    print("⏮️ Retrocediendo a la semana anterior...")
                time.sleep(3)
                intentos += 1
            except Exception as e:
                print(f"⚠️ Error al interpretar rango de fechas: {str(e)}")
                break
                
        print(f"❌ No se encontró la fecha {fecha} después de {max_intentos} intentos")
        return False
        
    except Exception as e:
        print(f"❌ Error navegando a la fecha: {str(e)}")
        return False
def parsear_fecha_es(fecha_str, meses_es):
    """Convierte fechas en español a objeto datetime"""
    try:
        if "de" in fecha_str:
            partes = fecha_str.split()
            dia = int(partes[0])
            mes = meses_es[partes[2].lower()]
            año = int(partes[4])
            return datetime(año, mes, dia)
        else:
            return datetime.strptime(fecha_str, "%d/%m/%Y")
    except Exception as e:
        print(f"⚠️ Error parseando fecha '{fecha_str}': {str(e)}")
        raise
def seleccionar_facultativo_por_horario(driver, fecha, hora, agenda_num):
    """Selecciona el facultativo según el día de la semana y horario especificado"""
    max_intentos = 3
    intentos = 0
    
    while intentos < max_intentos:
        try:
            fecha_dt = datetime.strptime(fecha, "%d-%m-%Y")
            dia_semana = fecha_dt.weekday()
            hora_dt = datetime.strptime(hora, "%H:%M")

            # Esperar a que el modal esté completamente estable
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#citaMotivo, #citaFacultativo")))
            time.sleep(0.5)  # Espera corta para estabilización

            # Localizar el select de facultativo FRESCO
            select_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "citaFacultativo")))
            
            select_facultativo = Select(select_element)

            # Obtener opciones actualizadas con manejo de errores
            opciones_validas = {}
            try:
                for op in select_facultativo.options:
                    try:
                        if op.text.strip() and "--" not in op.text:
                            opciones_validas[op.text.strip()] = op
                    except StaleElementReferenceException:
                        continue  # Si una opción se vuelve obsoleta, continuar con las demás
            except StaleElementReferenceException:
                print("⚠️ Las opciones cambiaron durante la iteración, reintentando...")
                continue  # Volver al inicio del bucle while

            if not opciones_validas:
                print("⚠️ No hay facultativos disponibles")
                return False

            # Lógica de selección según horario
            facultativo_seleccionado = None
            if dia_semana == 0:  # Lunes
                if agenda_num == 1:
                    facultativo_seleccionado = "Arnau Girones" if hora_dt.time() < datetime.strptime("14:00", "%H:%M").time() else "David Ibiza"
                else:
                    facultativo_seleccionado = "Jose Cabanes"
            elif dia_semana == 1:  # Martes
                facultativo_seleccionado = "Arnau Girones"
            elif dia_semana == 2:  # Miércoles
                facultativo_seleccionado = "David Ibiza" if datetime.strptime("10:00", "%H:%M").time() <= hora_dt.time() <= datetime.strptime("12:30", "%H:%M").time() else "Arnau Girones"
            elif dia_semana == 3:  # Jueves
                facultativo_seleccionado = "Jose Cabanes"
            elif dia_semana == 4:  # Viernes
                facultativo_seleccionado = "Arnau Girones"

            # Seleccionar el facultativo con manejo de errores
            try:
                for nombre, opcion in opciones_validas.items():
                    if facultativo_seleccionado and facultativo_seleccionado.lower() in nombre.lower():
                        opcion.click()  # Usar click() en lugar de select_by_visible_text para mayor robustez
                        print(f"👨‍⚕️ Facultativo seleccionado (según horario): {nombre}")
                        return True
            except StaleElementReferenceException:
                print("⚠️ El elemento cambió durante la selección, reintentando...")
                intentos += 1
                continue

            # Si no se encontró el facultativo esperado
            if opciones_validas:
                primera_opcion = list(opciones_validas.values())[0]
                primera_opcion.click()
                print(f"⚠️ Facultativo esperado no disponible. Seleccionado: {primera_opcion.text}")
                return True
            return False

        except StaleElementReferenceException:
            intentos += 1
            print(f"⚠️ Intento {intentos}: Elemento obsoleto, reintentando...")
            time.sleep(1)
        except Exception as e:
            print(f"⚠️ Error al seleccionar facultativo: {str(e)}")
            driver.save_screenshot(f"error_facultativo_intento_{intentos}.png")
            return False

    print("❌ No se pudo seleccionar el facultativo después de varios intentos")
    return False
def seleccionar_sala(driver, agenda_num, dia_semana):
    """Selecciona Box 2 para agenda 2 en lunes/jueves, Box 1 para agenda 1"""
    try:
        select_element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "modalRoom")))
        select_sala = Select(select_element)
        
        box = "Box 2" if agenda_num == 2 else "Box 1"  # Simplificado
        select_sala.select_by_visible_text(box)
        print(f"🏥 Sala seleccionada: {box} (Agenda {agenda_num})")
        return True
    except Exception as e:
        print(f"⚠️ Error al seleccionar sala: {str(e)}")
        return False
def modificar_hora_en_modal(driver, hora_deseada):
    try:
        print(f"🕒 Configurando hora: {hora_deseada}")
        horas, minutos = map(int, hora_deseada.split(':'))
        if not (0 <= horas <= 23 and 0 <= minutos <= 59):
            raise ValueError("Hora fuera de rango")
        minutos = (minutos // 5) * 5
        if minutos == 60:
            minutos = 55
        hora_deseada = f"{horas:02d}:{minutos:02d}"
        print(f"🔄 Hora ajustada a formato timepicker: {hora_deseada}")

        reloj_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".glyphicon-time"))
        )
        ActionChains(driver).move_to_element(reloj_icon).pause(0.3).click().perform()
        time.sleep(1)

        timepicker = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".bootstrap-timepicker-widget"))
        )
        input_hora = timepicker.find_element(By.CSS_SELECTOR, "input.bootstrap-timepicker-hour")
        input_minuto = timepicker.find_element(By.CSS_SELECTOR, "input.bootstrap-timepicker-minute")
        flecha_hora_up = timepicker.find_element(By.CSS_SELECTOR, "[data-action='incrementHour']")
        flecha_hora_down = timepicker.find_element(By.CSS_SELECTOR, "[data-action='decrementHour']")
        flecha_minuto_up = timepicker.find_element(By.CSS_SELECTOR, "[data-action='incrementMinute']")
        flecha_minuto_down = timepicker.find_element(By.CSS_SELECTOR, "[data-action='decrementMinute']")

        hora_actual = int(input_hora.get_attribute("value"))
        minuto_actual = int(input_minuto.get_attribute("value"))

        # Detectar si al bajar los minutos se va a restar 1 hora visualmente
        if minutos < minuto_actual:
            if horas == hora_actual:
                print("⚠️ Prevención: Añadiendo 1 a la hora porque vamos a bajar minutos")
                flecha_hora_up.click()
                time.sleep(0.2)

        # Ajustar la hora al valor deseado
        hora_actual = int(input_hora.get_attribute("value"))
        diferencia_horas = horas - hora_actual
        if diferencia_horas != 0:
            print(f"🕑 Ajustando horas de {hora_actual:02d} a {horas:02d}")
            flecha_hora = flecha_hora_up if diferencia_horas > 0 else flecha_hora_down
            for _ in range(abs(diferencia_horas)):
                flecha_hora.click()
                time.sleep(0.2)


        # Ajustar los minutos al valor deseado
        minuto_actual = int(input_minuto.get_attribute("value"))
        dif_normal = (minutos - minuto_actual) % 60
        dif_inversa = (minuto_actual - minutos) % 60

        if dif_normal <= dif_inversa:
            flecha = flecha_minuto_up
            pasos = dif_normal
            print(f"↑ Subir {pasos//5} clicks (de {minuto_actual:02d} a {minutos:02d})")
        else:
            flecha = flecha_minuto_down
            pasos = dif_inversa
            print(f"↓ Bajar {pasos//5} clicks (de {minuto_actual:02d} a {minutos:02d})")

        for _ in range(pasos // 5):
            flecha.click()
            time.sleep(0.15)

        # 🛠 Verificar si la hora ha bajado tras ajustar minutos y corregir
        hora_final_actual = int(input_hora.get_attribute("value"))
        if hora_final_actual < horas:
            print("⚠️ Se detectó que la hora bajó tras ajustar minutos, corrigiendo...")
            diferencia_horas_final = horas - hora_final_actual
            for _ in range(diferencia_horas_final):
                flecha_hora_up.click()
                time.sleep(0.2)

        # Verificación final
        hora_final = f"{input_hora.get_attribute('value')}:{input_minuto.get_attribute('value')}"
        if hora_final == hora_deseada:
            print(f"✅ Hora establecida correctamente: {hora_final}")
            return True
        else:
            print(f"⚠️ Hora resultante diferente: {hora_final} (esperado: {hora_deseada})")
            return False


    except Exception as e:
        print(f"❌ Error crítico: {str(e)}")
        driver.save_screenshot("error_timepicker_final.png")
        return False
    
def rellenar_modal_cita(driver, nombre_paciente, hora_deseada, fecha, agenda_num):
    """Versión mejorada para manejar el formulario de cita"""
    try:
        print("\n🖋️ Iniciando relleno de formulario...")
        # 1. Esperar a que el modal esté completamente cargado
        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.ID, "citaPaciente")))
        time.sleep(1)
        # 2. Autocompletado de paciente - Método mejorado
        input_paciente = driver.find_element(By.ID, "citaPaciente")
        # Limpiar el campo completamente
        for _ in range(3):
            input_paciente.clear()
            time.sleep(0.3)
        # Escribir el nombre con pausas naturales
        for i, char in enumerate(nombre_paciente):
            input_paciente.send_keys(char)
            # Pausa más larga después del primer nombre/apellido
            if char == ' ' or i == len(nombre_paciente.split()[0]):
                time.sleep(random.uniform(0.2, 0.4))
            else:
                time.sleep(random.uniform(0.05, 0.15))
        # Esperar y manejar sugerencias de autocompletado
        try:
            WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, ".autocomplete-list")))
            sugerencias = driver.find_elements(By.CSS_SELECTOR, ".autocomplete-list li")
            if sugerencias:
                # Buscar coincidencia exacta
                for sugerencia in sugerencias:
                    if nombre_paciente.lower() in sugerencia.text.lower():
                        ActionChains(driver).move_to_element(sugerencia).click().perform()
                        print(f"✅ Autocompletado seleccionado: {sugerencia.text}")
                        break
                else:
                    # Seleccionar la primera sugerencia si no hay coincidencia exacta
                    ActionChains(driver).move_to_element(sugerencias[0]).click().perform()
                    print(f"⚠️ Seleccionada primera sugerencia: {sugerencias[0].text}")
        except:
            print("ℹ️ No se encontraron sugerencias de autocompletado")
            # Continuar sin seleccionar sugerencia
        # 3. Configurar hora de la cita
        if not modificar_hora_en_modal(driver, hora_deseada):
            print("❌ No se pudo configurar la hora correctamente")
            return False
        # 5. Seleccionar facultativo según horario
        try:
            campo_observaciones = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "citaMotivo"))
            )
            campo_observaciones.click()
            time.sleep(0.5)
        except Exception as e:
            print(f"ℹ️ No se pudo hacer clic en observaciones: {str(e)}")
        time.sleep(1)
        
        if not seleccionar_facultativo_por_horario(driver, fecha, hora_deseada, agenda_num):
            print("⚠️ No se pudo seleccionar el facultativo según horario")
            return False
            
        # 6. Nueva: Seleccionar sala según reglas
        fecha_dt = datetime.strptime(fecha, "%d-%m-%Y")
        dia_semana = fecha_dt.weekday()
        if not seleccionar_sala(driver, agenda_num, dia_semana):
            print("⚠️ No se pudo seleccionar la sala, continuando igual...")
        # 6. Verificación final antes de guardar
        paciente_ingresado = driver.find_element(By.ID, "citaPaciente").get_attribute("value")
        if not paciente_ingresado:
            print("❌ Error: El campo paciente está vacío")
            return False
        print("✅ Formulario completado correctamente")
        return True
    except Exception as e:
        print(f"❌ Error crítico en formulario: {str(e)}")
        driver.save_screenshot("error_formulario.png")
        return False
def crear_cita_en_agenda(driver, fecha, hora, nombre_paciente, agenda):
    """Versión modificada para hacer clic siempre en el slot de 9:00 AM"""
    try:
        # Convertir hora a formato de 24 horas
        hora_dt = datetime.strptime(hora, "%H:%M")
        hora_24 = hora_dt.strftime("%H:%M")
        print(f"🕒 Hora confirmada: {hora_24}")
        
        if not navegar_a_fecha(driver, fecha):
            return False
            
        print("🔍 Buscando slot de 9:00 AM para abrir el modal...")
        time.sleep(2)
        
        # Estrategia específica para hacer clic en el slot de 9:00 AM
        try:
            # Buscar todas las celdas de hora
            celdas_hora = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td.fc-axis.fc-time.fc-widget-content")))
            
            # Encontrar la celda que contiene "9:00"
            slot_9am = None
            for celda in celdas_hora:
                if celda.text.strip() == "9:00":
                    slot_9am = celda
                    break
            
            if not slot_9am:
                print("❌ No se encontró el slot de 9:00 AM")
                return False
                
            # Hacer clic en la celda adyacente (la celda de agenda)
            # El HTML muestra que la celda de hora está seguida de una celda de contenido
            celda_agenda = slot_9am.find_element(By.XPATH, "./following-sibling::td[contains(@class, 'fc-widget-content')]")
            
            # Scroll y click preciso
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", celda_agenda)
            time.sleep(1)
            ActionChains(driver).move_to_element_with_offset(celda_agenda, 10, 10).click().perform()
            print("🖱️ Click realizado en el slot de 9:00 AM")
            time.sleep(3)
            
        except Exception as e:
            print(f"❌ Error al hacer clic en slot de 9:00 AM: {str(e)}")
            driver.save_screenshot("error_click_9am.png")
            return False
        # Verificar que el modal se abrió
        try:
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.ID, "citaPaciente")))
            print("✅ Modal de cita abierto correctamente")
        except:
            print("❌ El modal no apareció después del click")
            driver.save_screenshot("modal_no_aparecio.png")
            return False
            
        # Rellenar formulario (incluyendo la hora)
        if not rellenar_modal_cita(driver, nombre_paciente, hora_24, fecha, agenda):  # Añade agenda como parámetro
            return False
            
        # Guardar la cita
        try:
            guardar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "guardarCita")))
            driver.execute_script("arguments[0].scrollIntoView();", guardar)
            time.sleep(1)
            guardar.click()
            print("💾 Cita guardada correctamente")
            time.sleep(3)
            return True
        except Exception as e:
            print(f"❌ Error al guardar: {str(e)}")
            driver.save_screenshot("error_guardar.png")
            return False
            
    except Exception as e:
        print(f"❌ Error crítico: {str(e)}")
        driver.save_screenshot("error_final.png")
        return False
# Modificar la función actualizar_json_citas()
def actualizar_json_citas(fecha, hora, paciente, agenda):
    """Añade la nueva cita al JSON existente incluyendo el número de agenda"""
    try:
        fecha_dt = datetime.strptime(fecha, "%d-%m-%Y")
        fecha_iso = fecha_dt.strftime("%Y-%m-%d")
        
        # Leer el archivo actual
        try:
            with open(RUTA_CITAS, 'r', encoding='utf-8') as f:
                citas = json.load(f)
        except FileNotFoundError:
            citas = []
            
        # Calcular hora de fin (45 minutos después)
        hora_fin = (datetime.strptime(hora, "%H:%M") + timedelta(minutes=45)).strftime("%H:%M")
        
        # Añadir nueva cita
        nueva_cita = {
            "semana": 1 if len(citas) < 5 else 2,
            "dia": fecha_iso,
            "hora_inicio": hora,
            "hora_fin": hora_fin,
            "paciente": paciente,
            "agenda": str(agenda)  # Guardamos el número de agenda usado
        }
        citas.append(nueva_cita)
        
        # Guardar de nuevo
        with open(RUTA_CITAS, 'w', encoding='utf-8') as f:
            json.dump(citas, f, indent=2, ensure_ascii=False)
            
        print(f"📄 JSON actualizado con la nueva cita para {paciente} (Agenda {agenda}).")
    except Exception as e:
        print(f"⚠️ Error al actualizar JSON: {str(e)}")
def main():
    driver = configurar_navegador()
    try:
        if not login(driver):
            print("🚫 Error en el login, cerrando programa...")
            return
        
        print("\n=== VERIFICACIÓN DE PACIENTE ===")
        nombre_paciente = verificar_paciente()
        if not nombre_paciente:
            print("\n⚠️ Paciente no encontrado. Ejecutando script para crearlo...")
            subprocess.run([sys.executable, "Crear_usuario.py"])
            print("🔄 Vuelve a ejecutar gestion_citas.py después de crear el paciente.")
            return
        
        print("\n=== AGENDAMIENTO DE CITA ===")
        cita = seleccionar_cita()
        if cita:
            fecha, hora, agenda = cita  # Ahora recibimos también el número de agenda
            print(f"\n✅ Cita provisional agendada (Agenda {agenda}):")
            print(f"📅 Fecha: {fecha}")
            print(f"⏰ Hora: {hora}")
            
            if crear_cita_en_agenda(driver, fecha, hora, nombre_paciente, agenda):
                actualizar_json_citas(fecha, hora, nombre_paciente, agenda)
                print("\n🔄 Ejecutando actualización de Excel de citas...")
                subprocess.run([sys.executable, "descargar_excel.py"])
                
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        driver.save_screenshot("error.png")
    finally:
        driver.quit()
if __name__ == "__main__":
    main()
import pandas as pd
import json
import time
from datetime import datetime, timedelta
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
import subprocess
import sys

# Configuración global
load_dotenv("env/.env")
BASE_URL = "https://esiclinic.com/"
EXCEL_PACIENTES = "data/clientes/pacientes.xlsx"
JSON_CITAS = "data/citas_2_semanas.json"

# Configuración de horarios laborales
# Configuración de horarios
HORARIO_MANANA_AGENDA_1 = {
    'inicio': '10:15',
    'fin': '14:00'
}
HORARIO_MANANA_AGENDA_2 = {
    'inicio': '10:30',  # 15 mins después que Agenda 1
    'fin': '13:30'      # 30 mins antes que Agenda 1
}

HORARIO_TARDE_AGENDA_1 = {
    'inicio': '15:00',
    'fin': '20:15'
}
HORARIO_TARDE_AGENDA_2 = {
    'inicio': '15:15',  # 15 mins después que Agenda 1
    'fin': '19:45'      # 30 mins antes que Agenda 1
}

INTERVALO_CITAS = 45  # Duración de cada cita
DIAS_CON_TARDE = [0, 1, 2, 3, 4]  # Martes, Miércoles, Jueves, Viernes (ahora incluye jueves)
class GestorCitas:
    def __init__(self):
        self.driver = None
        self.datos_usuario = {
            "email": "",
            "nombre_completo": "",
            "dni": "",
            "citas": []
        }
        self.cita_seleccionada = None

    def configurar_navegador(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )

    def login(self):
        try:
            print("\n➡️ Iniciando sesión en esiclinic.com...")
            self.driver.get(BASE_URL)
            
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.ID, "esi_user"))
            ).send_keys(os.getenv("USUARIO_ESICLINIC"))
            
            self.driver.find_element(By.ID, "esi_pass").send_keys(os.getenv("PASSWORD_ESICLINIC"))
            
            btn_login = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "bt_acceder")))
            ActionChains(self.driver).move_to_element(btn_login).pause(0.5).click().perform()
            
            WebDriverWait(self.driver, 15).until(EC.url_contains("agenda.php"))
            print("✅ Sesión iniciada correctamente")
            return True
        except Exception as e:
            print(f"❌ Error en login: {str(e)}")
            self.driver.save_screenshot("error_login.png")
            return False

    def cargar_datos_paciente(self):
        try:
            pacientes_df = pd.read_excel(EXCEL_PACIENTES, engine="openpyxl")

            self.datos_usuario["email"] = input("\n✉️ Ingrese su email registrado: ").strip().lower()
            pacientes = pacientes_df[pacientes_df["E-Mail"].str.lower() == self.datos_usuario["email"]]

            if pacientes.empty:
                print("❌ Email no encontrado en la base de datos")
                opcion = input("¿Desea crear un nuevo paciente? (s/n): ").strip().lower()
                if opcion == "s":
                    print("🚀 Ejecutando Crear_usuario.py...")
                    subprocess.run([sys.executable, "Crear_usuario.py"])
                else:
                    print("\n👋 Programa finalizado, no se ha creado el paciente.")
                sys.exit(0)

            # Verificar si hay múltiples pacientes con el mismo correo
            if len(pacientes) > 1:
                print("\n⚠️ Este correo está asociado a múltiples pacientes:")
                print("----------------------------------------")
                
                # Mostrar lista de pacientes con este correo
                for i, (_, row) in enumerate(pacientes.iterrows(), 1):
                    print(f"{i}. {row['Nombre']} {row['Apellidos']} (DNI: {row['CIF']})")
                
                print("----------------------------------------")
                
                # Pedir selección por DNI o número de lista
                while True:
                    seleccion = input("🔢 Ingrese el número del paciente o el DNI completo: ").strip()
                    
                    # Si ingresa un número de la lista
                    if seleccion.isdigit() and 1 <= int(seleccion) <= len(pacientes):
                        paciente = pacientes.iloc[int(seleccion)-1]
                        break
                    # Si ingresa un DNI
                    elif seleccion.lower() in pacientes['CIF'].str.lower().values:
                        paciente = pacientes[pacientes['CIF'].str.lower() == seleccion.lower()].iloc[0]
                        break
                    else:
                        print("❌ Opción no válida. Intente nuevamente.")
                
                self.datos_usuario["nombre_completo"] = f"{paciente['Nombre']} {paciente['Apellidos']}"
                self.datos_usuario["dni"] = paciente['CIF']
                print(f"\n✅ Paciente seleccionado: {self.datos_usuario['nombre_completo']} (DNI: {self.datos_usuario['dni']})")
                return True
            
            # Solo un paciente con este correo
            else:
                paciente = pacientes.iloc[0]
                self.datos_usuario["nombre_completo"] = f"{paciente['Nombre']} {paciente['Apellidos']}"
                self.datos_usuario["dni"] = paciente['CIF']
                print(f"\n👤 Paciente encontrado: {self.datos_usuario['nombre_completo']} (DNI: {self.datos_usuario['dni']})")
                return True
                
        except Exception as e:
            print(f"❌ Error cargando datos: {str(e)}")
            return False

    def cargar_citas_desde_json(self):
        try:
            with open(JSON_CITAS, 'r', encoding='utf-8') as f:
                citas_data = json.load(f)
            
            # Generar variantes del nombre (apellidos primero)
            nombre_completo = self.datos_usuario["nombre_completo"]
            partes = nombre_completo.split()
            formatos_nombre = [
                nombre_completo.lower(),  # "sergi giovanni vique villafuerte"
                " ".join(partes[-2:] + partes[:-2]).lower(),  # "vique villafuerte sergi giovanni"
                " ".join(partes[-1:] + partes[:-1]).lower(),  # "villafuerte sergi giovanni vique"
            ]
            
            self.datos_usuario["citas"] = [
                {
                    "indice": i+1,
                    "fecha": c["dia"],
                    "hora": c["hora_inicio"],
                    "paciente": c["paciente"]
                }
                for i, c in enumerate(citas_data)
                if any(
                    formato in c["paciente"].lower() 
                    for formato in formatos_nombre
                )
            ]
            
            if not self.datos_usuario["citas"]:
                print("ℹ️ No hay citas registradas para este paciente")
                opcion = input("¿Desea crear una nueva cita? (s/n): ").strip().lower()
                if opcion == 's':
                    print("🚀 Ejecutando gestionar_citas.py...")
                    subprocess.run([sys.executable, "gestion_citas.py"])
                    sys.exit(0)
                else:
                    print("\n👋 Programa finalizado, no se ha creado ninguna cita.")
                    sys.exit(0)
                    
            print("\n📅 Tus citas registradas:")
            print("-" * 40)
            for cita in self.datos_usuario["citas"]:
                print(f"[{cita['indice']}] {cita['fecha']} a las {cita['hora']}")
            print("-" * 40)
            
            return True
        except Exception as e:
            print(f"❌ Error cargando citas desde JSON: {str(e)}")
            return False

    def buscar_paciente_por_dni(self):
        try:
            print("\n🔍 Buscando paciente en la agenda...")
            
            # Esperar y limpiar campo de búsqueda
            input_busqueda = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input#TpacienteWidget.form-control")))
            input_busqueda.clear()
            time.sleep(1)
            
            # Ingresar DNI carácter por carácter
            dni = str(self.datos_usuario["dni"]).upper()
            for char in dni:
                input_busqueda.send_keys(char)
                time.sleep(0.2)
            print(f"✓ DNI {dni} ingresado")
            time.sleep(3)
            
            # Construir múltiples variantes del nombre para búsqueda
            nombre_completo = self.datos_usuario["nombre_completo"]
            partes = nombre_completo.split()
            
            # Crear diferentes formatos de nombre
            formatos_nombre = [
                nombre_completo.lower(),  # "nombre apellido"
                " ".join(partes[-2:] + partes[:-2]).lower(),  # "apellido nombre"
                " ".join(partes[-1:] + partes[:-1]).lower(),  # "segundo_apellido nombre primer_apellido"
                " ".join(reversed(partes)).lower()  # "apellido apellido nombre"
            ]
            
            # Eliminar duplicados
            formatos_nombre = list(set(formatos_nombre))
            
            print(f"🔍 Probando formatos de nombre: {formatos_nombre}")
            
            # Intentar con cada formato hasta encontrar una coincidencia
            for formato in formatos_nombre:
                try:
                    # XPath más flexible que busca coincidencias parciales
                    opcion_xpath = f"""
                    //li[contains(
                        translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 
                        '{formato}'
                    )]
                    """
                    
                    opcion = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, opcion_xpath)))
                    
                    # Scroll y click con múltiples intentos
                    for intento in range(3):
                        try:
                            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", opcion)
                            time.sleep(0.5)
                            opcion.click()
                            print(f"✓ Paciente seleccionado con formato: {formato}")
                            time.sleep(2)
                            break
                        except Exception as e_click:
                            print(f"⚠️ Intento {intento+1} fallido: {str(e_click)}")
                            time.sleep(1)
                    else:
                        continue
                    
                    # Verificar que se seleccionó correctamente
                    try:
                        WebDriverWait(self.driver, 5).until(
                            lambda d: input_busqueda.get_attribute("value").lower() in formato)
                        
                        # Localizar el botón 'Ver citas' con múltiples selectores
                        selectores_boton = [
                            "button.btn.btn-info.menosmargen.pull-right.masCitas",
                            "button[onclick*='masCitas']",
                            "//button[contains(text(), 'Ver citas')]"
                        ]
                        
                        btn = None
                        for selector in selectores_boton:
                            try:
                                if selector.startswith("//"):
                                    btn = WebDriverWait(self.driver, 3).until(
                                        EC.element_to_be_clickable((By.XPATH, selector)))
                                else:
                                    btn = WebDriverWait(self.driver, 3).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                                break
                            except:
                                continue
                        
                        if not btn:
                            raise Exception("No se pudo encontrar el botón 'Ver citas'")
                        
                        # Intentar hacer clic de varias formas
                        try:
                            btn.click()
                        except:
                            self.driver.execute_script("arguments[0].click();", btn)
                            
                        print("✓ Botón 'Ver citas' clickeado")
                        time.sleep(2)
                        
                        return True
                        
                    except Exception as e_verificacion:
                        print(f"⚠️ No se verificó la selección: {str(e_verificacion)}")
                        continue
                        
                except Exception as e_opcion:
                    print(f"⚠️ Formato '{formato}' no funcionó: {str(e_opcion)}")
                    continue
                    
            print("❌ No se pudo seleccionar al paciente con ningún formato de nombre")
            self.driver.save_screenshot("error_seleccion_paciente.png")
            return False
                
        except Exception as e:
            print(f"❌ Error buscando paciente: {str(e)}")
            self.driver.save_screenshot("error_busqueda_paciente.png")
            return False
        
 
    def configurar_rango_fechas(self, fecha_cita):
        try:
            print("\n📅 Configurando rango de fechas...")
            
            fecha_cita_dt = datetime.strptime(fecha_cita, "%Y-%m-%d")
            fecha_inicio = (fecha_cita_dt - timedelta(days=1)).strftime("%d-%m-%Y")
            fecha_fin = (fecha_cita_dt + timedelta(days=15)).strftime("%d-%m-%Y")
            
            def establecer_fecha(campo_id, valor):
                campo = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, campo_id)))
                campo.clear()
                campo.send_keys(valor)
                campo.send_keys(Keys.RETURN)
                print(f"✓ Fecha {campo_id} establecida: {valor}")
            
            establecer_fecha("fecha", fecha_inicio)
            establecer_fecha("fecha2", fecha_fin)
            
            time.sleep(3)
            print("✓ Rango de fechas configurado correctamente")
            return True
        except Exception as e:
            print(f"❌ Error configurando fechas: {str(e)}")
            self.driver.save_screenshot("error_fechas.png")
            return False

    def buscar_y_seleccionar_cita(self, fecha_buscar, hora_buscar):
        try:
            print(f"\n🔍 Buscando cita: {fecha_buscar} a las {hora_buscar}")
            
            fecha_buscar_dt = datetime.strptime(fecha_buscar, "%Y-%m-%d")
            fecha_buscar_web = fecha_buscar_dt.strftime("%d-%m-%Y")
            
            # Debug: Mostrar qué estamos buscando exactamente
            print(f"Buscando fecha: {fecha_buscar_web} | Hora: {hora_buscar}")
            print(f"Nombre de referencia: {self.datos_usuario['nombre_completo']}")
            
            selectores_tabla = [
                "table.table",
                "table#tablaCitas",
                "div.table-container table",
                "//table[.//th[contains(text(),'Fecha')]]",
                "table.data-table"
            ]

            tabla = None
            for selector in selectores_tabla:
                try:
                    if selector.startswith("//"):
                        tabla = WebDriverWait(self.driver, 3).until(
                            EC.presence_of_element_located((By.XPATH, selector))
                        )
                    else:
                        tabla = WebDriverWait(self.driver, 3).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    print(f"✓ Tabla encontrada con selector: {selector}")
                    break
                except:
                    continue

            if not tabla:
                print("❌ No se pudo encontrar la tabla de citas con ningún selector")
                self.driver.save_screenshot("error_tabla_no_encontrada.png")
                return False

            filas = tabla.find_elements(By.XPATH, ".//tbody/tr[.//td]")
            if not filas:
                print("ℹ️ La tabla no contiene filas de datos")
                return False

            print(f"🔍 Analizando {len(filas)} citas...")
            
            for i, fila in enumerate(filas, 1):
                try:
                    celdas = fila.find_elements(By.TAG_NAME, "td")
                    if len(celdas) < 3:
                        continue
                        
                    fecha = celdas[0].text.strip()
                    hora = celdas[1].text.strip()
                    paciente = celdas[2].text.strip()

                    # Debug: Mostrar qué estamos comparando
                    print(f"[{i}] Comparando: {fecha} {hora} - {paciente}")
                    print(f"¿Coincide fecha? {fecha == fecha_buscar_web}")
                    print(f"¿Coincide hora? {hora == hora_buscar}")
                    print(f"¿Nombre contenido? {self.datos_usuario['dni'] in paciente or any(apellido.lower() in paciente.lower() for apellido in self.datos_usuario['nombre_completo'].split()[-2:])}")
                    
                    # Comparación más flexible
                    if (fecha == fecha_buscar_web and 
                        hora == hora_buscar and
                        (self.datos_usuario["dni"] in paciente or 
                        any(apellido.lower() in paciente.lower() 
                            for apellido in self.datos_usuario["nombre_completo"].split()[-2:]))):
                        
                        print(f"🎯 CITA ENCONTRADA en fila {i}")
                        
                        for intento in range(3):
                            try:
                                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", celdas[0])
                                time.sleep(0.5)
                                celdas[0].click()
                                print(f"✓ Clic en fecha realizado (intento {intento+1})")
                                
                                WebDriverWait(self.driver, 3).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, ".modal, .popup"))
                                )
                                return True
                            except Exception as e_click:
                                print(f"⚠️ Intento {intento+1} fallido: {str(e_click)}")
                                time.sleep(1)
                        
                        print("❌ No se pudo abrir el modal de edición")
                        return False
                        
                except Exception as e_fila:
                    print(f"⚠️ Error en fila {i}: {str(e_fila)}")
                    continue

            print("❌ No se encontró la cita especificada")
            return False
        except Exception as e:
            print(f"❌ Error crítico: {str(e)}")
            self.driver.save_screenshot("error_fatal_busqueda.png")
            return False
    def seleccionar_facultativo_por_horario(self, fecha, hora, agenda_num):
        """Selecciona el facultativo según el día de la semana y horario especificado"""
        max_intentos = 3
        intentos = 0
        
        while intentos < max_intentos:
            try:
                fecha_dt = datetime.strptime(fecha, "%d-%m-%Y")
                dia_semana = fecha_dt.weekday()
                hora_dt = datetime.strptime(hora, "%H:%M")

                # Esperar a que el modal esté completamente estable
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#citaMotivo, #citaFacultativo")))
                time.sleep(0.5)

                # Localizar el select de facultativo
                select_element = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "citaFacultativo")))
                
                select_facultativo = Select(select_element)

                # Obtener opciones actualizadas
                opciones_validas = {}
                try:
                    for op in select_facultativo.options:
                        try:
                            if op.text.strip() and "--" not in op.text:
                                opciones_validas[op.text.strip()] = op
                        except StaleElementReferenceException:
                            continue
                except StaleElementReferenceException:
                    print("⚠️ Las opciones cambiaron durante la iteración, reintentando...")
                    continue

                if not opciones_validas:
                    print("⚠️ No hay facultativos disponibles")
                    return False

                # Lógica de selección según horario y agenda
                facultativo_seleccionado = None
                if dia_semana == 0:  # Lunes
                    if agenda_num == 1:
                        facultativo_seleccionado = "Arnau Girones" if hora_dt.time() < datetime.strptime("14:00", "%H:%M").time() else "David Ibiza"
                    else:
                        facultativo_seleccionado = "Jose Cabanes"
                elif dia_semana == 1:  # Martes
                    facultativo_seleccionado = "Arnau Girones"
                elif dia_semana == 2:  # Miércoles
                    facultativo_seleccionado = "David Ibiza" if hora_dt.time() < datetime.strptime("12:30", "%H:%M").time() else "Arnau Girones"
                elif dia_semana == 3:  # Jueves
                    facultativo_seleccionado = "Jose Cabanes"
                elif dia_semana == 4:  # Viernes
                    facultativo_seleccionado = "Arnau Girones"

                # Seleccionar el facultativo
                for nombre, opcion in opciones_validas.items():
                    if facultativo_seleccionado and facultativo_seleccionado.lower() in nombre.lower():
                        opcion.click()
                        print(f"👨‍⚕️ Facultativo seleccionado (según horario): {nombre}")
                        
                        # Seleccionar sala correspondiente
                        self.seleccionar_sala(agenda_num, dia_semana)
                        return True

                # Si no se encontró el esperado, seleccionar el primero disponible
                if opciones_validas:
                    primera_opcion = list(opciones_validas.values())[0]
                    primera_opcion.click()
                    print(f"⚠️ Facultativo esperado no disponible. Seleccionado: {primera_opcion.text}")
                    return True

            except StaleElementReferenceException:
                intentos += 1
                print(f"⚠️ Intento {intentos}: Elemento obsoleto, reintentando...")
                time.sleep(1)
            except Exception as e:
                print(f"⚠️ Error al seleccionar facultativo: {str(e)}")
                self.driver.save_screenshot(f"error_facultativo_intento_{intentos}.png")
                return False

        print("❌ No se pudo seleccionar el facultativo después de varios intentos")
        return False
    def seleccionar_sala(self, agenda_num, dia_semana):
        """Selecciona Box 2 para agenda 2 en lunes/jueves, Box 1 para agenda 1"""
        try:
            select_element = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.ID, "modalRoom")))
            select_sala = Select(select_element)
            
            box = "Box 2" if agenda_num == 2 else "Box 1"
            select_sala.select_by_visible_text(box)
            print(f"🏥 Sala seleccionada: {box} (Agenda {agenda_num})")
            return True
        except Exception as e:
            print(f"⚠️ Error al seleccionar sala: {str(e)}")
            return False
    def ejecutar_descarga_excel(self):
        try:
            print("\n⬇️ Ejecutando script de descarga de Excel...")
            resultado = subprocess.run([sys.executable, "descargar_excel.py"], check=True, capture_output=True, text=True)
            print("✅ Script de descarga ejecutado correctamente")
            print(resultado.stdout)
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ Error al ejecutar descarga_excel.py: {e.stderr}")
            return False
        except Exception as e:
            print(f"❌ Error inesperado al ejecutar descarga_excel.py: {str(e)}")
            return False

    def cancelar_cita(self):
        try:
            if not self.cargar_citas_desde_json():
                return False
                
            seleccion = input("\nSeleccione el número de cita a cancelar (0 para salir): ").strip()
            
            if seleccion == "0":
                return False
                
            if not seleccion.isdigit() or int(seleccion) > len(self.datos_usuario["citas"]):
                print("❌ Selección inválida")
                return False
                
            cita = self.datos_usuario["citas"][int(seleccion)-1]
            self.cita_seleccionada = cita
            
            if not self.configurar_rango_fechas(cita["fecha"]):
                return False
                
            if not self.buscar_y_seleccionar_cita(cita["fecha"], cita["hora"]):
                return False
            
            confirmacion1 = input("\n⚠️ ¿Está seguro de que desea cancelar esta cita? (s/n): ").strip().lower()
            if confirmacion1 != 's':
                print("❌ Cancelación abortada por el usuario")
                return False
            
            try:
                btn_eliminar1 = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button.btn.btn-danger.lock.bt_eliminar")))
                
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_eliminar1)
                time.sleep(1)
                self.driver.execute_script("arguments[0].click();", btn_eliminar1)
                print("✓ Primer botón 'Eliminar' clickeado")
                time.sleep(2)
            except Exception as e:
                print(f"❌ Error al hacer clic en el primer botón Eliminar: {str(e)}")
                self.driver.save_screenshot("error_primer_eliminar.png")
                return False
            
            confirmacion2 = input("\n⚠️ CONFIRMACIÓN FINAL: ¿Está completamente seguro de eliminar esta cita? (s/n): ").strip().lower()
            if confirmacion2 != 's':
                print("❌ Cancelación abortada en última confirmación")
                try:
                    btn_cancelar = self.driver.find_element(By.CSS_SELECTOR, "button.btn-default[data-dismiss='modal']")
                    btn_cancelar.click()
                    print("✓ Modal de confirmación cerrado")
                except:
                    pass
                return False
            
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "div.jconfirm-box")))
                
                btn_eliminar2 = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div.jconfirm-box button.btn-danger")))
                
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_eliminar2)
                time.sleep(1)
                self.driver.execute_script("arguments[0].click();", btn_eliminar2)
                print("✓ Segundo botón 'Eliminar' clickeado")
                time.sleep(3)
            except Exception as e:
                print(f"❌ Error al hacer clic en el segundo botón Eliminar: {str(e)}")
                self.driver.save_screenshot("error_segundo_eliminar.png")
                return False
            
            if not self.actualizar_json_citas(eliminar=True):
                return False
                
            print("\n✅✅✅ CITA CANCELADA CORRECTAMENTE ✅✅✅")
            if not self.ejecutar_descarga_excel():
                print("⚠️ Se completó la cancelación pero falló la descarga del Excel")
            return True
            
        except Exception as e:
            print(f"❌ Error cancelando cita: {str(e)}")
            self.driver.save_screenshot("error_cancelar_cita.png")
            return False

    def seleccionar_cita_para_modificar(self):
        try:
            if not self.cargar_citas_desde_json():
                return False
                
            seleccion = input("\nSeleccione el número de cita a reagendar (0 para salir): ").strip()
            
            if seleccion == "0":
                return False
                
            if not seleccion.isdigit() or int(seleccion) > len(self.datos_usuario["citas"]):
                print("❌ Selección inválida")
                return False
            
            nueva_fecha = input("\n📅 Ingrese la nueva fecha (DD-MM-YYYY): ").strip()
            cita = self.datos_usuario["citas"][int(seleccion)-1]
            self.cita_seleccionada = cita
            
            if not self.configurar_rango_fechas(cita["fecha"]):
                return False
                
            if not self.buscar_y_seleccionar_cita(cita["fecha"], cita["hora"]):
                return False
            
            try:
                datetime.strptime(nueva_fecha, "%d-%m-%Y")
            except ValueError:
                print("❌ Formato de fecha incorrecto. Use DD-MM-YYYY")
                return False
            
            if not self.mostrar_horas_disponibles(nueva_fecha):
                return False
            
            if not self.modificar_campos_cita(nueva_fecha, nueva_hora):
                return False
                
            if not self.guardar_cambios_cita():
                return False
                
            if not self.actualizar_json_citas(nueva_fecha=nueva_fecha, nueva_hora=nueva_hora):
                return False
                
            print("\n✅✅✅ CITA REAGENDADA CORRECTAMENTE ✅✅✅")
            if not self.ejecutar_descarga_excel():
                print("⚠️ Se completó el reagendamiento pero falló la descarga del Excel")
            return True
            
        except Exception as e:
            print(f"❌ Error reagendando cita: {str(e)}")
            self.driver.save_screenshot("error_reagendar_cita.png")
            return False

    def mostrar_horas_disponibles(self, fecha):
        try:
            fecha_dt = datetime.strptime(fecha, "%d-%m-%Y")
            dia_semana = fecha_dt.weekday()  # 0=Lunes, 6=Domingo

            with open(JSON_CITAS, 'r', encoding='utf-8') as f:
                citas = json.load(f)

            citas_fecha = [c for c in citas if c['dia'] == fecha_dt.strftime("%Y-%m-%d")]

            print("\n🕒 Horarios disponibles (duración: 45 minutos):")
            self.horas_disponibles = []  # Reiniciar lista

            def hora_ocupada(hora_dt, agenda_id):
                for cita in citas_fecha:
                    if str(cita.get("agenda", "1")) != str(agenda_id):
                        continue
                    inicio = datetime.strptime(cita["hora_inicio"], "%H:%M")
                    fin = datetime.strptime(cita["hora_fin"], "%H:%M")
                    if inicio <= hora_dt < fin:
                        return True
                return False

            def mostrar_agenda(nombre, inicio_str, fin_str, agenda_id):
                print(f"\n📋 {nombre} ({inicio_str} - {fin_str})")
                hora_actual = datetime.strptime(inicio_str, "%H:%M")
                fin = datetime.strptime(fin_str, "%H:%M")

                while hora_actual + timedelta(minutes=INTERVALO_CITAS) <= fin:
                    hora_str = hora_actual.strftime("%H:%M")
                    if not hora_ocupada(hora_actual, agenda_id):
                        print(f"✅ {hora_str} - Disponible")
                        self.horas_disponibles.append((hora_str, agenda_id))
                    else:
                        print(f"❌ {hora_str} - Ocupado")
                    hora_actual += timedelta(minutes=INTERVALO_CITAS)

            # Mostrar Agenda 1 - Mañana
            mostrar_agenda("Agenda Principal (Mañana)", HORARIO_MANANA_AGENDA_1["inicio"], HORARIO_MANANA_AGENDA_1["fin"], "1")

            # Mostrar Agenda 2 - Mañana si lunes o jueves y agenda 1 llena
            if dia_semana in [0, 3]:
                agenda1_llena = all(
                    hora_ocupada(datetime.strptime(HORARIO_MANANA_AGENDA_1["inicio"], "%H:%M") + timedelta(minutes=INTERVALO_CITAS * i), "1")
                    for i in range(int((datetime.strptime(HORARIO_MANANA_AGENDA_1["fin"], "%H:%M") - datetime.strptime(HORARIO_MANANA_AGENDA_1["inicio"], "%H:%M")).seconds // 60 // INTERVALO_CITAS))
                )
                if agenda1_llena:
                    mostrar_agenda("Agenda Secundaria (Mañana)", HORARIO_MANANA_AGENDA_2["inicio"], HORARIO_MANANA_AGENDA_2["fin"], "2")

            # Mostrar Agenda 1 - Tarde
            if dia_semana in DIAS_CON_TARDE:
                mostrar_agenda("Agenda Principal (Tarde)", HORARIO_TARDE_AGENDA_1["inicio"], HORARIO_TARDE_AGENDA_1["fin"], "1")

                if dia_semana == 3:
                    agenda1_llena_tarde = all(
                        hora_ocupada(datetime.strptime(HORARIO_TARDE_AGENDA_1["inicio"], "%H:%M") + timedelta(minutes=INTERVALO_CITAS * i), "1")
                        for i in range(int((datetime.strptime(HORARIO_TARDE_AGENDA_1["fin"], "%H:%M") - datetime.strptime(HORARIO_TARDE_AGENDA_1["inicio"], "%H:%M")).seconds // 60 // INTERVALO_CITAS))
                    )
                    if agenda1_llena_tarde:
                        mostrar_agenda("Agenda Secundaria (Tarde)", HORARIO_TARDE_AGENDA_2["inicio"], HORARIO_TARDE_AGENDA_2["fin"], "2")

            if not self.horas_disponibles:
                print("❌ No hay horas disponibles.")
                return False

            hora_input = input("\n⏰ Ingrese la nueva hora deseada (HH:MM): ").strip()
            for hora, agenda in self.horas_disponibles:
                if hora_input == hora:
                    self.nueva_hora = hora
                    self.nueva_agenda = agenda
                    return True

            print("❌ Hora no válida o no disponible.")
            return False

        except Exception as e:
            print(f"❌ Error mostrando horarios: {str(e)}")
            return False
    def modificar_hora_en_modal(self, hora_deseada):
        """Versión robusta para ajustar hora con timepicker visual"""
        try:
            print(f"🕒 Configurando hora: {hora_deseada}")
            horas, minutos = map(int, hora_deseada.split(':'))

            # Activar el timepicker
            reloj_icon = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".glyphicon-time"))
            )
            ActionChains(self.driver).move_to_element(reloj_icon).pause(0.3).click().perform()
            time.sleep(1)

            timepicker = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, ".bootstrap-timepicker-widget"))
            )
            input_hora = timepicker.find_element(By.CSS_SELECTOR, "input.bootstrap-timepicker-hour")
            input_minuto = timepicker.find_element(By.CSS_SELECTOR, "input.bootstrap-timepicker-minute")
            flecha_hora_up = timepicker.find_element(By.CSS_SELECTOR, "[data-action='incrementHour']")
            flecha_minuto_up = timepicker.find_element(By.CSS_SELECTOR, "[data-action='incrementMinute']")
            flecha_hora_down = timepicker.find_element(By.CSS_SELECTOR, "[data-action='decrementHour']")
            flecha_minuto_down = timepicker.find_element(By.CSS_SELECTOR, "[data-action='decrementMinute']")

            # Calcular diferencia de hora
            hora_actual = int(input_hora.get_attribute("value"))
            diferencia_horas = horas - hora_actual
            flecha_hora = flecha_hora_up if diferencia_horas > 0 else flecha_hora_down
            for _ in range(abs(diferencia_horas)):
                flecha_hora.click()
                time.sleep(0.2)

            # Ajustar minutos al múltiplo de 5 más cercano
            minutos = (minutos // 5) * 5
            minuto_actual = int(input_minuto.get_attribute("value"))
            dif_normal = (minutos - minuto_actual) % 60
            dif_inversa = (minuto_actual - minutos) % 60
            if minutos == 0 and minuto_actual > 0:
                flecha = flecha_minuto_down
                pasos = dif_inversa
            else:
                flecha = flecha_minuto_up if dif_normal <= dif_inversa else flecha_minuto_down
                pasos = min(dif_normal, dif_inversa)
            for _ in range(pasos // 5):
                flecha.click()
                time.sleep(0.15)

            hora_final = f"{input_hora.get_attribute('value')}:{input_minuto.get_attribute('value')}"
            if hora_final == f"{horas:02d}:{minutos:02d}":
                print(f"✅ Hora establecida correctamente: {hora_final}")
                return True
            else:
                print(f"⚠️ Hora final no coincide exactamente: {hora_final} (esperado: {hora_deseada})")
                return False

        except Exception as e:
            print(f"❌ Error configurando hora visual: {str(e)}")
            self.driver.save_screenshot("error_timepicker_visual.png")
            return False

    def modificar_campos_cita(self, nueva_fecha, nueva_hora):
        try:
            print(f"\n⌛ Configurando nueva fecha: {nueva_fecha}")
            input_fecha = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "citaFecha")))
            
            self.driver.execute_script("arguments[0].value = '';", input_fecha)
            input_fecha.send_keys(nueva_fecha)
            print(f"✓ Fecha establecida: {nueva_fecha}")
            time.sleep(1)
            
            print(f"\n⌛ Configurando nueva hora: {nueva_hora}")
            
            try:
                # Método directo
                input_hora = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.ID, "citaHora")))
                
                self.driver.execute_script("arguments[0].value = arguments[1];", input_hora, nueva_hora)
                input_hora.send_keys(Keys.TAB)
                print("✓ Hora establecida directamente")

            except:
                # 🔄 Si falla el método directo, intenta con el timepicker visual
                print("ℹ️ Falló el método directo, usando timepicker visual...")
                if not self.modificar_hora_en_modal(nueva_hora):
                    print("❌ No se pudo modificar la hora con el timepicker visual")
                    return False
            
            # ✅ SELECCIONAR EL FACULTATIVO AQUÍ, SIEMPRE, después de poner la hora
            if not self.seleccionar_facultativo_por_horario(nueva_fecha, nueva_hora, self.nueva_agenda):
                print("⚠️ No se pudo seleccionar el facultativo según el nuevo horario")
                return False

            print(f"✓ Hora establecida: {nueva_hora}")
            return True

        except Exception as e:
            print(f"❌ Error crítico al modificar hora: {str(e)}")
            self.driver.save_screenshot("error_estableciendo_hora.png")
            return False

    def ajustar_hora_con_flechas(self, timepicker, hora_deseada, minuto_deseado):
        try:
            print("🔧 Usando método de flechas para ajustar hora...")
            
            flecha_hora_up = timepicker.find_element(By.CSS_SELECTOR, "[data-action='incrementHour']")
            flecha_hora_down = timepicker.find_element(By.CSS_SELECTOR, "[data-action='decrementHour']")
            flecha_minuto_up = timepicker.find_element(By.CSS_SELECTOR, "[data-action='incrementMinute']")
            flecha_minuto_down = timepicker.find_element(By.CSS_SELECTOR, "[data-action='decrementMinute']")
            
            hora_input = timepicker.find_element(By.CSS_SELECTOR, "input.bootstrap-timepicker-hour")
            minuto_input = timepicker.find_element(By.CSS_SELECTOR, "input.bootstrap-timepicker-minute")
            
            hora_actual = int(hora_input.get_attribute("value"))
            minuto_actual = int(minuto_input.get_attribute("value"))
            
            diferencia_horas = int(hora_deseada) - hora_actual
            if diferencia_horas != 0:
                flecha = flecha_hora_up if diferencia_horas > 0 else flecha_hora_down
                for _ in range(abs(diferencia_horas)):
                    flecha.click()
                    time.sleep(0.2)
            
            minuto_deseado = (int(minuto_deseado) // 5) * 5
            diferencia_minutos = minuto_deseado - minuto_actual
            
            if diferencia_minutos > 0:
                for _ in range(diferencia_minutos // 5):
                    flecha_minuto_up.click()
                    time.sleep(0.1)
            elif diferencia_minutos < 0:
                for _ in range(abs(diferencia_minutos) // 5):
                    flecha_minuto_down.click()
                    time.sleep(0.1)
            
            print("✓ Hora ajustada con flechas")
            return True
            
        except Exception as e:
            print(f"❌ Error ajustando hora con flechas: {str(e)}")
            return False

    def guardar_cambios_cita(self):
        try:
            btn_modificar = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "guardarCita")))
            
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_modificar)
            time.sleep(1)
            self.driver.execute_script("arguments[0].click();", btn_modificar)
            print("✓ Cambios guardados (clic en Modificar)")
            time.sleep(3)
            
            return True
            
        except Exception as e:
            print(f"❌ Error guardando cambios: {str(e)}")
            self.driver.save_screenshot("error_guardando_cambios.png")
            return False

    def actualizar_json_citas(self, nueva_fecha=None, nueva_hora=None, eliminar=False):
        try:
            with open(JSON_CITAS, 'r', encoding='utf-8') as f:
                citas = json.load(f)

            cita_original = self.cita_seleccionada
            dni_paciente = self.datos_usuario["dni"]

            citas_actualizadas = []
            modificada = False

            for cita in citas:
                coincide_fecha = cita["dia"] == cita_original["fecha"]
                coincide_hora = cita["hora_inicio"] == cita_original["hora"]
                coincide_paciente = any(nombre.lower() in cita.get("paciente", "").lower() 
                        for nombre in self.datos_usuario["nombre_completo"].split())

                if coincide_fecha and coincide_hora and coincide_paciente:
                    if eliminar:
                        print("🗑️ Eliminando cita del JSON...")
                        continue  # no se añade a la nueva lista
                    else:
                        cita["dia"] = datetime.strptime(nueva_fecha, "%d-%m-%Y").strftime("%Y-%m-%d")
                        cita["hora_inicio"] = nueva_hora
                        cita["hora_fin"] = (datetime.strptime(nueva_hora, "%H:%M") + timedelta(minutes=INTERVALO_CITAS)).strftime("%H:%M")
                        cita["agenda"] = str(self.nueva_agenda)  # 🆕 Actualizar número de agenda
                        print(f"♻️ Cita actualizada en JSON: {cita}")
                        modificada = True

                citas_actualizadas.append(cita)

            if not modificada and not eliminar:
                print("⚠️ No se encontró la cita a modificar en el JSON.")
                return False

            with open(JSON_CITAS, 'w', encoding='utf-8') as f:
                json.dump(citas_actualizadas, f, indent=2, ensure_ascii=False)

            print("💾 JSON actualizado correctamente.")
            return True

        except Exception as e:
            print(f"❌ Error actualizando JSON: {str(e)}")
            return False

    def mostrar_menu(self):
        print("\n=== MENÚ PRINCIPAL ===")
        print("1. Reagendar cita")
        print("2. Cancelar cita")
        print("0. Salir")
        return input("Seleccione una opción: ").strip()

    def ejecutar(self):
        try:
             # ✅ Ejecutar extracción de citas antes de todo
            print("\n📦 Actualizando JSON de citas (extraer_citas.py)...")
            subprocess.run([sys.executable, "extraer_citas.py"], check=True)
            print("✅ JSON actualizado correctamente.\n")
            self.driver = self.configurar_navegador()
            
            if not self.login():
                sys.exit(1)
                
            if not self.cargar_datos_paciente():
                sys.exit(1)
                
            if not self.buscar_paciente_por_dni():
                sys.exit(1)
            
            while True:
                opcion = self.mostrar_menu()
                
                if opcion == "1":
                    if self.seleccionar_cita_para_modificar():
                        sys.exit(0)
                elif opcion == "2":
                    if self.cancelar_cita():
                        sys.exit(0)
                elif opcion == "0":
                    print("\n👋 Programa finalizado por el usuario")
                    sys.exit(0)
                else:
                    print("❌ Opción no válida")       
        except Exception as e:
            print(f"❌ Error crítico: {str(e)}")
            sys.exit(1)
        finally:
            if hasattr(self, 'driver') and self.driver:
                self.driver.quit()

if __name__ == "__main__":
    gestor = GestorCitas()
    gestor.ejecutar()
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import traceback
import os

def subir_a_google_sheets(nombre_archivo, nombre_hoja):
    """""
    Sube un archivo Excel a Google Sheets, reemplazando los datos existentes
    Args:
        nombre_archivo (str): Ruta del archivo Excel a subir
        nombre_hoja (str): Nombre de la hoja de cálculo en Google Sheets
    """
    try:
        print(f"Subiendo {nombre_archivo} a Google Sheets...")
        # 1. Configuración de credenciales
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            "env/credentials.json", scope)
        client = gspread.authorize(creds)
        # 2. Leer el archivo Excel (con motor explícito)
        try:
            # Primero intentamos con openpyxl (para .xlsx)
            df = pd.read_excel(nombre_archivo, engine='openpyxl')
        except:
            # Si falla, probamos con xlrd (para .xls más antiguos)
            try:
                df = pd.read_excel(nombre_archivo, engine='xlrd')
            except Exception as e:
                print(f"No se pudo leer el archivo Excel: {str(e)}")
                return False
        # 3. Subir a Google Sheets
        try:
            spreadsheet = client.open(nombre_hoja)
        except gspread.SpreadsheetNotFound:
            print(f"Creando nueva hoja de cálculo: {nombre_hoja}")
            spreadsheet = client.create(nombre_hoja)
            # Compartir con el email autorizado
            spreadsheet.share(os.getenv("GMAIL_FROM"), perm_type='user', role='writer')
        worksheet = spreadsheet.sheet1
        # 4. Limpiar y actualizar datos
        worksheet.clear()
        # Convertir todos los datos a string para evitar problemas de formato
        data = [df.columns.values.tolist()] + df.astype(str).values.tolist()
        worksheet.update(
            data,
            value_input_option='USER_ENTERED'
        )
        print(f"Datos actualizados en Google Sheets: {nombre_hoja}")
        return True
    except Exception as e:
        print(f"Error al subir a Google Sheets: {str(e)}")
        print(traceback.format_exc())
        return False
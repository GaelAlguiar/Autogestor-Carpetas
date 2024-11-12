import os
import shutil
import tkinter as tk
from tkinter import simpledialog, messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import locale

# Configura el locale a español (esto puede variar dependiendo del sistema operativo)
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Para sistemas Unix/Linux/Mac
# Para Windows, podrías usar 'Spanish_Spain.1252'

def leer_datos_google_sheets():
    # Configura el acceso a Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/gaelalguiar/Downloads/credenciales.json', scope)
    client = gspread.authorize(creds)

    # Abre la hoja de Google Sheets
    try:
        sheet = client.open("ENEREY 2024").sheet1  # Cambia si tu hoja no es la principal
        records = sheet.get_all_records()
        return records
    except gspread.SpreadsheetNotFound:
        messagebox.showerror("Error", "No se encontró la hoja de Google Sheets 'ENEREY 2024'. Verifica el nombre.")
        return []
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al leer los datos: {e}")
        return []

def crear_txt(folder_path, nombre_archivo, contenido):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    archivo_path = os.path.join(folder_path, nombre_archivo)
    try:
        with open(archivo_path, 'w') as archivo:
            archivo.write(contenido)
        print(f'Se creó el archivo: {archivo_path}')
    except Exception as e:
        print(f'Error al crear el archivo {archivo_path}: {e}')


def crear_carpetas():
    records = leer_datos_google_sheets()
    if not records:
        return  # Sale si no hay registros

    # Solicita al usuario el rango de inicio y fin a través de cuadros de diálogo
    inicio = simpledialog.askinteger("Entrada", "Ingresa el número inicial del pedido:")
    fin = simpledialog.askinteger("Entrada", "Ingresa el número final del pedido:")

    if inicio is None or fin is None:
        messagebox.showerror("Error", "Debes ingresar un rango válido.")
        return
    if inicio > fin:
        messagebox.showerror("Error", "El número inicial debe ser menor o igual al número final.")
        return

    # Ruta a la carpeta de Descargas y creación de la carpeta contenedora
    downloads_path = os.path.expanduser('~/Downloads')
    carpeta_contenedora = os.path.join(downloads_path, 'Carpetas')

    # Elimina la carpeta si ya existe
    if os.path.exists(carpeta_contenedora):
        shutil.rmtree(carpeta_contenedora)
        print(f'Se eliminó la carpeta existente: {carpeta_contenedora}')

    os.makedirs(carpeta_contenedora)
    messagebox.showinfo("Información", f'Se creó la carpeta contenedora: {carpeta_contenedora}')

    # Filtra los registros según el rango de pedidos
    registros_filtrados = [
        record for record in records
        if isinstance(record.get('Pedido'), str) and record.get('Pedido').startswith('DIS')
        and record.get('Pedido')[3:].split('/')[0].strip().isdigit()  # Extrae el número después de 'DIS'
        and inicio <= int(record.get('Pedido')[3:].split('/')[0].strip()) <= fin
    ]

    # Crea las carpetas y archivos .txt
    for record in registros_filtrados:
        pedido = record.get('Pedido')
        fac_venta = record.get('FactVenta', '').replace(" ", "") 
        fact_flete = record.get('Fact Flete', '').replace(" ", "")
        fact_complemento = record.get('Fact Complemento', '').strip()
        fecha_pedido = record.get('Fecha Pedido')

        # Imprime la información de pedido y facturas para depuración
        print(f'Procesando pedido: {pedido}, FactVenta: {fac_venta}, Fact Flete: {fact_flete}, Fact Complemento: {fact_complemento}')

        if not pedido or not fecha_pedido:
            continue

        # Maneja el formato de la fecha asegurando que sea una cadena
        if isinstance(fecha_pedido, int):
            fecha_pedido = str(fecha_pedido)  # Convierte enteros a cadena
        elif not isinstance(fecha_pedido, str):
            print(f"Formato de fecha no válido para el pedido {pedido}: {fecha_pedido}")
            continue

        # Intenta convertir la fecha
        try:
            fecha_obj = datetime.strptime(fecha_pedido, "%d/%m/%Y")  # Ajustado para el formato correcto
            mes_num = fecha_obj.strftime("%m")  # Número del mes
            mes_str = fecha_obj.strftime("%B").upper()  # Nombre completo del mes en mayúsculas
            mes_folder_name = f'{mes_num} {mes_str}'  # Formato '01 ENERO', '02 FEBRERO', etc.
            dia_str = fecha_obj.strftime("%d")  # Día con formato de dos dígitos
        except ValueError:
            print(f"Formato de fecha inválido para el pedido {pedido}: {fecha_pedido}")
            continue

        # Crea la carpeta para el mes
        mes_folder = os.path.join(carpeta_contenedora, mes_folder_name)
        if not os.path.exists(mes_folder):
            os.makedirs(mes_folder)

        # Crea la carpeta para el día
        dia_folder = os.path.join(mes_folder, dia_str)
        if not os.path.exists(dia_folder):
            os.makedirs(dia_folder)

        # Crea la carpeta para el pedido
        if isinstance(pedido, str) and ('/ 1' in pedido or '/ 2' in pedido):
            pedido_base = pedido.split('/')[0].strip()
            sufijo = pedido.split('/')[-1].strip()
            pedido_folder = os.path.join(dia_folder, f'{pedido_base} {sufijo}')
        else:
            pedido_folder = os.path.join(dia_folder, f'{pedido}')

        if not os.path.exists(pedido_folder):
            os.makedirs(pedido_folder)
            print(f'Se creó la carpeta: {pedido_folder}')

        # Crea archivos .txt para FactVenta, Fact Flete, y Fact Complemento si aplican
        if fac_venta and fac_venta.upper() != 'NA':
            crear_txt(pedido_folder, f'Venta_{fac_venta}.txt', fac_venta)
        if fact_flete and fact_flete.upper() != 'NA':
            crear_txt(pedido_folder, f'Flete_{fact_flete}.txt', fact_flete)
        if fact_complemento and fact_complemento.upper() != 'NA':
            crear_txt(pedido_folder, f'Complemento_{fact_complemento}.txt', fact_complemento)

    messagebox.showinfo("Finalizado", f'Se han creado las carpetas y archivos desde DIS{inicio} hasta DIS{fin} en {carpeta_contenedora}')

# Configuración de la ventana principal de tkinter
root = tk.Tk()
root.withdraw()  # Oculta la ventana principal

crear_carpetas()

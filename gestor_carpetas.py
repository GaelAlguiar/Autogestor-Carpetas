import os
import shutil
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import locale
import re

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
    # Asegúrate de que el directorio esté creado antes de intentar escribir en el archivo
    os.makedirs(folder_path, exist_ok=True)  # Crea la ruta completa si no existe
    
    archivo_path = os.path.join(folder_path, nombre_archivo)
    try:
        with open(archivo_path, 'w') as archivo:
            archivo.write(contenido)
        print(f'Se creó el archivo: {archivo_path}')
    except Exception as e:
        print(f'Error al crear el archivo {archivo_path}: {e}')

def iniciar_creacion_carpetas(inicio, fin, root):
    records = leer_datos_google_sheets()
    if not records:
        return  # Sale si no hay registros

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
        if isinstance(record.get('Pedido'), str) and re.match(r'DIS\s*\d+', record.get('Pedido'))  # Verifica que empiece con 'DIS' seguido de un número
        and inicio <= int(re.search(r'DIS\s*(\d+)', record.get('Pedido')).group(1)) <= fin
    ]

    # Crea las carpetas y archivos .txt
    for record in registros_filtrados:
        pedido = record.get('Pedido').strip()  # Mantiene el pedido completo
        pedido = re.sub(r'\s+', ' ', pedido)  # Elimina espacios múltiples
        pedido = pedido.replace('/', '').replace('  ', ' ')  # Elimina '/' y dobles espacios
        fac_venta = str(record.get('FactVenta', '')).replace(" ", "")  # Elimina espacios de la FactVenta
        fact_flete = str(record.get('Fact Flete', '')).strip()
        fact_complemento = str(record.get('Fact Complemento', '')).strip()
        fact_compra = str(record.get('Fact Compra', '')).strip()  # Nueva línea para Fact Compra
        fecha_fact_venta = record.get('Fecha Fact Venta')  # Usa Fecha Fact Venta

        # Imprime la información de pedido y facturas para depuración
        print(f'Procesando pedido: {pedido}, FactVenta: {fac_venta}, Fact Flete: {fact_flete}, Fact Complemento: {fact_complemento}, Fact Compra: {fact_compra}')

        if not pedido or not fecha_fact_venta:
            continue

        # Maneja el formato de la fecha asegurando que sea una cadena
        if isinstance(fecha_fact_venta, int):
            fecha_fact_venta = str(fecha_fact_venta)  # Convierte enteros a cadena
        elif not isinstance(fecha_fact_venta, str):
            print(f"Formato de fecha no válido para el pedido {pedido}: {fecha_fact_venta}")
            continue

        # Intenta convertir la fecha
        try:
            fecha_obj = datetime.strptime(fecha_fact_venta, "%d/%m/%Y")  # Ajustado para el formato correcto
            mes_num = fecha_obj.strftime("%m")  # Número del mes
            mes_str = fecha_obj.strftime("%B").upper()  # Nombre completo del mes en mayúsculas
            mes_folder_name = f'{mes_num} {mes_str}'  # Formato '01 ENERO', '02 FEBRERO', etc.
            dia_str = fecha_obj.strftime("%d")  # Día con formato de dos dígitos
        except ValueError:
            print(f"Formato de fecha inválido para el pedido {pedido}: {fecha_fact_venta}")
            continue

        # Crea la carpeta para el mes
        mes_folder = os.path.join(carpeta_contenedora, mes_folder_name)
        os.makedirs(mes_folder, exist_ok=True)

        # Crea la carpeta para el día
        dia_folder = os.path.join(mes_folder, dia_str)
        os.makedirs(dia_folder, exist_ok=True)

        # Crea la carpeta para el pedido
        pedido_folder = os.path.join(dia_folder, pedido)
        os.makedirs(pedido_folder, exist_ok=True)

        # Crea archivos .txt para FactVenta, Fact Flete, Fact Complemento y Fact Compra si aplican
        if fac_venta and fac_venta.upper() != 'NA':
            crear_txt(pedido_folder, f'Venta_{fac_venta}.txt', fac_venta)
        if fact_flete and fact_flete.upper() != 'NA':
            crear_txt(pedido_folder, f'Flete_{fact_flete}.txt', fact_flete)
        if fact_complemento and fact_complemento.upper() != 'NA':
            crear_txt(pedido_folder, f'Complemento_{fact_complemento}.txt', fact_complemento)
        if fact_compra and fact_compra.upper() != 'NA':  # Nueva condición para Fact Compra
            crear_txt(pedido_folder, f'Compra_{fact_compra}.txt', fact_compra)

    messagebox.showinfo("Finalizado", f'Se han creado las carpetas y archivos desde DIS{inicio} hasta DIS{fin} en {carpeta_contenedora}')
    root.destroy()  # Cierra la ventana al finalizar

def mostrar_interfaz():
    root = tk.Tk()
    root.title("Gestor de Carpetas")
    root.geometry("600x350")  # Aumenta el tamaño de la ventana
    root.resizable(False, False)

    # Calcula la posición para centrar la ventana
    window_width = 600
    window_height = 350
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height/2 - window_height/2)
    position_right = int(screen_width/2 - window_width/2)

    root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

    # Crear un "AppBar" o encabezado con fondo verde
    header_frame = tk.Frame(root, bg="green", height=50)
    header_frame.pack(fill=tk.X)
    header_label = tk.Label(header_frame, text="Gestor de Carpetas", bg="green", fg="white", font=("Arial", 20))
    header_label.pack(pady=10)

    frame = ttk.Frame(root, padding=50)
    frame.pack(expand=True, fill=tk.BOTH)

    ttk.Label(frame, text="Rango inicial del pedido:", font=("Arial", 16)).grid(column=0, row=0, pady=10, padx=50, sticky=tk.W)
    inicio_entry = ttk.Entry(frame, font=("Arial", 13))
    inicio_entry.grid(column=1, row=0, pady=10, padx=10)

    ttk.Label(frame, text="Rango final del pedido:", font=("Arial", 16)).grid(column=0, row=1, pady=10, padx=50, sticky=tk.W)
    fin_entry = ttk.Entry(frame, font=("Arial", 13))
    fin_entry.grid(column=1, row=1, pady=10, padx=10)

    def iniciar_proceso():
        try:
            inicio = int(inicio_entry.get())
            fin = int(fin_entry.get())
            iniciar_creacion_carpetas(inicio, fin, root)
        except ValueError:
            messagebox.showerror("Error", "Por favor, ingresa números válidos.")

    ttk.Button(frame, text="Iniciar Proceso", command=iniciar_proceso).grid(column=0, row=2, columnspan=5, pady=50)

    root.mainloop()

mostrar_interfaz()

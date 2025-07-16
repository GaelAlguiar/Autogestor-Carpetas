import gc
import os
import re
import zipfile
import shutil
import tempfile
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

app = Flask(__name__, static_url_path='/static', static_folder='static')
app.secret_key = os.getenv("FLASK_SECRET_KEY", "clave_segura")

MESES_ES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

def leer_datos_google_sheets(json_path):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
    client = gspread.authorize(creds)
    try:
        sheet = client.open("ENEREY 2025").sheet1
        return sheet.get_all_records()
    except Exception as e:
        print(f"Error al leer la hoja: {e}")
        return []

def limpiar_nombre(nombre):
    return re.sub(r'[\/:*?"<>|]', '_', str(nombre))

def crear_txt(folder_path, nombre_archivo, contenido):
    os.makedirs(folder_path, exist_ok=True)
    archivo_path = os.path.join(folder_path, nombre_archivo)
    try:
        with open(archivo_path, 'w') as archivo:
            archivo.write(contenido)
            archivo.flush()
    except Exception as e:
        print(f'Error al crear archivo {archivo_path}: {e}')

def comprimir_zip(output_zip_path, folder_a_comprimir):
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_a_comprimir):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, folder_a_comprimir)
                zipf.write(abs_path, arcname=rel_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            inicio = int(request.form['inicio'])
            fin = int(request.form['fin'])
            file = request.files.get('credenciales')

            if not file or file.filename == '':
                flash("Debe subir el archivo de credenciales.", "danger")
                return redirect(url_for('index'))

            with tempfile.TemporaryDirectory() as temp_dir:
                cred_path = os.path.join(temp_dir, 'credenciales.json')
                file.save(cred_path)

                records = leer_datos_google_sheets(cred_path)
                if not records:
                    flash("No se pudieron leer datos de Google Sheets.", "danger")
                    return redirect(url_for('index'))

                carpeta_contenedora = os.path.join(temp_dir, 'Carpetas')
                os.makedirs(carpeta_contenedora, exist_ok=True)

                BATCH_SIZE = 100
                for i in range(0, len(records), BATCH_SIZE):
                    lote = records[i:i + BATCH_SIZE]

                    for record in lote:
                        pedido = str(record.get('Pedido', '')).strip()

                        if not pedido or not re.match(r'DIS\s*\d+', pedido):
                            continue

                        match = re.search(r'DIS\s*(\d+)', pedido)
                        if not match:
                            continue

                        num_pedido = int(match.group(1))
                        if num_pedido < inicio or num_pedido > fin:
                            continue

                        fecha_fact_venta = str(record.get('Fecha Fact Venta', '')).strip()
                        if not re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', fecha_fact_venta):
                            continue

                        try:
                            fecha_obj = datetime.strptime(fecha_fact_venta, "%d/%m/%Y")
                            mes_nombre = MESES_ES[fecha_obj.month - 1]
                            mes_folder = f"{fecha_obj.strftime('%m')} {mes_nombre}"
                            dia_str = fecha_obj.strftime('%d')
                        except Exception:
                            continue

                        pedido_folder = os.path.join(
                            carpeta_contenedora, mes_folder, dia_str, limpiar_nombre(pedido)
                        )
                        os.makedirs(pedido_folder, exist_ok=True)

                        def escribir_factura(tipo, valor):
                            if valor and valor.upper() != 'NA':
                                nombre = f"{tipo}_{limpiar_nombre(valor)}.txt"
                                crear_txt(pedido_folder, nombre, valor)

                        escribir_factura("Venta", record.get('FactVenta', '').replace(" ", ""))
                        escribir_factura("Flete", record.get('Fact Flete', ''))
                        escribir_factura("Complemento", record.get('Fact Complemento', ''))
                        escribir_factura("Compra", record.get('Fact Compra', ''))

                    gc.collect()  # Liberar memoria tras cada lote

                # Comprimir carpeta (compatible con macOS)
                zip_path = os.path.join(temp_dir, 'carpetas.zip')
                comprimir_zip(zip_path, carpeta_contenedora)

                if not os.path.exists(zip_path) or os.path.getsize(zip_path) == 0:
                    flash("El archivo ZIP está vacío o corrupto.", "danger")
                    return redirect(url_for('index'))

                return send_file(zip_path, as_attachment=True, download_name="carpetas.zip")

        except ValueError:
            flash("Por favor ingresa números válidos.", "danger")
        except Exception as e:
            flash(f"Ocurrió un error: {str(e)}", "danger")
        return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

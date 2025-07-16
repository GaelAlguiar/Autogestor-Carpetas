from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
import tempfile
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re
import zipfile

app = Flask(__name__, static_url_path='/static', static_folder='static')
app.secret_key = os.getenv("FLASK_SECRET_KEY")

# Lista manual de meses en español (mayúsculas)
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
        records = sheet.get_all_records()
        return records
    except Exception as e:
        print(f"Error al leer la hoja: {e}")
        return []


def crear_txt(folder_path, nombre_archivo, contenido):
    os.makedirs(folder_path, exist_ok=True)
    archivo_path = os.path.join(folder_path, nombre_archivo)
    try:
        with open(archivo_path, 'w') as archivo:
            archivo.write(contenido)
        print(f'Se creó el archivo: {archivo_path}')
    except Exception as e:
        print(f'Error al crear el archivo {archivo_path}: {e}')


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

            # Carpeta temporal
            with tempfile.TemporaryDirectory() as temp_dir:
                cred_path = os.path.join(temp_dir, 'credenciales.json')
                file.save(cred_path)

                records = leer_datos_google_sheets(cred_path)
                if not records:
                    flash("No se pudieron leer datos de Google Sheets.", "danger")
                    return redirect(url_for('index'))

                carpeta_contenedora = os.path.join(temp_dir, 'Carpetas')
                os.makedirs(carpeta_contenedora, exist_ok=True)

                for record in records:
                    pedido = record.get('Pedido', '').strip()

                    if not pedido or not re.match(r'DIS\s*\d+', pedido):
                        continue

                    num_pedido_match = re.search(r'DIS\s*(\d+)', pedido)
                    if not num_pedido_match:
                        continue

                    num_pedido = int(num_pedido_match.group(1))
                    if num_pedido < inicio or num_pedido > fin:
                        continue

                    fac_venta = str(record.get('FactVenta', '')).replace(" ", "")
                    fact_flete = str(record.get('Fact Flete', '')).strip()
                    fact_complemento = str(record.get('Fact Complemento', '')).strip()
                    fact_compra = str(record.get('Fact Compra', '')).strip()
                    fecha_fact_venta = record.get('Fecha Fact Venta', '').strip()

                    # Validar que tenga un patrón de fecha válido antes de parsear
                    if not re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', fecha_fact_venta):
                        continue

                    try:
                        fecha_obj = datetime.strptime(fecha_fact_venta, "%d/%m/%Y")
                        mes_nombre = MESES_ES[fecha_obj.month - 1]
                        mes_folder_name = f"{fecha_obj.strftime('%m')} {mes_nombre}"
                        dia_str = fecha_obj.strftime("%d")
                    except Exception:
                        continue

                    pedido_folder = os.path.join(carpeta_contenedora, mes_folder_name, dia_str, pedido)
                    os.makedirs(pedido_folder, exist_ok=True)

                    if fac_venta and fac_venta.upper() != 'NA':
                        crear_txt(pedido_folder, f'Venta_{fac_venta}.txt', fac_venta)
                    if fact_flete and fact_flete.upper() != 'NA':
                        crear_txt(pedido_folder, f'Flete_{fact_flete}.txt', fact_flete)
                    if fact_complemento and fact_complemento.upper() != 'NA':
                        crear_txt(pedido_folder, f'Complemento_{fact_complemento}.txt', fact_complemento)
                    if fact_compra and fact_compra.upper() != 'NA':
                        crear_txt(pedido_folder, f'Compra_{fact_compra}.txt', fact_compra)

                # Comprimir carpeta
                zip_path = os.path.join(temp_dir, 'carpetas.zip')
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, _, files in os.walk(carpeta_contenedora):
                        for file in files:
                            abs_path = os.path.join(root, file)
                            rel_path = os.path.relpath(abs_path, carpeta_contenedora)
                            zipf.write(abs_path, arcname=rel_path)

                return send_file(zip_path, as_attachment=True, download_name="carpetas.zip")

        except ValueError:
            flash("Por favor ingresa números válidos.", "danger")
            return redirect(url_for('index'))
        except Exception as e:
            flash(f"Ocurrió un error: {str(e)}", "danger")
            return redirect(url_for('index'))

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)

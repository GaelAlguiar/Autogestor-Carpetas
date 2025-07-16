from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
import tempfile
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re
import shutil
import traceback

app = Flask(__name__, static_url_path='/static', static_folder='static')
app.secret_key = os.getenv("FLASK_SECRET_KEY", "default-secret")

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
def copiar_archivos_factura(origen_dir, destino_dir, folio):
    folio_limpio = ''.join(filter(str.isdigit, folio))
    print(f"🔍 Buscando archivos con folio '{folio}' (limpio: '{folio_limpio}') en: {origen_dir}")
    
    if not os.path.exists(origen_dir):
        print("⚠️ Ruta no existe:", origen_dir)
        return

    for root, _, files in os.walk(origen_dir):
        for f in files:
            if f.endswith(('.pdf', '.xml')) and folio_limpio in f:
                origen = os.path.join(root, f)
                destino = os.path.join(destino_dir, f)
                shutil.copy(origen, destino)
                print(f"✅ Copiado: {origen} → {destino}")


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            inicio = int(request.form['inicio'])
            fin = int(request.form['fin'])
            year = request.form.get('year', '2025')
            mes = request.form.get('mes', '')  # por compatibilidad futura
            base_path = request.form['ruta_base'].strip()
            file = request.files.get('credenciales')

            print("📂 Ruta base:", base_path)
            if not file or file.filename == '' or not os.path.exists(base_path):
                flash("Debe subir el archivo de credenciales y una ruta válida.", "danger")
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

                archivos_copiados = 0

                for record in records:
                    pedido = record.get('Pedido', '').strip()
                    if not pedido or not re.match(r'DIS\s*\d+', pedido):
                        continue
                    num_pedido = int(re.search(r'DIS\s*(\d+)', pedido).group(1))
                    if num_pedido < inicio or num_pedido > fin:
                        continue

                    fac_venta = str(record.get('FactVenta', '')).replace(" ", "")
                    fact_flete = str(record.get('Fact Flete', '')).strip()
                    fact_complemento = str(record.get('Fact Complemento', '')).strip()
                    fact_compra = str(record.get('Fact Compra', '')).strip()
                    fecha_fact_venta = record.get('Fecha Fact Venta', '').strip()

                    if not re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', fecha_fact_venta):
                        continue

                    fecha_obj = datetime.strptime(fecha_fact_venta, "%d/%m/%Y")
                    mes_nombre = MESES_ES[fecha_obj.month - 1]
                    mes_folder_name = f"{fecha_obj.strftime('%m')} {mes_nombre}"  # esto está bien
                    dia_str = fecha_obj.strftime("%d")

                    pedido_folder = os.path.join(carpeta_contenedora, year, mes_folder_name, dia_str, pedido)
                    os.makedirs(pedido_folder, exist_ok=True)
                    
                    ruta_ventas = os.path.join(base_path, 'VENTAS', year, mes_folder_name)
                    ruta_myf = os.path.join(base_path, 'MOLECULAS Y FLETES', year, mes_folder_name, dia_str)


                    if fac_venta and fac_venta.upper() != 'NA':
                        copiar_archivos_factura(ruta_ventas, pedido_folder, fac_venta)
                        archivos_copiados += 1
                    if fact_flete and fact_flete.upper() != 'NA':
                        copiar_archivos_factura(ruta_myf, pedido_folder, fact_flete)
                        archivos_copiados += 1
                    if fact_complemento and fact_complemento.upper() != 'NA':
                        copiar_archivos_factura(ruta_myf, pedido_folder, fact_complemento)
                        archivos_copiados += 1
                    if fact_compra and fact_compra.upper() != 'NA':
                        copiar_archivos_factura(ruta_myf, pedido_folder, fact_compra)
                        archivos_copiados += 1

                if archivos_copiados == 0:
                    flash("No se encontraron archivos para copiar en el rango especificado.", "warning")
                    return redirect(url_for('index'))

                zip_path = os.path.join(temp_dir, 'carpetas.zip')
                shutil.make_archive(zip_path.replace('.zip', ''), 'zip', carpeta_contenedora)
                print("📦 ZIP generado con éxito:", zip_path)
                return send_file(zip_path, as_attachment=True, download_name="carpetas.zip")

        except Exception as e:
            traceback.print_exc()
            flash(f"Ocurrió un error: {str(e)}", "danger")
            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

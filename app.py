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

def copiar_factura_si_existe(origen_dir, destino_dir, numero_factura):
    """
    Copia archivos .pdf o .xml cuyo nombre contenga EXACTAMENTE el número de factura.
    Retorna la cantidad de archivos copiados.
    """
    copiados = 0
    if not numero_factura:
        return 0

    numero_limpio = ''.join(filter(str.isdigit, numero_factura))

    if not os.path.exists(origen_dir):
        print(f"⚠️ Ruta no encontrada: {origen_dir}")
        return 0

    for root, _, files in os.walk(origen_dir):
        for file in files:
            if not file.lower().endswith(('.pdf', '.xml')):
                continue
            nombre_base = os.path.splitext(file)[0]
            # Coincidencia exacta solo si el número está aislado (delimitado por no dígitos o al inicio/final)
            if re.search(rf'(?<!\d){re.escape(numero_limpio)}(?!\d)', nombre_base):
                shutil.copy(os.path.join(root, file), os.path.join(destino_dir, file))
                print(f"✅ Copiado exacto: {file}")
                copiados += 1

    return copiados

def crear_txt(destino_dir, prefijo, contenido):
    """
    Crea un archivo .txt con el número de factura como contenido si la factura existe.
    """
    if contenido and contenido.upper() != 'NA':
        os.makedirs(destino_dir, exist_ok=True)
        numero = ''.join(filter(str.isdigit, contenido))
        nombre = f"{prefijo.capitalize()}_{numero}.txt"
        ruta = os.path.join(destino_dir, nombre)
        with open(ruta, 'w', encoding='utf-8') as f:
            f.write(contenido)
        print(f"📄 TXT creado: {ruta}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            inicio = int(request.form['inicio'])
            fin = int(request.form['fin'])
            year = request.form.get('year', '2025')
            base_path = request.form['ruta_base'].strip()
            file = request.files.get('credenciales')

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
                total_archivos_copiados = 0

                for record in records:
                    pedido = record.get('Pedido', '').strip()
                    if not pedido or not re.match(r'DIS\s*\d+', pedido):
                        continue

                    num_pedido = int(re.search(r'DIS\s*(\d+)', pedido).group(1))
                    if num_pedido < inicio or num_pedido > fin:
                        continue

                    fac_venta = str(record.get('FactVenta', '')).strip()
                    fact_flete = str(record.get('Fact Flete', '')).strip()
                    fact_complemento = str(record.get('Fact Complemento', '')).strip()
                    fact_compra = str(record.get('Fact Compra', '')).strip()
                    fecha_fact_venta = record.get('Fecha Fact Venta', '').strip()

                    if not re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', fecha_fact_venta):
                        continue

                    fecha_obj = datetime.strptime(fecha_fact_venta, "%d/%m/%Y")
                    mes = fecha_obj.strftime("%m")
                    dia = fecha_obj.strftime("%d")

                    pedido_folder = os.path.join(carpeta_contenedora, year, mes, dia, pedido)
                    os.makedirs(pedido_folder, exist_ok=True)

                    ruta_ventas = os.path.join(base_path, 'VENTAS', year, mes, dia)
                    ruta_myf = os.path.join(base_path, 'MOLECULAS Y FLETES', year, mes)

                    # 1. Buscar archivos reales. Si no hay, crear el .txt
                    if fac_venta and fac_venta.upper() != 'NA':
                        copiados = copiar_factura_si_existe(ruta_ventas, pedido_folder, fac_venta)
                        total_archivos_copiados += copiados
                        if copiados == 0:
                            crear_txt(pedido_folder, "venta", fac_venta)

                    if fact_flete and fact_flete.upper() != 'NA':
                        copiados = copiar_factura_si_existe(ruta_myf, pedido_folder, fact_flete)
                        total_archivos_copiados += copiados
                        if copiados == 0:
                            crear_txt(pedido_folder, "flete", fact_flete)

                    if fact_complemento and fact_complemento.upper() != 'NA':
                        copiados = copiar_factura_si_existe(ruta_myf, pedido_folder, fact_complemento)
                        total_archivos_copiados += copiados
                        if copiados == 0:
                            crear_txt(pedido_folder, "complemento", fact_complemento)

                    if fact_compra and fact_compra.upper() != 'NA':
                        copiados = copiar_factura_si_existe(ruta_myf, pedido_folder, fact_compra)
                        total_archivos_copiados += copiados
                        if copiados == 0:
                            crear_txt(pedido_folder, "compra", fact_compra)

                if total_archivos_copiados == 0:
                    flash("No se encontraron archivos PDF/XML exactos para copiar. Solo se generaron los .txt.", "warning")
                else:
                    flash(f"Se copiaron {total_archivos_copiados} archivos correctamente.", "success")

                zip_base = os.path.join(temp_dir, 'carpetas')
                zip_path = shutil.make_archive(zip_base, 'zip', carpeta_contenedora)
                return send_file(zip_path, as_attachment=True, download_name="carpetas.zip")

        except Exception as e:
            traceback.print_exc()
            flash(f"Ocurrió un error: {str(e)}", "danger")
            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

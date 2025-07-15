from flask import Flask, render_template, request, redirect, url_for, flash
import os
import shutil
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import locale
import re

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

app = Flask(__name__)
app.secret_key = 'MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCasjike1KyWFj2\nObCXPWdq8S6+q40EksilUL4uJ3uSNjrRSTmDYGfiJZfWWrDGCiN7ji6czLlyAIjK\n1Jbgjf5hGd4aPZcPT6qn88hC8eqIb0igHZZtLGDxeahG6IUzIi7djKLboqLdUr/x\nTOkQtSuFdnbClGRrdJ1/LKniy0D4wFsL9zXWIKLFOCxZvc7QlH0u5adhb/cI7VXP\n/nE241TNzkrdCawZV/+cxk5BUas/k1Q9ef6m8dOf0F4+ktPMP4YLwrL4JGbTk7Tg\n2qL2DSaNZljdLDxFMywUdymuUNOYkOA+aI3b7sF90QJJRqOb+hjRW3O38LzZaoSm\nkTZLHM6dAgMBAAECggEABUos16sD4+dTe2/zkuhdnfGLWKfDFbHzPHvrVOayuggy\nsK9hURW916TTcVf+jXcRSYtOGryBZt2Pz+e/FQSl+yoIRztt6+8cdcvHQErHa0zq\n3dbFKwyGcBtp3qrayynTEm5Zr0r7aLgIqjaoDZM0XsbzPqoWWPpO3Gdpk8DBgv/7\nMtfUNPami4FtWf3ClRuRxt2Ux+6Q5exA2nbDmWSrrdLu3pevej3+Y9awMz50D6Hr\nVNukLHBkHsEWh4TEYXZ95f5v+AxRvxt1vnAp8joGsm4AEVe/m8hPmqy1nOd5FZW5\nhxP6gZ5Z3YKtMd9tsxQtIAjKWMj4/A4hoPw2LyQe4QKBgQDZt60kNbrDgAggCQR8\nbdBvfXmdlN144XD4gYMI6zSuYHBSHC/sHyPH42D+oFVYusEVi0UbtAAr+e7k3qlB\nVy/Z2EGj529p1JNul25A8dCnCTfc++EdMLqh0OSBb0dNzbwbOJl58RVOblf7nYXY\nLPWdqlOH7JEwf4HSOoG3lNyZuQKBgQC15bT8oY57TAuCYt/eHCbD9kPA+RWKF9F+\nsgYyBCP1/ghUpvhbWEeQu3FqXhVrTEb/NQIppq9VuZ6dSg9H8zDfqmJVfhxk5qU/\nGgbIyH6mE1ot0YzP2sqhdIsy+KinJGm4BcBGV0ZChslAo6961xl6GazVHDcFb9Qr\nud9h/6s+BQKBgFEOB+DWPAzyypOap9fnTlVjonZfaMDLNbLfDLiyUG+nKcn4AoNm\n6HxPk9nYOU4KYT4zFmyE7BdzOlRs7RiNbBwvXei2jg0ZfjYLHJoDLQoy7WBRSfMD\nJEiAK8Jgemxl7uU3gjQa5DLJ8+mSMLVVr6+eLPEKytcCcYOiEo8VVbfJAoGAGpHc\nNw7ORjpccAZLVyFblEJTsUtxFwPVqSEOAJ5UNmmOA/eDzav+gCixL21gyZSRxlOS\n5kyfzfDYN3eK9eKTIAi+ZmiOczqxpp8BoLCQt2eaQ5kZbX8zHBRvBNoHoKqT+rp0\nVJIJBEy19wgx6MqkwQ4hDdwaOWQVZPG4rJLxC5UCgYEA1idxW1hXuZDtXP6NbmWN\nqUDGuxM8Ew6q7UrtH8y0GJFabJuvVuZWuFXT83SPEs6DERst73QczeHm6jw+qsw6\n74mUX96bDsrFK6AJCsEDUoArfopqbLxJQkKv2Sml8NzfKH/Zw3skNaFAT7yhO2mL\nKuOuRpHSSvs7F1ksplHkzJg'

def leer_datos_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/gaelalguiar/Downloads/credenciales.json', scope)
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
            if inicio > fin:
                flash("El número inicial debe ser menor o igual al número final.", "danger")
                return redirect(url_for('index'))

            records = leer_datos_google_sheets()
            if not records:
                flash("No se pudieron leer datos de Google Sheets.", "danger")
                return redirect(url_for('index'))

            downloads_path = os.path.expanduser('~/Downloads')
            carpeta_contenedora = os.path.join(downloads_path, 'Carpetas')

            if os.path.exists(carpeta_contenedora):
                shutil.rmtree(carpeta_contenedora)

            os.makedirs(carpeta_contenedora)

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
                fecha_fact_venta = record.get('Fecha Fact Venta', '')

                if not fecha_fact_venta:
                    continue

                try:
                    fecha_obj = datetime.strptime(fecha_fact_venta, "%d/%m/%Y")
                    mes_folder_name = fecha_obj.strftime("%m %B").upper()
                    dia_str = fecha_obj.strftime("%d")
                except ValueError:
                    continue

                mes_folder = os.path.join(carpeta_contenedora, mes_folder_name)
                dia_folder = os.path.join(mes_folder, dia_str)
                pedido_folder = os.path.join(dia_folder, pedido)

                os.makedirs(pedido_folder, exist_ok=True)

                if fac_venta and fac_venta.upper() != 'NA':
                    crear_txt(pedido_folder, f'Venta_{fac_venta}.txt', fac_venta)
                if fact_flete and fact_flete.upper() != 'NA':
                    crear_txt(pedido_folder, f'Flete_{fact_flete}.txt', fact_flete)
                if fact_complemento and fact_complemento.upper() != 'NA':
                    crear_txt(pedido_folder, f'Complemento_{fact_complemento}.txt', fact_complemento)
                if fact_compra and fact_compra.upper() != 'NA':
                    crear_txt(pedido_folder, f'Compra_{fact_compra}.txt', fact_compra)

            flash(f'Carpetas y archivos creados desde DIS{inicio} hasta DIS{fin} en {carpeta_contenedora}', "success")
            return redirect(url_for('index'))

        except ValueError:
            flash("Por favor ingresa números válidos.", "danger")
            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
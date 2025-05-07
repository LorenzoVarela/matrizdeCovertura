from flask import Flask, render_template, request
import pandas as pd
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def procesar_ficheros(ruta_requisitos, ruta_pruebas):
    """
    Lee y procesa los ficheros ODS para generar la matriz de cobertura.

    Args:
        ruta_requisitos (str): Ruta al fichero ODS de requisitos.
        ruta_pruebas (str): Ruta al fichero ODS de pruebas.

    Returns:
        tuple: Una tupla conteniendo:
            - lista_ids_requisitos (list): Lista de IDs de requisitos.
            - lista_codigos_prueba (list): Lista de códigos de prueba.
            - matriz_cobertura (dict): Diccionario representando la matriz de cobertura.
    """
    try:
        df_requisitos = pd.read_excel(ruta_requisitos,engine='odf', sheet_name='Requisitos_de_Sistema', header=7)
        df_pruebas = pd.read_excel(ruta_pruebas, engine='odf', sheet_name='Resultados_Pruebas', skiprows=8)

        if 'ID' not in df_requisitos.columns:
            df_requisitos['ID'] = pd.Series(dtype=str)
        if 'Código de Prueba' not in df_pruebas.columns:
            df_pruebas['Código de Prueba'] = pd.Series(dtype=str)

        lista_ids_requisitos = df_requisitos['ID'].astype(str).tolist()
        lista_codigos_prueba = df_pruebas['Código de Prueba'].astype(str).tolist()

        lista_ids_requisitos = [id_req for id_req in lista_ids_requisitos if pd.notna(id_req)]
        lista_codigos_prueba = [codigo for codigo in lista_codigos_prueba if pd.notna(codigo)]

        matriz_cobertura = {}
        for codigo_prueba in lista_codigos_prueba:
            matriz_cobertura[codigo_prueba] = {id_requisito: "" for id_requisito in lista_ids_requisitos}

        for index, row in df_requisitos.iterrows():
            id_requisito = str(row.get('ID', '')).strip()
            cubierto_por_str = str(row.get('Cubierto por', '')).strip()

            if cubierto_por_str and cubierto_por_str != 'nan':
                pruebas_cubriendo = [p.strip() for p in cubierto_por_str.split(',')]
                for prueba in pruebas_cubriendo:
                    if prueba in matriz_cobertura and id_requisito in matriz_cobertura[prueba]:
                        matriz_cobertura[prueba][id_requisito] = "X"

        return lista_ids_requisitos, lista_codigos_prueba, matriz_cobertura

    except Exception as e:
        print(f"Error al procesar los ficheros: {e}")
        return None, None, None

@app.route('/', methods=['GET', 'POST'])
def index():
    matriz_html = ""
    lista_ids_guardada = []
    lista_pruebas_guardada = []
    matriz_guardada = {}

    if request.method == 'POST':
        if 'fichero_requisitos' not in request.files or 'fichero_pruebas' not in request.files:
            return render_template('index.html', error='Por favor, sube ambos ficheros.')

        fichero_requisitos = request.files['fichero_requisitos']
        fichero_pruebas = request.files['fichero_pruebas']

        if fichero_requisitos.filename == '' or fichero_pruebas.filename == '':
            return render_template('index.html', error='Por favor, selecciona ambos ficheros.')

        if fichero_requisitos and fichero_pruebas:
            filename_requisitos = secure_filename(fichero_requisitos.filename)
            filename_pruebas = secure_filename(fichero_pruebas.filename)
            ruta_requisitos = os.path.join(app.config['UPLOAD_FOLDER'], filename_requisitos)
            ruta_pruebas = os.path.join(app.config['UPLOAD_FOLDER'], filename_pruebas)
            fichero_requisitos.save(ruta_requisitos)
            fichero_pruebas.save(ruta_pruebas)

            lista_ids, lista_pruebas, matriz = procesar_ficheros(ruta_requisitos, ruta_pruebas)

            if lista_ids is not None and lista_pruebas is not None and matriz is not None:
                matriz_html = "<table><tr><th>Código de Prueba</th>"
                for id_req in lista_ids:
                    matriz_html += f"<th>{id_req}</th>"
                matriz_html += "</tr>"
                for codigo_prueba, cobertura in matriz.items():
                    matriz_html += f"<tr><td>{codigo_prueba}</td>"
                    for id_req in lista_ids:
                        matriz_html += f"<td>{cobertura.get(id_req, '')}</td>"
                    matriz_html += "</tr>"
                matriz_html += "</table>"
                lista_ids_guardada = lista_ids
                lista_pruebas_guardada = lista_pruebas
                matriz_guardada = matriz
            else:
                return render_template('index.html', error='Error al procesar los ficheros. Asegúrate de que tienen el formato correcto.')

    return render_template('index.html', matriz_html=matriz_html,
                           lista_ids=lista_ids_guardada,
                           lista_pruebas=lista_pruebas_guardada,
                           matriz=matriz_guardada)

@app.route('/exportar_excel')
def exportar_excel():
    lista_ids = request.args.getlist('lista_ids')
    lista_pruebas = request.args.getlist('lista_pruebas')
    matriz_data = {}
    for prueba in request.args.getlist('lista_pruebas'):
        cobertura = {}
        for id_req in request.args.getlist(f'cobertura_{prueba}'):
            partes = id_req.split(':')
            if len(partes) == 2:
                req_id, valor = partes
                cobertura[req_id] = valor
        matriz_data[prueba] = cobertura

    if not matriz_data:
        return "No hay datos para exportar."

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    df = pd.DataFrame.from_dict(matriz_data, orient='index', columns=lista_ids)
    df.index.name = 'Código de Prueba'
    df.to_excel(writer, sheet_name='Matriz de Cobertura')

    writer.close()
    output.seek(0)

    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='matriz_cobertura.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
import os
from flask import Flask, render_template, request, redirect, url_for
from Funcion_final import run_task

app = Flask(__name__)

# Ruta donde se guardarán los archivos de datos
data_folder = 'data'
output_folder = 'output'

# Si los directorios de datos y salida no existen, créalos
#if not os.path.exists(data_folder):
    #os.makedirs(data_folder)

#if not os.path.exists(output_folder):
    #os.makedirs(output_folder)

@app.route('/')
def index():
    return """
    <html>
    <head>
        <title>Cálculo de Pasivos Pensionales</title>
    </head>
    <body>
        <h1><img src="/static/Vela.jpg" height="60px" width="60px"> Cálculo de Pasivos Pensionales</h1>
        <form method="POST" action="/upload" enctype="multipart/form-data">
            <input type="file" name="file1" accept=".csv"><br><br>
            <input type="submit" name="saveBtn" value="Generar Informe">
        </form>
    </body>
    </html>
    """

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file1' not in request.files:
        return "No se seleccionó ningún archivo."

    file = request.files['file1']
    if file.filename == '':
        return "Nombre de archivo vacío."

    if file:
        filename1 = os.path.join(data_folder, file.filename)
        file.save(filename1)

        # Ejecutar la función utilizando el archivo cargado
        run_task(filename1)

        return "Informe generado correctamente."

if __name__ == '__main__':
    app.run(debug=True)
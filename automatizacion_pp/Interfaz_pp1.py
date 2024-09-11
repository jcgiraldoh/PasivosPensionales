import os
from flask import Flask, render_template, request, redirect, url_for
from Funcion_final1 import run_task

app = Flask(__name__)

# Ruta donde se guardarán los archivos de datos
data_folder = 'data'
output_folder = 'salida'

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
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css">
    </head>
    <body>
        <h1><img src="/static/Vela.jpg" height="60px" width="60px"> Cálculo de Pasivos Pensionales</h1>
        <form method="POST" action="/upload" enctype="multipart/form-data">
            <input type="file" name="file1" accept=".csv"><br><br>
            <input type="file" name="file2" accept=".csv"><br><br>
            <input type="submit" name="saveBtn" value="Generar Informe">
        </form>
        
        <!-- Información de contacto -->
        <h2>Información de Contacto</h2>
        <ul class="elementor-icon-list-items elementor-inline-items">
            <li class="elementor-icon-list-item elementor-inline-item">
                <span class="elementor-icon-list-icon">
                    <i aria-hidden="true" class="fas fa-map-marker-alt"></i>
                </span>
                <span class="elementor-icon-list-text">Calle 95 No 13-55 Of 414</span>
            </li>
            <li class="elementor-icon-list-item elementor-inline-item">
                <span class="elementor-icon-list-icon">
                    <i aria-hidden="true" class="fas fa-mobile-alt"></i>
                </span>
                <span class="elementor-icon-list-text">C. (57) 316 3748679 - 315 8406111 / T. (601) 621 4056 </span>
            </li>
            <li class="elementor-icon-list-item elementor-inline-item">
                <span class="elementor-icon-list-icon">
                    <i aria-hidden="true" class="fas fa-envelope"></i>
                </span>
                <span class="elementor-icon-list-text">
                    <a href="mailto:hola@vela.com.co">hola@vela.com.co</a>
                </span>
            </li>
        </ul>
    </body>
    </html>
    """

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file1' not in request.files or 'file2' not in request.files:
        return "No se seleccionaron ambos archivos."

    file1 = request.files['file1']
    
    file2 = request.files['file2']

    if file1.filename == '' or file2.filename == '':
        return "Nombre de archivo vacío."

    if file1 and file2:
        filename1 = os.path.join(data_folder, file1.filename)
        filename2 = os.path.join(data_folder, file2.filename)

        file1.save(filename1)
        file2.save(filename2)

        # Ejecutar la función utilizando los archivos cargados
        run_task(filename1, filename2)

        return "Informes generados correctamente."

if __name__ == '__main__':
    app.run(debug=True)
    
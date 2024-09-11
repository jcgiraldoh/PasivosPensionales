import sys
import os

# Agrega el directorio "source" al path de Python
source_dir = os.path.join(os.path.dirname(__file__), 'source')
sys.path.append(source_dir)

# Importa la función jub_escribir desde jub_funcion.py
from jub_funcion import jub_escribir

# Obtiene el directorio actual
main_path = os.getcwd()

# Define la ruta de destino y la ubicación del formato
ruta_destino = os.path.join(main_path, "output/")
ubicacion_formato = os.path.join(main_path, "plantilla", "jubilados.xlsx")

# Verifica si la carpeta de destino existe, si no, la crea
if not os.path.exists(ruta_destino):
    os.makedirs(ruta_destino)

# Define el nombre del archivo
nombre_archivo = "Cálculo_Pasivos_Pensionales_Jub"

# Función para ejecutar la tarea
def run_task(input_jub_path):
    try:
        print(f"Archivo de entrada: {input_jub_path}")
        print("Escribiendo datos Jub...")

        # Verifica si el archivo de plantilla existe antes de continuar
        if not os.path.exists(ubicacion_formato):
            print(f"Error: El archivo de plantilla no existe en {ubicacion_formato}")
            return

        # Llama a la función jub_escribir con los argumentos necesarios
        jub_escribir(input_jub_path, ruta_destino, nombre_archivo, ubicacion_formato)

        # Muestra la ubicación del archivo actualizado
        ubicacion_formato_actualizada = os.path.join(ruta_destino, f"{nombre_archivo}.xlsx")
        print(f"Ubicación del formato actualizado: {ubicacion_formato_actualizada}")

    except Exception as e:
        print(f"Ocurrió un error durante la ejecución: {e}")

import os
import sys
import pandas as pd
import openpyxl

# Agrega el directorio "source" al path de Python
source_dir = os.path.join(os.path.dirname(__file__), 'source')
sys.path.append(source_dir)

# Importa la función beneficio_pleno_escribir desde beneficio_pleno_funcion.py
from beneficio_pleno_funcion import beneficio_pleno_escribir

# Obtiene el directorio actual
main_path = os.getcwd()

# Define la ruta de destino y la ubicación del formato
ruta_destino = os.path.join(main_path, "output/")
ubicacion_formato = os.path.join(main_path, "plantilla", "jubilados.xlsx")

# Define el nombre del archivo de salida
nombre_archivo = "Cálculo_Pasivos_Pensionales_Beneficio"

# Verificar si el directorio de destino existe, si no, crearlo
if not os.path.exists(ruta_destino):
    os.makedirs(ruta_destino)
    print(f"Directorio de salida creado: {ruta_destino}")

# Verificar si el archivo de plantilla existe
if not os.path.exists(ubicacion_formato):
    print(f"Error: El archivo de plantilla {ubicacion_formato} no existe.")
    sys.exit(1)

# Función para ejecutar la tarea
def run_task(input_beneficio_path):
    # Verificar si el archivo de entrada existe
    if not os.path.exists(input_beneficio_path):
        print(f"Error: El archivo {input_beneficio_path} no existe.")
        return

    try:
        print(f"Procesando el archivo: {input_beneficio_path}")
        
        # Verificar que el archivo de plantilla Excel es válido
        if not verificar_excel(ubicacion_formato):
            print("Error: El archivo de plantilla no es válido.")
            return

        print("Escribiendo datos Beneficio...")

        # Llama a la función beneficio_pleno_escribir con los argumentos necesarios
        try:
            beneficio_pleno_escribir(input_beneficio_path, ruta_destino, nombre_archivo, ubicacion_formato)
        except Exception as e:
            print(f"Error al escribir los datos en el archivo: {e}")
            return
        
        ubicacion_formato_actualizada = os.path.join(ruta_destino, f"{nombre_archivo}.xlsx")
        print(f"Archivo actualizado correctamente: {ubicacion_formato_actualizada}")
    
    except Exception as e:
        print(f"Error durante la ejecución de la tarea: {e}")

# Función para verificar que el archivo Excel es un archivo .xlsx válido
def verificar_excel(archivo_excel):
    try:
        # Intentamos abrir el archivo Excel usando openpyxl para verificar su validez
        openpyxl.load_workbook(archivo_excel)
        print("El archivo Excel está en buen estado.")
        return True
    except Exception as e:
        print(f"Error al cargar el archivo Excel: {e}")
        return False

# Ejemplo de ejecución
if __name__ == "__main__": 
    input_beneficio_path = 'D:/PP/automatizacion_pp/datos_beneficio.csv'  # Ajusta la ruta al archivo de datos de entrada
    run_task(input_beneficio_path)

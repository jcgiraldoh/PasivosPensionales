import sys
import os

# Agrega el directorio "source" al path de Python
source_dir = os.path.join(os.path.dirname(__file__), 'source')
sys.path.append(source_dir)

# Importa la función evasion_escribir desde evasion_funcion.py
from plenas_funcion import plenas_escribir  

# Obtiene el directorio actual
main_path = os.getcwd()

# Define el nombre del archivo
nombre_archivo = "Cálculo_Pasivos_Pensionales_Plenas"

# Define la ruta de destino y la ubicación del formato
ruta_destino = os.path.join(main_path, "output/")
ubicacion_formato = os.path.join(main_path, "plantilla", "jubilados.xlsx")

# Verificar si la carpeta 'output' existe, si no, crearla
if not os.path.exists(ruta_destino):
    os.makedirs(ruta_destino)

# Función para ejecutar la tarea
def run_task(input_plenas_path):
    try:
        print(f"Ruta del archivo CSV: {input_plenas_path}")
        print(f"Ruta del archivo de plantilla: {ubicacion_formato}")
        
        # Verificar si el archivo de entrada y la plantilla existen
        if not os.path.exists(input_plenas_path):
            print(f"Error: El archivo {input_plenas_path} no existe.")
            return
        
        if not os.path.exists(ubicacion_formato):
            print(f"Error: El archivo de plantilla {ubicacion_formato} no existe.")
            return
        
        print("Escribiendo datos Plenas")
        
        # Llama a la función plenas_escribir con los argumentos necesarios
        plenas_escribir(input_plenas_path, ruta_destino, nombre_archivo, ubicacion_formato)
        
        ubicacion_formato_actualizada = os.path.join(ruta_destino, f"{nombre_archivo}.xlsx")
        print(f"Ubicación del formato actualizado: {ubicacion_formato_actualizada}")
    
    except Exception as e:
        print(f"Ha ocurrido un error: {str(e)}")

# if __name__ == '__main__':
#     run_task("C:/Proyectos/PP/automatizacion_pp/datos_plenas.csv")

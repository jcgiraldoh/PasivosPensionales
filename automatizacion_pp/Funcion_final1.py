import os
import sys
import pandas as pd

# Agrega el directorio "source" al path de Python
source_dir = os.path.join(os.path.dirname(__file__), 'source')
sys.path.append(source_dir)

from beneficio_pleno_funcion import beneficio_pleno_escribir
from compartidas_funcion import compartidas_escribir

# Establece las rutas y nombres de archivo
main_path = os.getcwd()
nombre_archivo = "pasivos.xlsx"
ruta_destino = os.path.join(main_path, "salida/")
ubicacion_formato = os.path.join(main_path, "plantilla", "jubilados.xlsx")


def run_task(input_beneficio_path, input_compartida_path):
    print("Escribiendo datos Beneficio")
    
    # Llama a la función beneficio_pleno_escribir con los argumentos necesarios
    beneficio_pleno_escribir(input_beneficio_path, ruta_destino, nombre_archivo, ubicacion_formato,nombre_hoja='BENEF PLENO VIT.(4)')
    
    # Imprime información de ubicación
    print(f'{ruta_destino}/{nombre_archivo}')
    
    print("Escribiendo datos Compartidas")
    
    # Llama a la función compartidas_escribir con los argumentos necesarios
    compartidas_escribir(input_compartida_path, ruta_destino, nombre_archivo, ubicacion_formato,nombre_hoja='CALCULO COMPARTIDAS(2)')
    
    # Actualiza la ubicación del formato actualizado después de la segunda llamada
    ubicacion_formato_actualizada = os.path.join(ruta_destino, nombre_archivo)
    print(f"Ubicación del formato actualizado: {ubicacion_formato_actualizada}")




if __name__ == '__main__':
    run_task("C:/Proyectos/PP/automatizacion_pp/datos_beneficio.csv",
              "C:/Proyectos/PP/automatizacion_pp/datos_compartidas.csv")
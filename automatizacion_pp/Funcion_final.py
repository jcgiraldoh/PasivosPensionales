import sys
import os


# Agrega el directorio "source" al path de Python
source_dir = os.path.join(os.path.dirname(__file__), 'source')
sys.path.append(source_dir)

# Importa la función evasion_escribir desde evasion_funcion.py
from beneficio_pleno_funcion import beneficio_pleno_escribir
from compartidas_funcion import compartidas_escribir   
 
# Obtiene el directorio actual
main_path = os.getcwd()

# Define la ruta de destino y la ubicación del formato
ruta_destino = os.path.join(main_path,"output")
ubicacion_formato = os.path.join(main_path, "plantilla", "jubilados.xlsx")

# Define el nombre del archivo
nombre_archivo = "Cálculo_Pasivos_Pensionales"

# Función para ejecutar la tarea
def run_task(input_beneficio_path,input_compartidas_path):
    print(input_beneficio_path)
    print("Escribiendo datos Beneficio")
    
    # Llama a la función beneficio_pleno_escribir con los argumentos necesarios
    beneficio_pleno_escribir(input_beneficio_path, ruta_destino, nombre_archivo, ubicacion_formato)
    
    print(input_compartidas_path)
    print("Escribiendo datos Compartidas")
    
    # Llama a la función beneficio_pleno_escribir con los argumentos necesarios
    compartidas_escribir(input_compartidas_path, ruta_destino, nombre_archivo, ubicacion_formato)
    
    ubicacion_formato_actualizada = os.path.join(ruta_destino, f"{nombre_archivo}.xlsx")
    print(f"Ubicación del formato actualizado: {ubicacion_formato_actualizada}")



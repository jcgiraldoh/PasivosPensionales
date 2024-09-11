import sys
import os

# Agrega el directorio "source" al path de Python
source_dir = os.path.join(os.getcwd(), 'source')  # Cambié __file__ por os.getcwd()
sys.path.append(source_dir)

# Importa la función evasion_escribir desde evasion_funcion.py
try:
    from compartidas_funcion import compartidas_escribir
except ImportError as e:
    print(f"Error al importar la función: {e}")
    sys.exit(1)

# Obtiene el directorio actual
main_path = os.getcwd()

# Define el nombre del archivo (sin caracteres especiales para evitar problemas en Windows)
nombre_archivo = "Calculo_Pasivos_Pensionales_Compartidas"

# Define la ruta de destino y la ubicación del formato
ruta_destino = os.path.join(main_path, "output/")
ubicacion_formato = os.path.join(main_path, "plantilla", "jubilados.xlsx")

# Verifica si la carpeta de destino existe, si no, la crea
if not os.path.exists(ruta_destino):
    os.makedirs(ruta_destino)
    print(f"Carpeta de destino creada: {ruta_destino}")

# Verifica si el archivo de plantilla existe
if not os.path.exists(ubicacion_formato):
    print(f"Error: La plantilla no existe en la ubicación: {ubicacion_formato}")
    sys.exit(1)  # Salir del script si no se encuentra la plantilla

# Función para ejecutar la tarea
def run_task(input_compartidas_path):
    try:
        # Verifica si el archivo de datos compartidas existe antes de continuar
        if not os.path.exists(input_compartidas_path):
            print(f"Error: El archivo {input_compartidas_path} no existe.")
            sys.exit(1)  # Salir si el archivo no existe
        
        print(f"Ruta del archivo de entrada: {input_compartidas_path}")
        print("Escribiendo datos Compartidas")
        
        # Llama a la función compartidas_escribir con los argumentos necesarios
        compartidas_escribir(input_compartidas_path, ruta_destino, nombre_archivo, ubicacion_formato)
        
        ubicacion_formato_actualizada = os.path.join(ruta_destino, f"{nombre_archivo}.xlsx")
        print(f"Ubicación del formato actualizado: {ubicacion_formato_actualizada}")
    
    except Exception as e:
        print(f"Error al ejecutar la tarea: {e}")
        sys.exit(1)  # Salir del script si hay un error grave

# Ejecución del script principal
if __name__ == "__main__":
    input_compartidas_path = 'data/datos_compartidas.csv'  # Ruta al archivo CSV de entrada
    
    # Verifica si el archivo de entrada existe antes de ejecutar
    if not os.path.exists(input_compartidas_path):
        print(f"Error: El archivo de datos compartidas no existe: {input_compartidas_path}")
        sys.exit(1)
    
    run_task(input_compartidas_path)

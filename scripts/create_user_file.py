import shutil
import os

def create_user_submission_file(output_filename="user_evaluacion.xlsx", source_path="../data/base_datos_original.xlsx", output_dir="../user_submissions/"):
    """
    Crea una copia del archivo base para que el usuario la modifique.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    destination_path = os.path.join(output_dir, output_filename)

    try:
        shutil.copy(source_path, destination_path)
        print(f"Archivo para el usuario '{output_filename}' creado en '{output_dir}'.")
        print(f"Por favor, indica al usuario que modifique este archivo según las instrucciones.")
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo base en '{source_path}'.")
    except Exception as e:
        print(f"Ocurrió un error al crear el archivo para el usuario: {e}")

if __name__ == "__main__":
    # Puedes cambiar el nombre del archivo de salida para cada usuario
    user_name = input("Introduce el nombre del usuario para el archivo (ej: Juan_Perez): ")
    if user_name:
        create_user_submission_file(f"{user_name}_evaluacion.xlsx")
    else:
        create_user_submission_file() # Usa el nombre por defecto
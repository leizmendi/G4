"""
    Convierte el archivo pdf2b64.pdf que debe estar en la carpeta especificada
    en base 64 y se guarda en pdf2b64.txt en la carpeta especificada.
    Uso: python nombre_del_script.py [carpeta_de_entrada_salida]
"""

import base64
import os
import sys
from pathlib import Path

def convert_pdf_to_base64(file_path: str) -> str:
    with open(file_path, 'rb') as pdf_file:
        pdf_content = pdf_file.read()
        return base64.b64encode(pdf_content).decode('utf-8')

def save_base64_to_txt(content: str, output_path: str):
    #with open(output_path, 'w', encoding='utf-8') as txt_file:
    #    txt_file.write(content)

# Ruta del archivo de texto donde deseas guardar el string
    archivo_texto = Path(output_path)
    archivo_texto.write_text(content)


def main():
    try:
        if len(sys.argv) < 2:
            print("Por favor, proporciona una carpeta de entrada y salida.")
            print("Uso: python nombre_del_script.py [carpeta_de_entrada_salida]")
            sys.exit(1)

        directory = sys.argv[1]

        # Definir rutas
        pdf_file_path = Path(directory) / 'pdf2b64.pdf'
        output_file_path = Path(directory) / 'pdf2b64.txt'

        if not pdf_file_path.exists():
            print(f"No se encontró el archivo 'pdf2b64.pdf' en {directory}.")
            sys.exit(1)

        # Convertir y guardar
        pdf_base64 = convert_pdf_to_base64(pdf_file_path)
        save_base64_to_txt(pdf_base64, output_file_path)

        #print(f"Archivo 'pdf2b64.txt' guardado en {directory}")
        print(pdf_base64)
        # tu código aquí
        sys.exit(0)  # todo salió bien
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)  # ocurrió un error    

if __name__ == "__main__":
    main()

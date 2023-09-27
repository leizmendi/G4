import base64
import sys

def convert_base64_to_pdf(sBase64, output_pdf):
    #with open(base64_file, 'r', encoding='utf-8') as file:
    #    data_base64 = file.read()
    #    data_bytes = base64.b64decode(data_base64)

    with open(output_pdf, 'wb', encoding='utf-8') as file:
    #    file.write(data_bytes)
        file.write(sBase64)

if __name__ == '__main__':
    try:
        if len(sys.argv) != 3:
            print("Uso: script.py <ruta_al_archivo_base64> <ruta_al_output_pdf>")
            sys.exit(1)

        sBase64 = sys.argv[1]
        output_file = sys.argv[2]

        convert_base64_to_pdf(sBase64, output_file)
        print(f"PDF guardado como '{output_file}'")
        sys.exit(0)
    except Exception as e :
        print(f"Error: {e}")
        sys.exit(1)  # ocurrió un error    


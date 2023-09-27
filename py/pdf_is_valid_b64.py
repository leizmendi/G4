import base64

def is_valid_pdf(b64_string):
    try:
        # 1. Decodificar el string base64 a bytes
        decoded_bytes = base64.b64decode(b64_string)

        # 2. Verificar el encabezado y el final del PDF decodificado
        # Los PDFs suelen comenzar con %PDF- y terminar con %%EOF
        if not decoded_bytes.startswith(b'%PDF-') or not decoded_bytes.endswith(b'%%EOF'):
            return False
        
        # 3. Opcional: Usar una biblioteca para asegurarte de que el PDF es completamente válido
        # Puedes usar PyPDF2 o cualquier otra biblioteca de manejo de PDFs que prefieras
        # from PyPDF2 import PdfFileReader
        # from io import BytesIO
        #
        # pdf = PdfFileReader(BytesIO(decoded_bytes))
        # if pdf.getNumPages() < 1:
        #     return False

        return True
    except Exception as e:
        print(e)
        return False

def pdf_is_valid_b64(filename):
    # Definimos la ruta al archivo
    #filename = "ruta_del_archivo.txt"

    # Abrimos el archivo en modo lectura ('r')
    with open(filename, 'r') as file:
        # Leemos el contenido del archivo
        content = file.read()

    # Test
    b64_pdf_string = content  # Aquí va tu string en base64
    print(is_valid_pdf(b64_pdf_string))


pdf_is_valid_b64('pdf2b64.txt')
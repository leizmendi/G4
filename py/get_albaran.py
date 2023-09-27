import requests
import argparse

def get_albaran(api_key, url, file_pdf):
    headers = {'Authorization': f'Bearer {api_key}'}
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.text.encode('utf-8')  # Convierte el texto a binario usando codificación UTF-8
        # Guardar los datos en un archivo binario
        with open(file_pdf, 'wb') as file:
            file.write(data)
        msg = 'OK'
    else:
        msg=response.text
    return response.status_code, msg
    



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Hacer una solicitud web.')
    parser.add_argument('--api-key', required=True, help='API key para la solicitud.')
    parser.add_argument('--url', required=True, help='URL a la cual hacer la solicitud.')
    parser.add_argument('--file_pdf', required=True, help='Ruta y nombre del archivo a recuperar.') 
    args = parser.parse_args()

    codigo, texto = get_albaran(args.api_key, args.url, args.file_pdf)
    print(f'Código de Respuesta: {codigo}')
    print(f'Body de la Respuesta:\n{texto}')

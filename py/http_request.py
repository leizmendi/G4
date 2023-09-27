import requests
import argparse

def hacer_solicitud(api_key, url, metodo):
    headers = {'Authorization': f'Bearer {api_key}'}
    
    if metodo.lower() == "get":
        response = requests.get(url, headers=headers)
    elif metodo.lower() == "post":
        response = requests.post(url, headers=headers)
    # Puedes agregar otros métodos (PUT, DELETE, etc.) de ser necesario

    return response.status_code, response.text

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Hacer una solicitud web.')
    parser.add_argument('--api-key', required=True, help='API key para la solicitud.')
    parser.add_argument('--url', required=True, help='URL a la cual hacer la solicitud.')
    parser.add_argument('--metodo', required=True, choices=['get', 'post'], help='Método HTTP a utilizar.')  # Puedes agregar otros métodos a "choices" si lo deseas
    args = parser.parse_args()

    codigo, texto = hacer_solicitud(args.api_key, args.url, args.metodo)
    print(f'Código de Respuesta: {codigo}')
    print(f'Body de la Respuesta:\n{texto}')

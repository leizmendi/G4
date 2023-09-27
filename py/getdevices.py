import requests
import json

# Define la URL y la APIKey
url = "https://restapi.firmar.online/SignFromApp/v40/devices"
APIKey = "de637e97-0dae-4d2b-83db-aa11b7ecf016"  # Reemplaza TU_APIKEY_AQUI por tu clave real

# Define los encabezados para el esquema de autenticación "Bearer"
headers = {
    "Authorization": f"Bearer {APIKey}"
}

# Realiza la solicitud GET
response = requests.get(url, headers=headers)

# Verifica si la solicitud fue exitosa
if response.status_code == 200:
    # Imprime la respuesta como texto
    print("Respuesta en formato texto:")
    print(response.text)
    
    # Intenta analizar y mostrar la respuesta como un objeto JSON
    try:
        data = response.json()
        print("\nRespuesta en formato JSON:")
        print(json.dumps(data, indent=4))
    except json.JSONDecodeError:
        print("La respuesta no es un JSON válido")
else:
    print(f"Error {response.status_code}: {response.text}")


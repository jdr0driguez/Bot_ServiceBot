

from config import USERNAME, PASSWORDWISER, UNIT_ID, BASE_URLWISER, CSV_FILENAME, CSV_HEADERS_JSON
import requests

AUTH_URL = f"{BASE_URLWISER}/api/auth/security/signin"
UPLOAD_URL = f"{BASE_URLWISER}/cargues-propia/cargues/compromisos-genericos"

def obtener_token():
    auth_payload = {
        "username": USERNAME,
        "password": PASSWORDWISER,
        "unitId": UNIT_ID
    }

    try:
        response = requests.post(AUTH_URL, json=auth_payload, verify=False)
        response.raise_for_status()
        token = response.json().get("token")
        token = response.json().get("token")
        if not token:
            print(" No se encontró el token en la respuesta.")
            return None
        print(" Token obtenido correctamente.")
        return token
    except Exception as e:
        print(f" Error durante autenticación: {e}")
        return None

def enviar_csv(token):
    headers = {"Authorization": f"Bearer {token}"}
    form_data = {
        "unidad_id": UNIT_ID,
        "headers": CSV_HEADERS_JSON,
        "delimitador": ","
    }

    try:
        with open(CSV_FILENAME, "rb") as f:
            files = {"datos": f}
            response = requests.post(UPLOAD_URL, headers=headers, data=form_data, files=files, verify=False)
            response.raise_for_status()
            print(" Archivo enviado exitosamente.")
            print(" Respuesta del servidor:")
            print(response.text)
    except FileNotFoundError:
        print(f" El archivo '{CSV_FILENAME}' no fue encontrado.")
    except Exception as e:
        print(f" Error al enviar archivo: {e}")

def enviar_archivo_api():
    print("\n Autenticando y enviando archivo por API...")
    token = obtener_token()
    if token:
        enviar_csv(token)

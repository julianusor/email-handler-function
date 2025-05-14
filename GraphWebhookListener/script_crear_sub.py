# script_para_crear_suscripcion.py
import os
import sys

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)


# cargar env variables de .env
from dotenv import load_dotenv
load_dotenv()
# Configura estas variables de entorno antes de ejecutar el script
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

FUNCTION_APP_NAME ="new-email-handler-function" # Ej: "myemailhandlerfunction"
TARGET_USER_ID = "julianusu@outlook.com" # Ej: "usuario@tudominio.com" o el ID de usuario

from GraphWebhookListener import create_graph_subscription, get_graph_token # Asumiendo que __init__.py está en GraphWebhookListener

if __name__ == "__main__":
    print("Intentando obtener token de Graph...")
    token = get_graph_token()
    if token:
        print("Token obtenido.")
        print("Creando suscripción a Graph...")
        subscription_details = create_graph_subscription()
        if subscription_details:
            print(f"Suscripción creada con ID: {subscription_details.get('id')}")
            print(f"Expira en: {subscription_details.get('expirationDateTime')}")
        else:
            print("Fallo al crear la suscripción.")
    else:
        print("Fallo al obtener el token de Graph.")

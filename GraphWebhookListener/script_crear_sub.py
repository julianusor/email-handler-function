# script_para_crear_suscripcion.py
import os
import sys
import logging # Added for better output

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# cargar env variables de .env
from dotenv import load_dotenv
load_dotenv()

# Configure logging for the script
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuration for the subscription to be created by this script
FUNCTION_APP_NAME_FOR_SUBSCRIPTION = "new-email-handler-function" # Ej: "myemailhandlerfunction"
TARGET_USER_ID_FOR_SUBSCRIPTION = "julianusu@outlook.com" # Ej: "usuario@tudominio.com" or the Object ID of the user

# The following environment variables should be set in your .env file for get_graph_token to work:
# TENANT_ID, CLIENT_ID, CLIENT_SECRET
# OPENAI_API_KEY (not directly used by this script but good practice if other parts of the module need it)

from GraphWebhookListener import create_graph_subscription, get_graph_token

if __name__ == "__main__":
    logging.info("Intentando obtener token de Graph...")
    token = get_graph_token() # Relies on TENANT_ID, CLIENT_ID, CLIENT_SECRET from environment
    if token:
        logging.info("Token obtenido.")
        
        notification_url = f"https://{FUNCTION_APP_NAME_FOR_SUBSCRIPTION}.azurewebsites.net/api/GraphWebhookListener"
        
        logging.info(f"Creando suscripción a Graph para el usuario: {TARGET_USER_ID_FOR_SUBSCRIPTION}")
        logging.info(f"URL de notificación: {notification_url}")
        
        # Call create_graph_subscription with specific parameters
        subscription_details = create_graph_subscription(
            subscription_target_user_id=TARGET_USER_ID_FOR_SUBSCRIPTION,
            subscription_notification_url=notification_url
        )
        
        if subscription_details:
            logging.info(f"Suscripción creada con ID: {subscription_details.get('id')}")
            logging.info(f"Expira en: {subscription_details.get('expirationDateTime')}")
        else:
            logging.error("Fallo al crear la suscripción.")
    else:
        logging.error("Fallo al obtener el token de Graph.")

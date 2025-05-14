import logging, json, os
import azure.functions as func
from msal import ConfidentialClientApplication
import requests
from openai import OpenAI

# Variables de entorno (app settings)
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

FUNCTION_APP_NAME ="new-email-handler-function" # Ej: "myemailhandlerfunction"
TARGET_USER_ID = "julianusu@outlook.com" # Ej: "usuario@tudominio.com" o el ID de usuario

def get_graph_token():
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result["access_token"]

# Placeholder for OCR function to avoid NameError during notification processing
def ocr_from_bytes(content_bytes):
    logging.warning("OCR functionality is not implemented. Returning placeholder text.")
    # In a real scenario, this would involve sending bytes to an OCR service.
    return "OCR placeholder text for attachment"

def create_graph_subscription():
    """
    Crea una suscripción a notificaciones de Microsoft Graph para nuevos correos.
    Esta función normalmente se ejecutaría una vez para configurar la suscripción,
    o periódicamente para renovarla.
    """
    if not FUNCTION_APP_NAME:
        logging.error("La variable de entorno FUNCTION_APP_NAME no está configurada.")
        return None
    if not TARGET_USER_ID:
        logging.error("La variable de entorno TARGET_USER_ID no está configurada.")
        return None

    token = get_graph_token()
    if not token:
        logging.error("No se pudo obtener el token de Graph.")
        return None

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # Asegúrate de que la URL de notificación sea la correcta para tu Azure Function
    # El nombre del listener es 'GraphWebhookListener' según tu estructura de carpetas.
    notification_url = f"https://{FUNCTION_APP_NAME}.azurewebsites.net/api/GraphWebhookListener"

    # La fecha de expiración debe estar en formato ISO 8601 y en el futuro.
    # Graph permite un máximo de 3 días para algunas suscripciones si no se renuevan.
    # Para este ejemplo, usamos la fecha que proporcionaste.
    # Considera hacerla dinámica, ej: datetime.utcnow() + timedelta(days=2, hours=23)
    expiration_datetime = "2025-06-20T11:00:00.000Z" # Ajusta según sea necesario

    subscription_payload = {
        "changeType": "created",
        "notificationUrl": notification_url,
        "resource": f"/users/{TARGET_USER_ID}/mailFolders('Inbox')/messages",
        "expirationDateTime": expiration_datetime,
        "clientState": "secret-webhook-state-string" # Puedes cambiar esto
    }

    try:
        response = requests.post(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
            json=subscription_payload
        )
        response.raise_for_status()  # Lanza una excepción para códigos de error HTTP (4xx o 5xx)
        subscription_details = response.json()
        logging.info(f"Suscripción creada exitosamente: {subscription_details.get('id')}")
        return subscription_details
    except requests.exceptions.RequestException as e:
        logging.error(f"Error al crear la suscripción: {e}")
        if hasattr(e, 'response') and e.response is not None:
            logging.error(f"Detalles del error: {e.response.text}")
        return None
    except json.JSONDecodeError:
        logging.error(f"Error al decodificar la respuesta JSON de la creación de suscripción. Respuesta: {response.text}")
        return None
    
def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Webhook received")

    # --- VALIDACIÓN DE SUSCRIPCIÓN (GET) ---
    # Microsoft Graph typically uses POST for validation, but GET can be supported.
    if req.method == "GET" and "validationToken" in req.params:
        validation_token = req.params["validationToken"]
        logging.info(f"GET validation request received. Token: {validation_token}")
        return func.HttpResponse(
            validation_token,
            status_code=200,
            mimetype="text/plain"
        )

    # --- Handle POST requests (Graph validation or notifications) ---
    if req.method == "POST":
        try:
            body = req.get_json()
        except ValueError:
            logging.error("Failed to parse JSON body or request is not JSON.")
            return func.HttpResponse("Request body must be valid JSON.", status_code=400)

        # --- VALIDACIÓN DE SUSCRIPCIÓN (POST) ---
        # This is the primary validation mechanism for Graph API subscriptions.
        if isinstance(body, dict) and "validationToken" in body:
            validation_token = body["validationToken"]
            logging.info(f"POST validation token received: {validation_token}")
            return func.HttpResponse(
                validation_token,
                status_code=200,
                mimetype="text/plain"  # Graph expects the token back as plain text.
            )

        # --- PROCESAR NOTIFICACIONES ---
        # If it's a POST request and not a validation request, then it's a notification.
        logging.info("Processing notification(s).")
        for notification in body.get("value", []):
            message_id = notification.get("resourceData", {}).get("id")
            sender_email_info = notification.get("resourceData", {}).get("from", {}).get("emailAddress", {})
            user_id = sender_email_info.get("address")

            if not message_id or not user_id:
                logging.warning(f"Could not extract message_id or user_id from notification: {notification}")
                continue

            logging.info(f"Processing message ID: {message_id} for user (sender): {user_id}")

            # 3) Traer el correo completo
            token = get_graph_token()
            headers = {"Authorization": f"Bearer {token}"}
            mail = requests.get(
                f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}",
                headers=headers
            ).json()

            # 4) Descargar adjuntos y OCR (si corresponde)
            adj_texts = []
            if mail.get("hasAttachments"):
                attachments_url = mail.get("@odata.id", "") + "/attachments"
                if not mail.get("@odata.id"):
                    logging.error(f"Missing '@odata.id' in mail object for message {message_id}")
                else:
                    attachments_response = requests.get(attachments_url, headers=headers)
                    if attachments_response.status_code == 200:
                        attachments = attachments_response.json().get("value", [])
                        for att in attachments:
                            if "contentBytes" in att and att["contentBytes"] is not None:
                                text = ocr_from_bytes(att["contentBytes"])
                                adj_texts.append(text)
                            else:
                                logging.warning(f"Attachment '{att.get('name', 'N/A')}' for message {message_id} has no contentBytes.")
                    else:
                        logging.error(f"Failed to get attachments for message {message_id}. Status: {attachments_response.status_code}, Response: {attachments_response.text}")

            # 5) Llamar a OpenAI para estructurar
            openai = OpenAI(api_key=OPENAI_API_KEY)
            prompt = f"""
                Eres un parser de emails. Devuelve un JSON con:
                nombre, cedula, texto_original, adjuntos (lista de textos OCR).
                Email completo:
                \"\"\"{mail['body']['content']}\"\"\" 
                Adjuntos OCR:
                \"\"\"{json.dumps(adj_texts)}\"\"\" 
                """
            completion = openai.chat.completions.create(
                model="gpt-4o-mini", temperature=0,
                messages=[{"role":"user","content":prompt}]
            )
            data = json.loads(completion.choices[0].message.content)

            # 6) Insertar fila en Excel (OneDrive/Graph API)
            excel_row = [
                data["nombre"], data["cedula"],
                data["texto_original"], ", ".join(data["adjuntos"])
            ]
            body_insert = {"values": [excel_row]}
            requests.post(
                "https://graph.microsoft.com/v1.0/me/drive/root:/datos/emails.xlsx:/"
                "workbook/tables/Table1/rows/add",
                headers=headers, json=body_insert
            )

        return func.HttpResponse(status_code=202)

    # If not GET (with validationToken) or POST, method is not allowed.
    logging.warning(f"Unhandled request method: {req.method}")
    return func.HttpResponse("Method not allowed.", status_code=405)

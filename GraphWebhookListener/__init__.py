import logging, json, os, base64
import azure.functions as func
from msal import ConfidentialClientApplication
import requests
from openai import OpenAI

# Variables de entorno (app settings)
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

FUNCTION_APP_NAME = "new-email-handler-function"  # Ej: "myemailhandlerfunction"
TARGET_USER_ID = "julianusu@outlook.com"  # Ej: "usuario@tudominio.com" o el ID de usuario

def get_graph_token():
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result["access_token"]

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

    notification_url = f"https://{FUNCTION_APP_NAME}.azurewebsites.net/api/GraphWebhookListener"
    expiration_datetime = "2025-06-20T11:00:00.000Z"  # Ajusta según sea necesario

    subscription_payload = {
        "changeType": "created",
        "notificationUrl": notification_url,
        "resource": f"/users/{TARGET_USER_ID}/mailFolders('Inbox')/messages",
        "expirationDateTime": expiration_datetime,
        "clientState": "secret-webhook-state-string"
    }

    try:
        response = requests.post(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
            json=subscription_payload
        )
        response.raise_for_status()
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
    logging.info(f"Webhook received a {req.method} request.")

    validation_token_from_param = req.params.get("validationToken")
    if validation_token_from_param:
        logging.info(f"Validation token found in query parameter: {validation_token_from_param}")
        return func.HttpResponse(
            validation_token_from_param,
            status_code=200,
            mimetype="text/plain"
        )

    if req.method == "POST":
        req_body_bytes = None
        try:
            req_body_bytes = req.get_body()
        except Exception as e:
            logging.error(f"Could not read request body: {e}")
            return func.HttpResponse("Error reading request body.", status_code=400)

        body = None
        if req_body_bytes:
            try:
                body = json.loads(req_body_bytes.decode('utf-8'))
            except json.JSONDecodeError:
                logging.warning("POST request body is not valid JSON. Raw body (first 200 chars): %s", req_body_bytes[:200])
                return func.HttpResponse("Request body must be valid JSON for notifications or JSON-based validation.", status_code=400)
        else:
            logging.info("POST request with empty body.")
            return func.HttpResponse("POST request with empty body and no validationToken in query param.", status_code=400)

        if isinstance(body, dict):
            if "validationToken" in body:
                validation_token_from_body = body["validationToken"]
                logging.info(f"POST validation token received from simple JSON body: {validation_token_from_body}")
                return func.HttpResponse(
                    validation_token_from_body,
                    status_code=200,
                    mimetype="text/plain"
                )
            elif "value" in body and isinstance(body["value"], list):
                for item in body["value"]:
                    if isinstance(item, dict) and "validationTokens" in item and \
                       isinstance(item["validationTokens"], list) and len(item["validationTokens"]) > 0:
                        validation_token_from_complex_body = item["validationTokens"][0]
                        logging.info(f"POST validation token received from complex JSON body: {validation_token_from_complex_body}")
                        return func.HttpResponse(
                            validation_token_from_complex_body,
                            status_code=200,
                            mimetype="text/plain"
                        )

        if isinstance(body, dict) and "value" in body and isinstance(body["value"], list):
            logging.info("Processing notification(s).")
            for notification in body["value"]:
                resource_data = notification.get("resourceData", {})
                message_id = resource_data.get("id")

                if not message_id:
                    logging.warning(f"Could not extract message_id from notification: {notification}")
                    continue

                logging.info(f"Processing message ID: {message_id} for target user: {TARGET_USER_ID}")

                token = get_graph_token()
                if not token:
                    logging.error("Failed to get Graph token for processing notification.")
                    continue
                
                headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
                
                mail_url = f"https://graph.microsoft.com/v1.0/users/{TARGET_USER_ID}/messages/{message_id}"
                mail_response = requests.get(mail_url, headers=headers)
                
                try:
                    mail_response.raise_for_status()
                    mail = mail_response.json()
                except requests.exceptions.HTTPError as e:
                    logging.error(f"Error fetching email {message_id} for user {TARGET_USER_ID}: {e}. Response: {mail_response.text}")
                    continue
                except json.JSONDecodeError:
                    logging.error(f"Error decoding JSON for email {message_id}. Response: {mail_response.text}")
                    continue

                adj_texts = []
                if mail.get("hasAttachments"):
                    attachments_url = f"https://graph.microsoft.com/v1.0/users/{TARGET_USER_ID}/messages/{message_id}/attachments"
                    attachments_response = requests.get(attachments_url, headers=headers)
                    
                    if attachments_response.status_code == 200:
                        attachments = attachments_response.json().get("value", [])
                        for att in attachments:
                            if att.get("contentBytes"): 
                                try:
                                    decoded_bytes = base64.b64decode(att["contentBytes"])
                                    text = ocr_from_bytes(decoded_bytes)
                                    adj_texts.append(text)
                                except base64.binascii.Error as b64e:
                                    logging.error(f"Base64 decoding failed for attachment '{att.get('name', 'N/A')}': {b64e}")
                                except Exception as e:
                                    logging.error(f"OCR processing failed for attachment '{att.get('name', 'N/A')}': {e}")
                            else:
                                logging.warning(f"Attachment '{att.get('name', 'N/A')}' for message {message_id} has no contentBytes or it's null.")
                    else:
                        logging.error(f"Failed to get attachments for message {message_id}. Status: {attachments_response.status_code}, Response: {attachments_response.text}")

                if not OPENAI_API_KEY:
                    logging.error("OPENAI_API_KEY environment variable is not set. Skipping OpenAI processing.")
                    continue

                openai = OpenAI(api_key=OPENAI_API_KEY)
                email_body_content = mail.get("body", {}).get("content", "")
                
                prompt = f"""
                    Eres un parser de emails. Tu tarea es extraer información específica del siguiente correo electrónico y sus adjuntos.
                    Debes devolver la información en formato JSON. El JSON debe contener estrictamente las siguientes claves:
                    - "nombre": El nombre de la persona mencionada en el correo. Si no se encuentra, usa null o una cadena vacía.
                    - "cedula": El número de cédula o identificación. Si no se encuentra, usa null o una cadena vacía.
                    - "texto_original": El cuerpo principal del correo electrónico tal como se recibió.
                    - "adjuntos": Una lista de strings, donde cada string es el texto extraído (OCR) de un adjunto. Si no hay adjuntos o no se pudo extraer texto, usa una lista vacía [].

                    Email completo:
                    \"\"\"{email_body_content}\"\"\" 
                    Textos de adjuntos (OCR):
                    \"\"\"{json.dumps(adj_texts)}\"\"\" 
                    """
                
                data = None
                try:
                    completion = openai.chat.completions.create(
                        model="gpt-4o-mini", 
                        temperature=0,
                        messages=[{"role": "user", "content": prompt}],
                        response_format={"type": "json_object"}
                    )
                    
                    if completion.choices and completion.choices[0].message and completion.choices[0].message.content:
                        data_str = completion.choices[0].message.content
                        logging.info(f"OpenAI raw response string: {data_str}")
                        try:
                            data = json.loads(data_str)
                        except json.JSONDecodeError as jde:
                            logging.error(f"Failed to parse JSON from OpenAI response: {jde}. Raw response: {data_str}")
                            continue 
                    else:
                        logging.error("OpenAI response is empty or not in the expected format.")
                        if completion:
                            logging.error(f"OpenAI full completion object for debugging: {completion}")
                        continue

                except Exception as e: 
                    logging.error(f"Error calling OpenAI API: {e}")
                    if hasattr(e, 'response') and e.response:
                         logging.error(f"OpenAI API error details: {e.response.text}")
                    continue 
                
                if not data: 
                    logging.error("Failed to obtain structured data from OpenAI. Skipping Excel insertion for this email.")
                    continue

                nombre = data.get("nombre", "")
                cedula = data.get("cedula", "")
                texto_original_from_ai = data.get("texto_original", email_body_content)
                
                adjuntos_from_ai = data.get("adjuntos", [])
                if not isinstance(adjuntos_from_ai, list) or not all(isinstance(item, str) for item in adjuntos_from_ai):
                    logging.warning(f"OpenAI 'adjuntos' field was not a list of strings: {adjuntos_from_ai}. Using empty list instead.")
                    adjuntos_from_ai = []

                excel_row = [
                    nombre, 
                    cedula,
                    texto_original_from_ai,
                    ", ".join(adjuntos_from_ai)
                ]
                
                body_insert = {"values": [excel_row]}
                
                excel_path = f"https://graph.microsoft.com/v1.0/users/{TARGET_USER_ID}/drive/root:/datos/emails.xlsx:/workbook/tables/Table1/rows/add"
                
                insert_response = requests.post(excel_path, headers=headers, json=body_insert)
                if insert_response.status_code >= 300:
                    logging.error(f"Error inserting row into Excel: {insert_response.status_code}. Response: {insert_response.text}")

            return func.HttpResponse(status_code=202)
        else:
            logging.warning(f"POST request with unhandled JSON body structure: {body}")
            return func.HttpResponse("Bad Request: Unhandled JSON structure.", status_code=400)

    logging.warning(f"Unhandled request method: {req.method}. Only GET with validationToken or POST is allowed.")
    return func.HttpResponse("Method not allowed.", status_code=405)

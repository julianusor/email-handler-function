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

def get_graph_token():
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result["access_token"]

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Webhook recibido")

    # --- 1) VALIDACIÓN DE SUSCRIPCIÓN ---------------------------------
    if req.method == "GET" and "validationToken" in req.params:
        # devolver el token tal cual, texto plano y 200 OK
        return func.HttpResponse(
            req.params["validationToken"],
            status_code=200,
            mimetype="text/plain"
        )

    # a partir de aquí, todo lo demás son POST con notificaciones --------
    try:
        body = req.get_json()
    except ValueError:
        return func.HttpResponse("Bad request", status_code=400)
    
    # 1) Validación inicial de Graph (cuando crea el subscription)
    if "validationToken" in body:
        return func.HttpResponse(body["validationToken"], status_code=200)

    # 2) Procesar notificaciones
    for notification in body.get("value", []):
        message_id = notification["resourceData"]["id"]
        user_id    = notification["resourceData"]["from"]["emailAddress"]["address"]

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
            attachments = requests.get(
                mail["@odata.id"] + "/attachments", headers=headers
            ).json()["value"]
            for att in attachments:
                # Lógica OCR aquí; por ejemplo enviarlo a un endpoint de Computer Vision
                text = ocr_from_bytes(att["contentBytes"])
                adj_texts.append(text)

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

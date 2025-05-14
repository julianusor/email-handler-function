import os, requests
from msal import ConfidentialClientApplication

# set .env file to load environment variables
from dotenv import load_dotenv
load_dotenv()
# 1) Variables de entorno
TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
FUNC_URL      = "https://new-email-handler-function.azurewebsites.net/api/"
USER_EMAIL    = "julianusu@outlook.com"

# 2) Obtener token
app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET
)
token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])["access_token"]

# 3) Crear suscripción
url = "https://graph.microsoft.com/v1.0/subscriptions"
headers = {
    "Authorization": f"Bearer {token}",
    "Content-Type": "application/json"
}
body = {
  "changeType": "created",
  "notificationUrl": FUNC_URL,
  "resource": f"/users/{USER_EMAIL}/mailFolders('Inbox')/messages",
  "expirationDateTime": "2025-06-20T11:00:00Z",
  "clientState": "mi-secreto-para-validar"
}

r = requests.post(url, headers=headers, json=body)
print(r.status_code, r.json())

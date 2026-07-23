import msal
from app.core.config import settings

AUTHORITY = f"https://login.microsoftonline.com/{settings.TENANT_ID}"

SCOPES = ["https://graph.microsoft.com/.default"]

msal_app = msal.ConfidentialClientApplication(
    client_id = settings.CLIENT_ID, 
    client_credential = settings.CLIENT_SECRET, 
    authority = AUTHORITY
)

def get_graph_token():
    
    result = msal_app.acquire_token_for_client(
        scopes = SCOPES
    )

    if "access_token" not in result:
        raise Exception(result)
    
    return result["access_token"]
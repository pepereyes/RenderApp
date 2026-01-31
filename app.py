
import os
import json
import msal
import requests
import logging
from flask import Flask, request, redirect, session, send_file
from datetime import datetime

# Logging
logging.basicConfig(level=logging.INFO)

# Flask App
app = Flask(__name__)
app.secret_key = os.urandom(24)

# Load config.json
with open("config.json") as f:
    config = json.load(f)

CLIENT_ID = config["client_id"]
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = config["tenant_id"]  # Should be 'common'
AUTHORITY = config["authority"]
REDIRECT_URI = config["redirect_uri"]
SCOPES = config["scopes"]

ITEM_ID = os.getenv("ITEM_ID")   # This you store in Render

# MSAL: Create Confidential Client
def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists("token_cache.bin"):
        cache.deserialize(open("token_cache.bin", "r").read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        open("token_cache.bin", "w").write(cache.serialize())

def build_msal_app(cache=None):
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache
    )

def get_token():
    cache = load_cache()
    app_msal = build_msal_app(cache)

    # Try refresh token
    accounts = app_msal.get_accounts()
    if accounts:
        result = app_msal.acquire_token_silent(SCOPES, account=accounts[0])
        if result:
            save_cache(cache)
            return result["access_token"]

    # If no token exists, redirect user to login
    return None

# -----------------------
# ROUTES
# -----------------------

@app.route("/")
def index():
    # Check if authenticated
    token = get_token()
    if not token:
        return redirect("/login")

    return send_file("formulario.html")


@app.route("/login")
def login():
    cache = load_cache()
    app_msal = build_msal_app(cache)

    auth_url = app_msal.get_authorization_request_url(
        SCOPES,
        redirect_uri=REDIRECT_URI
    )
    save_cache(cache)
    return redirect(auth_url)


@app.route("/auth/callback")
def callback():
    cache = load_cache()
    app_msal = build_msal_app(cache)

    result = app_msal.acquire_token_by_authorization_code(
        request.args["code"],
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )

    save_cache(cache)

    if "access_token" in result:
        return redirect("/")
    else:
        return f"Error during authentication: {result}"


@app.route("/guardar", methods=["POST"])
def guardar():
    try:
        token = get_token()
        if not token:
            return redirect("/login")

        data = request.json
        fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        values = [[
            fecha_hora,
            data.get("NombreCompleto"),
            data.get("Correo"),
            data.get("Telefono"),
            data.get("Preferencia"),
            data.get("Peticion"),
            data.get("Responsable"),
            data.get("Observaciones")
        ]]

        # INSERT ROW INTO TABLE1 (auto-append)
        url = (
            f"https://graph.microsoft.com/v1.0/me/drive/items/{ITEM_ID}"
            "/workbook/worksheets/SEGUIMIENTO/tables/Table1/rows/add"
        )

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        body = {"values": values}

        logging.info(f"POST to {url} with body {body}")
        response = requests.post(url, headers=headers, json=body)
        logging.info(f"GRAPH RESPONSE: {response.status_code} {response.text}")

        response.raise_for_status()
        return {"status": "ok"}

    except Exception as e:
        logging.error(f"Error en /guardar: {e}")
        return {"status": "error", "message": str(e)}, 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))



import os
import logging
import requests
from flask import Flask, request
from datetime import datetime

logging.basicConfig(level=logging.INFO)
app = Flask(__name__)

# Leer valores desde variables de entorno
ACCESS_TOKEN = os.getenv('ACCESS_TOKEN')
ITEM_ID = os.getenv('ITEM_ID')

@app.route('/')
def index():
    return open('formulario.html').read()

@app.route('/guardar', methods=['POST'])
def guardar():
    try:
        data = request.json
        fecha_hora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        
values = [[
    fecha_hora,
    data.get('NombreCompleto'),
    data.get('Correo'),
    data.get('Telefono'),
    data.get('Preferencia'),
    data.get('Peticion'),
    data.get('Responsable'),
    data.get('Observaciones')
]]

url = f"https://graph.microsoft.com/v1.0/me/drive/items/{ITEM_ID}/workbook/worksheets/SEGUIMIENTO/tables/Table1/rows/add"
headers = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Content-Type': 'application/json'
}
body = {"values": values}

logging.info(f"POST to {url} with body: {body}")
response = requests.post(url, headers=headers, json=body)

        logging.info(f"Response status: {response.status_code}, body: {response.text}")

        response.raise_for_status()
        return {'status': 'ok', 'response': response.json()}
    except Exception as e:
        logging.error(f"Error en /guardar: {e}")
        return {'status': 'error', 'message': str(e)}, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))

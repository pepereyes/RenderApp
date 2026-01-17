import requests
from flask import Flask, request, send_file
from datetime import datetime
import os
import logging
logging.basicConfig(level=logging.INFO)


app = Flask(__name__)

# Reemplaza estos valores con tu token y el ID del archivo en OneDrive
ACCESS_TOKEN = 'EwCYBMl6BAAUu4TQbLz/EdYigQnDPtIo76ZZUKsAAYMi1O89F0WcqaiG3962hxB4Z8HOVrhvxd0AkDguYyHhLILxIAlM7YVI/IbTNkKTBIMZd9A3U7bPNnw5zge3WShcCCmg91hcUIw6WPx1W8wPi+nmq02utgPTyd93jB7qbZawd54b7UckvsToU9584yyQSzRGHm7171941Ho3Q1aQjX6EkpwSNfygYnRYn0VTlIpV4yU5gpCP8sYgmKRYkFlvc2yt0nF8iR6gk++DqFUtPF9K99YOvIWm8mt3OR2ToXDCvEfoNRVh9f/VALJVfoQdsn5vtY2ktXuJYGgdFLjyoTwjDS0yG5gQNhjAfr+aIfYTGPh22rhPCD+uCvUHOGsQZgAAEONITvxr+4KGYKH/V7k4ed5gAz1bN40LiNW20GqKyu5hUQli+shsjEwUDrUxAhkdpbUTBtRFVTJPzyCln1dAdGsRFuHLpNkjr5ID8NGDYKjIXueqUrV3waMp6R7nOADe/Iy91ymtWzy0VB7ezm9v/IAfL8rppumbQ8rgmbFsM3lbm/K0/XLBikuABmXCwEkECKw2NrRYXCWlb0nYmqYLqL2g2RDOClopeTiR7u2X7ar3atEkYC3hLJpzFqVh71YH248HGA6bKcMI6vrXv01/QlXbLG9KkqABh+aYgjbc9Brj9UokPvNmoL6vlrwu1HF95JqGBeB4B0gO19siJiTJFX8biNf9thyCvBEC0kP5TdfTTDI3TSzmLY9CLUkU2iBUfDNYrz6sDjvEL7rBYsFkf4s4WKdSESYubfj4gB7MclwjrBdLNrEyIzGa5odJwFB9O3DNY/17F4SFNYC2cqKzT3KZwHNnkU359HAXz2IXkzxhJmDcZ3gNSjVbnlgG54g7T+OfcGlytwPhIHI1XKE0n7AqHaWrrjFGzKGrBq/NLJ8Q6Fkya1SvPU6zGFNxjJDjhzb05VFwc2zxxXEXeL8iBvRBsqx2y0qIXrjtbwISO3aVt40YgFUSZ+/J6Yee03kLs0uU6aJdUV9XP+cnvxVl4wnkf1Gl4Lw2hyVm73RmX2alQgtSfPIyyWZy2sFkYENGNdVIJRFkCBkdXcyyjhTKABCdyfp1dlQcY24nbNiAfmxq/ApEGaXKaXz9ND0j0BTbDzN/sszjIginPHDcaSUXaP8itBBbfKfiQzyHRHWKdvVGXuoIK+7j5sPldLXm5Zk61Qb37+/FgDxsBcgXaFrBr1WdfAaJWu3/APrCMc5pgBdrVFrMxfb8ECz+wcMu+iaLCLUWY/5uIGUdvZBC75M3Zi+vaaBEBbfBwKvMP9Jjw1X9w8QkFwqYsJeh3UvyUcSOIVYR8QmChysTG5TeWL8E4DavBN2EiCrR7vT8Y6t10ILsUC8awZtv73ClxuFMHuC995cnPUu2ZOD98tRVcRrtKrKqV72gKGNJ5h2xyrCKEgQeYqPStifgC858vUaWmPWag0yIK7+ogkdDKf3DEIFG9azQdQHV1Wj51wLkmzevZ+m5Md4xChm+OlG2lST18n+GcSPoFc+bSyGoy/+2Ygv8v4J1J5ID'

@app.route('/')
def index():
    return send_file('formulario.html')

@app.route('/guardar', methods=['POST'])
def guardar():
    data = request.json
    fecha_hora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Datos a enviar al Excel en OneDrive
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

    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{ITEM_ID}/workbook/worksheets/Hoja1/tables/Table1/rows/add"
    headers = {
        'Authorization': f'Bearer {ACCESS_TOKEN}',
        'Content-Type': 'application/json'
    }
    body = {"values": values}
    logging.info(f"POST to {url} with body: {body}")

    response = requests.post(url, headers=headers, json=body)
    
    logging.info(f"Response status: {response.status_code}, body: {response.text}")

    response.raise_for_status()  # Lanza error si la respuesta no es 200
    return {'status': 'ok', 'response': response.json()}
    except Exception as e:
        logging.error(f"Error en /guardar: {e}")
        return {'status': 'error', 'message': str(e)}, 50


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))


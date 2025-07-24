import requests
import sys
import json
from bs4 import BeautifulSoup
import re

# Datos del request
payload_request = sys.argv[1]
payload_request = payload_request.replace("\n", "")
payload_request = json.loads(payload_request.replace("'",'"'))

# Crear una sesión para manejar las cookies automáticamente
session = requests.Session()

# URL de login
url_login = "https://oet-avansat4.intrared.net:8083/ap/spyme_transc/session.php"

# Datos de la solicitud
payload = {
    "usuario": "mateo.zapata",
    "clave": "transcomercol18",
    "app": "1",
    "op": "1",
    "bd": "spyme_transc",
    "standa": "sate_standa",
    "fondo": "Array",
    "img_log": "../sate_standa/imagenes/avansat-empresarial.png",
    "img_fon": "../sate_standa/imagenes/login/9.jpg",
    "col_fon": "#b80404",
    "col_bot": "#b80404"
}

# Encabezados
headers = {
    "Content-Type": "application/x-www-form-urlencoded"
}

# Hacer login
response_login = session.post(url_login, headers=headers, data=payload)

# Verificar si se obtuvo una cookie de sesión
print("Cookie:", session.cookies)

# URL para aprobar pedido
url_pedido = "https://oet-avansat4.intrared.net:8083/ap/spyme_transc/index.php"


# Enviar solicitud con la sesión autenticada
response_pedido = session.post(url_pedido, headers=headers, data=payload_request)

print(response_pedido.status_code)
print(response_pedido.text)
# Ver la respuesta



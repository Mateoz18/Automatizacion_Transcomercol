import requests
import sys
import json
from bs4 import BeautifulSoup
import re

# Datos del request
payload_request = {
    "window": "central",
    "cod_servic": "337"
}

# Crear una sesión para manejar las cookies automáticamente
session = requests.Session()

# URL de login
url_login = "https://oet-avansat4.intrared.net:8083/ap/spyme_transc/session.php"

# Datos de la solicitud
payload = {
    "usuario": "transc@0912",
    "clave": "soporte@0981",
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
'''print("Cookie:", session.cookies)'''

# URL para consultar manifiestos
url_pedido = "https://oet-avansat4.intrared.net:8083/ap/spyme_transc/index.php"


# Enviar solicitud con la sesión autenticada
response_pedido = session.post(url_pedido, headers=headers, data=payload_request)

#Obtener HTML
html = response_pedido.text

#Parsear HTML
soup = BeautifulSoup(html, 'html.parser')

# Obtener la última tabla (la que contiene los manifiestos)
tabla = soup.find_all('table')[-1]

# Obtener todas las filas de la tabla
filas = tabla.find_all('tr')

# Ignorar la primera fila que es el encabezado
datos = filas[1:]

# Obtener la primera fila con manifiesto
primera_fila = datos[0]
columnas = primera_fila.find_all('td')

# Extraer manifiesto
manifiesto = columnas[0].text.strip()

manifiesto = int(manifiesto) + 1

print(manifiesto)
# Ver la respuesta



import requests
import sys
import json
from bs4 import BeautifulSoup
from html import unescape
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
url_pedido = "https://oet-avansat4.intrared.net:8083/ap/spyme_transc/index.php?cod_servic=38"


# Enviar solicitud con la sesión autenticada
response_pedido = session.post(url_pedido, headers=headers, data=payload_request)
print("StatusCode:", response_pedido.status_code)

#Extraer Pedido
html_response = response_pedido.text

soup = BeautifulSoup(html_response, "html.parser")

mensaje = soup.find("p").text.strip()
print("Resultado: "+mensaje)

# Extraer y limpiar el texto
tds = soup.find_all("td")
for td in tds:
    if "Se han encontrado errores" in td.decode_contents():
        # Primero lo encontramos con decode_contents (más flexible)
        raw_html = td.decode_contents()
        raw_html = raw_html.replace("<br>", "\n")  # respetamos saltos de línea
        mensajeRNDC = BeautifulSoup(raw_html, "html.parser").get_text(separator=" ", strip=True)
        mensajeRNDC = unescape(mensajeRNDC)  # decodifica entidades HTML
        print("Resultado RNDC:\n" + mensajeRNDC)
        break

if "Pedido N" in mensaje:
    numero_pedido = re.search(r'\d+', mensaje).group()
    print("N Pedido: "+str(numero_pedido)+".")
elif "Orden de Cargue" in mensaje:
    orden_cargue = re.search(r'\d+', mensaje).group()
    print("N Orden Cargue: "+str(orden_cargue)+".")
elif "Se corre del consecutivo" in mensaje:
    consecutivos_oc = re.findall(r'\d+', mensaje)
    print(f"Orden Cargue {consecutivos_oc[1]}.")
elif "remesa" in mensaje:
    remesa = re.search(r'\d+', mensaje).group()
    print("N Remesa: "+str(remesa)+".")    


# Ver la respuesta



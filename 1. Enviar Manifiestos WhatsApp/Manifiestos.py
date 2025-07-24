import time
import re
import os
import pandas as pd
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Capturar el argumento
raw_data = sys.argv[1]

# Dividir en filas y luego en columnas
rows = raw_data.split("|")
data = [row.split(",") for row in rows]

# Crear DataFrame
df = pd.DataFrame(data, columns=["Celular", "Archivo"])
df["Celular"] = df["Celular"].apply(lambda x: f'+57{x}')
df["Archivo"] = df["Archivo"].apply(lambda x: x.strip())

# Ruta de geckodriver (Firefox)
geckodriver_path = r"C:\geckodriver\geckodriver.exe"

# Configurar opciones de Firefox
firefox_options = Options()
firefox_options.add_argument("--profile")
firefox_options.add_argument(r"C:\Users\Mateo Zapata\AppData\Roaming\Mozilla\Firefox\Profiles\5edrwgjo.default-release")  

# Inicializar WebDriver
service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=firefox_options)
wait = WebDriverWait(driver, 30)

# Iterar sobre cada fila del archivo Excel
for index, row in df.iterrows():
    celular = row['Celular']
    ruta_archivo = row['Archivo']

    # Extraer el número de manifiesto con regex
    match = re.search(r"Manifiesto Nro\. (\d+)", os.path.basename(ruta_archivo))
    manifiesto_num = match.group(1) if match else "Desconocido"

    #Extraer el número de manifiesto
    match = re.search(r"Manifiesto Nro\. (\d+)", os.path.basename(ruta_archivo))
    manifiesto_num = match.group(1) if match else "Desconocido"
    
    # Abrir chat de WhatsApp Web
    url = f"https://web.whatsapp.com/send?phone={celular}"
    driver.get(url)
    time.sleep(15)

    try:
        # Esperar a que el chat esté disponible
        chat_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"]')))
        time.sleep(2)

        # Verificar si el botón de adjuntar está presente
        clip_button = wait.until(EC.presence_of_element_located((By.XPATH, '//button[@title="Adjuntar"]')))
        driver.execute_script("arguments[0].click();", clip_button)
        time.sleep(2)

        # Subir archivo
        file_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="file"]')))
        file_input.send_keys(ruta_archivo)
        time.sleep(3)

        # Enviar archivo
        send_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="wds-ic-send-filled"]')))
        driver.execute_script("arguments[0].click();", send_button)
        time.sleep(5)

        print(f"Manifiesto Nro {manifiesto_num} enviado a {celular}")
        os.remove(ruta_archivo)

    except Exception as e:
        print(f"Error enviando el Manifiesto Nro {manifiesto_num} a {celular}")

driver.quit()

import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import sys

# Carga el JSON de datos
datos = json.load(open('datos.json'))

# Cargar datos de la session
sesion = webdriver.ChromeOptions()
sesion.add_argument(r'user-data-dir=C:\Users\gilberto.castrillon\AppData\Local\Google\Chrome\User Data\Default')
sesion.add_argument('--profile-directory=Default')

# Iniciar el navegador con la sesión cargada
driver = webdriver.Chrome(r'C:\driver\chromedriver.exe', options=sesion)
driver.get('https://web.whatsapp.com/')

# Esperar a que se cargue la página y buscar un contacto
driver.implicitly_wait(20)
driver.find_element("xpath", '//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').click()
driver.find_element("xpath", '//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').send_keys(datos['contacto'])

# Clic en el contacto
driver.find_element("xpath", '//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').send_keys(Keys.ENTER)

# Adjuntar el video
driver.find_element("xpath", '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div').click()
driver.implicitly_wait(20)
image_box = driver.find_element("xpath", '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]').send_keys(datos['video'])

# Añadir un mensaje
driver.implicitly_wait(30) # time.sleep(10)
driver.find_element("xpath", '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[1]/div[2]').click()
driver.find_element("xpath", '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[1]/div[2]').send_keys(datos['mensaje'])
driver.find_element("xpath", '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div[2]/div[1]/div[2]').send_keys(Keys.ENTER)
print('mensaje enviado')
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time
import os
import pandas as pd
from openpyxl import load_workbook

opciones_chrome = Options()
download_dir = r"C:\Users\primo\Desktop\Ciencia de Datos\Proyectos\Open"  
prefs = {
    "download.default_directory": download_dir,  
    "download.prompt_for_download": False,      
    "directory_upgrade": True,                  
    "safebrowsing.enabled": True                
}
opciones_chrome.add_experimental_option("prefs", prefs)

# Inicializar el navegador con las opciones de descarga configuradas
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opciones_chrome)

# Ir a la URL de login
url_login = "https://www.herrafe.com/acceso"
driver.get(url_login)

# Busca los campos de usuario y contraseña usando ID
campo_usuario = driver.find_element(By.ID, "username")  
campo_contraseña = driver.find_element(By.ID, "password")  

# Ingresa datos en los campos
campo_usuario.send_keys("******")  # Reemplazar con datos de usuario
campo_contraseña.send_keys("******")  # Remplazar con datos de contraseña
campo_contraseña.send_keys(Keys.RETURN)  

# Espera a que el enlace 'Mis Listas' esté presente e ingresa
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//a[@href='/acceso/mis-listas-de-precios']"))
)
mis_listas_link = driver.find_element(By.XPATH, "//a[@href='/acceso/mis-listas-de-precios']")
mis_listas_link.click()

# Espera a que el ícono de Excel esté presente y descarga
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//la-icon[@icon='file-excel']"))
)
download_icon = driver.find_element(By.XPATH, "//la-icon[@icon='file-excel']")
download_icon.click()

# Mensaje de confirmación
print("Archivo descargado con éxito.")
time.sleep(7)
# Cierra el navegador
driver.quit()

# Encuentra el archivo más reciente en la carpeta de descargas y carga el df
def find_latest_file(download_dir):
    files = [os.path.join(download_dir, f) for f in os.listdir(download_dir)]
    latest_file = max(files, key=os.path.getctime) 
    return latest_file

latest_file = find_latest_file(download_dir)
print(f"Archivo más reciente encontrado: {latest_file}")
df_precios = pd.read_excel(latest_file, skiprows=7)

# Define los DF de bases de datos existentes y el nuevo para comprobar que haya habido modificaciones
bdd_precios = r"C:\Users\primo\Desktop\Ciencia de Datos\Proyectos\Open\Lista de precios OPEN.xlsx"
hoja_destino = "Herrafe"
df_comparacion = pd.read_excel(latest_file, skiprows=8)
df_bdd = pd.read_excel(bdd_precios, skiprows=1, sheet_name= hoja_destino)


# Compara el archivo de Excel existente y si hubo modificaciones actualiza la base de datos de precios
if df_comparacion.equals(df_bdd):
    print("No hubo modificación en los precios")
else:
    with pd.ExcelWriter(bdd_precios, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_precios.to_excel(writer, index=False, sheet_name=hoja_destino)
    print("Archivo con actualización de los precios cargado exitosamente")


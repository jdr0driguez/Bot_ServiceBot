# main.py
from config import EMAIL, PASSWORD, BASE_URL
from login import login_selenium
from download_csv import download_csv_selenium, obtener_ruta_documentos
from browser_config import configurar_driver
from sendAPI import enviar_archivo_api

def run():
    
    driver = configurar_driver()

    try:
        
        login_selenium(driver, EMAIL, PASSWORD, BASE_URL)

        
        obtener_ruta_documentos()

        
        download_csv_selenium(driver)

        enviar_archivo_api()

    finally:
        
        driver.quit()

if __name__ == "__main__":
    run()

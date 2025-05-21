

import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pywinauto import Desktop
import time
import ctypes.wintypes
import os
from config import CSV_FILENAME, BASE_URL, COMPLE_BASEURL

def obtener_ruta_documentos():
    CSIDL_PERSONAL = 5  
    SHGFP_TYPE_CURRENT = 0

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)

    return buf.value

ruta_documentos = obtener_ruta_documentos()


def download_csv_selenium(driver):
    try:
        wait = WebDriverWait(driver, 20)

        
        driver.get(BASE_URL + COMPLE_BASEURL)
        time.sleep(6)


        
        wait = WebDriverWait(driver, 15)
        iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)

       
        reporte_unicos = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[title="Reporte Unicos"]'))
        )

        driver.execute_script("document.querySelector('span[title=\"Reporte Unicos\"]').click();")

        time.sleep(5)

        driver.execute_script("document.getElementById('sheet_control_panel_header').click();")

        time.sleep(5)


        

        
        rel_date_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[data-automation-id="sheet-control-relative-date"]'))
        )
        rel_date_input.click()
        time.sleep(1)

        
        relative_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='dateType'][value='relative']"))
        )
        driver.execute_script("arguments[0].click();", relative_input)
        time.sleep(0.5)

        
        thisday_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='Relative by'][value='This day']"))
        )
        driver.execute_script("arguments[0].click();", thisday_input)
        time.sleep(0.5)
        
        thisday_input.send_keys(Keys.ESCAPE)

        

        wait = WebDriverWait(driver, 10)
        actions = ActionChains(driver)

        
        tabla = wait.until(EC.presence_of_element_located(
            (By.ID, 'block-e0f7648e-bf39-4d90-8a1a-042933ba4471_ffe0b6d9-3fa0-46da-af01-b9b50ed55d9f')
        ))

        
        actions.move_to_element(tabla).perform()
        time.sleep(1)

       
        boton_menu = tabla.find_element(
            By.CSS_SELECTOR, 'button[data-automation-id="analysis_visual_dropdown_menu_button"]'
        )
        driver.execute_script("arguments[0].click();", boton_menu)
        time.sleep(0.5)

        
        export_csv = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '[data-automation-id="dashboard_visual_dropdown_export"]')
        ))
        driver.execute_script("arguments[0].click();", export_csv)

        
        nombre_archivo = CSV_FILENAME

        
        ruta_archivo = os.path.join(ruta_documentos, nombre_archivo)

       
        time.sleep(3)

       
        desktop = Desktop(backend="win32")

        
        ventanas = desktop.windows(class_name="#32770")
        ventana_guardar = None

        for v in ventanas:
            titulo = v.window_text().lower()
            if "guardar como" in titulo or "save as" in titulo:
                ventana_guardar = v
                break

        if ventana_guardar:
            print("✅ Ventana detectada:", ventana_guardar.window_text())
            ventana_guardar.set_focus()

            
            try:
                edits = ventana_guardar.descendants(class_name="Edit")
                for idx, edit in enumerate(edits):
                    try:
                        edit.set_edit_text(ruta_archivo)
                        print(f"✍️ Ruta escrita en Edit #{idx+1}")
                        break
                    except Exception as e:
                        print(f"❌ No se pudo escribir en Edit #{idx+1}: {e}")
            except Exception as e:
                print("❌ No se encontraron campos Edit:", e)

            
            try:
                botones = ventana_guardar.descendants(class_name="Button")
                guardado = False
                for btn in botones:
                    if "guardar" in btn.window_text().lower() or "save" in btn.window_text().lower():
                        btn.double_click_input()
                        print("✅ Botón 'Guardar/Save' presionado")
                        guardado = True
                        break
                if not guardado:
                    print("❌ No se encontró botón Guardar/Save")
            except Exception as e:
                print("❌ No se pudo hacer clic en Guardar:", e)

            
            time.sleep(1)

            ventanas_conf = desktop.windows(class_name="#32770")
            for v in ventanas_conf:
                titulo = v.window_text().lower()
                if "confirmar" in titulo or "confirm save as" in titulo:
                    try:
                        v.set_focus()
                        botones_conf = v.descendants(class_name="Button")
                        for btn in botones_conf:
                            if "sí" in btn.window_text().lower() or "yes" in btn.window_text().lower():
                                btn.click()
                                print("🟢 Confirmación de sobrescritura aceptada")
                                break
                    except Exception as e:
                        print("❌ No se pudo confirmar sobrescritura:", e)

        else:
            print("❌ No se encontró la ventana 'Guardar como' / 'Save As'")

    except Exception as e:
        print(f"❌ Error durante la automatización: {e}")


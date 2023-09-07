from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import os
import time
import traceback
import shutil
import glob
from pathlib import Path
import zipfile
from selenium.common.exceptions import StaleElementReferenceException
import sys
from dotenv import load_dotenv

logs_directory = os.path.join(os.path.dirname(__file__), '..', 'logs')
sys.path.append(logs_directory)

from config_log import *

ruta_descargas = os.path.join(os.path.expanduser('~'), 'Downloads')
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

load_dotenv()

# Email Credentials
EMAIL = os.getenv('EMAIL_USER')
PASSWORD = os.getenv('EMAIL_PASSWORD')
OUTLOOK_SITE = os.getenv('EMAIL_URL')

logger = configure_logger('descarga_mail')

def encontrar_elemento_karpay(wait, fecha):
    buscador_email = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="topSearchInput"]')))
    elemento_a_buscar = f"A/B {fecha}"
    buscador_email.send_keys(elemento_a_buscar)
    time.sleep(3)
    btn_buscador = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button[aria-label="Buscar"]')))
    
    btn_buscador.click()
    
    time.sleep(5)

    try:
        elementos_spans = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'span')))
        
        for elemento in elementos_spans:
            texto = elemento.get_attribute('innerHTML')
            if texto.startswith('ALERTA ROJA') and fecha in texto:
                return elemento
            
    except Exception as e:
        traceback.print_exc()
        logger.error(f'Error al encontrar el elemento Karpay: {str(e)}')

    return None

def obtener_archivos_karpay(wait, fecha):
    logger.info('Buscando archivos Karpay. La fecha a consultar es: ')
    logger.info(fecha)
    ruta_destino_karpay = os.path.join(desktop_path, r'file_download\REPORTES KARPAY')

    # Verificar si la carpeta de destino existe, y si no, crearla
    if not os.path.exists(ruta_destino_karpay):
        os.makedirs(ruta_destino_karpay)

    try:
        asunto_correo = encontrar_elemento_karpay(wait, fecha)
    
        if asunto_correo is not None:
            contenido_asunto = asunto_correo.get_attribute("innerHTML")
            
            asunto_correo.click()
            time.sleep(10)

            dato_adjunto = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//span[contains(@class, 'ms-Button-label') and text()='Descargar todo']")))

            time.sleep(5)

            if dato_adjunto:
                logger.info('Existe el botón descargar todo')
                dato_adjunto.click()

                time.sleep(15)

                nombre_zip = None

                # Buscar archivos zip que empiecen con "ALERTA ROJA" en la carpeta de descargas
                archivos_descargas = glob.glob(os.path.join(ruta_descargas, "ALERTA ROJA*.zip"))

                if archivos_descargas:
                    nombre_zip = archivos_descargas[0]
                    logger.info(f"Se encontró el archivo zip: {nombre_zip}")

                    # Mover el archivo zip a la ruta de destino
                    ruta_destino_archivo_zip = os.path.join(ruta_destino_karpay, os.path.basename(nombre_zip))
                    Path(nombre_zip).rename(ruta_destino_archivo_zip)
                    logger.info(f"Se movió el archivo zip a: {ruta_destino_archivo_zip}")

                    with zipfile.ZipFile(ruta_destino_archivo_zip, 'r') as zip_ref:
                        zip_ref.extractall(ruta_destino_karpay)
                    logger.info("Archivos descomprimidos correctamente.")

                    # Eliminar el archivo zip
                    os.remove(ruta_destino_archivo_zip)
                    logger.info("Archivo zip eliminado.")

                else:
                    logger.error("No se encontró ningún archivo zip con nombre que empiece por 'ALERTA ROJA'.")


            else:
                logger.error("No se encontró el botón de descargar archivos adjuntos en el correo.")

        else:
            logger.error("El asunto del correo no contiene la palabra 'ALERTA ROJA SPEI'.")


    except Exception as e:
        traceback.print_exc()
        logger.error(f'Error al obtener los archivos de la instancia KARPAY: {str(e)}')


def encontrar_elemento_flex(wait, fecha):
    buscador_email = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="topSearchInput"]')))
    elemento_a_buscar = f"FLEX {fecha}"
    buscador_email.send_keys(elemento_a_buscar)
    time.sleep(3)
    btn_buscador = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button[aria-label="Buscar"]')))
    
    btn_buscador.click()
    time.sleep(5)
    try:
        elementos_spans = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'span')))

        for elemento in elementos_spans:
            try:
                texto = elemento.get_attribute('innerHTML')
                if texto == 'FLEX':
                    return elemento
            except StaleElementReferenceException:
                continue

    except Exception as e:
        traceback.print_exc()
        logger.error(f'Error al encontrar el elemento FLEX: {str(e)}')

    return None

def obtener_archivo_flex(wait, fecha):
    logger.info('Buscando el archivo flex. La fecha a buscar es: ')
    logger.info(fecha)

    ruta_destino_flex = os.path.join(desktop_path, r'file_download\REPORTES FLEX')

    # Verificar si la carpeta de destino existe, y si no, crearla
    if not os.path.exists(ruta_destino_flex):
        os.makedirs(ruta_destino_flex)

    try:
        asunto_correo = encontrar_elemento_flex(wait, fecha)
        
        if asunto_correo is not None:
            contenido_asunto = asunto_correo.get_attribute("innerHTML")

            asunto_correo.click()
            time.sleep(5)

            dato_adjunto = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//span[contains(@class, 'ms-Button-label') and text()='1 dato adjunto']")))

            if dato_adjunto:
                mas_acciones_button = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//button[contains(@class, 'ms-Button') and @aria-label='Más acciones']")))
                time.sleep(2)
                mas_acciones_button.click()
                time.sleep(1)
                descargar_button = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//span[contains(@class, 'ms-ContextualMenu-itemText') and text()='Descargar']")))
                descargar_button.click()

                time.sleep(5)

                try:
                    # Buscar archivos en la ruta de descargas
                    archivos_descargas = os.listdir(ruta_descargas)

                    for archivo in archivos_descargas:
                        if "Reports" in archivo and archivo.endswith(".xls"):
                            ruta_archivo_descarga = os.path.join(ruta_descargas, archivo)
                            ruta_archivo_destino = os.path.join(ruta_destino_flex, archivo)

                            shutil.copy2(ruta_archivo_descarga, ruta_archivo_destino)

                            logger.info(f"Se copió el archivo a: {ruta_archivo_destino}")

                            os.remove(ruta_archivo_descarga)

                            logger.info(f"Se eliminó el archivo de descargas: {ruta_archivo_descarga}")

                            break

                    else:
                        logger.error("No se encontró ningún archivo válido en la ruta de descargas.")

                except Exception as e:
                    traceback.print_exc()
                    logger.error(f'Error al obtener el archivo FLEX: {str(e)}')
            else:
                logger.error("No se encontró el botón de descargar archivos adjuntos en el correo.")

        else:
            logger.error("No se encontró ningún elemento con asunto 'FLEX'.")

    except Exception as e:
        traceback.print_exc()
        logger.error(f'Error al obtener el archivo FLEX: {str(e)}')



def ingresar_al_email(tipo_archivo='', fecha=''):
    email = EMAIL
    contrasenia = PASSWORD

    driver = webdriver.Firefox()

    try:
        wait = WebDriverWait(driver, 20)

        driver.get(OUTLOOK_SITE)
        
        iniciar_sesion_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='mectrl_headerPicture']")))
        time.sleep(1)
        iniciar_sesion_btn.click()
        
        email_input = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='email']")))
        email_input.send_keys(email)


        siguiente_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='idSIButton9']")))
        siguiente_btn.click()

        try:
            time.sleep(5)
            contrasenia_div = wait.until(EC.visibility_of_element_located(
                (By.XPATH, "//*[@id='loginHeader']/div")))
            if contrasenia_div:
                time.sleep(2)
                try:
                    contrasenia_input = wait.until(
                        EC.presence_of_element_located((By.ID, "i0118")))
                    contrasenia_input.send_keys(contrasenia)
                except Exception as e:
                    logger.error(f'Error al encontrar el elemento de contraseña: {str(e)}')

            if contrasenia_input:
                time.sleep(2)
                ingresar_btn = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//*[@id='idSIButton9']")))
                ingresar_btn.click()

        except Exception as e:
            logger.error(f'Error al querer ingresar la contraseña: {str(e)}')

        # Verificar si se requiere acción adicional después de ingresar la contraseña
        try:
            wait.until(EC.visibility_of_element_located(
                (By.XPATH, "/html/body/div/form/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[1]")))
            # Si el elemento aparece, apretar el botón con el identificador "idBtn_Back"
            btn_back = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[@id='idBtn_Back']")))
            btn_back.click()

        except Exception as e:
            logger.error(f'Error al querer cerrar el modal: {str(e)}')

        time.sleep(5)
        tenant_logo_img = wait.until(EC.visibility_of_element_located(
            (By.XPATH, "//*[@id='O365_MainLink_TenantLogoImg']")))
        if tenant_logo_img.is_displayed():
            logger.info("Se ha iniciado sesión correctamente")
            if tipo_archivo == 'flex':
                obtener_archivo_flex(wait, fecha)
            elif tipo_archivo == 'karpay':
                obtener_archivos_karpay(wait, fecha)
            else:
                logger.error('No se seleccionó el tipo de archivo')
    except Exception as e:
        logger.error(f'Error al ingresar al correo electrónico: {str(e)}')

    finally:
        driver.quit()


#ingresar_al_email(tipo_archivo='flex', fecha='15-08-2023')
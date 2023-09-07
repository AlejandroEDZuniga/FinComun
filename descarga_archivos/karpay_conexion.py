from selenium import webdriver
import time
from dotenv import load_dotenv
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import shutil
import sys

logs_directory = os.path.join(os.path.dirname(__file__), '..', 'logs')
sys.path.append(logs_directory)

from config_log import *

load_dotenv()

# Karpay Credentials
USERNAME = os.getenv('KARPAY_USER')
PASSWORD = os.getenv('KARPAY_PASSWORD')
KARPAY_SITE = os.getenv('KARPAY_URL')


logger = configure_logger('karpay_conexion')

def cerrar_sesion(driver):
    try:
        span_logout_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH , "//span[@class='link-salir item-navegation']")))
        # Encuentra el enlace dentro del span
        logout_btn = span_logout_element.find_element(By.TAG_NAME, 'a')
        logout_btn.click()

        time.sleep(5)

    except TimeoutException:
        logger.error("El enlace 'Salir' no se encontró o no pudo ser clicado.")

def mover_archivos_instancia_A(driver):
    ruta_descargas = os.path.join(os.path.expanduser('~'), 'Downloads')
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    ruta_destino_karpay = os.path.join(desktop_path, r'file_download\REPORTES KARPAY')

    contenido = os.listdir(ruta_descargas)

    try:
        for elemento in contenido:
            if elemento.startswith("EnvioOrdenes") or elemento.startswith("OrdenesPendientes") or elemento.startswith("OrdenesRecibidas"):
                if elemento.endswith(".xlsx"):
                    origen = os.path.join(ruta_descargas, elemento)
                    destino = os.path.join(ruta_destino_karpay, elemento)
                    shutil.move(origen, destino)
                    logger.info(f"Se movió {elemento} a {ruta_destino_karpay}")

        cerrar_sesion(driver)

    except Exception as e:
        logger.error(e)    


def descarga_ordenes(driver):
    try:
        ordenes_de_pago_element = driver.find_element(By.XPATH, "//span[text()='ORDENES DE PAGO']")
        ordenes_de_pago_element.click()

        estado_ordenes_de_pago_element = driver.find_element(By.XPATH, "//span[text()='Estado de Ordenes']")
        estado_ordenes_de_pago_element.click()

        time.sleep(5)

        check_element = driver.find_element(By.XPATH, "//font[contains(@class, 'subtitle-black-general') and contains(text(), 'Envío de Órdenes (Autorizada, Enviando, Enviada, Liquidada Sólo al Filtrar)')]")

        if check_element:
            logger.info("Se accedió al estado de las ordenes correctamente.")
            time.sleep(3)

            try:
                input_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR , "input[type='text'][name='IniOpFechaOper']")))
                driver.execute_script("arguments[0].removeAttribute('readonly')", input_element)
                time.sleep(3)

                input_element.clear()
                input_element.send_keys('15/08/2023')

                time.sleep(5)

                consultar_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR , "input[type='button'][name='Aplicar']")))
                consultar_btn.click()

                time.sleep(3)

                descargar_enviadas_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR , "input[type='button'][name='ExportaLibExcel']")))
                descargar_enviadas_btn.click()

                time.sleep(3)

                descargar_pendientes_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR , "input[type='button'][name='ExportaEnvExcel']")))
                descargar_pendientes_btn.click()

                time.sleep(3)

                descargar_recibidas_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR , "input[type='button'][name='ExportaDevExcel']")))
                descargar_recibidas_btn.click()

                time.sleep(3)

                mover_archivos_instancia_A(driver)

            except Exception as e:
                logger.error(e)
        else:
            logger.error("No se pudo acceder al estado de las ordenes.")

    except Exception as e:
        logger.error(e)

def login_to_karpay_A():
    driver = webdriver.Firefox()
    try:
        driver.get(KARPAY_SITE)

        time.sleep(3)

        j_username = WebDriverWait(driver, 10).until(
            lambda driver: driver.find_element(By.NAME, "j_username")
        )
        
        j_username.send_keys(USERNAME)
        j_password = driver.find_element(By.NAME, "j_password")
        j_password.send_keys(PASSWORD)

        time.sleep(3)

        entrar_button = driver.find_element(By.NAME, "aceptar")
        entrar_button.click()

        time.sleep(10)

        # Cambia al frame fmeCuerpo
        frame_cuerpo = driver.find_element(By.NAME, "fmeCuerpo")
        driver.switch_to.frame(frame_cuerpo)

        try:
            ordenes_de_pago_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[@class='menu-text font-menu' and text()='ORDENES DE PAGO']"))
            )

        except TimeoutException:
            logger.error("El elemento no se pudo encontrar después de esperar 10 segundos.")

        if ordenes_de_pago_element:
            logger.info("Sesión iniciada correctamente.")
            descarga_ordenes(driver)
        else:
            logger.error("No se pudo iniciar sesión.")

        driver.switch_to.default_content()

    except Exception as e:
        logger.error(e)
    
    finally:
        driver.quit()

#login_to_karpay_A()
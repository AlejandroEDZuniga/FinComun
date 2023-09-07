import sys
from logs.config_log import *

logger = configure_logger('run')

def excepthook(exc_type, exc_value, exc_traceback):
    logger.error("Excepción no capturada:", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = excepthook

import menu_opciones
menu_opciones.login()
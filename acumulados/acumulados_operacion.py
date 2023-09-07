import os
import shutil
import sys
import dateparser
import logging
import traceback
from logs.config_log import *

sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'reportes'))
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from utilidades import *
from karpay_operacion import open_files
from importar_archivos import *

desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
path =  os.path.join(desktop_path, 'Acumulados')
ruta_destino = os.path.join(desktop_path, 'file_download')
ruta_procesados = os.path.join(desktop_path, 'Procesados')

current_directory = os.getcwd()
ruta_file_operacion = os.path.join(current_directory, 'file_operacion')

def obtener_archivos_para_generar_reportes_acumulados(fecha_inicio, fecha_fin):

    try:
        lista_archivos = os.listdir(path)

        fecha_inicio_dt = dateparser.parse(fecha_inicio)
        fecha_fin_dt = dateparser.parse(fecha_fin)

        lista_carpetas = [
            archivo for archivo in lista_archivos
            if os.path.isdir(os.path.join(path, archivo)) and
            fecha_inicio_dt <= dateparser.parse(archivo) <= fecha_fin_dt
        ]

        if os.listdir(ruta_destino):
            shutil.rmtree(ruta_destino)
            os.mkdir(ruta_destino)

        if os.listdir(ruta_file_operacion):
            shutil.rmtree(ruta_file_operacion)
            os.mkdir(ruta_file_operacion)

        for carpeta in lista_carpetas:
            ruta_carpeta_origen = os.path.join(path, carpeta)
            ruta_carpeta_destino = os.path.join(ruta_destino, carpeta)
            ruta_acumulados = os.path.join(ruta_procesados, f"Acumulados del {fecha_inicio} al {fecha_fin}")
            ruta_acumulados = ruta_acumulados.replace("\\", "/")
            os.makedirs(ruta_acumulados, exist_ok=True)

            # lista de archivos en la carpeta
            lista_archivos = os.listdir(ruta_carpeta_origen)

            # Ejecutar las acciones por cada carpeta dentro de Acumulados
            for subcarpeta in lista_archivos:
                ruta_subcarpeta_origen = os.path.join(ruta_carpeta_origen, subcarpeta)
                ruta_subcarpeta_destino = os.path.join(ruta_destino, subcarpeta)
                shutil.copytree(ruta_subcarpeta_origen, ruta_subcarpeta_destino)
            normalize_and_save_karpay_files()
            normalize_and_save_sibamex_files()
            save_flex_folder()
            open_files(ruta_acumulados)
            shutil.rmtree(ruta_destino)
            os.mkdir(ruta_destino)
            shutil.rmtree(ruta_file_operacion)
            os.mkdir(ruta_file_operacion)
            logging.info("Finalización exitosa de la función obtener_archivos_para_generar_reportes_acumulados.")
    except Exception as e:
        acumulados_obtener_archivos_logger = configure_logger('acumulados_obtener_archivos', 'obtener_archivos_acumulados.log')
        error_message = f"Ocurrió un error durante la descarga de los archivos para los reportes acumulados: {str(e)}"
        acumulados_obtener_archivos_logger.error(error_message)
        traceback.print_exc()

#obtener_archivos_para_generar_reportes_acumulados(fecha_inicio = '10 enero 2023', fecha_fin = '11 enero 2023')


def unificar_guardar_reportes_acumulados():
    carpeta_acumulados = [nombre for nombre in os.listdir(ruta_procesados) if os.path.isdir(os.path.join(ruta_procesados, nombre)) and nombre.startswith('Acumulados')]

    carpeta_acumulados.sort(key=lambda x: os.path.getmtime(os.path.join(ruta_procesados, x)), reverse=True)

    ruta_subcarpeta = None
    subcarpeta_eliminar = None 
    subcarpetas_a_eliminar = []
    
    for carpeta_acumulado in carpeta_acumulados:
        ruta_carpeta_acumulado = os.path.join(ruta_procesados, carpeta_acumulado)
        
        subcarpetas = [nombre for nombre in os.listdir(ruta_carpeta_acumulado) if os.path.isdir(os.path.join(ruta_carpeta_acumulado, nombre))]
        
        #print(f"Contenido de la carpeta {carpeta_acumulado}:")

        archivos_RCOBO = []
        archivos_RTERO = []
        archivos_RDSERO = []
        archivos_RDSERO2 = []
        
        for subcarpeta in subcarpetas:
            ruta_subcarpeta = os.path.join(ruta_carpeta_acumulado, subcarpeta)
            contenido_subcarpeta = os.listdir(ruta_subcarpeta)
            
            #print(f"Contenido de la subcarpeta {subcarpeta}:")

            archivos_a_eliminar = ['CCOCC.xlsx', 'CKCO.xlsx', 'Flex.xlsx', 'Local.xlsx']

            subcarpetas_a_eliminar.append(ruta_subcarpeta)

            for archivo in contenido_subcarpeta:
                ruta_archivo = os.path.join(ruta_subcarpeta, archivo)

                if archivo not in archivos_a_eliminar:
                    #print('Archivo: ' + archivo)
                    if archivo.startswith('RCOBO'):
                        archivos_RCOBO.append((archivo, ruta_archivo))
                    if archivo.startswith('RTERO'):
                        archivos_RTERO.append((archivo, ruta_archivo))
                    if archivo == 'RDSERO.xlsx':
                        archivos_RDSERO.append((archivo, ruta_archivo))
                    if archivo == 'RDSERO2.xlsx':
                        archivos_RDSERO2.append((archivo, ruta_archivo))
                else:
                    # Eliminar el archivo
                    os.remove(ruta_archivo)

        

         # Invocar funciones que unifican los archivos
        if archivos_RCOBO:
            unificar_guardar_archivos_RCOBO(ruta_carpeta_acumulado, archivos_RCOBO)
        if archivos_RTERO:
            unificar_guardar_archivos_RTERO(ruta_carpeta_acumulado, archivos_RTERO)
        if archivos_RDSERO:
            unificar_guardar_archivos_RDSERO(ruta_carpeta_acumulado, archivos_RDSERO)
        if archivos_RDSERO2:
            unificar_guardar_archivos_RDSERO2(ruta_carpeta_acumulado, archivos_RDSERO2)

        subcarpetas_a_eliminar.append(ruta_subcarpeta)

    # Eliminar las subcarpetas y su contenido
    for subcarpeta_eliminar in subcarpetas_a_eliminar:
        if subcarpeta_eliminar and os.path.exists(subcarpeta_eliminar):
            shutil.rmtree(subcarpeta_eliminar)

#unificar_guardar_reportes_acumulados()

def buscar_y_unificar_reportes(fecha_inicio, fecha_fin):
    fecha_inicio_dt = dateparser.parse(fecha_inicio)
    fecha_fin_dt = dateparser.parse(fecha_fin)

    carpetas_en_carpeta_procesados = os.listdir(ruta_procesados)

    carpetas_coincidentes = []
    for carpeta in carpetas_en_carpeta_procesados:
        try:
            # Obtener solo el comienzo de la fecha desde el nombre de la carpeta
            fecha_carpeta_comienzo = carpeta.split('_')[0]

            # Utilizar dateparser para analizar la fecha del comienzo de la carpeta
            fecha_carpeta_dt = dateparser.parse(fecha_carpeta_comienzo, settings={'DATE_ORDER': 'YMD'})
            if fecha_carpeta_dt and fecha_inicio_dt.date() <= fecha_carpeta_dt.date() <= fecha_fin_dt.date():
                carpetas_coincidentes.append(carpeta)
        except Exception as e:
            buscar_unificar_reportes_logger = configure_logger('buscar_unificar_reportes', 'buscar_unificar_reportes.log')
            error_message = f"Ocurrió un error durante la descarga y unificacion de los reportes acumulados: {str(e)}"
            buscar_unificar_reportes_logger.error(error_message)
            traceback.print_exc()

    carpeta_acumulados = f"Acumulados del {fecha_inicio} al {fecha_fin}"
    ruta_carpeta_acumulados = os.path.join(ruta_procesados, carpeta_acumulados)
    os.makedirs(ruta_carpeta_acumulados, exist_ok=True)

    archivos_RCOBO = []
    archivos_RTERO = []
    archivos_RDSERO = []
    archivos_RDSERO2 = []

    for carpeta in carpetas_coincidentes:
        ruta_carpeta = os.path.join(ruta_procesados, carpeta)

        for archivo in os.listdir(ruta_carpeta):
            if archivo.startswith("RCOBO"):
                archivos_RCOBO.append((archivo, os.path.join(ruta_carpeta, archivo)))
            elif archivo.startswith("RTERO"):
                archivos_RTERO.append((archivo, os.path.join(ruta_carpeta, archivo)))
            elif archivo == "RDSERO.xlsx":
                archivos_RDSERO.append((archivo, os.path.join(ruta_carpeta, archivo)))
            elif archivo == "RDSERO2.xlsx":
                archivos_RDSERO2.append((archivo, os.path.join(ruta_carpeta, archivo)))

    unificar_guardar_archivos_RCOBO(ruta_carpeta_acumulados, archivos_RCOBO)
    unificar_guardar_archivos_RTERO(ruta_carpeta_acumulados, archivos_RTERO)
    unificar_guardar_archivos_RDSERO(ruta_carpeta_acumulados, archivos_RDSERO)
    unificar_guardar_archivos_RDSERO2(ruta_carpeta_acumulados, archivos_RDSERO2)


#buscar_y_unificar_reportes(fecha_inicio = '09 Enero 2023', fecha_fin = '13 Enero 2023')
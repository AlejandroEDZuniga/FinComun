import subprocess
import os
import fileinput
import shutil
import locale
from datetime import timedelta, datetime
import sys
import logging
import time

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from utiles import *

documents_path = os.path.join(os.path.expanduser('~'), 'Documents')

def ejecutar_sikulix(new_fday_value):
    sikulix_path = r'C:\jdk19\bin'
    
    script_paths = {
        'unIngreso': os.path.join(documents_path, 'unIngreso.sikuli', 'unIngreso.py'),
        'mov_cuenta': os.path.join(documents_path, 'MovCuentaOperativa.sikuli', 'MovCuentaOperativa.py')
    }

    # Sobreescribir la variable fday en cada archivo de script
    for script_name, script_path in script_paths.items():
        with fileinput.FileInput(script_path, inplace=True) as file:
            for line in file:
                if line.startswith('fday ='):
                    print(f'fday = date({new_fday_value.year}, {new_fday_value.month}, {new_fday_value.day})')
                else:
                    print(line, end='')
    
    # Cambiar al directorio de SikuliX
    os.chdir(sikulix_path)
    
    # Ejecutar los scripts actualizados
    for script_path in script_paths.values():
        time.sleep(3)
        subprocess.call(['java', '-jar', r'..\sikulixide.jar', '-r', script_path])

    logging.info("La ejecución de SikuliX ha finalizado.")



def mover_archivos_reportes():
    ruta_origen = os.path.join(os.path.expanduser('~'), 'Desktop')
    ruta_destino = os.path.join(ruta_origen, r'file_download\REPORTES SIBAMEX')
    ruta_reporte_mov_cuenta = os.path.join(ruta_origen, r'Sibamex_80\Sibamex_80')

    # Verificar si la carpeta de destino existe, si no, crearla
    if not os.path.exists(ruta_destino):
        os.makedirs(ruta_destino)

    archivos_buscar = ['SPEI-CAPTURADOS.xls', 'SPEI-RECIBIDAS.csv', 'PAGOS Y CAPTACIONES.csv']

    archivos_faltantes = []

    for archivo in archivos_buscar:
        ruta_archivo_origen = os.path.join(ruta_origen, archivo)
        ruta_archivo_destino = os.path.join(ruta_destino, archivo)

        if os.path.exists(ruta_archivo_origen):
            shutil.move(ruta_archivo_origen, ruta_archivo_destino)
            #print(f"Archivo '{archivo}' movido correctamente.")
        else:
            #print(f"Archivo '{archivo}' no encontrado en la ruta de origen.")
            archivos_faltantes.append(archivo)

    archivos_mov_cuenta = os.listdir(ruta_reporte_mov_cuenta)
    archivos_rep_mov = [archivo for archivo in archivos_mov_cuenta if archivo.startswith('Rep_Mov')]

    if archivos_rep_mov:
        for archivo in archivos_rep_mov:
            ruta_archivo_origen = os.path.join(ruta_reporte_mov_cuenta, archivo)
            ruta_archivo_destino = os.path.join(ruta_destino, archivo)
            shutil.move(ruta_archivo_origen, ruta_archivo_destino)
            #print(f"Archivo '{archivo}' movido correctamente.")
    else:
        #print("Ningún archivo 'Rep_Mov' encontrado en la ruta de origen.")
        archivos_faltantes.append("MOVIMIENTOS DE CUENTA.txt")

    return archivos_faltantes

'''
def mover_archivos_de_los_acumulados(date):
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_TIME, 'Spanish_Spain')
        except locale.Error:
            print("No se pudo establecer la localización en español.")
            return
        
    nombre_carpeta = date.strftime('%d %B %Y')

    ruta_origen = 'C:/Users/svcmerr/Desktop'
    ruta_destino = 'C:/Users/svcmerr/Desktop/Acumulados'
    ruta_reporte_mov_cuenta = 'C:/Users/svcmerr/Desktop/Sibamex_80/Sibamex_80'

    ruta_destino_carpeta = os.path.join(ruta_destino, nombre_carpeta)
    os.makedirs(ruta_destino_carpeta, exist_ok=True)
    
    # Crear la carpeta "REPORTES SIBAMEX" dentro de la carpeta creada
    ruta_destino_reportes = os.path.join(ruta_destino_carpeta, 'REPORTES SIBAMEX')
    os.makedirs(ruta_destino_reportes, exist_ok=True)
    
    # Mover los archivos desde ruta_origen a ruta_destino_reportes
    archivos_buscar = ['SPEI-CAPTURADOS.xls', 'SPEI-RECIBIDAS.csv', 'PAGOS Y CAPTACIONES.csv']
    for archivo in archivos_buscar:
        ruta_archivo_origen = os.path.join(ruta_origen, archivo)
        ruta_archivo_destino = os.path.join(ruta_destino_reportes, archivo)
        if os.path.exists(ruta_archivo_origen):
            shutil.move(ruta_archivo_origen, ruta_archivo_destino)
            print(f"Archivo '{archivo}' movido correctamente.")
        else:
            print(f"Archivo '{archivo}' no encontrado en la ruta de origen.")
    
    # Mover el archivo desde ruta_reporte_mov_cuenta si comienza con "Rep_Mov"
    archivos_reporte_mov_cuenta = os.listdir(ruta_reporte_mov_cuenta)
    for archivo in archivos_reporte_mov_cuenta:
        if archivo.startswith('Rep_Mov'):
            ruta_archivo_origen = os.path.join(ruta_reporte_mov_cuenta, archivo)
            ruta_archivo_destino = os.path.join(ruta_destino_reportes, archivo)
            if os.path.exists(ruta_archivo_origen):
                shutil.move(ruta_archivo_origen, ruta_archivo_destino)
                print(f"Archivo '{archivo}' movido correctamente.")

#mover_archivos_de_los_acumulados(date=date(2023, 2, 8))

#ejecutar_sikulix(new_fday_value = date(2023, 2, 8))


def gestionar_descarga_de_archivos_acumulados(fecha_inicio, fecha_fin):
    try:
        fecha_actual = convertir_fecha(fecha_inicio)
        fecha_fin = convertir_fecha(fecha_fin)

        while fecha_actual <= fecha_fin:
            fecha_actual_str = fecha_actual.strftime("%Y-%m-%d")
            fecha_actual_date = datetime.strptime(fecha_actual_str, "%Y-%m-%d").date()
            ejecutar_sikulix(new_fday_value=fecha_actual_date)
            mover_archivos_de_los_acumulados(date=fecha_actual_date)
            fecha_actual += timedelta(days=1)
    except Exception as e:
        print("Error en gestionar_descarga_de_archivos_acumulados:", str(e))

#gestionar_descarga_de_archivos_acumulados(fecha_inicio='07 Febrero 2023', fecha_fin='08 Febrero 2023')
'''

def chequear_descarga_de_archivos():
    ruta_absoluta = os.getcwd()
    ruta_destino = os.path.join(ruta_absoluta, r'file_operacion\REPORTES SIBAMEX')
    archivos_faltantes = []

    archivos_buscar = ['OPERACIONES CAPTURADAS.xls', 'OPERACIONES RECIBIDAS.csv', 'PAGOS Y CAPTACIONES.csv', 'MOVIMIENTOS DE CUENTA OPERATIVA.txt']

    for archivo in archivos_buscar:
        ruta_archivo_destino = os.path.join(ruta_destino, archivo)
        if not os.path.exists(ruta_archivo_destino):
            archivos_faltantes.append(archivo)

    return archivos_faltantes

#chequear_descarga_de_archivos()
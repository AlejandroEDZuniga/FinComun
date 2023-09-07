import shutil
import os
import re
#import descarga_archivos.descarga_desde_mail as mail_operacion

desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
ruta_origen = os.path.join(desktop_path, 'file_download')

current_directory = os.getcwd()
ruta_destino = os.path.join(current_directory, 'file_operacion')

def normalize_and_save_karpay_files():
    ruta_karpay_origen = os.path.join(ruta_origen, 'REPORTES KARPAY')
    ruta_karpay_destino = os.path.join(ruta_destino, 'REPORTES KARPAY')

    if os.path.exists(ruta_karpay_destino):
        shutil.rmtree(ruta_karpay_destino)
    shutil.copytree(ruta_karpay_origen, ruta_karpay_destino)
    archivos = os.listdir(ruta_karpay_destino)
    
    archivos_filtrados = [archivo for archivo in archivos if not (archivo.startswith('A') and archivo.endswith('.xlsx'))]

    # Eliminar los archivos que comienzan con "A" y tienen extensión ".xlsx"
    for archivo in archivos:
        if archivo not in archivos_filtrados:
            ruta_archivo = os.path.join(ruta_karpay_destino, archivo)
            os.remove(ruta_archivo)

    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_karpay_destino, archivo)
        
        nuevo_nombre = None
        
        if re.match(r'EnvioOrdenes.*\.xlsx$', archivo):
            nuevo_nombre = 'Envio Ordenes instancia A.xlsx'
        elif re.match(r'^B.*EnvioOrdenes.*\.xlsx$', archivo):
            nuevo_nombre = 'Envio Ordenes instancia B.xlsx'
        elif re.match(r'OrdenesRecibidas.*\.xlsx$', archivo):
            nuevo_nombre = 'Ordenes Recibidas instancia A.xlsx'
        elif re.match(r'^B.*OrdenesRecibidas.*\.xlsx$', archivo):
            nuevo_nombre = 'Ordenes Recibidas instancia B.xlsx'
        elif re.match(r'OrdenesPendientes.*\.xlsx$', archivo):
            nuevo_nombre = 'Ordenes Pendientes.xlsx'
        
        # Renombrar el archivo si se encontró un nuevo nombre
        if nuevo_nombre:
            nuevo_ruta_archivo = os.path.join(ruta_karpay_destino, nuevo_nombre)
            os.rename(ruta_archivo, nuevo_ruta_archivo)
    
    archivos_actualizados = os.listdir(ruta_karpay_destino)


def normalize_and_save_sibamex_files():
    ruta_sibamex_origen = os.path.join(ruta_origen, 'REPORTES SIBAMEX')
    ruta_sibamex_destino = os.path.join(ruta_destino, 'REPORTES SIBAMEX')
    if os.path.exists(ruta_sibamex_destino):
        shutil.rmtree(ruta_sibamex_destino)
    shutil.copytree(ruta_sibamex_origen, ruta_sibamex_destino)
    archivos = os.listdir(ruta_sibamex_destino)
    for archivo in archivos:
        nombre_archivo_sin_ext, extension = os.path.splitext(archivo)
        # eliminar cualquier número al final
        nombre_archivo_sin_ext = re.sub(r'\s*\d+$', '', nombre_archivo_sin_ext)
        nuevo_nombre_archivo = nombre_archivo_sin_ext + extension
        if nuevo_nombre_archivo.startswith("Rep_Mov"):
            nuevo_nombre_archivo = "MOVIMIENTOS DE CUENTA OPERATIVA" + extension
        elif nuevo_nombre_archivo == "SPEI-RECIBIDAS.csv":
            nuevo_nombre_archivo = "OPERACIONES RECIBIDAS.csv"
        elif nuevo_nombre_archivo == "SPEI-CAPTURADOS.xls":
            nuevo_nombre_archivo = "OPERACIONES CAPTURADAS.xls"
        ruta_archivo_antiguo = os.path.join(ruta_sibamex_destino, archivo)
        ruta_archivo_nuevo = os.path.join(ruta_sibamex_destino, nuevo_nombre_archivo)
        os.rename(ruta_archivo_antiguo, ruta_archivo_nuevo)
        
        # Verificar y eliminar la primera línea si contiene "IF"
        if nuevo_nombre_archivo == "MOVIMIENTOS DE CUENTA OPERATIVA.txt":
            with open(ruta_archivo_nuevo, 'r', encoding='utf-8') as archivo_txt:
                lineas = archivo_txt.readlines()
            
            if lineas and lineas[0].startswith("IF"):
                lineas.pop(0)  # Eliminar la primera línea si comienza con "IF"
            
            with open(ruta_archivo_nuevo, 'w', encoding='utf-8') as archivo_txt:
                archivo_txt.writelines(lineas)
        
    archivos = os.listdir(ruta_sibamex_destino)


def save_flex_folder():
    #mail_operacion.ingresar_al_email(tipo_archivo='flex')

    ruta_flex_origen = os.path.join(ruta_origen, 'REPORTES FLEX')
    ruta_flex_destino = os.path.join(ruta_destino, 'REPORTES FLEX')
    if not os.path.exists(ruta_flex_destino):
        os.makedirs(ruta_flex_destino)
    archivos = os.listdir(ruta_flex_origen)
    archivo_origen = os.path.join(ruta_flex_origen, archivos[0])
    archivo_destino = "Reporte_flex.xls"  # Nuevo nombre del archivo
    ruta_archivo_destino = os.path.join(ruta_flex_destino, archivo_destino)
    shutil.copy2(archivo_origen, ruta_archivo_destino)


#normalize_and_save_karpay_files()
#normalize_and_save_sibamex_files()
import pandas as pd
from datetime import datetime
import os


def flex_conciliacion():
    ruta_absoluta = "C:\\Users\\svcmerr\\Desktop\\pw\\Conciliacion"
    
    df_text = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES FLEX\Reporte_flex.xls', header=None)
    fecha_hora = df_text.iloc[6, 4]  # Obtener el valor de la fila 7, columna E
    fecha, hora = fecha_hora.split()

    # Fecha y hora actual
    fecha_actual = fecha
    hora_actual = hora

    # ________ KARPAY _______________
    df_Envio_Ordenes_A = pd.read_excel(ruta_absoluta + r"\file_operacion\REPORTES KARPAY\Envio Ordenes instancia A.xlsx",engine='openpyxl', header=1)
    df_Envio_Ordenes_A['Instancia'] = 'A'
    df_Envio_Ordenes_B = pd.read_excel(ruta_absoluta + r"\file_operacion\REPORTES KARPAY\Envio Ordenes instancia B.xlsx",engine='openpyxl', header=1)
    df_Envio_Ordenes_B['Instancia'] = 'B'
    df_Ordenes_Pendientes = pd.read_excel(ruta_absoluta + r"\file_operacion\REPORTES KARPAY\Ordenes Pendientes.xlsx",engine='openpyxl', header=1)

    df_Ordenes_Recibidas_A = pd.read_excel(ruta_absoluta + r"\file_operacion\REPORTES KARPAY\Ordenes Recibidas instancia A.xlsx",engine='openpyxl', header=1)
    df_Ordenes_Recibidas_A['Instancia'] = 'A'
    df_Ordenes_Recibidas_B = pd.read_excel(ruta_absoluta + r"\file_operacion\REPORTES KARPAY\Ordenes Recibidas instancia B.xlsx",engine='openpyxl', header=1)
    df_Ordenes_Recibidas_B['Instancia'] = 'B'
    df_Envio_Ordenes = pd.concat([df_Envio_Ordenes_A,df_Envio_Ordenes_B], axis=0, ignore_index=True)
    df_karpay_recibidas = pd.concat([df_Ordenes_Recibidas_A,df_Ordenes_Recibidas_B], axis=0, ignore_index=True)
    # _____________________________ # 

    #_________Flexcube_________ #
    df_Movimientos_Flex = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES FLEX\Reporte_flex.xls', header=4)
    df_flex_filter = df_Movimientos_Flex[df_Movimientos_Flex['DESCRIPCION']=='Transferencia de Cuenta a Cuenta']
    df_flex_credito = df_flex_filter[df_flex_filter['NATURALEZA']=='CREDITO']
    df_flex_debito = df_flex_filter[df_flex_filter['NATURALEZA']=='DEBITO']

    # Defino el nombre del la carpeta y el archivo de cada df
    escritorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    ruta_procesados = os.path.join(escritorio, 'Procesados')
    if not os.path.exists(ruta_procesados):
        os.mkdir(ruta_procesados)

    fecha_hora_actual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    ruta_fecha_hora = os.path.join(ruta_procesados, fecha_hora_actual)
    os.mkdir(ruta_fecha_hora)


    #RECIBIDAS
    df_karpay_recibidas['Cuenta Beneficiario'] = df_karpay_recibidas['Cuenta Beneficiario'].astype(str)
    df_flex_credito['CUENTA'] = df_flex_credito['CUENTA'].astype(str)
    
    df_recibidas = pd.DataFrame({'CUENTA': df_flex_credito['CUENTA'], 'Cuenta Beneficiario': df_karpay_recibidas['Cuenta Beneficiario']})
    archivo_excel = 'Flex_operacion_recibidas.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')
    df_fecha_hora = pd.DataFrame([('Fecha de Operación',fecha_actual),('Hora de Operación',hora_actual)])
    df_recibidas = df_recibidas.dropna(axis=1, how='all')

    df_fecha_hora.to_excel(writer, sheet_name='Flex_operacion_recibidas.xlsx', index=False, startrow=1)
    df_recibidas.to_excel(writer, sheet_name='Flex_operacion_recibidas.xlsx', index=False, startrow=4)

    workbook = writer.book
    worksheet = writer.sheets['Flex_operacion_recibidas.xlsx']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    writer._save()

    coincidencias = df_karpay_recibidas[df_karpay_recibidas['Cuenta Beneficiario'].str.contains('|'.join(df_flex_credito['CUENTA'].tolist()))]

    #CAPTURADAS
    df_Envio_Ordenes_Flex = df_Envio_Ordenes[df_Envio_Ordenes['Medio Entrega']=='FLEXCUBE']
    df_flex_debito.sort_values(by='CANTIDAD')
    df_Envio_Ordenes_Flex = df_Envio_Ordenes_Flex.rename(columns={'Importe': 'CANTIDAD', 'Rastreo':'NUMERO DE TRANSACCION'})
    df_diff = pd.merge(df_flex_debito, df_Envio_Ordenes_Flex, how='outer',indicator=True)
    df_diferencias = df_diff[df_diff['_merge'] == 'left_only']


    #REPORTES

    Total_registro, _ = df_flex_filter.shape
    Total_credito, _ = df_flex_credito.shape
    Total_debito, _ = df_flex_debito.shape
    Total_diferencias, _ = df_diferencias.shape

    # General
    
    df_concepto = pd.DataFrame({'CONCEPTO': ['Total Registros', 'Total CREDITO', 'Total DEBITO'],'CANTIDAD':[Total_registro, Total_credito, Total_debito]})
    archivo_excel = 'Flex_operacion_gral.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')
    df_fecha_hora = pd.DataFrame([('Fecha de Operación',fecha_actual),('Hora de Operación',hora_actual)])

    df_flex_merge = pd.concat([df_flex_credito, df_flex_debito], axis=0, ignore_index=True)
    #Elimino columnas vacias
    df_flex_merge = df_flex_merge.dropna(axis=1, how='all')

    df_fecha_hora.to_excel(writer, sheet_name='Flex_operacion_gral', index=False, startrow=1, header=False)
    df_concepto.to_excel(writer, sheet_name='Flex_operacion_gral', index=False, startrow=4)
    df_flex_merge.to_excel(writer, sheet_name='Flex_operacion_gral', index=False, startrow=9)

    workbook = writer.book
    worksheet = writer.sheets['Flex_operacion_gral']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    writer._save()
    
    # Diferencias
    
    df_concepto = pd.DataFrame({'CONCEPTO':['Total Diferencia'],'Cant':[Total_diferencias]})
    archivo_excel = 'Flex_operacion_diferencias.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    #Elimino columnas vacias
    df_diferencias = df_diferencias.dropna(axis=1, how='all')

    df_fecha_hora.to_excel(writer, sheet_name='Flex_operacion_diferencias', index=False, startrow=1, header=False)
    df_concepto.to_excel(writer, sheet_name='Flex_operacion_diferencias', index=False, startrow=4)
    df_diferencias.to_excel(writer, sheet_name='Flex_operacion_diferencias', index=False,startrow=(7),header=True)

    workbook = writer.book
    worksheet = writer.sheets['Flex_operacion_diferencias']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    writer._save()


#flex_conciliacion()
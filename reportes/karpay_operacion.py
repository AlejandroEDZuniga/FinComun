import pandas as pd
from datetime import datetime
import os
import numpy as np


def convertir_hora(hora_str, formato_original='%I:%M %p', formato_deseado='%H:%M:%S'):
    try:
        hora_dt = datetime.strptime(hora_str, formato_original)
        hora_corregida = hora_dt.strftime(formato_deseado)
        return hora_corregida
    except ValueError as e:
        print(f"Error al convertir la hora: {hora_str}. Detalles: {str(e)}")
        return None


#Abrir Archivos
def open_files(ruta_acumulados=''):
    ruta_absoluta = ruta_absoluta = "C:\\Users\\svcmerr\\Desktop\\pw\\Conciliacion"
    ruta_imagen = os.path.join(ruta_absoluta, 'FC_Logo.png')

    df_text = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES KARPAY\Envio Ordenes instancia B.xlsx',engine='openpyxl', header=None, dtype={"Cuenta Beneficiario": str, "Cuenta Ordenante": str}) 
    time = df_text.at[0,0]
    time = ' '.join(time.split())
    hora_str = time.split('Hora: ')[1][:11]
    fecha_str = time.split('Fecha: ')[1][:10]
    #hora_limit = datetime.strptime(hora_str, '%I:%M:%S %p').time()
    hora_limit = hora_str
    fecha_actual = datetime.strptime(fecha_str, "%d/%m/%Y").date()

    # ________ KARPAY _______________
    df_Envio_Ordenes_A = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES KARPAY\Envio Ordenes instancia A.xlsx',engine='openpyxl', header=1, dtype={"Cuenta Beneficiario": str, "Cuenta Ordenante": str})
    df_Envio_Ordenes_A['Instancia'] = 'A'
    df_Envio_Ordenes_B = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES KARPAY\Envio Ordenes instancia B.xlsx',engine='openpyxl', header=1, dtype={"Cuenta Beneficiario": str, "Cuenta Ordenante": str})
    df_Envio_Ordenes_B['Instancia'] = 'B'
    df_Ordenes_Pendientes = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES KARPAY\Ordenes Pendientes.xlsx',engine='openpyxl', header=1, dtype={"Cuenta Beneficiario": str, "Cuenta Ordenante": str})

    df_Ordenes_Recibidas_A = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES KARPAY\Ordenes Recibidas instancia A.xlsx',engine='openpyxl', header=1, dtype={"Cuenta Beneficiario": str, "Cuenta Ordenante": str})
    df_Ordenes_Recibidas_A['Instancia'] = 'A'
    df_Ordenes_Recibidas_B = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES KARPAY\Ordenes Recibidas instancia B.xlsx',engine='openpyxl', header=1, dtype={"Cuenta Beneficiario": str, "Cuenta Ordenante": str})
    df_Ordenes_Recibidas_B['Instancia'] = 'B'
    df_Envio_Ordenes = pd.concat([df_Envio_Ordenes_A,df_Envio_Ordenes_B], axis=0, ignore_index=True)
    df_karpay_recibidas = pd.concat([df_Ordenes_Recibidas_A,df_Ordenes_Recibidas_B], axis=0, ignore_index=True)
    # _____________________________ # 
    #__________ SIBAMEX ____________
    #Movimiento Operativa
    df_Movimiento_Operativa = pd.read_csv(ruta_absoluta + r'\file_operacion\REPORTES SIBAMEX\MOVIMIENTOS DE CUENTA OPERATIVA.txt',sep='|', header=1)
    df_Movimiento_Operativa = df_Movimiento_Operativa.iloc[1:,1:11]
    df_Estado_Operativa = df_Movimiento_Operativa
    df_Estado_Operativa.columns = ["CUENTA", "FOLIO", "NATURALEZA",'FECHA1','DESCRIPCION','Rastreo', 'CANTIDAD','USUARIO','FECHA2','SUCURSAL']
    df_Estado_Operativa['Rastreo']= df_Estado_Operativa['Rastreo'].str.strip()
    df_Estado_Operativa['NATURALEZA']=df_Estado_Operativa['NATURALEZA'].astype(str).str.strip()
    #df_Estado_Operativa['CANTIDAD'] = pd.to_numeric(df_Estado_Operativa['CANTIDAD'], errors='coerce')
    df_Estado_Operativa['CANTIDAD'] = df_Estado_Operativa['CANTIDAD'].str.replace(',', '').str.replace('$', '').astype(float)

    #Operaciones Capturadas
    df_Operaciones_Capturadas = pd.read_excel(ruta_absoluta + r'\file_operacion\REPORTES SIBAMEX\OPERACIONES CAPTURADAS.xls',engine='xlrd', header=None)
    df_Operaciones_Capturadas = df_Operaciones_Capturadas.iloc[9:,:]
    df_completa_Capturadas = df_Operaciones_Capturadas.iloc[9:,:]
    filter_Op= df_Operaciones_Capturadas[9].isna()
    df_filter = df_Operaciones_Capturadas[~filter_Op]
    df_filter = df_filter.dropna(axis=1, how='all')
    df_filter.columns=['Cuenta','Cliente','# Suc.','Cta Beneficiario','Beneficiario','Banco','# Banco','Usuario','Folio envio','Modulo','Hora Proceso','Status','Transac Sistema','Importe','Comisión','Iva','Total'] 
    df_Operaciones_Capturadas = df_filter
    #Operaciones Recibidas
    df_Operaciones_Recibidas = pd.read_csv(ruta_absoluta + r'\file_operacion\REPORTES SIBAMEX\OPERACIONES RECIBIDAS.csv',header=None, encoding='latin-1')
    df_Operaciones_Recibidas = df_Operaciones_Recibidas.iloc[:,10:20]
    headers_op = ['Sucursal','Cuenta','Cliente','Fecha','Hora','Nom. Ordenante','No. Banco','Banco','Rastreo','Importe']
    df_Operaciones_Recibidas.columns=headers_op
    #Pagos y Captaciones
    df_Pagos_Captaciones = pd.read_csv(ruta_absoluta + r'\file_operacion\REPORTES SIBAMEX\PAGOS Y CAPTACIONES.csv',header=None, encoding='latin-1')
    df_Pagos_Captaciones = df_Pagos_Captaciones.iloc[:,18:33]
    df_filter_pc = df_Pagos_Captaciones.dropna(axis=1, how='all')
    #_________________________________________________________________________
    
    #Capturados
    df_Operativa_2= df_Estado_Operativa[df_Estado_Operativa['NATURALEZA']=='2']
    #df_Operativa_2['CANTIDAD'] = df_Operativa_2['CANTIDAD'].apply(lambda x: float(x.replace(',', '')))
    df_Operativa_2['CANTIDAD'] = pd.to_numeric(df_Operativa_2['CANTIDAD'], errors='coerce')
    df_Operativa_2['DESCRIPCION'] = df_Operativa_2['DESCRIPCION'].str.strip()
    df_Operativa_2_exp = df_Operativa_2[df_Operativa_2['DESCRIPCION'].str.contains('SPEI recibido de')]
    df_Operativa_2 = df_Operativa_2[~df_Operativa_2['DESCRIPCION'].str.contains('SPEI recibido de')]
    df_Operativa_2['DESCRIPCION'] = df_Operativa_2['DESCRIPCION'].str.replace('SPEI enviado a cargo de', '')
    
    df_Envio_Ordenes_Capturadas = df_Envio_Ordenes
    df_KP_SBMX = df_Envio_Ordenes_Capturadas[df_Envio_Ordenes_Capturadas['Medio Entrega'].str.contains('SIBAMEX')]
    df_no_acreditadas = df_KP_SBMX[df_Envio_Ordenes_Capturadas['Tipo Pago'].str.contains('Dev')]
    df_devuelta = df_KP_SBMX[df_Envio_Ordenes_Capturadas['Estado']!='Liquidada']
    df_KP_FLEX = df_Envio_Ordenes_Capturadas[df_Envio_Ordenes_Capturadas['Medio Entrega'].str.contains('FLEXCUBE')]
    df_KP_LOCAL = df_Envio_Ordenes_Capturadas[df_Envio_Ordenes_Capturadas['Medio Entrega'].str.contains('LOCAL')]
    
    #Recibidos
    df_Operativa_1= df_Estado_Operativa[df_Estado_Operativa['NATURALEZA']=='1']
    df_Operativa_1 = df_Operativa_1[~df_Operativa_1['Rastreo'].str.contains('SPEI')]
    df_KP_recibidas = df_karpay_recibidas[~df_karpay_recibidas['Tipo Pago'].str.contains('Dev')]
    df_devueltas = df_KP_recibidas[df_KP_recibidas['Estado']!='Liquidada']
    Suma_Devueltas = df_devueltas['Importe'].sum()
    df_merge = pd.merge(df_Operativa_1,df_KP_recibidas, on=['Rastreo'],how='outer', indicator=True)
    df_merge_right = df_merge[df_merge['_merge'] == 'right_only']
    df_merge_left = df_merge[df_merge['_merge'] == 'left_only']

    Suma_diferencia = df_merge_right['Importe'].sum()

    df_differencias_recibidas_KP = df_KP_recibidas[df_KP_recibidas["Rastreo"].isin(df_merge_right["Rastreo"])].sort_values(['Importe'], ascending=True)
    df_no_conciliadas_A = df_differencias_recibidas_KP[df_differencias_recibidas_KP['Instancia'] == 'A']
    df_no_conciliadas_B = df_differencias_recibidas_KP[df_differencias_recibidas_KP['Instancia'] == 'B']

    Total_no_conciliadas, _ = df_differencias_recibidas_KP.shape
    no_conciliadas_kp_A, _ = df_no_conciliadas_A.shape 
    no_conciliadas_kp_B, _ = df_no_conciliadas_B.shape 

    df_conciliadas_KP_Total = df_merge[df_merge['_merge'] == 'both']
    df_conciliadas_KP = df_KP_recibidas[df_KP_recibidas["Rastreo"].isin(df_conciliadas_KP_Total["Rastreo"])]
    df_conciliadas_KP_A = df_conciliadas_KP[df_conciliadas_KP['Instancia'] == 'A']
    df_conciliadas_KP_B = df_conciliadas_KP[df_conciliadas_KP['Instancia'] == 'B']

    Total_conciliadas, _ = df_conciliadas_KP.shape
    conciliadas_kp_A, _ = df_conciliadas_KP_A.shape 
    conciliadas_kp_B, _ = df_conciliadas_KP_B.shape 

    df_karpay_CKCO = df_conciliadas_KP_Total[['CANTIDAD','Rastreo','Instancia']]
    df_karpay_CKCO.insert(2,'Estatus','Conciliado')
    df_karpay_CKCO['Recibido o Enviado']='Recibido'

    #REPORTES
    # Defino el nombre del la carpeta y el archivo de cada df

    escritorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    ruta_procesados = os.path.join(escritorio, 'Procesados')
    if not os.path.exists(ruta_procesados):
        os.mkdir(ruta_procesados)

    fecha_hora_actual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    ruta_fecha_hora = ''

    if ruta_acumulados:
        ruta_procesados = os.path.join(ruta_procesados, ruta_acumulados)
    
    ruta_fecha_hora = os.path.join(ruta_procesados, fecha_hora_actual)
    if(ruta_fecha_hora):
        os.mkdir(ruta_fecha_hora)

    # Karpay Recibidas para operación ventanilla
    '''
    columnas_deseadas = ['Usuario', 'Fecha Operacion', 'Rastreo', 'Importe', 'Cuenta Beneficiario', 'Instancia']
    df_karpay_recibidas_seleccionado = df_karpay_recibidas[columnas_deseadas]

    archivo_excel = 'karpay_recibidas.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    #df_karpay_recibidas_seleccionado['Cuenta Beneficiario'] = df_karpay_recibidas_seleccionado['Cuenta Beneficiario'].apply(convert_to_str_with_zeros)
    df_karpay_recibidas_seleccionado.to_excel(writer, sheet_name='Karpay_recibidas', index=False,header=True)

    workbook = writer.book
    worksheet = writer.sheets['Karpay_recibidas']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas_seleccionado.columns)-1, width=25)

    writer._save()
    '''

    ##RTERO
    A_Enviadas_rows, A_Enviadas_cols = df_Envio_Ordenes_A.shape
    B_Enviadas_rows, B_Enviadas_cols = df_Envio_Ordenes_B.shape
    A_Recibidas_rows, A_Enviadas_cols = df_Ordenes_Recibidas_A.shape
    B_Recibidas_rows, B_Enviadas_cols = df_Ordenes_Recibidas_B.shape
    Total_Enviadas_rows, Total_Enviadas_cols = df_Envio_Ordenes.shape
    Total_Recibidas_rows, Total_Recibidas_cols = df_karpay_recibidas.shape
    Total_KP = Total_Enviadas_rows + Total_Recibidas_rows
    df_fecha_hora = pd.DataFrame([('Fecha de consulta',fecha_actual),('Hora de consulta',hora_limit)])
    #df_concepto = pd.DataFrame({'CONCEPTO':['Total de registros','Total instancia A','Total instancia B'],'ENVIADAS':[Total_Enviadas_rows,A_Enviadas_rows,B_Enviadas_rows],'RECIBIDAS':[Total_Recibidas_rows,A_Recibidas_rows,B_Recibidas_rows]})
    df_concepto = pd.DataFrame({'CONCEPTO':['Total de registros'],'ENVIADAS':[Total_Enviadas_rows],'RECIBIDAS':[Total_Recibidas_rows]})
    archivo_excel = 'RTERO.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')
    df_Envio_Ordenes = df_Envio_Ordenes.drop(columns=['Empresa'], axis=1)
    df_Envio_Ordenes = df_Envio_Ordenes.rename(columns={'Area': 'Area/Empresa'})

    df_fecha_hora.to_excel(writer, sheet_name='RTERO', index=False, startrow=2, header=False)
    df_concepto.to_excel(writer, sheet_name='RTERO', index=False, startrow=8)
    df_Envio_Ordenes.to_excel(writer, sheet_name='RTERO', index=False,startrow=15, header=True)
    df_karpay_recibidas.to_excel(writer, sheet_name='RTERO', index=False,startrow=Total_Enviadas_rows+16, header=False)

    workbook = writer.book
    worksheet = writer.sheets['RTERO']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'valign': 'vcenter'})
    worksheet.set_row(0, 35)
    worksheet.insert_image('A1', ruta_imagen, {'x_offset': -1, 'y_offset': -6, 'x_scale': 0.35, 'y_scale': 0.25})
    worksheet.write('D1', 'Reporte de transacciones Enviadas y Recibidas por Operación', title_format)

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    writer._save()
    
    #CKCO

    A_Enviadas_rows, A_Enviadas_cols = df_Envio_Ordenes_A.shape
    B_Enviadas_rows, B_Enviadas_cols = df_Envio_Ordenes_B.shape
    A_Recibidas_rows, A_Enviadas_cols = df_Ordenes_Recibidas_A.shape
    B_Recibidas_rows, B_Enviadas_cols = df_Ordenes_Recibidas_B.shape
    Total_Enviadas_rows, Total_Enviadas_cols = df_Envio_Ordenes.shape
    Total_Recibidas_rows, Total_Recibidas_cols = df_karpay_recibidas.shape
    Total_KP = Total_Enviadas_rows + Total_Recibidas_rows
    total_registros_rows, total_registros_cols = df_Estado_Operativa.shape
    cv_rastreo = df_Envio_Ordenes['Rastreo']
    montos = df_Envio_Ordenes['Importe']
    instancia= df_Envio_Ordenes['Instancia']

    df_fecha_hora = pd.DataFrame([('Fecha de Operación',fecha_actual),('Hora de Operación',hora_limit)])
    df_concepto = pd.DataFrame({'CONCEPTO':['Total de registros'],'KARPAY':[Total_KP],'COB':[total_registros_rows]})
    df_resumen = pd.DataFrame({'CLASIFICACIÓN':['Total de registros conciliados','Total de registros no conciliados'], 'A':[conciliadas_kp_A,no_conciliadas_kp_A],'B':[conciliadas_kp_B,no_conciliadas_kp_B]})
    df_rep = pd.DataFrame({'Monto':montos,'Clave de rastreo':cv_rastreo})
    df_rep['Estatus']='Conciliada'
    df_rep['Instancia']=instancia
    df_rep['Recibido o Enviado']='Enviado'

    df_karpay_CKCO = df_karpay_CKCO.rename(columns={'CANTIDAD':'Monto'})
    #df_karpay_CKCO['Monto'] = df_karpay_CKCO['Monto'].apply(lambda x: float(x.replace(',', '')))
    df_karpay_CKCO['Monto'] = pd.to_numeric(df_karpay_CKCO['Monto'], errors='coerce')
    df_karpay_CKCO = df_karpay_CKCO.rename(columns={'Rastreo':'Clave de rastreo'})
    result = pd.concat([df_rep, df_karpay_CKCO], axis=0)
    archivo_excel = 'CKCO.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    df_fecha_hora.to_excel(writer, sheet_name='CKCO', index=False, startrow=2, header=False)
    df_concepto.to_excel(writer, sheet_name='CKCO', index=False, startrow=8)
    df_resumen.to_excel(writer, sheet_name='CKCO', index=False, startrow=10)
    result.to_excel(writer, sheet_name='CKCO', index=False,startrow=(15),header=True)

    workbook = writer.book
    worksheet = writer.sheets['CKCO']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'valign': 'vcenter'})
    worksheet.set_row(0, 35)
    worksheet.insert_image('A1', ruta_imagen, {'x_offset': -1, 'y_offset': -6, 'x_scale': 0.35, 'y_scale': 0.25})
    worksheet.write('C1', 'Conciliación KARPAY Cuenta Operativa', title_format)

    writer._save()


    archivo_excel = 'medio_entrega_Local.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    #Elimino columnas vacias
    df_KP_LOCAL = df_KP_LOCAL.dropna(axis=1, how='all')
    df_KP_LOCAL.to_excel(writer, sheet_name='Local', index=False,header=True)

    workbook = writer.book
    worksheet = writer.sheets['Local']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    writer._save()

    archivo_excel = 'medio_entrega_Flex.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    #Elimino columnas vacias
    df_KP_FLEX = df_KP_FLEX.dropna(axis=1, how='all')
    df_KP_FLEX.to_excel(writer, sheet_name='Flex', index=False,header=True)

    workbook = writer.book
    worksheet = writer.sheets['Flex']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    writer._save()

    

    #RDSERO
    df_fecha_hora = pd.DataFrame([('Fecha de operacion', fecha_actual), ('Hora de operacion', hora_limit)])
    archivo_excel = 'RDSERO.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    df_fecha_hora.to_excel(writer, sheet_name='RDSERO', index=False, startrow=2, header=False)
    df_concepto.to_excel(writer, sheet_name='RDSERO', index=False, startrow=8)

    df_KP_FLEX_LOCAL = pd.concat([df_KP_FLEX, df_KP_LOCAL], axis=0, ignore_index=True)

    df_combined = pd.concat([df_differencias_recibidas_KP, df_KP_FLEX_LOCAL], ignore_index=True)
    df_combined = df_combined.assign(Enviadas_Recibidas=np.where(df_combined.index < len(df_differencias_recibidas_KP), 'Recibida', 'Enviada'))

    df_combined.to_excel(writer, sheet_name='RDSERO', index=False, startrow=(12), header=True)

    workbook = writer.book
    worksheet = writer.sheets['RDSERO']

    # ajustar el ancho de todas las columnas de df_combined
    worksheet.set_column(0, len(df_combined.columns)-1, width=25)

    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'valign': 'vcenter'})
    worksheet.set_row(0, 35)
    worksheet.insert_image('A1', ruta_imagen, {'x_offset': -1, 'y_offset': -6, 'x_scale': 0.35, 'y_scale': 0.25})
    worksheet.write('C1', 'Reporte de Diferencias en SPEI´s Enviadas y Recibidas por Operación', title_format)

    writer._save()



    #Cuenta Operativa

    df_canceladas = df_Estado_Operativa.dropna()
    df_Op_canceladas = df_canceladas[df_canceladas['DESCRIPCION'].str.contains('Cancelacion')]
    df_Op_recibidas = df_Estado_Operativa[df_Estado_Operativa['NATURALEZA'] == '1']
    df_Op_enviadas = df_Estado_Operativa[df_Estado_Operativa['NATURALEZA'] == '2']
    Total_Operativa, _ = df_Estado_Operativa.shape
    Total_Op_recibidas, _ = df_Op_recibidas.shape
    Total_Op_enviadas, _ = df_Op_enviadas.shape
    Total_Op_canceladas,_ = df_Op_canceladas.shape

    #RCOBO
    df_fecha_hora = pd.DataFrame([('Fecha de operacion',fecha_actual),('Hora de operacion',hora_limit)])
    df_total_reg = pd.DataFrame([('RESUMEN', ''),('Total Registro', Total_Operativa)])
    df_concepto = pd.DataFrame({'CONCEPTO':['Total de SPEI recibidos','Total de SPEI envidados','Total de SPEI cancelados', 'Total de SPEI devolucion'],'CANT':[Total_Op_recibidas,Total_Op_enviadas,Total_Op_canceladas, '0']})

    archivo_excel = 'RCOBO.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    df_fecha_hora.to_excel(writer, sheet_name='RCOBO', index=False, startrow=2, header=False)
    df_total_reg.to_excel(writer, sheet_name='RCOBO', index=False, startrow=5,header=False)
    df_concepto.to_excel(writer, sheet_name='RCOBO', index=False, startrow=10)
    df_canceladas.to_excel(writer, sheet_name='RCOBO', index=False,startrow=17, header=True)


    workbook = writer.book
    worksheet = writer.sheets['RCOBO']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'valign': 'vcenter'})
    worksheet.set_row(0, 35)
    worksheet.insert_image('A1', ruta_imagen, {'x_offset': -1, 'y_offset': -6, 'x_scale': 0.35, 'y_scale': 0.25})
    worksheet.write('C1', 'Reporte de la Cuenta Operativa Banxico por Operación', title_format)


    writer._save()

    #Capturadas

    df_Operaciones_Capturadas['Status'] = df_Operaciones_Capturadas['Status'].str.strip()
    df_canceladas = df_Operaciones_Capturadas[df_Operaciones_Capturadas['Status']=='Cancelada']
    df_procesadas = df_Operaciones_Capturadas[df_Operaciones_Capturadas['Status']=='Procesada']
    df_autorizadas = df_Operaciones_Capturadas[df_Operaciones_Capturadas['Status']=='Autorizada']
    filter = df_completa_Capturadas[28] == 'O'
    df_no_autorizadas = df_completa_Capturadas[filter]
    diff_cuenta_operativa = df_no_autorizadas[39].sum()

    df_no_autorizadas = df_no_autorizadas.dropna(axis=1, how='all')

    #Recibidas

    df_Operativa_1['Rastreo'] = df_Operativa_1['Rastreo'].str.strip()
    df_Operaciones_Recibidas['Rastreo'] = df_Operaciones_Recibidas['Rastreo'].str.strip()

    df_merge_sbmx = pd.merge(df_Operativa_1,df_Operaciones_Recibidas, on=['Rastreo'], how='outer', indicator=True)
    df_conciliadas_recibidas_sbmx = df_merge_sbmx[df_merge_sbmx['_merge']=='both']
    df_no_conciliadas_recibidas_sbmx = df_Operaciones_Recibidas[~df_Operaciones_Recibidas['Rastreo'].isin(df_conciliadas_recibidas_sbmx['Rastreo'])]

    #Pagos y Captaciones

    df_filter_pc.columns =['Sucursal','Caja','Cheque','Usuario','Fecha','Hora','Folio','Cuenta 2','Cuenta 3','Cuenta 4','Nombre','Deposito','Retiro','Void','Concepto'] 
    df_filter_pc=df_filter_pc.drop(['Cuenta 3','Cuenta 4','Void'], axis=1)

    df_pc_Retiro=df_filter_pc[df_filter_pc['Deposito']=='$0.00']
    df_sucursales = df_pc_Retiro[(df_pc_Retiro['Sucursal']==1) | (df_pc_Retiro['Sucursal']==2) | (df_pc_Retiro['Sucursal']==77)]
    df_PC_SPEI = df_sucursales[df_sucursales['Concepto'].str.contains('SPEI ENVIADO A CARGO')]


    df_PC_SPEI['Retiro'] = df_PC_SPEI['Retiro'].str.replace(',', '').str.replace('$', '').astype(float)
    df_PC_SPEI['Nombre'] = df_PC_SPEI['Nombre'].str.strip()

    ''''
    hora_limit = hora_limit.strip()

    # Convertir la hora_limit usando el formato especificado
    hora_limit_dt = convertir_hora(hora_limit)

    print('hora limit  es: ---------------------')
    print(hora_limit)
    print('df_PC_SPEI[Hora] es: ---------------------')
    print(df_PC_SPEI['Hora'])
    
    df_PC_SPEI_k = df_PC_SPEI[df_PC_SPEI['Hora'] < hora_limit_dt]
    df_fuera_hora = df_PC_SPEI[df_PC_SPEI['Hora'] > hora_limit_dt]
    df_PC_SPEI= df_PC_SPEI[df_PC_SPEI['Hora'] < hora_limit_dt]

    suma_fuera = df_fuera_hora['Retiro'].sum()
    '''

    df_New_Operativa_2 = df_Operativa_2.rename(columns={"DESCRIPCION": "Nombre", "CANTIDAD": "Retiro", 'FOLIO':'Folio'})
    df_New_Operativa_2['Nombre'] = df_New_Operativa_2['Nombre'].str.strip()
    df_PC_SPEI['Folio'] = df_PC_SPEI['Folio'].astype('int64')
    df_New_Operativa_2['Folio'] = df_New_Operativa_2['Folio'].astype('int64')
    df_merge_pc = pd.merge(df_PC_SPEI, df_New_Operativa_2, on=['Folio'], how='outer', indicator=True)
    df_both = df_merge_pc[df_merge_pc['_merge']=='both']
    df_missing = df_New_Operativa_2.loc[df_New_Operativa_2.index.difference(df_both.index)]
    df_missing_2 = df_New_Operativa_2.loc[df_New_Operativa_2.index.difference(df_missing.index)]
    
    df_diferencias_PC = df_merge_pc[df_merge_pc['_merge']=='right_only']
    df_diferencias_PC = df_diferencias_PC[df_diferencias_PC['Nombre_y'] != 'Abono']

    df_diferencias_PC=df_New_Operativa_2[df_New_Operativa_2['Folio'].isin(df_diferencias_PC['Folio'])]


    #RDSERO2
    Total_diferencias_PC, _ =df_diferencias_PC.shape
    df_fecha_hora = pd.DataFrame([('Fecha de Operación',fecha_actual),('Hora de Operación',hora_limit)])
    df_concepto = pd.DataFrame({'CONCEPTO':['Total de diferencias'],'Cant':[Total_diferencias_PC]})

    archivo_excel = 'RDSERO2.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    df_fecha_hora.to_excel(writer, sheet_name='RDSERO2', index=False, startrow=2, header=False)
    df_concepto.to_excel(writer, sheet_name='RDSERO2', index=False, startrow=6)
    df_diferencias_PC.to_excel(writer, sheet_name='RDSERO2', index=False,startrow=(10),header=True)

    workbook = writer.book
    worksheet = writer.sheets['RDSERO2']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'valign': 'vcenter'})
    worksheet.set_row(0, 35)
    worksheet.insert_image('A1', ruta_imagen, {'x_offset': -1, 'y_offset': -6, 'x_scale': 0.35, 'y_scale': 0.25})
    worksheet.write('C1', 'Reporte de Diferencias en SPEI´s, Enviadas y Recibidas por Operación 2', title_format)

    writer._save()

    #CCOCC

    Total_Registros_PC, _ = df_PC_SPEI.shape
    df_fecha_hora = pd.DataFrame([('Fecha de Operación',fecha_actual),('Hora de Operación',hora_limit)])
    df_concepto = pd.DataFrame({'CONCEPTO':['Total de registros'],'Cant':[Total_Registros_PC]})
    archivo_excel = 'CCOCC.xlsx'
    ruta_archivo = os.path.join(ruta_fecha_hora, archivo_excel)
    writer = pd.ExcelWriter(ruta_archivo, engine='xlsxwriter')

    df_fecha_hora.to_excel(writer, sheet_name='CCOCC', index=False, startrow=2, header=False)
    df_concepto.to_excel(writer, sheet_name='CCOCC', index=False, startrow=6)
    df_PC_SPEI.to_excel(writer, sheet_name='CCOCC', index=False,startrow=(10),header=True)

    workbook = writer.book
    worksheet = writer.sheets['CCOCC']

    # ajustar el ancho de todas las columnas
    worksheet.set_column(0, len(df_karpay_recibidas.columns)-1, width=25)

    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'valign': 'vcenter'})
    worksheet.set_row(0, 35)
    worksheet.insert_image('A1', ruta_imagen, {'x_offset': -1, 'y_offset': -6, 'x_scale': 0.35, 'y_scale': 0.25})
    worksheet.write('C1', 'Conciliación Cuenta Operativa – Cuenta de Cliente', title_format)

    writer._save()

#open_files()
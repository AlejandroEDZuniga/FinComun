from datetime import datetime

def convertir_fecha(fecha_str):
    meses = {
        'Enero': 1,
        'Febrero': 2,
        'Marzo': 3,
        'Abril': 4,
        'Mayo': 5,
        'Junio': 6,
        'Julio': 7,
        'Agosto': 8,
        'Septiembre': 9,
        'Octubre': 10,
        'Noviembre': 11,
        'Diciembre': 12
    }

    partes = fecha_str.split()
    dia = int(partes[0])
    mes = meses[partes[1]]
    anio = int(partes[2])

    fecha = datetime(anio, mes, dia).date()
    return fecha
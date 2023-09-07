import tkinter as tk
import tkinter.messagebox as messagebox
import reportes.flex_operacion as flex
import reportes.karpay_operacion as karpay
from PIL import Image, ImageTk
import reportes.importar_archivos as archivos
import acumulados.acumulados_operacion as acumulados
import descarga_archivos.sibamex_conexion as sibamex
import descarga_archivos.karpay_conexion as karpay
import descarga_archivos.descarga_desde_mail as process_email
from datetime import date
from logs.config_log import *
import traceback
from dotenv import load_dotenv

load_dotenv()

# Admin Credentials
ADMIN_USERNAME = os.getenv('ADMIN_USER')
ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD')
# Josue Credentials
JOSUE_USERNAME = os.getenv('JOSUE_USER')
JOSUE_PASSWORD = os.getenv('JOSUE_PASSWORD')
# Tonatiu Credentials
TONATIU_USERNAME = os.getenv('TONATIU_USER')
TONATIU_PASSWORD = os.getenv('TONATIU_PASSWORD')


def login():
    if  (user_entry.get() == JOSUE_USERNAME and password_entry.get() == JOSUE_PASSWORD) or (user_entry.get() == TONATIU_USERNAME and password_entry.get() == TONATIU_PASSWORD) or (user_entry.get() == ADMIN_USERNAME and password_entry.get() == ADMIN_PASSWORD):
        root.withdraw()  # Oculta la ventana de inicio de sesión
        message_label.config(text="Sesión iniciada correctamente")
        main_window.deiconify()
    else:
        login_logger = configure_logger('login_menu')
        login_logger.error("Error en el inicio de sesión: Usuario o contraseña incorrectos")
        messagebox.showinfo("Error", "Usuario y/o contraseña incorrectos")

def karpay_function():
    try:
        process_email.ingresar_al_email(tipo_archivo='karpay', fecha='15-08-2023')
        karpay.login_to_karpay_A()
        archivos.normalize_and_save_karpay_files()
        messagebox.showinfo("Completado", "Archivos del sistema Karpay descargados correctamente.")
    except Exception as e:
        karpay_normalize_logger = configure_logger('descarga_sist_karpay')
        error_message = f"Ocurrió un error durante la descarga de los archivos del sistema Karpay: {str(e)}"
        karpay_normalize_logger.error(error_message)
        messagebox.showinfo("Error", error_message)

def flex_function():
    try:
        process_email.ingresar_al_email(tipo_archivo='flex', fecha='15-08-2023')
        archivos.save_flex_folder()
        messagebox.showinfo("Completado", "Archivos del sistema Flex descargados correctamente.")
    except Exception as e:
        flex_normalize_logger = configure_logger('normalize_files')
        error_message = f"Ocurrió un error durante la descarga de los archivos del sistema flex: {str(e)}"
        flex_normalize_logger.error(error_message)
        messagebox.showinfo("Error", error_message)

def sibamex_function():
    new_fday_value = date(2023, 1, 9)
    try:
        sibamex.ejecutar_sikulix(new_fday_value)
        archivos_faltantes = sibamex.mover_archivos_reportes()
        archivos.normalize_and_save_sibamex_files()

        if not archivos_faltantes:
            messagebox.showinfo("Completado", "Archivos del Sistema Sibamex descargados correctamente.")
        else:
            archivos_faltantes_str = "\n".join(archivos_faltantes)
            mensaje_error = f"Algunos archivos no se descargaron correctamente:\n{archivos_faltantes_str}"
            messagebox.showinfo("Error", mensaje_error)

    except Exception as e:
        sibamex_normalize_logger = configure_logger('normalize_files', '_sibamex.log')
        error_message = f"Ocurrió un error durante la descarga de los archivos del sistema sibamex:\n{str(e)}"
        traceback_str = traceback.format_exc()
        error_message += f"\n\nTraceback:\n{traceback_str}"
        sibamex_normalize_logger.error(error_message)
        messagebox.showinfo("Error", error_message)

system_functions = {
    "Karpay": karpay_function,
    "Flex": flex_function,
    "Sibamex": sibamex_function
    #"Isi": isi_function
}


def aplicar_function():
    checkboxes = [("Karpay", karpay_var), ("Sibamex", sibamex_var), ("Flex", flex_var), ("Isi", isiLoans_var)]
    selected_systems = []
    for checkbox in checkboxes:
        if checkbox[1].get() == 1:
            selected_systems.append(checkbox[0])
    if len(selected_systems) == 0:
        messagebox.showinfo("Error", "Por favor seleccione al menos una opción.")
    else:
        for system in selected_systems:
            system_functions[system]()


def crear_comparativos_karpay_flex():
    try:
        flex.flex_conciliacion()
        messagebox.showinfo("Completado", "Los comparativos Karpay vs. Flex se encuentran en la carpeta Procesados.")
    except Exception as e:
        karpay_flex_comparativo_logger = configure_logger('comparativos_operacion', '_karpay_flex.log')
        error_message = f"Ocurrió un error durante la descarga de los reportes Karpay vs. Flex: {str(e)}"
        karpay_flex_comparativo_logger.error(error_message)
        messagebox.showinfo("Error", error_message)


def crear_comparativos_karpay_sibamex():
    try:
        karpay.open_files(ruta_acumulados='')
        messagebox.showinfo("Completado", "Los comparativos Karpay vs. Sibamex se encuentran en la carpeta Procesados.")
    except Exception as e:
        karpay_sibamex_comparativo_logger = configure_logger('comparativos_operacion', '_karpay_sibamex.log')
        error_message = f"Ocurrió un error durante la descarga de los reportes Karpay vs. Sibamex: {str(e)}"
        karpay_sibamex_comparativo_logger.error(error_message)
        messagebox.showinfo("Error", error_message)

comparativos_functions = {
    "Sibamex": crear_comparativos_karpay_sibamex,
    "Flex": crear_comparativos_karpay_flex,
    #"Isi": isi_function
}


def aplicar_function_comparativos():
    checkboxes = [("Sibamex", compare_karpay_sibamex_var), ("Flex", compare_karpay_flex_var)]
    selected_comparative = []
    for checkbox in checkboxes:
        if checkbox[1].get() == 1:
            selected_comparative.append(checkbox[0])
    if len(selected_comparative) == 0:
        messagebox.showinfo("Error", "Por favor seleccione al menos una opción.")
    else:
        for system in selected_comparative:
            comparativos_functions[system]()


def crear_reportes_acumulados():
    dia_inicio = dia_entry_start.get()
    mes_inicio = mes_var_start.get()
    año_inicio = anio_entry_start.get()
    fecha_inicio = f"{dia_inicio} {mes_inicio} {año_inicio}"

    dia_fin = dia_entry_end.get()
    mes_fin = mes_var_end.get()
    año_fin = anio_entry_end.get()
    fecha_fin = f"{dia_fin} {mes_fin} {año_fin}"

    # Condicional para verificar el rango que tenenmos actualmente de archivos
    if fecha_inicio not in ['09 Enero 2023', '10 Enero 2023', '11 Enero 2023', '12 Enero 2023', '13 Enero 2023'] or fecha_fin not in ['09 Enero 2023', '10 Enero 2023', '11 Enero 2023', '12 Enero 2023', '13 Enero 2023']:
        messagebox.showinfo("Error", "Las fechas ingresadas se encuentran fuera del alcance.")
    else:
        try:
            #acumulados.buscar_y_unificar_reportes(fecha_inicio, fecha_fin)
            mensaje = f"Los acumulados del {fecha_inicio} al {fecha_fin} se encuentran en la carpeta Procesados."
            messagebox.showinfo("Correcto", mensaje)
        except Exception as e:
            acumulados_creacion_logger = configure_logger('creacion_acumulados', '_creacion_acumulados.log')
            error_message = f"Ocurrió un error durante la creación de los reportes acumulados: {str(e)}"
            acumulados_creacion_logger.error(error_message)
            messagebox.showinfo("Error", error_message)


## DEFINIMOS LA INTERFAZ DEL LOGIN

global_photo = None

root = tk.Tk()
root.title("Automatización de procesos Fincomun")

window_width = 400
window_height = 400

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)

root.geometry(f"{window_width}x{window_height}+{x}+{y}")

current_directory = os.getcwd()
ruta_imagen = os.path.join(current_directory, 'FC_Logo.png')
#ruta_imagen = r'C:\Users\svcmerr\Desktop\pw\Conciliacion\FC_Logo.png'
image = Image.open(ruta_imagen)
image = image.resize((200, 100))
global_photo = ImageTk.PhotoImage(image) 

logo_label = tk.Label(root, image=global_photo)
logo_label.pack(side="top", pady=5)

title_label = tk.Label(root, text="Inicio de sesión", font=("Helvetica", 16))
title_label.pack(pady=(10))

user_label = tk.Label(root, text="Usuario")
user_label.pack()

user_entry = tk.Entry(root, width=30, justify='center')
user_entry.pack(pady=5, padx=10)

password_label = tk.Label(root, text="Contraseña")
password_label.pack()

password_entry = tk.Entry(root, show="*", width=30, justify='center')
password_entry.pack(pady=5, padx=10)

login_button = tk.Button(root, text="Iniciar sesión", command=login)
login_button.pack()

message_label = tk.Label(root, text="")
message_label.pack()


# VENTANA DE OPCIONES
main_window = tk.Toplevel(root)
main_window.title("Área de Operaciones")
window_width = 750
window_height = 650

screen_width = main_window.winfo_screenwidth()
screen_height = main_window.winfo_screenheight()

x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)

main_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

main_window.withdraw()

frame_logo = tk.Frame(main_window)
frame_logo.pack(side="top", fill="x", padx=10, pady=10)

title_label_main = tk.Label(frame_logo, text="Automatización de procesos", font=("Helvetica", 20))
title_label_main.pack(side="bottom", padx=10)

current_directory = os.getcwd()
ruta_imagen = os.path.join(current_directory, 'FC_Logo.png')
#ruta_imagen = r'C:\Users\svcmerr\Desktop\pw\Conciliacion\FC_Logo.png'
image = Image.open(ruta_imagen)
image_resized = image.resize((100, 50))  # cambiar tamaño a 300x150
global_photo_resized = ImageTk.PhotoImage(image_resized)

logo_label_main = tk.Label(frame_logo, image=global_photo_resized)
logo_label_main.pack(side="left")

title_label_main2 = tk.Label(main_window, text="Conciliación del SPEI", font=("Helvetica", 15))
title_label_main2.pack(side="top", padx=10)

title_label_main = tk.Label(main_window, text="1.0 - Descarga de información de los diferentes sistemas", font=("Helvetica", 12))
title_label_main.pack(side="top", padx=10, anchor="w", pady=20)

# Checkboxes de seleccion de los sistemas

karpay_var = tk.IntVar()
frame_sist_karpay = tk.Frame(main_window)
frame_sist_karpay.pack(side="top", padx=48, anchor="w")
tk.Label(frame_sist_karpay, text="1.1 - Sistema Karpay").grid(row=0, column=0, sticky="w")
karpay_button = tk.Checkbutton(frame_sist_karpay, variable=karpay_var)
karpay_button.grid(row=0, column=1, padx=180)

flex_var = tk.IntVar()
frame_sist_flex = tk.Frame(main_window)
frame_sist_flex.pack(side="top", padx=48, anchor="w")
tk.Label(frame_sist_flex, text="1.3 - Sistema Flex").grid(row=0, column=0, sticky="w")
flex_button = tk.Checkbutton(frame_sist_flex, variable=flex_var)
flex_button.grid(row=0, column=1, padx=196)

sibamex_var = tk.IntVar()
frame_sist_sibamex = tk.Frame(main_window)
frame_sist_sibamex.pack(side="top", padx=48, anchor="w")
tk.Label(frame_sist_sibamex, text="1.2 - Sistema Sibamex").grid(row=0, column=0, sticky="w")
sibamex_button = tk.Checkbutton(frame_sist_sibamex, variable=sibamex_var)
sibamex_button.grid(row=0, column=1, padx=172)

isiLoans_var = tk.IntVar()
frame_sist_isi = tk.Frame(main_window)
frame_sist_isi.pack(side="top", padx=48, anchor="w")
tk.Label(frame_sist_isi, text="1.4 - Sistema IsiLoans").grid(row=0, column=0, sticky="w")
isiLoans_button = tk.Checkbutton(frame_sist_isi, variable=isiLoans_var)
isiLoans_button.grid(row=0, column=1, padx=174)

aplicar_button = tk.Button(main_window, text="Aplicar", command=aplicar_function)
aplicar_button.pack(pady=5)

title_label_main = tk.Label(main_window, text="2.0 - Comparativos de Pagos y Aplicación de Diferencias", font=("Helvetica", 12))
title_label_main.pack(side="top", padx=10, anchor="w", pady=20)

# checkboxes de aplicacion de los comparativos

compare_karpay_flex_var = tk.IntVar()
frame1 = tk.Frame(main_window)
frame1.pack(side="top", padx=48, anchor="w")
tk.Label(frame1, text="2.1 - Karpay vs. Flex").grid(row=0, column=0, sticky="w")
compare_button_karpay_flex = tk.Checkbutton(frame1, variable=compare_karpay_flex_var)
compare_button_karpay_flex.grid(row=0, column=1, padx=185)

compare_karpay_sibamex_var = tk.IntVar()
frame2 = tk.Frame(main_window)
frame2.pack(side="top", padx=48, anchor="w")
tk.Label(frame2, text="2.2 - Karpay vs. Sibamex").grid(row=0, column=0, sticky="w")
compare_button_karpay_sibamex = tk.Checkbutton(frame2, variable=compare_karpay_sibamex_var)
compare_button_karpay_sibamex.grid(row=0, column=1, padx=161)

aplicar_comparativos_button = tk.Button(main_window, text="Aplicar", command=aplicar_function_comparativos)
aplicar_comparativos_button.pack(pady=5)

# Seccion Acumulados
title_label_main = tk.Label(main_window, text="3.0 - Descarga de archivos y obtención de Reportes Acumulados", font=("Helvetica", 12))
title_label_main.pack(side="top", padx=10, anchor="w", pady=20)

frame_label = tk.Frame(main_window)
frame_label.pack(side="top", padx=48, anchor="w")
tk.Label(frame_label, text="3.1 - Seleccione las fechas a tener en cuenta para la descarga de los archivos.").pack(side="top")

frame_dates = tk.Frame(main_window)
frame_dates.pack(side="top", padx=48, anchor="w")

frame_start_date = tk.Frame(frame_dates)
frame_start_date.pack(side="left")

tk.Label(frame_start_date, text="Fecha inicial:").pack(side="left")
dia_entry_start = tk.Entry(frame_start_date, width=4)
dia_entry_start.pack(side="left", padx=5)

mes_var_start = tk.StringVar()
mes_var_start.set("Mes")
mes_options = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes_dropdown_start = tk.OptionMenu(frame_start_date, mes_var_start, *mes_options)
mes_dropdown_start.pack(side="left", padx=5)

anio_entry_start = tk.Entry(frame_start_date, width=6)
anio_entry_start.pack(side="left", padx=5)

frame_end_date = tk.Frame(frame_dates)
frame_end_date.pack(side="left")

tk.Label(frame_end_date, text="Fecha final:").pack(side="left")
dia_entry_end = tk.Entry(frame_end_date, width=4)
dia_entry_end.pack(side="left", padx=5)

mes_var_end = tk.StringVar()
mes_var_end.set("Mes")
mes_dropdown_end = tk.OptionMenu(frame_end_date, mes_var_end, *mes_options)
mes_dropdown_end.pack(side="left", padx=5)

anio_entry_end = tk.Entry(frame_end_date, width=8)
anio_entry_end.pack(side="left", padx=5)

obtener_fecha_button = tk.Button(frame_dates, text="Crear Reportes Acumulados", command=crear_reportes_acumulados, width=23)
obtener_fecha_button.pack(side="left", padx=5)

'''
frame_label_acumulados = tk.Frame(main_window)
frame_label_acumulados.pack(side="top", padx=48, anchor="w", pady=5)
tk.Label(frame_label_acumulados, text="3.2 - Obtener reportes acumulados.").pack(side="top")

crear_acumulados_button = tk.Button(frame_label_acumulados, text="Crear Reportes Acumulados", command=crear_reportes_acumulados)
crear_acumulados_button.pack(pady=5)
'''

main_window.mainloop()
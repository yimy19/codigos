from os.path import join,exists
import json

path = GetVar("workfolder_path")

# Nombres de los archivos de configuración y logging de eventos
logfile = "log.txt"

with open(join(path, "config.json"), mode='r', encoding='utf8') as fp:
    data = json.load(fp)

# # Rutas de almacenamiento de los archivos
# config_path = join(path, config_file)
# log_path = join(path, logfile)

# # Verificar si el archivo log existe, sino lo crea.
# try:
#     file = open(log_path, mode='r')
# except FileNotFoundError:
#     file = open(log_path, mode='w')
# finally:
#     file.close()

# Ruta del archivo con las rutas de los reportes a generar en Siga.
reports_paths = join(path, "Rutas.json")
# Rutas de los archivos a descargar
with open(reports_paths, mode='r', encoding='utf8') as file:
    reports = json.load(file)
# Credenciales
with open(join(path, "robot", "credentials.json"), mode='r', encoding='utf8') as file:
    credentials = json.load(file)
# Crear una lista de chequeo para controlar que archivos se han descargado.
checklist = [{'report': report.get('report_name'), 'status': 'NA'}
             for report in reports]

folder = join(path, "email")
if not exists(folder):
    mkdir(folder)


with open(join(folder, data['email']['email_message']), encoding='utf8') as file:
    msg = ""
    for line in file.readlines():
        msg += line

with open(join(folder, data['email']['email_settings']), encoding='utf8') as fp:
    email_settings = json.load(fp)

SetVar("smtp_server", data['email']['smtp_server'])
SetVar("smtp_port", data['email']['smtp_port'])
SetVar("smtp_username", data['email']['smtp_username'])
SetVar("smtp_password", data['email']['smtp_password'])
SetVar("email_settings", email_settings)
SetVar("email_message", msg)

SetVar("reports", reports)
SetVar("credentials", credentials)
SetVar("checklist", checklist)

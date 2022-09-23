from email.mime.text import MIMEText

from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import smtplib
import sys
from os.path import join,exists,isfile,isdir
from os import listdir
import pandas as pd

hora1= '{Hora}'
fechacarpte= '{fechacarpeta}'
path = r"D:\Users\gruporeditos\OneDrive - GANA S.A\GIROS\Auto Reportes Giros\Archivos Cruce Sured"
Path = GetVar("DirReport")
path1= r"D:\Users\gruporeditos\OneDrive - GANA S.A\GIROS\Auto Reportes Giros\Archivos Cruce Sured\Enviar"
NomFile = GetVar("NomFileCruce")
zipfolder = GetVar("zip")
main_file = "GANA "+hora1+".xlsx"
main_file1 = "NOVEDADES HORA "+hora1+".xlsx"
novedades = join(path, main_file1)

zip_path = join(zipfolder, 'Resultado.zip')
archivosadjuntos = []
for file in listdir(path1):
     print(file)
     if not isdir(join(path, file)) and not file.endswith('xlsx'):
         archivosadjuntos.append(file)



# print(df_test)

email_user = 'automata.mercadeo@gruporeditos.com'
email_password = 'Agosto2022*'
recipients = ['danilo.ortiz@gruporeditos.com','daniela.bedoya@gruporeditos.com','gilberto.castrillon@gruporeditos.com','alejandra.olarte@gruporeditos.com','cesar.cabeza@sured.com.co','jorge.abella@sured.com.co','fausto.ramirez@sured.com.co','astrid.correa@gruporeditos.com','alberto.osorio@gruporeditos.com', 'jeniffer.montoya@gruporeditos.com','diana.otalvaro@gruporeditos.com','cesar.naranjo@gruporeditos.com','maria.manjarres@gruporeditos.com'] 
emaillist = [elem.strip().split(',') for elem in recipients]
msg = MIMEMultipart()
msg['Subject'] = 'Novedades  Transaccional Giros'
msg['From'] = 'automata.mercadeo@gruporeditos.com'




body = """
       """

# body = "\n".join([tittle_table,df_test.values(), body])
try:
  with open(join(path, main_file1), mode='rb') as fp:
      df_test = pd.read_excel(fp, sheet_name=None, engine="openpyxl")

except:
  body = """<tittle><strong>
              <br><br> No se presentan novedades en las bases de datos.
              <br>             
            
              </strong></tittle>"""
else:
    for sheet_name, df in df_test.items():
      tittle_table = """<h4><strong>
                          <br><br> &var1 que no están en &var2 <br>
                        </strong></h4>"""
      tittle_table = tittle_table.replace("&var1", sheet_name)
      if "Reditos" in tittle_table:
        tittle_table = tittle_table.replace("&var2", "Matrix")
      else:
        tittle_table = tittle_table.replace("&var2", "Reditos")
      body = body + tittle_table + df.to_html(index=False, classes='mystyle') + "\n"
html = """
<!DOCTYPE html>
<html>
  <head>
    <style>
      table, th, td {{font-size:8pt; border:1px solid black; border-collapse:collapse; text-align:justify; min-width: 200px;}}
      th, td {{padding: 3px;}}
    </style></head>
  <body>
    
    <div>
      A continuación se generan las novedades de las distintas bases de datos,Hora de corte: {1}
    </div>
    {0}
    </div>
    <div>Cualquier duda o inquietud con la información reportada contacta al siguiente correo
              adjuntando la evidencia correspondiente alejandra.olarte@gruporeditos.com.
              <br><br></div>
              <div>
Por favor no responder ni enviar correos de respuesta a 
la cuenta automata.mercadeo@gruporeditos.com.  <br></div>
 <br><br>
  </body>
</html>
""".format(body, hora1)

print(html)

part1 = MIMEText(html, 'html')
msg.attach(part1)



for file in archivosadjuntos:
    # Inserción del consolidado excel en el mensaje.
    with open(join(Path,'Enviar',file), mode='rb') as part:
        excel_file = MIMEBase('application', 'octet-stream')
        excel_file.set_payload(part.read())

    encoders.encode_base64(excel_file)
    excel_file.add_header('Content-Disposition', 'attachment',
                        filename= file)
    msg.attach(excel_file)

server = smtplib.SMTP('smtp.office365.com', 587)
server.starttls()
server.login(email_user,email_password)
server.sendmail(msg['From'], emaillist , msg.as_string())
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

sender_email_address = 'aprendiz.automatiza@gruporeditos.com'
sender_email_password = 'Aut0m4t1z43'
receiver_email_address = 'aprendiz4.automatiza@gruporeditos.com'

email_subject_line = 'python '

msg = MIMEMultipart()
msg['From'] = sender_email_address
msg['To'] = receiver_email_address
msg['Subject'] = email_subject_line

email_body = "Buenas tardes,s archivos"

msg.attach(MIMEText(email_body, 'plain'))

#ruta de archivos
filename = "Historial_Anulaciones.xlsx"
attachment_file = open('C:\\Users\\aprendiz.automatiza\\OneDrive - GANA S.A (1)\\Anulacion_Colillas\\Historial_De_Correos\\Historial_Anulaciones.xlsx','rb')
attachment_file1 = open('C:\\Users\\aprendiz.automatiza\\OneDrive - GANA S.A (1)\\Anulacion_Colillas\\Historial_De_Correos\\Historial_Anulaciones.xlsx','rb')

#archivo 1
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment_file).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment_file; filename = "+filename)
msg.attach(part)



#archivo 2
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment_file1).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment_file; filename = "+filename)
msg.attach(part)

#conexion
email_body_content = msg.as_string()
server = smtplib.SMTP('smtp-mail.outlook.com:587')
server.starttls()
server.login(sender_email_address, sender_email_password)
server.sendmail(sender_email_address, receiver_email_address, email_body_content)
server.quit()
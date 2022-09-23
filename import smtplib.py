import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

sender_email_address = 'aprendiz.automatiza@gruporeditos.com'
sender_email_password = 'Aut0m4t1z4'
receiver_email_address = 'aprendiz.automatiza@gruporeditos.com'

email_subject_line = 'subject'

msg = MIMEMultipart()
msg['From'] = sender_email_address
msg['To'] = receiver_email_address
msg['Subject'] = email_subject_line

email_body = 'Hello World. This is Python email sender application with Attachments.'
msg.attach(MIMEText(email_body, 'plain'))

filename = "texto.txt"
attachment_file = open('C:\Users\aprendiz.automatiza\Desktop\texto.txt')
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment_file).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment_file; filename = "+filename)

msg.attach(part)

email_body_content = msg.as_string()
server = smtplib.SMTP('smtp-mail.outlook.com:587')
server.starttls()
server.login(sender_email_address, sender_email_password)
server.sendmail(sender_email_address, receiver_email_address, email_body_content)
server.quit()
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

username = "aprendiz.automatiza@gruporeditos.com"
password = "Aut0m4t1z4"
mail_from = "aprendiz.automatiza@gruporeditos.com"
mail_to = "aprendiz.automatiza@gruporeditos.com"
mail_subject = "Test Subject"
mail_body = "This is a test message"
mail_attachment="C:\Users\aprendiz.automatiza\Desktop\hola.txt"
mail_attachment_name="hola.txt"

mimemsg = MIMEMultipart()
mimemsg['From']=mail_from
mimemsg['To']=mail_to
mimemsg['Subject']=mail_subject
mimemsg.attach(MIMEText(mail_body, 'plain'))

with open(mail_attachment, "rb") as attachment:
    mimefile = MIMEBase('application', 'octet-stream')
    mimefile.set_payload((attachment).read())
    encoders.encode_base64(mimefile)
    mimefile.add_header('Content-Disposition', "attachment; filename= %s" % mail_attachment_name)
    mimemsg.attach(mimefile)
    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(username,password)
    connection.send_message(mimemsg)
    connection.quit()

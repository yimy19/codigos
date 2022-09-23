import win32com.client as win32
import warnings
import sys
import pythoncom
 
reload(sys)
sys.setdefaultencoding('utf8')
warnings.filterwarnings('ignore')
pythoncom.CoInitialize()
def sendmail():
    sub = 'outlook python mail test'
    body = 'my test\r\n my python mail'
    outlook = win32.Dispatch('outlook.application')
    receivers = ['xxx']
    mail = outlook.CreateItem(0)
    mail.To = receivers[0]
    mail.Subject = sub.decode('utf-8')
    mail.Body = body.decode('utf-8')
    mail.Attachments.Add('C:\Users\aprendiz.automatiza\Desktop\alertas\BAJO CAUCA\Alertas_Zona_BAJO CAUCA_21-02-2022.xlsx)
    mail.Send()
 
sendmail()
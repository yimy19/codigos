import pandas as pd
from openpyxl import workbook

Control = pd.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\alerta.xlsx")
    #Data = pd.DataFrame(Control)
filtro= ['Jineteo']
#filtrado= Control['Alerta']
filtrado= Control[Control['Alerta'].isin(filtro)]

print(filtrado)
print(Control)


Control1 = pd.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\alerta.xlsx",sheet_name='CARTERA INDIRECTA')
    #Data = pd.DataFrame(Control)
#filtro= ['Jineteo']
filtrado1= Control1['PRERROGATIVA']
#filtrado= Control1[Control1['Alerta'].isin(filtro)]

print(filtrado1)
print(Control1)

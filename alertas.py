import pandas as pd
from openpyxl import workbook



productos = ['BETPLAY','CARTERA INDIRECTA',

'CARTERA DIRECTA',

'ADULTO MAYOR',

'LINEA BUSES',

'RECARGAS GANA',

'INGRESO SOLIDARIO',

'GIROS',

'ANULACIÓN GIROS',

'ANULACIÓN EPM',

'RIS']


for x in productos:
    
    Control = pd.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\alerta.xlsx",sheet_name=x)
    #Data = pd.DataFrame(Control)
    #filtrado= Control['Zona']
    if(x=='BETPLAY'):
        filtro= ['SUR']
        filtrado= Control[Control['Zona'].isin(filtro)]
        #filtrado= Control['Zona']
        print(filtrado)
        

    if(x=='LINEA BUSES'):
        #filtro= ['SUR']
        #filtrado= Control[Control['Zona'].isin(filtro)]
        filtrado1= Control['Corte']
        print(filtrado1)

print("______________________________________")

print('|              '+x+'          |')

print("______________________________________")

print("----------------------------------------------------------------------")

print(Control)

import pandas
from openpyxl import workbook
  
data = pandas.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.xls") 
dfdata =  pandas.DataFrame(data)
print (len(dfdata))
str.encode('UTF-8')
nuevo = pandas.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.xlsx")
dfdata1 =  pandas.DataFrame(nuevo)
print (dfdata1)
print (len(dfdata1))


historico = pandas.concat([dfdata1,dfdata]) 
historico.to_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.xlsx")
print (len(historico))
wb = load_workbook(r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.xlsx")
hoja = wb.active
hoja.delete_cols(1)
wb.save(r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.xlsx")

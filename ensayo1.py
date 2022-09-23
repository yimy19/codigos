import pandas
from openpyxl import workbook

r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.xls"
r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.csv"
r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.csv"

exc = pandas.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.xls",header=0)
dfExc = pandas.DataFrame(exc)

dfExc.to_csv(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.csv",sep=';',index=None)

f = open(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.csv",mode='r',encoding='utf-8')

fRead = f.read()
fRead = fRead.replace('N\x1a','Ñ')
fRead = fRead.replace('E\xc3\xad','É')
fRead = fRead.replace('I\xc3\xad','Í')
fRead = fRead.replace('O\xc3\xad','Ó')


f.close()

ff = open(r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.csv",mode='w',encoding='latin1')
ff.write(fRead)

ff.close()
pandas.read_csv(r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.csv",encoding='latin1').to_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos.xls", index=False, engine='xlwt')

#pandas.read_csv(r"C:\Users\aprendiz.automatiza\Desktop\reportehistorico.csv",sep=';',encoding='latin1').to_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos.xls",encoding='latin1')

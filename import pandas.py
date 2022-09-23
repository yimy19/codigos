import pandas 
from openpyxl import load_workbook
  
  
data = pandas.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporteGirosPagados.xls") 
  
filtro= ['900039533']

#AlertOpen = data.loc[data['Documento Remi.'] ==filtro]
filtrado= data[data['Documento'].isin(filtro)]
print(filtrado)
filtrado.to_excel(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporte1.xls", index=None)


data1 = pandas.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporteGirosPagados.xls") 
  
filtro1= ['9000395338']

#AlertOpen = data.loc[data['Documento Remi.'] ==filtro]
filtrado1= data[data['Documento'].isin(filtro1)]
print(filtrado1)
filtrado1.to_excel(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporte2.xls", index=None)




#concatenar
#ANULACIÓN GIROS
data1 = pandas.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporte2.xls") 
dfdata3 =  pandas.DataFrame(data1)
nuevo = pandas.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporte1.xls")
dfdata2 =  pandas.DataFrame(nuevo) 
historico1 = pandas.concat([dfdata2,dfdata3]) 
historico1.to_excel(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporte.xls", index=None)
wb = load_workbook(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporte.xls")
hoja = wb.active
hoja.delete_cols(1)
wb.save(r"C:\Users\aprendiz.automatiza\Desktop\bot subsidios del estado\reporte.xls")
wb.close()

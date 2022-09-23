import pandas as pd

pd.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.xls").to_csv(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.csv",sep=';', index=False,encoding='utf-16')



pd.read_csv(r"C:\Users\aprendiz.automatiza\Desktop\reportePagos20.csv",sep=';',encoding='utf-16').to_excel(r"C:\Users\aprendiz.automatiza\Desktop\reportePagoshi.xls",index=None)




# -*- coding: utf-8 -*-
 
import pandas as pd
import sys

 
read_data = pd.read_excel (r'C:\Users\aprendiz.automatiza\Desktop\simbolo.xlsx')
print(read_data)
read_data = read_data [~ read_data ['hola']. str.contains ('Unknown')]
print(read_data)
read_data.to_excel (r'C:\Users\aprendiz.automatiza\Desktop\data_al.xlsx', index = False) 

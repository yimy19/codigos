import pandas as pd


data = pd.read_csv(r'C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\RetirosEingresos\contacts.csv', sep=',', engine='python')
print(data)
data.to_excel(r'C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\RetirosEingresos\contacts.xlsx', index=None)
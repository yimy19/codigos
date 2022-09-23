import csv
import openpyxl

input_file = 'C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\Anulacion_Colillas\Historial_De_Correos\Mensaje_Anu_2022-03-28_07_28_.txt'
output_file = 'C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\Anulacion_Colillas\FF.xlsx'

wb = openpyxl.Workbook()
ws = wb.worksheets[0]

with open(input_file, 'r') as data:
    reader = csv.reader(data, delimiter='\t')
    for row in reader:
        ws.append(row)

wb.save(output_file)
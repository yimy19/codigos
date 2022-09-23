import csv

txt_file = r"C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\Anulacion_Colillas\Historial_De_Correos\Mensaje_Anu_2022-03-28_07_28_.txt"
csv_file = r"C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\Anulacion_Colillas\FF.csv"

in_txt = csv.reader(open(txt_file, "r"), delimiter = "\n")
out_csv = csv.writer(open(csv_file, 'w'))

out_csv.writerows(in_txt)

del out_csv

print(in_txt)
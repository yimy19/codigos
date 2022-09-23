from os.path import join
from numpy import empty
import pandas as pd
from datetime import datetime



hora1= '{Hora}'
fechacarpte= '{fechacarpeta}'
listahojas = {nombreHojaHora}
#ruta2 = r"D:\Users\gruporeditos\OneDrive - GANA S.A\GIROS\Auto Reportes Giros\Archivos Cruce Sured\Enviar"
ruta1 = r"D:\Users\gruporeditos\OneDrive - GANA S.A\GIROS\reports"
path = r"D:\Users\gruporeditos\OneDrive - GANA S.A\GIROS\Auto Reportes Giros\Archivos Cruce Sured"
path1 = r""+ruta1+"\\"+fechacarpte+""



reports = listahojas

name_report = [ "Anulados Matrix",
                "Enviados Matrix",
                "Pagados Matrix",                
                "Anulados Reditos",
                "Enviados Reditos",
                "Pagados Reditos"]
col_filter = [  "Fecha Confirmación",#columna de anulados r
                "Fecha Giro",#columnas de enviados r
                "Fecha Pago",#columnas de pagados r
                "HORA_ANULACION",#hora anulacion M
                "HORA_IMPOSICION",#hora enviados M 
                "HORA_PAGO"#hora pagadas M
                
             ]
global fecha_corte
hora2 =int(hora1)
if hora2%2:   
	hora2 -=1
hora1 = str(hora2) 
SetVar('Hora', hora1)
hora2 =int(hora1)
if hora2 == int(8):
  print(hora2)  
  hora2 = '08'
  print(hora2) 
hora1 = str(hora2)
print(hora1) 

SetVar('Hora', hora1)
fecha_corte = datetime.strptime(""+hora1+":00:00", '%H:%M:%S')
main_file = "GANA "+hora1+".xls"
main_file1 = "NOVEDADES HORA "+hora1+".xls"
new_main_file = main_file.replace("xls", "xlsx")
new_main_file1 = main_file1.replace("xls", "xlsx")
SetVar('NomFileCruce',new_main_file1)
new_df = {}
pines_novedad = {}
print(hora1)

def numSheet(num):
    if num == 0:
        num_sheet = 2
    elif num == 1:
        num_sheet = 0
    else:
        num_sheet = 1
    return num_sheet


def diferenciasPin(df_1, df_2):
    novedades = []
    for pin in df_2.index:
        if not pin in df_1.index:
            novedades.append(pin)
    return novedades

def date_filter(fecha):

    if isinstance(fecha, str):

        fecha = fecha.split()[-1]

        fecha = datetime.strptime(fecha, '%H:%M:%S')

        return fecha < fecha_corte

 
for num, sheet in enumerate(reports):
    
    print(num, sheet)
    num_sheet = numSheet(num)
    with open(join(path, main_file), mode='rb') as fp:
        df_matrix = pd.read_excel(fp, sheet_name=num_sheet, index_col="PIN_GIRO")
        df_matrix.index = df_matrix.index.astype(str)
        print("with 1")

    with open(join(path1, sheet), mode='rb') as fp:
        df_reditos = pd.read_excel(fp, index_col="PIN")
        df_reditos.index = df_reditos.index.astype(str)
        print("with 2")
    if len(df_matrix):   
      pin_reditos = diferenciasPin(df_matrix, df_reditos)
      print("diferencia")

      print(len(pin_reditos))
      if len(pin_reditos) > 0:
          df_filtered = df_reditos.loc[pin_reditos]
          bool = df_filtered[col_filter[num]].apply(date_filter)
          bool = bool[bool[pin_reditos] == True]
          if not bool.empty:
              new_df[name_report[num+3]] = df_reditos.loc[bool.index]
              pines_novedad[name_report[num+3]] = bool.index
              print(df_reditos.loc[bool.index])       
    if len(df_reditos):
      pin_matrix = diferenciasPin(df_reditos, df_matrix)
      print("diferencia 2")

      print(len(pin_matrix))
      if len(pin_matrix) > 0:
          df_filtered = df_matrix.loc[pin_matrix]
          bool = df_filtered[col_filter[num+3]].apply(date_filter)
          bool = bool[bool[pin_matrix] == True]
          if not bool.empty:
              new_df[name_report[num]] = df_matrix.loc[bool.index]
              pines_novedad[name_report[num]] = bool.index
              print(df_matrix.loc[bool.index])           


print(new_df)
if len(new_df):
  with pd.ExcelWriter(join(path, new_main_file1), engine = 'openpyxl') as wr:
      for sheet_names, new_df in new_df.items():        
          new_df.to_excel(wr, sheet_name = sheet_names)
        
        
        

        
from os.path import join
from numpy import empty
import pandas as pd
from datetime import datetime

path = r"C:\Users\John Eduard\Documents\Trabajo\BlackSmith Reasearch\RPA\Reporte Giros\Reportes"
main_file = "Cruce Matrix JUL 03 (8 PM).xls"
reports = [ "reporteGirosEnviadosEstados03-Jul-2022-20-02.xls",
            "reporteGirosPagados03-Jul-2022-20-02.xls",
            "reporteGirosAnulados03-Jul-2022-20-02.xls"]
# reports = [ "reporteGirosEnviadosEstados05-Jul-2022-18-03.xls" ]
name_report = [ "Enviados Matrix",
                "Pagados Matrix",
                "Anulados Matrix",
                "Enviados Reditos",
                "Pagados Reditos",
                "Anulados Reditos"]
col_filter = [  "Fecha Giro",
                "HORA_IMPOSICION",
                "HORA_PAGO",
                "HORA_ANULACION"]

fecha_corte = datetime.strptime("20:00:00", '%H:%M:%S')
new_main_file = main_file.replace("xls", "xlsx")
new_df = {}
pines_novedad = {}
def diferenciasPin(df_1, df_2):
    novedades = []
    for pin in df_2.index:
        if not pin in df_1.index:
            novedades.append(pin)
    return novedades

def date_filter(fecha):
    fecha = fecha.split()[1]
    fecha = datetime.strptime(fecha, '%H:%M:%S')
    #print(fecha, "\t", fecha_corte)
    return fecha < fecha_corte

def date_filter2(fecha):
    #fecha = fecha.split()[1]
    fecha = datetime.strptime(fecha, '%H:%M:%S')
    #print(fecha, "\t", fecha_corte)
    return fecha < fecha_corte

for num, sheet in enumerate(reports):
    with open(join(path, main_file), mode='rb') as fp:
        df_matrix = pd.read_excel(fp, sheet_name=num, index_col="PIN_GIRO")
        df_matrix.index = df_matrix.index.astype(str)

    with open(join(path, sheet), mode='rb') as fp:
        df_reditos = pd.read_excel(fp, index_col="PIN")
        df_reditos.index = df_reditos.index.astype(str)

    pin_reditos = diferenciasPin(df_matrix, df_reditos)
    if len(pin_reditos) > 0:
        df_filtered = df_reditos.loc[pin_reditos]
        bool = df_filtered[col_filter[0]].apply(date_filter)
        bool = bool[bool[pin_reditos] == True]
        if not bool.empty:
            new_df[name_report[num+3]] = df_reditos.loc[bool.index]
            pines_novedad[name_report[num+3]] = bool.index
            print(df_reditos.loc[bool.index])
            #print(f"novedad {name_report[num+3]}: {len(bool.index)}")

    pin_matrix = diferenciasPin(df_reditos, df_matrix)
    if len(pin_matrix) > 0:
        df_filtered = df_matrix.loc[pin_matrix]
        bool = df_filtered[col_filter[num+1]].apply(date_filter2)
        bool = bool[bool[pin_matrix] == True]
        if not bool.empty:
            new_df[name_report[num]] = df_reditos.loc[bool.index]
            pines_novedad[name_report[num]] = bool.index
            print(df_matrix.loc[bool.index])
            #print(f"novedad {name_report[num]}: {len(bool.index)}")


print(new_df)

with pd.ExcelWriter(join(path, new_main_file), engine = 'openpyxl') as wr:
    for sheet_names, new_df in new_df.items():
        new_df.to_excel(wr, sheet_name = sheet_names)

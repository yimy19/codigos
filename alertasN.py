import pandas as Pd
from openpyxl import workbook

Productos = ['BETPLAY',
'CARTERA INDIRECTA',
'CARTERA DIRECTA',
'ADULTO MAYOR',
'LINEA BUSES',
'RECARGAS GANA',
'INGRESO SOLIDARIO',
'GIROS',
'ANULACIÓN GIROS',
'ANULACIÓN EPM',
'RIS']

DirectorioRegionales = [
 ['SUR','danilo.ortiz@gruporeditos.com'],
 ['NORTE', 'aprendiz.automatiza@gruporeditos.com'], 
 ['MAGDALENA MEDIO', 'danilo.ortiz@gruporeditos.com'], 
 ['NORDESTE', 'aprendiz.automatiza@gruporeditos.com'], 
 ['OCCIDENTE', 'danilo.ortiz@gruporeditos.com'], 
 ['CENTRO OCCIDENTAL','aprendiz.automatiza@gruporeditos.com'], 
 ['URABA', 'danilo.ortiz@gruporeditos.com'], 
 ['CENTRO', 'aprendiz.automatiza@gruporeditos.com'], 
 ['BAJO CAUCA', 'gilberto.castrillon@gruporeditos.com'], 
 ['SUROESTE', 'aprendiz.automatiza@gruporeditos.com'], 
 ['CENTRO ORIENTAL','gilberto.castrillon@gruporeditos.com'], 
 ['ORIENTE','danilo.ortiz@gruporeditos.com'],
 ['Linea de buses','danilo.ortiz@gruporeditos.com']
 ]



for w in range(len(DirectorioRegionales)):
    for y in  range(len(DirectorioRegionales[w])):

        Zona = DirectorioRegionales[w][y]
        NombreArchivo = 'Alertas_Zona_'+ Zona + '.xlsx'
        break
            

    writer = Pd.ExcelWriter(NombreArchivo)

    for x in Productos:
        
        FileAlertasControl = Pd.read_excel(r"C:\Users\aprendiz.automatiza\Desktop\Consolidado automatizacion de alertas.xlsx",sheet_name= x,header=0)
        
        Data = Pd.DataFrame(FileAlertasControl)
        
        if x == 'BETPLAY':
            
            Filter = Data[['Consecutivo alerta',
                            'Priorizacion',
                            'Estado',
                            'Fecha Seguimiento',
                            'Cedula Vendedor',
                            'Oficina',
                            'Zona',
                            'Producto',
                            'Valor alerta',
                            'Alerta']]
            
            Data = Pd.DataFrame(Filter)
            ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
            if len(ZonaDf) > 0:   
                print(x, Zona) 
                print("-----------------") 
                print(ZonaDf)    
                if len(Filter) > 0:    
                    ZonaDf.to_excel(writer,sheet_name=x,index = False)
                                
        elif x == 'CARTERA INDIRECTA':
            
            Filter = Data[['Consecutivo alerta',
                        'Priorizacion',
                        'ESTADO',
                        'FECHA ALERTA',
                        'CEDULA','NOMBRE',
                        'CARGO',
                        'Zona',
                        'Oficina',
                        'cartera al dia de la alerta',
                        'ALERTA',
                        'Estado vendero',
                        'Reindicide',
                        'PRERROGATIVA']] 
            
            Data = Pd.DataFrame(Filter)
            ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
            
            if len(ZonaDf) > 0:   
                print(x,Zona)
                print("-----------------") 
                print(ZonaDf)   
                if len(Filter) > 0: 
                    ZonaDf.to_excel(writer,sheet_name=x,index = False)           
        elif x == 'CARTERA DIRECTA':
            
            Filter = Data[['Consecutivo alerta',
                        'Priorizacion',
                        'ESTADO',
                        'FECHA ALERTA',
                        'CEDULA',
                        'NOMBRE',
                        'CARGO',
                        'Zona',
                        'Oficina',
                        'Reindicide',
                        'OBSERVACION RETIRO',
                        'PRERROGATIVA']]
            
            Data = Pd.DataFrame(Filter)
            ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
            if len(ZonaDf) > 0:   
                print(x,Zona)
                print("-----------------") 
                print(ZonaDf)   
            
                if len(Filter) > 0:  
                    ZonaDf.to_excel(writer,sheet_name=x,index = False)        
        elif x == 'ADULTO MAYOR':
                    Filter = Data[['Consecutivo alerta',
                        'Priorizacion',
                        'FECHA ALERTA',
                        'CEDULA',
                        'Oficina',
                        'Zona',
                        'cartera al dia de la alerta',
                        'ALERTA',
                        'Estado vendero']]
        
                    Data = Pd.DataFrame(Filter)
                    ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
                    if len(ZonaDf) > 0:   
                        print(x, Zona) 
                        print("-----------------") 
                        print(ZonaDf)    
                        if len(Filter) > 0:    
                            ZonaDf.to_excel(writer,sheet_name=x,index = False)   
        elif x == 'LINEA BUSES':
                    Filter = Data[['Consecutivo alerta',
                    'Priorizacion',
                    #'Estado',
                    #'Seguimiento',
                    'Cedula',
                    'Oficina',
                    'Zona',
                    'valor cancelaciones',
                    'Estado vendedor']]

                    Data = Pd.DataFrame(Filter)
                    ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
                    if len(ZonaDf) > 0:   
                        print(x, Zona) 
                        print("-----------------") 
                        print(ZonaDf)    
                        if len(Filter) > 0:    
                            ZonaDf.to_excel(writer,sheet_name=x,index = False)
        elif x == 'RECARGAS GANA':
            
                Filter = Data[['Consecutivo alerta',
                'Priorizacion',
                #'Estado',
                'Fecha de Señal',
                'Cedula',
                'Oficina',
                'Zona',
                'Estado de vendedor',
                'Observaciones',
                'Tipo de Señal']]

                Data = Pd.DataFrame(Filter)
                ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
                if len(ZonaDf) > 0:   
                    print(x, Zona) 
                    print("-----------------") 
                    print(ZonaDf)    
                    if len(Filter) > 0:    
                        ZonaDf.to_excel(writer,sheet_name=x,index = False)
        elif x == 'INGRESO SOLIDARIO':
                Filter = Data[['Consecutivo alerta',	 
                'Priorizacion',
                'ESTADO',
                'FECHA ALERTA',
                'CEDULA',
                'OFICINA',
                'Zona',
                'VALOR ALERTA',
                'ALERTA']]
            
                Data = Pd.DataFrame(Filter)
                ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
                if len(ZonaDf) > 0:   
                    print(x, Zona) 
                    print("-----------------") 
                    print(ZonaDf)    
                    if len(Filter) > 0:    
                        ZonaDf.to_excel(writer,sheet_name=x,index = False)
        elif x == 'GIROS':
                Filter = Data[['Consecutivo alerta',
                'Priorizacion',
                'Estado',
                #'Fecha Seguimiento',
                'Cedula Vendedor',
                'Oficina',
                'Zona',
                'Producto',
                'Valor alerta',
                'Alerta']]
        
                Data = Pd.DataFrame(Filter)
                ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
                if len(ZonaDf) > 0:   
                    print(x, Zona) 
                    print("-----------------") 
                    print(ZonaDf)    
                    if len(Filter) > 0:    
                        ZonaDf.to_excel(writer,sheet_name=x,index = False)
        elif x == 'ANULACIÓN GIROS':
                Filter = Data[['Consecutivo alerta',
                'Priorizacion',
                'Estado',
                'Cedula Vendedor',
                'Oficina',
                'Zona',
                'Producto',
                'Valor alerta',
                'Alerta']]
        
                Data = Pd.DataFrame(Filter)
                ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
                if len(ZonaDf) > 0:   
                    print(x, Zona) 
                    print("-----------------") 
                    print(ZonaDf)    
                    if len(Filter) > 0:    
                        ZonaDf.to_excel(writer,sheet_name=x,index = False)
        elif x == 'RIS':
            
                Filter = Data[['Consecutivo alerta',
                'Priorizacion',
                'Estado',
                
                'Cedula Vendedor',
                'Oficina',
                'Zona',
                'Producto',
                'Valor alerta',
                'Alerta']]
        
                Data = Pd.DataFrame(Filter)
                ZonaDf = Data.loc[Data['Zona'] == str(Zona)]
                if len(ZonaDf) > 0:   
                    print(x, Zona) 
                    print("-----------------") 
                    print(ZonaDf)    
                    if len(Filter) > 0:    
                        ZonaDf.to_excel(writer,sheet_name=x,index = False)
    
        

    writer.save()
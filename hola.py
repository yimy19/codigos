import pandas as pd   

dataframe1 = pd.read_csv("C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\Anulacion_Colillas\Historial_De_Correos\Mensaje_Anu_2022-03-28_07_28_.txt") 
  
dataframe1.to_csv("C:\Users\aprendiz.automatiza\OneDrive - GANA S.A (1)\Anulacion_Colillas\FF.csv", index = None)
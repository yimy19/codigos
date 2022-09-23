import pandas as pd
  

sr = pd.Series(['ALONSO DEJESUS GALVIS CASTANO'])
  
print(sr)

result = sr.str.decode(encoding ='utf-16-le')

print(result)



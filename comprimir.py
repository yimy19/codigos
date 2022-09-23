from os import listdir
from os.path import join, isdir, relpath, exists, isfile
import datetime as dt
import zipfile

path = r"C:\Users\aprendiz.automatiza\Desktop\PRUEBA\configuracion\historico"


folder = path



files = []
for file in listdir(folder):
    if not isdir(join(folder, file)) and file.endswith('png'):
        files.append(file)

zip_path = join(folder, 'Resultado.zip')

if not isfile(zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zfp:
        for file in files:
                        
            zfp.write(filename=join(folder, file), arcname=file)
    print("Zip file created.")
else:
    print("Zip file already exists.")



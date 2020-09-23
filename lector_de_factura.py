# Librerías necesarias
import pandas as pd
import collections
from datetime import datetime

# Lectura de la BD Manifiestos
URL = 'BASE DATOS MANIFIESTOS.xlsx'
form_df = pd.read_excel(URL)

# Ubicación del la factura 1719 1887 2162 2169
URL = '2169.xlsx'
data = pd.read_excel(URL)

# Elegir columna de referencias y eliminar nulos
data = data["SAGUI IMPORT SAS"]
data = data.dropna()
# Convertir referencias en una lista
items = data.tolist()

rows = []
# Obtenemos las filas para cada referencia
for i in range(0, len(items)):
    search = form_df.loc[form_df['Referencias ']== items[i]].index.tolist()
    if len(search) > 1:
        rows.append(search[1])
    elif len(search) == 1:
        rows.append(search[0])

# Obtener lista de columnas con la información de los manifiestos
output = []
for i in range(0, len(rows)):
    output.append(form_df.loc[rows[i], ['Numero de Formulario', 'PAGINA ', 'FECHA']].tolist())
    
# Convertir en caracteres todos los outputs para el pre-procesado final
for i in range(0, len(output)):
    output[i][0] = str(output[i][0])
    output[i][1] = str(output[i][1])
    output[i][2] = output[i][2].strftime("%d-%b-%Y")
    
# Obtener información de final de manifiestos, paginas y fechas
final = []
for i in range(0, len(output)):
    final.append(" - ".join(output[i]))
        
final = [x for x, y in collections.Counter(final).items() if y > 1]

print(final)


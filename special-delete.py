import pandas as pd
import re

excel = input('Ingrese el nombre del archivo Excel: ')

df = pd.read_excel(excel, sheet_name = 'Sheet1')

dfe = pd.read_excel(excel, sheet_name = 'dominios')

for i in ['S/NRO', 'S/N','*','SIN NOMBRE','S/ CALLE', 'S/NO', 'S/No','S/No.', 'S/Ndeg', 'SN', ' S/C ', 'S/C ', 'SIN CALLE', 'S/','SIN INFORMAR','/', 'o', 'PB', 'deg', 'Deg', '00000', 'S/Ndeg|']:
    df['CALLE'] = df['CALLE'].str.replace(i, '', regex = True)
  
for i in ['C.A.B.A.']:
    df['LOCALIDAD'] = df['LOCALIDAD'].str.replace(i, 'CIUDAD AUTONOMA BUENOS AIRES', regex = True)

for i in ['CIUDAD AUTONOMA']:
    df['PROVINCIA'] = df['PROVINCIA'].str.replace(i, 'CAPITAL FEDERAL', regex = True)

df.loc[df['CP'] == 1661, 'PARTIDO'] = 'SAN MIGUEL'

#llenar celdas vacias de calle con la palabra calle
df.loc[df['CALLE'].isnull(), 'CALLE'] = 'CALLE'

# for i in len(df):
#     df.loc[i,'CAP'] = df.loc[i, 'PROVINCIA']
#     df['PROVINCIA'] = df['PROVINCIA'].str.replace(i, 'CAPITAL FEDERAL', regex = True)
    
#for i in ['0']:
  #  df['COLUMNAS'] = df['COLUMNAS'].str.replace(i, '')

#Leer
with pd.ExcelWriter('special-delete-asignado.xlsx') as writer:  

    dfe.to_excel(writer, sheet_name = 'dominios', index = False)

    df.to_excel(writer, sheet_name = 'Sheet1', index = False)

print ('¡SPECIAL DELETE HAS BEEN INJECTED!')
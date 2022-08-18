
#!/usr/bin/env python
# coding: utf-8
import pandas
from unidecode import unidecode


excel = input('excel: ')

#traigo mi csv de partidos
df_ciudades = pandas.read_excel("data/partidos0.xlsx")

#abro mi excel de dominios
df  = pandas.read_excel(excel, sheet_name = 'Sheet1')
dfd = pandas.read_excel(excel, sheet_name = 'dominios')

df['PARTIDO'] = df['PARTIDO']
df  = df[df['PARTIDO'].isnull()]
dfa = df[df['PARTIDO'].isnull()]

lista = df.index.values.tolist() 

df['LOCALIDAD'] = df['LOCALIDAD'].astype(str)

for col in ['LOCALIDAD','PROVINCIA']:
    df[col] = df[col].apply(unidecode)

simbolos = ['+','-','.','/','@',' ','=']

# for simbolo in simbolos:
#     df['LOCALIDAD'] = df['LOCALIDAD'].str.replace(simbolo,'', regex = True)
#     df_ciudades['LOCALIDAD'] = df_ciudades['LOCALIDAD'].str.replace(simbolo,'', regex = True)
#     df['PROVINCIA'] = df['PROVINCIA'].str.replace(simbolo,'')
#     df_ciudades['PROVINCIA'] = df_ciudades['PROVINCIA'].str.replace(simbolo,'' , regex = True)

for i in lista: 
        

    localidad = df.loc[i,'LOCALIDAD']
    
    provincia = df.loc[i,'PROVINCIA']
    
    # if provincia in localidad:
    #     localidad.replace(provincia,'', regex = True)
    
    try:
        #ESTUDIO POR LOCALIDAD
        dfciudad = df_ciudades[df_ciudades['LOCALIDAD'] == localidad]
        
        dfciudad = dfciudad[dfciudad['PROVINCIA'] == provincia]    
        
        dfciudad = dfciudad.drop_duplicates(subset = ['LOCALIDAD'])
        
        dfciudad = dfciudad.set_index('LOCALIDAD')
        

        #compara la localidad del df original con la localidad del cp.xlsx
        partido = dfciudad.loc[localidad,'PARTIDO']
        #si coincide guarda el partido
        df.loc[i,'PARTIDO'] = partido
        
        print(partido,'asignado por BBDD', localidad, provincia)
    except Exception as e:
        print('no se pudo asignar el partido a',localidad, provincia)
    

original_df = pandas.read_excel(excel, sheet_name = 'Sheet1')
original_df = original_df.drop(original_df.index[lista])


df['LOCALIDAD'] = dfa['LOCALIDAD']
df['PROVINCIA'] = dfa['PROVINCIA']

original_df = pandas.concat([original_df, df])

original_df = original_df.sort_values('PARTIDO')

with pandas.ExcelWriter('partido-asignado0.xlsx') as writer:  

    dfd.to_excel(writer, sheet_name = 'dominios', index = False)

    original_df.to_excel(writer, sheet_name = 'Sheet1', index = False)
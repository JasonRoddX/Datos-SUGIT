#!/usr/bin/env python
# coding: utf-8
import pandas
import re
import jellyfish
from unidecode import unidecode


excel = input('excel: ')
#traigo mi csv de partidos
df_modelos = pandas.read_excel("data/vehiculos.xlsx")

#abro mi excel de dominios
df  = pandas.read_excel(excel, sheet_name = 'Sheet1')
dfd = pandas.read_excel(excel, sheet_name = 'dominios')


df['MARCAm'] = df['MARCA'].astype(str)
df['MODELOm'] = df['MODELO'].astype(str)
df_modelos['MODELO'] = df_modelos['MODELO'].astype(str)
df_modelos['MARCA'] = df_modelos['MARCA'].astype(str)

simbolos = ['+','-','.','/','@',' ','=']

for simbolo in simbolos:
    
    df['MARCAm'] = df['MARCAm'].str.replace(simbolo,'', regex = True)
    df['MODELOm'] = df['MODELOm'].str.replace(simbolo,'', regex = True)
    
    df_modelos['MODELO'] = df_modelos['MODELO'].str.replace(simbolo,'', regex = True)
    df_modelos['MARCA'] = df_modelos['MARCA'].str.replace(simbolo,'', regex = True)

for col in ['MARCAm','MODELOm']:
    df[col] = df[col].apply(unidecode)
    
for col in ['MARCA','MODELO']:
    df_modelos[col] = df_modelos[col].apply(unidecode)
    
    
for i in range(len(df)):
    
        #VARIABLES A COMPARAR
    marca = df.loc[i,'MARCAm']
    try:
        modelo = df.loc[i,'MODELOm'].replace(marca,'')
    except:
        modelo = df.loc[i,'MODELOm']
        
    dominio = df.loc[i,'DOMINIO']
    
    if re.search(r'\w\d\d\d\w\w\w',dominio) or re.search(r'\d\d\d\w\w\w',dominio):
        df.loc[i,'TIPO'] = 4
    else:
        try:
            #ORDENO VALORES POR TIPO
            dfi = df_modelos[df_modelos['MODELO'] == modelo]
            dfi = dfi.set_index('MODELO')
            dfi = dfi[dfi['MARCA'] == marca] 
            #compara la marca del df original con la marca de la data agregada
            tipo = dfi.loc[modelo,'TIPO']

            df.loc[i,'TIPO'] = tipo

            print(f'tipo {tipo} asignado por BBDD {marca} - {modelo}')
            
        except:
            # try:
            #     for m in df_modelos['MODELO'].to_list():
            #         if jellyfish.jaro_distance(m, modelo) > 0.9:
                        
            #             dfi = df_modelos.set_index('MODELO')
                        
            #             dfi = dfi[dfi['MARCA'] == marca] 

            #             tipo = dfi.loc[m,'TIPO']
                    
            #             df.loc[i,'TIPO'] = tipo
                        
            #             print(f'{tipo} asignado por comparacion {m} y {modelo}')
            #             break
            # except Exception as e:
        
                #print(e)
                pass
df = df.drop(['MARCAm','MODELOm'], axis = 1)
df['TIPO'] = df['TIPO'].astype(str)
df = df.sort_values(['TIPO'])

with pandas.ExcelWriter('tipo-asignado.xlsx') as writer:  

    dfd.to_excel(writer, sheet_name = 'dominios', index = False)

    df.to_excel(writer, sheet_name = 'Sheet1', index = False)
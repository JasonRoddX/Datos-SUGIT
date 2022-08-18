import pandas as pd
import random as rd
import re

excel = input("Ingrese el nombre del archivo Excel: ")

df  = pd.read_excel(excel, sheet_name = 'Sheet1')
dfd = pd.read_excel(excel, sheet_name = 'dominios')

df['NUMERO'] = df['NUMERO'].fillna(0)
for i in range(len(df)): 
    
    calle = str(df.loc[i,'CALLE'])
    
    calle_list = calle.split(' ')
    if 'S/N' not in calle:
        try:
            if all(isinstance(int(i), int) for i in calle_list):
                if len(calle_list) == 3:
                    df.loc[i,'CALLE'] = calle_list[0]
                    df.loc[i,'NUMERO'] = calle_list[1]
                    df.loc[i,'PISO'] = calle_list[2]
                elif len(calle_list) == 2:
                    df.loc[i,'CALLE'] = calle_list[0]
                    df.loc[i,'NUMERO'] = calle_list[1]
                
        except:
            if calle_list[-1].isdigit():
                
                numero = int(calle_list[-1])
                
                if  numero > 15:
                    
                    df.loc[i,'NUMERO'] = numero
                                    
                    calle_list.pop(-1)
                    
                    calle = ' '.join(str(e) for e in calle_list)
                    
                    df.loc[i,'CALLE'] = calle
                    df.loc[i,'PISO'] = ''
                    df.loc[i,'DEPARTAMENTO'] = ''
                else:
                    try:
                        if calle_list[-2].isdigit():
                            
                            df.loc[i,'NUMERO'] = calle_list[-2]
                            df.loc[i,'PISO'] = calle_list[-1] if int(calle_list[-1]) > 0 else ''
                            
                            calle_list.pop(-1)
                            calle_list.pop(-1)
                            
                            calle = ' '.join(str(e) for e in calle_list)
                            
                            df.loc[i,'CALLE'] = calle
                            df.loc[i,'DEPARTAMENTO'] = ''
                        
                    except:
                        df.loc[i,'NUMERO'] = calle_list[-1]
                        
                        calle_list.pop(-1)
                        
                        calle = ' '.join(str(e) for e in calle_list)
                        
                        df.loc[i,'CALLE'] = calle
                        df.loc[i,'PISO'] = ''
                        df.loc[i,'DEPARTAMENTO'] = ''

            elif calle_list[-1] == 'PB':
                if calle_list[-2].isdigit():
                    
                    df.loc[i,'NUMERO'] = calle_list[-2]
                                    
                    calle_list.pop(-2)
                    
                    calle = ' '.join(str(e) for e in calle_list)
                    
                    df.loc[i,'CALLE'] = calle
                    df.loc[i,'PISO'] = ''
                    df.loc[i,'DEPARTAMENTO'] = ''
                            
        if df.loc[i,'NUMERO'] == 0:
            df.loc[i,'NUMERO'] = rd.randint(100,200)

    if re.search('S/N', calle):
        numero = rd.randint(100,200)
        df.loc[i,'NUMERO'] = numero
        df.loc[i,'CALLE'] = calle.replace('S/N', '')
        df.loc[i,'PISO'] = ''
        df.loc[i,'DEPARTAMENTO'] = ''
        
with pd.ExcelWriter('numero-asignado.xlsx') as writer:  

    dfd.to_excel(writer, sheet_name = 'dominios', index = False)

    df.to_excel(writer, sheet_name = 'Sheet1', index = False)
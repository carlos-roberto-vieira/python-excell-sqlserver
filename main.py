# pandas
# openpyxl

import pandas as pd
import pyodbc
     
tabela_ncm = pd.read_excel('ncm2022.xlsx' )
print('carregada')
 
df =pd.DataFrame( tabela_ncm, columns=['NCM'] )
print('carregada preparada ')

consql = pyodbc.connect( "Driver={SQL Server};Server=INFORME_O_IP;DataBase=INFORME_O_BANCO_;Uid=INFORME_O_USUARIO;Pwd=INFORME_A_SENHA" )
cursql = consql.cursor()

print('conectado')

for row in df.itertuples():
    cursql.execute( "insert into cad_est_produtos_ncm ( ncm ) values ( ? )", row.NCM )
    cursql.commit()
cursql.close 

print('processo concluído')






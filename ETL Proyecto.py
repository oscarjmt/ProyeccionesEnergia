"""
Jeyner Arango
Elisa Samayoa
Oscar Méndez
"""

from io import BytesIO#, StringIO
from zipfile import ZipFile
from urllib.request import urlopen
import datetime
import pandas as pd
#import numpy as np

# def read_zip(zip_fn, extract_fn=None):
#     zf = ZipFile(zip_fn)
#     if extract_fn:
#         return zf.read(extract_fn)
#     else:
#         return {name:zf.read(name) for name in zf.namelist()}

df = pd.DataFrame()
fecha_inicio = input('Ingrese la fecha inicial en formato AAAA-MM-DD: ')
anio1, mes1, dia1 = map(int, fecha_inicio.split('-'))
fecha1 = datetime.date(anio1, mes1, dia1)
fecha_final = input('Ingrese la fecha   final en formato AAAA-MM-DD: ')
anio2, mes2, dia2 = map(int, fecha_final.split('-'))
fecha2 = datetime.date(anio2, mes2, dia2)
poe_rango_fechas= pd.date_range(str(anio1)+'-'+str(mes1)+'-'+str(dia1),
                                str(anio2)+'-'+str(mes2)+'-'+str(dia2), \
                                    freq='D')
lista_fechas = poe_rango_fechas.strftime("%Y%m%d").tolist()
meses_dict = {1: '01_ENERO', 2:'02_FEBRERO', 3: '03_MARZO',
         4: '04_MARZO', 5: '05_MAYO', 6: '06_JUNIO',
         7: '07_JUJIO', 8: '08_AGOSTO', 9: '09_SEPTIEMBRE',
         10: '10_OCTUBRE', 11: '11_NOVIEMBRE', 12: '12_DICIEMBRE'}

for i in range(len(poe_rango_fechas)):
    url_cadena = 'https://www.amm.org.gt/pdfs2/post_despacho/' + \
        'POSDESPACHO_DIARIO/'  + \
            str(poe_rango_fechas[i].year)+'/'+ \
                meses_dict[poe_rango_fechas[i].month] + '/' + \
                    'PD'+str(lista_fechas[i]) + '.zip'
    print('probando: '+lista_fechas[i])
    resp = urlopen(url_cadena)
    myzip = ZipFile(BytesIO(resp.read()))
    xlfile = myzip.open(myzip.filelist[0])
    df[lista_fechas[i]] = pd.read_excel(xlfile, sheet_name='POE',skiprows = 4, 
                       nrows=24,  usecols= 'F', header=1)
    #df.rename(columns={'POE (US$/MWh)': lista_fechas[i]}, inplace=True)
    xlfile.close()
    myzip.close()

df['promedio'] = df.mean(axis=1)
if (df.shape[1]-1) < 7:
    df['promedio_7_dias'] = df['promedio']
else:
    df['promedio_7_dias'] = df.iloc[:, -8:-1].mean(axis=1)

valle = df.iloc[[0,1,2,3,4,5,22,23],:-2].to_numpy().mean()
diurno = df.iloc[6:18,:-2].mean().to_numpy().mean()
pico = df.iloc[18:22,:-2].mean().to_numpy().mean()
promedio = df.iloc[:,:-2].to_numpy().mean()
maximo = df.iloc[:,:-2].to_numpy().max()
minimo = df.iloc[:,:-2].to_numpy().min()
actual = [valle, diurno, pico, promedio, maximo, minimo]
df_resumen = pd.DataFrame(actual, columns=['Actual'])

valle_7d = df.iloc[[0,1,2,3,4,5,22,23],-1].mean()
diurno_7d = df.iloc[6:18,-1].mean()
pico_7d = df.iloc[18:22,-1].mean()
promedio_7d = df.iloc[:,-1].mean()
maximo_7d = df.iloc[:,-1].max()
minimo_7d = df.iloc[:,-1].min()
ultimos_7d = [valle_7d, diurno_7d, pico_7d, 
              promedio_7d, maximo_7d, minimo_7d]
df_resumen['ultimos_7_dias'] = ultimos_7d
df_resumen.rename(index={0:'valle', 1:'diurno', 2:'pico',3:'promedio',
                                4:'maximo',5:'minimo'}, inplace=True)

df_stacked = df.iloc[:, :-2].unstack().reset_index()
df_stacked.columns = ["Fecha", "Hora", "Precio"]

writer = pd.ExcelWriter('precios_energia.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
df.to_excel(writer, sheet_name='precios por fecha')
df_resumen.to_excel(writer, sheet_name='resumen')
df_stacked.to_excel(writer, sheet_name='precios tidy', index=False)

writer.save()
writer.close()
from openpyxl import load_workbook
from datetime import datetime
rut=r'BaseCrudColabPersona.xlsx'

def leer(ruta:str, extraer:str):
  archivoExcel=load_workbook(ruta)

  hojaDatos=archivoExcel['persona']
  hojaDatos=hojaDatos['A2':'F'+str(hojaDatos.max_row)]

  info={}

  for i in hojaDatos:
    if isinstance(i[0].value,int):
      info.setdefault(i[0].value,{'nombre':i[1].value, 'edad':i[2].value,'telefono':i[3].value,'correo':i[4].value,'nacimiento':i[5].value})

  if not(extraer=='todo'):
    info=filtrar(info,extraer)
  for i in info:
    print('********** Tarea ***********')
    print('id:'+str(i)+'\n'+'nombre: '+str(info[i]['nombre'])+'\n'+'edad: '+str(info[i]['edad'])+'\n'+'telefono: '+str(info[i]['telefono'])+'\n'+'correo: '+str(info[i]['correo'])+'\n'+'nacimiento: '+str(info[i]['nacimiento']))
    print()
  return info

def filtrar(info:dict, filtro:str):
  aux={}
  for i in info:
    if info[i]['edad']==filtro:
      aux.setdefault(i,info[i])
  return aux
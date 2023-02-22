from openpyxl import load_workbook
from datetime import datetime
rut=r'BaseCrudColabPersona.xlsx'

def agregar(ruta:int, datos:dict):
  archivoExcel=load_workbook(ruta)

  hojaDatos=archivoExcel['persona']
  hojaDatos=hojaDatos['A2':'F'+str(hojaDatos.max_row+1)]
  hoja=archivoExcel.active

  nombre=2
  edad=3
  telefono=4
  correo=5
  nacimiento=6
  for i in hojaDatos:
    if not(isinstance(i[0].value,int)):
      identificador=i[0].row
      hoja.cell(row=identificador,column=1).value=identificador-1
      hoja.cell(row=identificador,column=nombre).value=datos['nombre']
      hoja.cell(row=identificador,column=edad).value=datos['edad']
      hoja.cell(row=identificador,column=telefono).value=datos['telefono']
      hoja.cell(row=identificador,column=correo).value=datos['correo']
      hoja.cell(row=identificador,column=nacimiento).value=datos['nacimiento']
      break
  archivoExcel.save(ruta)
  return
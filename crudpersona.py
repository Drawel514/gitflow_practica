from openpyxl import load_workbook
from datetime import datetime
rut=r'BaseCrudColabPersona.xlsx'

def actualizar(ruta:str,identificador:int,datosActualizados:dict):
  archivoExcel=load_workbook(ruta)

  hojaDatos=archivoExcel['persona']
  hojaDatos=hojaDatos['A2':'F'+str(hojaDatos.max_row)]
  hoja=archivoExcel.active

  nombre=2
  edad=3
  telefono=4
  correo=5
  nacimiento=6
  encontro=False
  for i in hojaDatos:
    if i[0].value==identificador:
      fila=i[0].row
      encontro=True
      for d in datosActualizados:
        if d=='nombre' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=nombre).value=datosActualizados[d]
        elif d=='edad' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=edad).value=datosActualizados[d]
        elif d=='telefono' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=telefono).value=datosActualizados[d]
        elif d=='correo' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=correo).value=datosActualizados[d]
        elif d=='nacimiento' and not(datosActualizados[d]==''):
          hoja.cell(row=fila,column=nacimiento).value=datosActualizados[d]
  archivoExcel.save(ruta)
  if encontro == False:
    print('Error: no existe una persona con ese id')
    print()
  return

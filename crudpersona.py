from openpyxl import load_workbook
from datetime import datetime
rut=r'BaseCrudColabPersona.xlsx'

while True:
  print('Indique la accion que desea realizar: \nConsultar: 1\nActualizar: 2\nRegistrar: 3\nBorrar: 4')
  accion =int(input('Escriba la opcion: '))
  if accion<1 or accion>4:
    print('Comando invalido, por favor eliga una opcion valida')
  elif accion==1:
    opcConsulta=''
    print('Indique la opcion que desee:\nTodas las personas: 1\nMayores o iguales 18: 2\nMenores a 18: 3')
    opcConsulta=input('Escriba la persona que see consultar: ')
    if opcConsulta=='1':
      print('\n\n** Consultado todas las personas **')
      leer(rut,'todo')
    elif opcConsulta=='2':
      print('\n\n** Consultado todas las personas **')
      leer(rut,'mayor')
    elif opcConsulta=='3':
      print('\n\n** Consultado todas las personas **')
      leer(rut,'menor')
  elif accion==2:
    datosActualizados={'nombre':'','edad':'','telefono':'','correo':'','nacimiento':''}
    print('** Actualizar persona **\n')
    idActualizar=int(input('Indique el ID de la persona que desea actualizar: '))
    print('\n** Nuevo nombre **\n** Nota: si no desea actualizar el nombre solo oprima ENTER **')
    datosActualizados['nombre']=input('Indique el nuevo nombre de la persona: ')
    print('\n** Nueva edad **\n** Nota: si no desea actualizar la edad solo oprima ENTER **')
    datosActualizados['edad']=input('Indique la nueva edad de la persona: ')
    print('\n** Nuevo telefono **')
    datosActualizados=input('Indique el nuevo telefono de la persona: ')
    print('\n** Nuevo correo **')
    datosActualizados=input('Indique el nuevo correo de la persona: ')
    print('\n** Nuevo nacimiento **')
    datosActualizados=input('Indique el nuevo nacimiento de la persona: ')
    actualizar(rut,idActualizar, datosActualizados)
    print()
  elif accion==3:
    datosActualizados={'nombre':'','edad':'','telefono':'','correo':'','nacimiento':''}
    print('** Crear nueva persona **\n')
    print('** nombre **\n')
    datosActualizados['nombre']=input('Indique el nombre de la persona: ')
    print('\n** edad **')
    datosActualizados['edad']=input('Indique la edad de la persona: ')
    print()
    print('\n** telefono **')
    datosActualizados['telefono']=input('Indique el telefono de la persona: ')
    print()
    print('\n** correo **')
    datosActualizados['correo']=input('Indique el correo de la persona: ')
    print()
    print('\n** nacimiento **')
    datosActualizados['nacimiento']=input('Indique el nacimiento de la persona: ')
    agregar(rut,datosActualizados)

  elif accion==4:
    print('\n** Eliminar persona **')
    iden=int(input('Indique el ID de la persona que desea eliminar: '))
    borrar(rut,iden)

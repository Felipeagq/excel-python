import os 
os.system('pip3 install openpyxl')
from openpyxl import Workbook
from openpyxl import load_workbook

##########################
### CREAR OBJETO EXCEL ###
##########################
wb = Workbook() # creamos objeto de Excel
wb.save('Doc_Excel.xlsx') 
# se guarda un documento de excel en blanco con
# una sola ws llamada "sheet"

###########################
### CARGAR OBJETO EXCEL ###
###########################
# Para cargar una ws de excel ya existente
wb = load_workbook('Doc_Excel.xlsx')

### CREAMOS UNA HOJA DE TRABAJO ###
wb.create_sheet('Primera_ws', # nombre de la hoja
0) # posición de la hoja

# Cambiarle el estilo a la pestaña de la ventana
ws = wb['Primera_ws']
ws.sheet_properties.tabColor = "1072AB"

###########################
### DATOS DE LAS CELDAS ###
###########################
# Forma 1:
ws['B1'] = 1
# Escribiremos valores en la primero columna
for i in range(10):
    wb['Primera_ws']['A{}'.format(i+1)] = i 
# Forma 2:
ws = wb['Primera_ws']
ws.cell(row=6, column=1, value='hola')

### GUARDADMOS LOS CAMBIOS EN EL ARCHIVO ###
wb.save('Doc_Excel.xlsx')

# Acceder al valor de una celda
print(wb['Primera_ws']['A6'].value)

# se puede acceder a un rango de valores
celdas = ws['A2':'A7'] # guardamos el rengo de celdas
# se guarda como una matrix
print(celdas[5][0].value) # accedemos a cada valor individual 
# print(celdas[5]) --> (<Cell 'Primera_ws'.A7>,)
# print(celdas[5][0]) <Cell 'Primera_ws'.A7>
# de nuestra matrix de celdas

# Tambien podemos acceder a varios valores de rows
for row in ws.iter_rows(min_row=1,max_col=3,max_row=2):
    for cell in row: # interamos las celdas de cada ro2
        print(cell)
print(' ')
for col in ws.iter_cols(min_row=1, max_col=3, max_row=2):
    for cell in col:
        print(cell)

#############################
### COPIA DE LA WORKSHEET ###
#############################
# Crear copias de Worksheet sobre la que estamos trabajando (activa)
source = wb.active
target = wb.copy_worksheet(source)

print(' ')
# Para iterar sobre todas las rows y cols 
ws['C11'] = 'Fin de las celdas'
print(tuple(ws.rows),end='\n\n')
ws['D15'] = 'Fs'
# para imprimir el valor de cada celda
for col in tuple(ws.rows):
    for cell in col:
        print(cell.value)
    print('-')

print(' ')

# Otra forma de visualizar solo valores por row
for row in ws.values:
    for value in row:
        print(value)
    print('-')
print(' ')
#############################
### VALORES DEL WORKSHEET ###
#############################
# iter_rows() e iter_cols() admiten values_only
for row in ws.iter_rows(min_row=1, max_col=3, max_row=15, values_only=True):
    print(row)


wb.save('Doc_Excel.xlsx')

os.system('libreoffice Doc_Excel.xlsx')

# python3 openpyxl/excel_work.py
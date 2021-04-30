import os 
from openpyxl import Workbook
from openpyxl import load_workbook

wb = Workbook() # creamos objeto de Excel
wb.save('Doc_Excel.xlsx') 
# se guarda un documento de excel en blanco con
# una sola ws llamada "sheet"

# Para cargar una ws de excel ya existente
wb = load_workbook('Doc_Excel.xlsx')
wb.create_sheet('Primera_ws', # nombre de la hoja
0) # posición de la hoja

# Escribiremos valores en la primero columna
for i in range(10):
    wb['Primera_ws']['A{}'.format(i+1)] = i 
# Otra forma de cambiar los valores
ws = wb['Primera_ws']
ws.cell(row=6, column=1, value='hola')

ws['B1'] = 1

wb.save('Doc_Excel.xlsx')

# Cambiarle el estilo a la pestaña de la ventana
ws = wb['Primera_ws']
ws.sheet_properties.tabColor = "1072AB"

# Acceder al valor de una celda
print(wb['Primera_ws']['A6'].value)

# se puede acceder a un rango de valores
celdas = ws['A2':'A7'] # guardamos el rengo de celdas
# se guarda como una matrix
print(celdas[5][0].value) # accedemos a cada valor individual 
# de nuestra matrix de celdas

wb.save('Doc_Excel.xlsx')

os.system('libreoffice Doc_Excel.xlsx')
from openpyxl import Workbook
from openpyxl import load_workbook

excel = Workbook() # creamos objeto de Excel
excel.save('Doc_Excel.xlsx') 
# se guarda un documento de excel en blanco con
# una sola hoja llamada "sheet"

# Para cargar una hoja de excel ya existente
excel2 = load_workbook('Doc_Excel.xlsx')
excel2.create_sheet('Primera_hoja',0)

# Escribiremos valores en la primero columna
for i in range(10):
    excel2['Primera_hoja']['A{}'.format(i+1)] = i 
# Otra forma de cambiar los valores
hoja = excel2['Primera_hoja']
hoja.cell(row=6, column=1, value=30)

excel2.save('Doc_Excel.xlsx')

# Cambiarle el estilo a la pestaña de la ventana
hoja = excel2['Primera_hoja']
hoja.sheet_properties.tabColor = "1072AB"

# Acceder al valor de una celda
print(excel2['Primera_hoja']['A6'].value)

# se puede acceder a un rango de valores
celdas = hoja['A2':'A7'] # guardamos el rengo de celdas
# se guarda como una matrix
print(celdas[5][0].value) # accedemos a cada valor individual 
# de nuestra matrix de celdas

excel2.save('Doc_Excel.xlsx')

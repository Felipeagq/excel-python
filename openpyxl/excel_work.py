import os 
import datetime
os.chdir(os.getcwd())
print(os.getcwd())
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
for row in ws.iter_rows(min_row=1, max_col=4, max_row=16, values_only=True):
    print(row)

# imprimir nombres de las hojas
print(wb.sheetnames)
['Sheet2', 'New Title', 'Sheet1']


##########################
### SAVING AS A STREAM ###
##########################
# If you want to save the file to a stream,
# e.g. when using a web application such as Pyramid,
# Flask or Django then you can simply provide a NamedTemporaryFile():
from tempfile import NamedTemporaryFile
from openpyxl import Workbook
wb2 = Workbook()
with NamedTemporaryFile() as tmp:
    wb2.save(tmp.name)
    tmp.seek(0)
    stream = tmp.read()

wb.save('Doc_Excel.xlsx')

# You can specify the attribute template=True,
# to save a workbook as a template:
wb = load_workbook('Doc_Excel.xlsx')
wb.template = True
wb.save('document_template.xltx')

# or set this attribute to False (default), to save as a document:
'''wb = load_workbook('document_template.xltx')
wb.template = False
wb.save('document.xlsx', as_template=False)'''



#####################
## DATETIME FORMAT ##
#####################
# set date using a Python datetime
ws = wb.active
ws['B3'] = datetime.datetime(2010, 7, 21)
print(ws['A1'].number_format)


###################
## USING FORMULA ##
###################
# add a simple formula
ws["B4"] = "=SUM(A2, A4)"


###########################
## MERGE / UNMERGE CELLS ##
###########################
ws.merge_cells('C11:D11')
# ws.unmerge_cells('C11:D11')
# or equivalently
#ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=4)
#ws.unmerge_cells(start_row=2, start_column=1, end_row=4, end_column=4)


########################
## INSERTING AN IMAGE ##
########################
from openpyxl.drawing.image import Image
ws['B5'] = 'You should see three logos below'
img = Image('logo.png')
ws.add_image(img, 'H6')


####################
## FOLD (OUTLINE) ##
####################
ws.column_dimensions.group('E','F', hidden=True)
ws.row_dimensions.group(16,17, hidden=True)


####################
## READ-ONLY MODE ##
####################
'''from openpyxl import load_workbook
wb = load_workbook(filename='large_file.xlsx', read_only=True)
ws = wb['big_data']
# Reading the data
for row in ws.rows:
    for cell in row:
        print(cell.value)
# Close the workbook after reading
wb.close()'''


# This operation will overwrite existing files without warning.



#################################################################################
################################ WRITE-ONLY MODE ################################
#################################################################################
'''from openpyxl import Workbook
wb2 = Workbook(write_only=True)
ws2 = wb2.create_sheet()
# now we'll fill it with 100 rows x 200 columns
for irow in range(100):
    ws2.append(['%d' % i for i in range(200)])
# save the file
wb2.save('new_big_file.xlsx') # doctest: +SKIP
#If you want to have cells with styles or comments then use a'''
'''
Unlike a normal workbook,
a newly-created write-only workbook does not contain any worksheets;
a worksheet must be specifically created with the create_sheet() method.

In a write-only workbook,
rows can only be added with append().
It is not possible to write (or read) cells at arbitrary locations with cell() or iter_rows().

It is able to export unlimited amount of data (even more than Excel can handle actually),
while keeping memory usage under 10Mb.

A write-only workbook can only be saved once.
After that, every attempt to save the workbook or append() to an existing worksheet will raise an openpyxl.
utils.exceptions.WorkbookAlreadySaved exception.

Everything that appears in the file before the actual cell data must be
created before cells are added because it must written to the file before then. 
For example, freeze_panes should be set before cells are added.
'''

'''wb = Workbook(write_only = True)
ws3 = wb.create_sheet()
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.styles import Font
cell = WriteOnlyCell(ws3, value="hello world")
cell.font = Font(name='Courier', size=36)
cell.comment = Comment(text="A comment", author="Author's Name")
ws3.append([cell, 3.14, None])
wb.save('write_only_file.xlsx')'''



####################
## INSERTING ROWS ##
####################
ws.insert_rows(7)



###############################
## DELETING ROWS AND COLUMNS ##
###############################
ws.delete_cols(6,3)
ws.delete_rows(10)


############################
## MOVING RANGES OF CELLS ##
############################
ws.move_range("A4:B4", rows=0, cols=2)
# movemos ese rango de celdas 0 rows y dos columnas hacia la derecha


#######################
## DATAFRAME TO ROWS ##
#######################
'''from openpyxl.utils.dataframe import dataframe_to_rows
wb = Workbook()
ws = wb.active

for r in dataframe_to_rows(df, index=True, header=True):
    ws.append(r)'''


############################
## WORKSHEET TO DATAFRAME ##
############################
'''df = DataFrame(ws.values)'''




wb.save('Doc_Excel.xlsx')
os.system('libreoffice Doc_Excel.xlsx')

# python3 openpyxl/excel_work.py
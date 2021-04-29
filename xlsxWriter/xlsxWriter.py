# pip3 install XlsxWriter
import os
import xlsxwriter

# Creamos un archivo de Excel
workbook = xlsxwriter.Workbook('Hello.xlsx') 

# Creamos una hoja de trabajo
worksheet1 = workbook.add_worksheet('Hoja1')
worksheet2 = workbook.add_worksheet('Hoja2')
worksheet3 = workbook.add_worksheet('Hoja3')

#escribimos los nombre de los Worksheet
print(workbook.sheetnames.keys())

# escribimos en la hoja de trabajo
worksheet1.write('A1','Hola mundo')
worksheet1.write(2,3,'Escribirmos algo')

for i in range(2,10):
    worksheet1.write(i, 0,i+5)

# Podemos agregar funciones
worksheet1.write(10,0,'=SUM(A3, A10)')
worksheet1.write('A11',"=SUM(A3, A10)")

# Podemos cambiar el formateo de la celda
f_cell = workbook.add_format({
    "bg_color":"#00ff00", # color de fondo
    "font": "Century", # Fuente del texto
    "font_size":15, # tamaño de la fuente
    "border":3, # grosor del borde
    "bold":True
})

# Podemos modificar el ancho de la columna
worksheet1.set_column(
    "D:F",  # las columnas que vamos ampliar inicio:final
    # n1,n2  siendo n1 el numero de la columna de inicio y n2 la final
    20      # Amplitud de la columna
)

worksheet1.write('B11',"=PI()", f_cell)
worksheet1.write(3,3,'Escribirmos algo',f_cell)

# Criterios al momento de escribir
worksheet1.conditional_format(
    "B1:B9",{
        "type":"cell",
        "criteria":"<=",
        "value":5,
        "format":f_cell
    }
)
# Podemos escribir en una columna
worksheet1.write_column('B1', [1,2,3,4,5,500,7,8,9])

# Podemos escribir en filas
worksheet1.write_row('C2',[1,2,3,4,5])

# Trabajando con TABLAS
full_border = workbook.add_format({
    "border":1,
    "border_color": '#000000'
})
worksheet1.write('D6',None, full_border)

# Agregando formato al bottom de la celda
full_border2 = workbook.add_format({
    "bottom":5, # puede ser top,left,right : el numero del tipo de borde que queremos
    "bottom_color": '#ff0000'
})
worksheet1.write('D7',None, full_border2)

header_border = workbook.add_format({
    # Los bordes en general
    "border":5,
    "border_color": '#000000',

    # solamente para el top
    'top':5,
    'top_color': '#ff0000',

    # Solamente para el right
    'right':5,
    'right_color':'#00ff00'
})
worksheet1.write('D8','celda', header_border)

# CREANDO UNA TABLA
columnas = ['c1','c2','c3','c4','c5']
row_index=9
column_index = 4
fin_col = column_index + len(columnas)
i = 0
for col in range(column_index,fin_col):
    worksheet1.write(row_index, col, columnas[col-column_index],header_border)

# tablas
data =[
    ['apples',1,2,4,5],
    ['bananas',1,2,7,8],
    [12,34,6,6,],
    [45,6,2,8,0]
]
worksheet1.add_table('A14:D18',{
    'name':'tabla_1',
    'data':data,
    'header_row':True,
    'banded_rows':False,
    'first_column':True,
    'style':'Table Style Light 11'
})



# cerramos y guardamos automaticamente
workbook.close()
os.system('libreoffice Hello.xlsx')


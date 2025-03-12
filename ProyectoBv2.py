# Biblioteca
from asyncio.windows_events import NULL
from pickle import FALSE, TRUE
import pandas
import numpy as np
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
from fuzzywuzzy import fuzz
import re

def normalize_name(name):
    if not isinstance(name, str):
        return ""
    name = name.lower().strip()  # Convertir a minúsculas y quitar espacios extras
    name = re.sub(r'[^\w\s]', '', name)  # Eliminar signos de puntuación
    words_to_remove = ['sa', 'cv', 'de', 'ltd', 'ltda', 'srl', 'sapi', 'sas', 'the']  
    name = ' '.join([word for word in name.split() if word not in words_to_remove])  
    
    # Si el nombre tiene más de una palabra, tratamos de reordenarlo
    name_parts = name.split()
    if len(name_parts) > 1:
        reordered_name = ' '.join(sorted(name_parts))  # Ordenamos las palabras
        return reordered_name
    
    return name

def are_names_similar(name1, name2, threshold=85):
    return fuzz.ratio(normalize_name(name1), normalize_name(name2)) >= threshold

#Functions
def eliminateDuplicates_Sort_Lists(lista):
    lista = list(set(lista))
    lista.sort()
    return lista

def titleRowExcel(excelObj):
    excelObj['Proveedor'].append('Proveedor')
    excelObj['RFC'].append('RFC')
    excelObj['# Comprobante'].append('Comprobante')
    excelObj['Moneda'].append('Moneda')
    excelObj['Tipo de Cambio'].append('Tipo de Cambio')
    excelObj['Importe'].append('Importe')
    excelObj['0%'].append('0%')
    excelObj['IVA'].append('IVA')
    excelObj['IVA RETENIDO'].append('IVA RETENIDO')
    excelObj['Total'].append('Total')
    excelObj['# Cheque o transacción'].append('Cheque o transacción')
    excelObj['Fecha cargos'].append('Fecha cargos')

def emptyRowsExcel(excelObj):
    for i in range(2):
        excelObj['Proveedor'].append('')
        excelObj['RFC'].append('')
        excelObj['# Comprobante'].append('')
        excelObj['Moneda'].append('')
        excelObj['Tipo de Cambio'].append('')
        excelObj['Importe'].append('')
        excelObj['0%'].append('')
        excelObj['IVA'].append('')
        excelObj['IVA RETENIDO'].append('')
        excelObj['Total'].append('')
        excelObj['# Cheque o transacción'].append('')
        excelObj['Fecha cargos'].append('')



#Vars
listNames = [] #Is neceseary for obtain the providers names
listProviders = {} #This is the most important var, this dictionary have all names, rfc, bills, etc...

#____________________________________________________________________________
#MAIN
#____________________________________________________________________________
#Open the file "DIOT" in the window PAGOS
base_path = Path(__file__).parent
file_path = (base_path/"resources/Diots/FORMATO DIOT JULIO 2024 1.xlsx").resolve()
excelarchive = pandas.read_excel(file_path, sheet_name = 'PAGOS')

#Columns 
#Obtain the columns of the excel 
providersNamesColumn = excelarchive.columns[18]
SheetColumn = excelarchive.columns[23]
tasa16Column = excelarchive.columns[5]
tasa0Column = excelarchive.columns[4]
ivaColumn = excelarchive.columns[13]
ivaRetColumn = excelarchive.columns[6]
Total_FilePurpleColumn = excelarchive.columns[14]
ChequeColumn = excelarchive.columns[0]
DateColumn = excelarchive.columns[17]


#Iterate all row in DIOT excel and with the column obtain the info of the providersNamesColumn and pass to listNames
for index, row in excelarchive.iterrows():
    listNames.append(str(row[providersNamesColumn]))


#Eliminate the duplicates and sort the list
listNames = eliminateDuplicates_Sort_Lists(listNames) 


#Iterate all names, and for each provider, create a object with the name of the provider, the rfc, and two lists
for name in listNames:
    listProviders.update({name:{
                              'nombre': name, 
                              'rfc': NULL,
                              'foliosData': [],
                              }})

#Iterate the excel again, and add the rest info in the listProviders, if find a provider, add the folios, the date, the cheque, and the object for pagosTodos
for index, row in excelarchive.iterrows():
    listProviders[row[providersNamesColumn]]['foliosData'].append({
        'folio': row[SheetColumn],
        'moneda': "MN",
        'tipocambio': "1", 
        'importe': row[tasa16Column], 
        '0%': row[tasa0Column],
        'iva': row[ivaColumn],
        'ivaretenido': row[ivaRetColumn], 
        'total': row[Total_FilePurpleColumn],
        'cheque': row[ChequeColumn],
        'fecha': row[DateColumn].strftime("%m/%d/%Y"),
        })


# Creamos una lista de objetos, el cual contendra el nombre del archivo, la pagina en donde se encuentra la info y las columnas que necesitamos
listaFilesInfo = [
    {
        # Archivo 4 2024
        "filename": "04 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
    },
    {
        # Archivo 5 2024
        "filename": "05 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
    },
]


for fileData in listaFilesInfo:
    print("Analizando el archivo: " + fileData['filename'])
    # Creamos el nombre del archivo que vamos a buscar 
    nombreArchivo = (base_path / "resources/XML" / fileData['filename']).resolve()
    # Abrimos el archivo
    excelarchiveNew = pandas.read_excel(nombreArchivo, sheet_name = fileData['sheetName'])


    # Columnas Archivos XML
    columnProvedor = excelarchiveNew.columns[fileData['columnProvedor']]
    columnRFC = excelarchiveNew.columns[fileData['columnRFC']]


    # For que itera todo el excel 
    for j, row in excelarchiveNew.iterrows():
        #nombre de la empresa en los archivos xlsx
        for provedor in listProviders:
            if are_names_similar(provedor, row[columnProvedor]):  # Usar comparación con fuzzy
                listProviders[provedor]["rfc"] = row[columnRFC]


print("Creando archivo de Excel...")
    

# PARTE DE IMPRESION EN EXCEL PUTOSSSSSSS
# crear objeto necesario  para DataFrame
def createExcelData():
    return {
        'Proveedor': [], 
        'RFC': [], 
        '# Comprobante': [], 
        'Moneda': [], 
        'Tipo de Cambio': [], 
        'Importe': [], 
        '0%': [], 
        'IVA': [], 
        'IVA RETENIDO': [],
        'Total': [], 
        '# Cheque o transacción': [], 
        'Fecha cargos': [], 
    }






excelObj = createExcelData()
numFilas = 1
listRows_titles = []
for provedor in listProviders:
    titleRowExcel(excelObj)
    listRows_titles.append(numFilas-1)
    numFilas = numFilas + 1
    inicioProveedor = numFilas
    importe = 0
    ceroporcentaje = 0 
    iva = 0
    ivaretenido = 0
    total = 0
    for folioData in listProviders[provedor]["foliosData"]:
        excelObj['Proveedor'].append(listProviders[provedor]['nombre'])
        excelObj['RFC'].append(listProviders[provedor]['rfc'])
        excelObj['# Comprobante'].append(folioData['folio'])
        excelObj['Moneda'].append(folioData['moneda'])
        excelObj['Tipo de Cambio'].append(folioData['tipocambio'])
        excelObj['Importe'].append(folioData['importe'])
        excelObj['0%'].append(folioData['0%'])
        excelObj['IVA'].append(folioData['iva'])
        excelObj['IVA RETENIDO'].append(folioData['ivaretenido'])
        excelObj['Total'].append(folioData['total'])
        excelObj['# Cheque o transacción'].append(folioData['cheque'])
        excelObj['Fecha cargos'].append(folioData['fecha'])
        importe = importe + (folioData['importe'])
        ceroporcentaje = ceroporcentaje + (folioData['0%'])
        iva = iva + (folioData['iva'])
        ivaretenido = ivaretenido + (folioData['ivaretenido'])
        total = total + (folioData['total'])
        numFilas = numFilas + 1
    finProveedor = numFilas - 1
    excelObj['Proveedor'].append('')
    excelObj['RFC'].append('')
    excelObj['# Comprobante'].append('')
    excelObj['Moneda'].append('')
    excelObj['Tipo de Cambio'].append('')
    excelObj['Importe'].append("$ " + str(importe))
    excelObj['0%'].append("$ " + str(ceroporcentaje))
    excelObj['IVA'].append("$ " + str(iva))
    excelObj['IVA RETENIDO'].append("$ " + str(ivaretenido))
    excelObj['Total'].append("$ " + str(total))
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    emptyRowsExcel(excelObj)
    numFilas = numFilas + 3


df = pandas.DataFrame(excelObj)
# crear el objeto ExcelWriter
escrito = pandas.ExcelWriter('Anexo 1 proveedores.xlsx', engine='xlsxwriter')
# escribir el DataFrame en excel
df.to_excel(escrito,'Try', startrow=0, startcol=1, header=False, index=False)
worksheet = escrito.sheets['Try']
cell_format = escrito.book.add_format({
    'align':    'center',
    'valign':   'vcenter',
    'bg_color':   '#92D050'})
cell_format_yellow = escrito.book.add_format()
cell_format_yellow.set_bg_color('#FFFF00')
currency_format = escrito.book.add_format({'num_format': '_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-'})


print("proovedor length")
print(len(excelObj['Proveedor']))


worksheet.set_column('A:A', 1)
worksheet.set_column('B:B', 15)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 30)
worksheet.set_column('E:E', 25)
worksheet.set_column('F:F', 25)
worksheet.set_column('G:G', 20)
worksheet.set_column('H:H', 15)  
worksheet.set_column('I:I', 15, currency_format)  #number style
worksheet.set_column('J:J', 15, currency_format)  #number style
worksheet.set_column('K:K', 20, currency_format)  #number style
worksheet.set_column('L:L', 15, currency_format)  #number style
worksheet.set_column('M:M', 15, currency_format)  #number style
worksheet.set_column('N:N', 30, currency_format)  #number style
worksheet.set_column('O:O', 30)
worksheet.set_column('P:P', 25)
worksheet.set_column('Q:Q', 20)

#set row height in titles headers
for i in listRows_titles:
    worksheet.set_row(i, 25, cell_format)


worksheet.conditional_format(0, 0, len(excelObj['Proveedor']), 16 ,{ 
    'type': 'text', 'criteria': 'containing', 'value': '$ ', 
    'format': cell_format_yellow
})
    
#Save the excel
escrito.close()
print("Archivo de Excel creado Exitosamente.")
# Biblioteca
from asyncio.windows_events import NULL
from pickle import FALSE, TRUE
import pandas
import numpy as np
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

#Functions
def eliminateDuplicates_Sort_Lists(lista):
    lista = list(set(lista))
    lista.sort()
    return lista

def createFolioObj(folio, fecha, cheque):
    date = fecha.strftime("%m/%d/%Y")
    return {
        'folio': folio, 
        'ffiscal': NULL, 
        'concepto': NULL, 
        'moneda': "MN", 
        'tipocambio': "1", 
        'importe': "0", 
        '0%': "0", 
        'iva': NULL, 
        'ivaretenido': NULL, 
        'totalIEPS': NULL,
        'total': NULL, 
        'cheque': cheque, 
        'fecha': date, 
        'banco': "BANORTE", 
    }

def createPagosTodosObj(totalmorado):
    return {
        'totalmorado': totalmorado,
        'Pfolio': NULL, 
        'Pffiscal': NULL, 
        'Pcoincidencia': FALSE, 
    }

def titleRowExcel(excelObj):
    excelObj['Proveedor'].append('Proveedor')
    excelObj['RFC'].append('RFC')
    excelObj['Folio Fiscal'].append('Folio Fiscal')
    excelObj['# Comprobante'].append('Comprobante')
    excelObj['Concepto facturado'].append('Concepto facturado')
    excelObj['Moneda'].append('Moneda')
    excelObj['Tipo de Cambio'].append('Tipo de Cambio')
    excelObj['Importe'].append('Importe')
    excelObj['0%'].append('0%')
    excelObj['IVA'].append('IVA')
    excelObj['IVA RETENIDO'].append('IVA RETENIDO')
    excelObj['IEPS'].append('IEPS')
    excelObj['Total'].append('Total')
    excelObj['# Cheque o transacción'].append('Cheque o transacción')
    excelObj['Fecha cargos'].append('Fecha cargos')
    excelObj['Nombre banco'].append('Nombre banco')
    excelObj['Referencia'].append('Referencia')

def emptyRowsExcel(excelObj):
    for i in range(2):
        excelObj['Proveedor'].append('')
        excelObj['RFC'].append('')
        excelObj['Folio Fiscal'].append('')
        excelObj['# Comprobante'].append('')
        excelObj['Concepto facturado'].append('')
        excelObj['Moneda'].append('')
        excelObj['Tipo de Cambio'].append('')
        excelObj['Importe'].append('')
        excelObj['0%'].append('')
        excelObj['IVA'].append('')
        excelObj['IVA RETENIDO'].append('')
        excelObj['IEPS'].append('')
        excelObj['Total'].append('')
        excelObj['# Cheque o transacción'].append('')
        excelObj['Fecha cargos'].append('')
        excelObj['Nombre banco'].append('')
        excelObj['Referencia'].append('')


#Vars
listNames = [] #Is neceseary for obtain the providers names
listProviders = {} #This is the most important var, this dictionary have all names, rfc, bills, etc...

#____________________________________________________________________________
#MAIN
#____________________________________________________________________________
#Open the file "DIOT" in the window PAGOS
base_path = Path(__file__).parent
file_path = (base_path/"resources/Diots/FORMATO DIOT JULIO 2024 2.xlsx").resolve()
excelarchive = pandas.read_excel(file_path, sheet_name = 'PAGOS')

#Columns 
#Obtain the columns of the excel 
providersNamesColumn = excelarchive.columns[18]
SheetColumn = excelarchive.columns[23]
ChequeColumn = excelarchive.columns[0]
DateColumn = excelarchive.columns[17]
Total_FilePurpleColumn = excelarchive.columns[14]

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
                              'folios': [], 
                              'pagosTodos': []}})

#Iterate the excel again, and add the rest info in the listProviders, if find a provider, add the flios, the date, the cheque, and the object for pagosTodos
for index, row in excelarchive.iterrows():
    listProviders[row[providersNamesColumn]]['folios'].append({
        'folio': row[SheetColumn],
        'fecha': row[DateColumn],
        'cheque': row[ChequeColumn],
        })
    pagosTodosObj = createPagosTodosObj(row[Total_FilePurpleColumn])
    listProviders[row[providersNamesColumn]]["pagosTodos"].append(pagosTodosObj)







for proveedorNombre in listProviders:
    #print(proveedorNombre) #// cada proveedor 
    Numfolios = len(listProviders[proveedorNombre]["folios"])
    listProviders[proveedorNombre]["foliosData"] = []
    for j in listProviders[proveedorNombre]["folios"]:
        #print(j) // cada folio del proveedor[i]
        listFolios = str(j['folio']).split("-")
        if(listFolios[0]== "F"):
            del listFolios[0]

        # Añadimos a FoliosTodos nuestro folio sin "-" y ademas le agrgamos los campos necesarios que dependen del folio
        for folio in listFolios:
            folioObj = createFolioObj(folio, j['fecha'], j['cheque'])
            listProviders[proveedorNombre]["foliosData"].append(folioObj)

    # Por ultimo eliminamos la lista de folios del provedor que ya no nos sirve 
    del listProviders[proveedorNombre]["folios"]

# Creamos una lista de objetos, el cual contendra el nombre del archivo, la pagina en donde se encuentra la info y las columnas que necesitamos
listaFilesInfo = [
    {
        # Archivo 4 2024
        "filename": "04 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 14,
        "columnIVARetenido": 20,
        "IEPScolumns": [13, 15, 16, 17, 18],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 5 2024
        "filename": "05 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 13,
        "columnIVARetenido": 20,
        "IEPScolumns": [14, 15, 16, 17, 18],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 6 2024
        "filename": "06 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 14,
        "columnIVARetenido": 20,
        "IEPScolumns": [13, 15, 16, 17, 18],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 7 2024
        "filename": "07 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 14,
        "columnIVARetenido": 21,
        "IEPScolumns": [13, 15, 16, 17, 18],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 8 2024
        "filename": "08 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 14,
        "columnIVARetenido": 22,
        "IEPScolumns": [13, 15, 16, 17, 19, 20],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 9 2024
        "filename": "09 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 14,
        "columnIVARetenido": 22,
        "IEPScolumns": [13, 15, 16, 17, 18, 20],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 10 2024
        "filename": "10 2024.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 14,
        "columnIVARetenido": 21,
        "IEPScolumns": [13, 15, 16, 18, 19],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
]


#listaFiles = []
for fileData in listaFilesInfo:
    print("Analizando el archivo: " + fileData['filename'])
    # Creamos el nombre del archivo que vamos a buscar 
    nombreArchivo = (base_path / "resources/XML" / fileData['filename']).resolve()
    #nombreArchivo = "C:\\Users\\User\\Desktop\\ProyectoB\\resources\\DEVOLUCION\\XML\\" + fileData['filename']
    # Abrimos el archivo
    excelarchiveNew = pandas.read_excel(nombreArchivo, sheet_name = fileData['sheetName'])
    
    # Columnas Archivos XML
    # Columna 1 (Nombre)
    columnProvedor = excelarchiveNew.columns[fileData['columnProvedor']]
    columnFolio = excelarchiveNew.columns[fileData['columnFolio']]
    columnRFC = excelarchiveNew.columns[fileData['columnRFC']]
    columnFFiscal = excelarchiveNew.columns[fileData['columnFFiscal']]
    columnConceptos = excelarchiveNew.columns[fileData['columnConceptos']]
    columnIVA = excelarchiveNew.columns[fileData['columnIVA']]
    columnIVARetenido = excelarchiveNew.columns[fileData['columnIVARetenido']]
    columnTOTAL = excelarchiveNew.columns[fileData['columnTOTAL']]
    columnsIEPS = []
    for i in range(len(fileData['IEPScolumns'])):
        columnsIEPS.append(excelarchiveNew.columns[fileData['IEPScolumns'][i]])
    

    # For que itera todo el excel 
    for j, row in excelarchiveNew.iterrows():
        #print(row[columnProvedor])  nombre de la empresa en los archivos xlsx
        for provedor in listProviders:
            # print(provedor) #// cada proveedor
            if(str(provedor).lower() == str(row[columnProvedor]).lower()):
                # Añadimos el RFC a todos 
                listProviders[provedor]["rfc"] = row[columnRFC]
                for folioData in listProviders[provedor]["foliosData"]:
                    if(str(folioData["folio"]).strip() == str(row[columnFolio]).strip() ):
                        # print("Provedor:" + provedor )
                        # print("Folio del DIOT:" + folioData["folio"] )
                        # print("Folio de archivo " + fileData['filename'] + ":" + row[columnFolio] +"\n")
                        # Agregamos los datos cuando encontro una coincidencia de folios
                        folioData["ffiscal"] = row[columnFFiscal]
                        folioData["concepto"] = row[columnConceptos]
                        folioData["iva"] = row[columnIVA]
                        folioData["ivaretenido"] = row[columnIVARetenido]
                        folioData["total"] = row[columnTOTAL]

                        #Calcular total ieps
                        folioData["totalIEPS"] = 0
                        for i in range(len(columnsIEPS)):
                            folioData["totalIEPS"] += row[columnsIEPS[i]]

                        

    # print("Analizando el archivo: " + fileData['filename'])
    # Creamos el nombre del archivo que vamos a buscar 
    nombreArchivo = (base_path / "resources/XML" / fileData['filename']).resolve()
    # Abrimos el archivo
    excelarchiveNew2 = pandas.read_excel(nombreArchivo, sheet_name = fileData['sheetNameP'])
    
    # Columnas Archivos XML
    columnPRFC = excelarchiveNew2.columns[fileData['columnPRFC']]
    columnPFolio = excelarchiveNew2.columns[fileData['columnPFolio']]
    columnPFFiscal = excelarchiveNew2.columns[fileData['columnPFFiscal']]
    columnPTOTAL = excelarchiveNew2.columns[fileData['columnPTOTAL']]

    # For que itera todo el excel 
    for j, row in excelarchiveNew2.iterrows():
        #print(row[columnProvedor])  nombre de la empresa en los archivos xlsx
        for provedor in listProviders:
            # print(provedor) #// cada proveedor
            # print(listProviders[provedor]["rfc"])
            # print(row[columnPRFC])
            # print(listProviders[provedor]["nombre"])
            # print(listProviders[provedor]["rfc"])
            # if(listProviders[provedor]["rfc"] == "TMM720509PYA"):
            #     print("Folio de archivo " + fileData['filename'] + ":" + row[columnFolio] +"\n")
            #     print(listProviders[provedor]["rfc"].lower())
            #     print(row[columnPRFC].lower())
            #     print(listProviders["3M MEXICO SA DE CV"]["pagosTodos"]["totalmorado"])
            #     print(row[columnPTOTAL])
            if(str(listProviders[provedor]["rfc"]).lower() == row[columnPRFC].lower()):
                for pagosTodosData in listProviders[provedor]["pagosTodos"]:
                    if(pagosTodosData["totalmorado"] - 3 <= row[columnPTOTAL]) and (row[columnPTOTAL] <= pagosTodosData["totalmorado"] + 3):
                    #if(97 <= 99) or (99 <= 103):
                        # print("ENTRO AL IF")
                        pagosTodosData["Pffiscal"] = row[columnPFFiscal]
                        pagosTodosData["Pfolio"] = row[columnPFolio]
                        pagosTodosData['Pcoincidencia'] = TRUE

                        
    

print("Creando archivo de Excel...")
    

# PARTE DE IMPRESION EN EXCEL PUTOSSSSSSS
# crear objeto necesario  para DataFrame
def createExcelData():
    return {
        'Proveedor': [], 
        'RFC': [], 
        'Folio Fiscal': [], 
        '# Comprobante': [], 
        'Concepto facturado': [], 
        'Moneda': [], 
        'Tipo de Cambio': [], 
        'Importe': [], 
        '0%': [], 
        'IVA': [], 
        'IVA RETENIDO': [],
        'IEPS': [], 
        'Total': [], 
        '# Cheque o transacción': [], 
        'Fecha cargos': [], 
        'Nombre banco': [],
        'Referencia': []
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
    ieps = 0
    total = 0
    for folioData in listProviders[provedor]["foliosData"]:
        excelObj['Proveedor'].append(listProviders[provedor]['nombre'])
        excelObj['RFC'].append(listProviders[provedor]['rfc'])
        excelObj['Folio Fiscal'].append(folioData['ffiscal'])
        excelObj['# Comprobante'].append(folioData['folio'])
        excelObj['Concepto facturado'].append(folioData['concepto'])
        excelObj['Moneda'].append(folioData['moneda'])
        excelObj['Tipo de Cambio'].append(folioData['tipocambio'])
        #excelObj['Importe'].append(folioData['iva']/0.16)
        excelObj['Importe'].append('=K' + str(numFilas) + '/0.16')
        excelObj['0%'].append('=N' + str(numFilas) + '-I' + str(numFilas) + '-K' + str(numFilas))
        excelObj['IVA'].append(folioData['iva'])
        excelObj['IVA RETENIDO'].append(folioData['ivaretenido'])
        excelObj['IEPS'].append(folioData['totalIEPS'])
        excelObj['Total'].append(folioData['total'])
        excelObj['# Cheque o transacción'].append(folioData['cheque'])
        excelObj['Fecha cargos'].append(folioData['fecha'])
        excelObj['Nombre banco'].append(folioData['banco'])
        excelObj['Referencia'].append('')
        importeactual = folioData['iva']/0.16
        importe = importe + (importeactual)
        ceroporcentaje = ceroporcentaje + (folioData['total'] - importeactual - folioData['iva'])
        iva = iva + (folioData['iva'])
        ivaretenido = ivaretenido + (folioData['ivaretenido'])
        ieps = ieps + (folioData['totalIEPS'])
        total = total + (folioData['total'])
        numFilas = numFilas + 1
    finProveedor = numFilas - 1
    excelObj['Proveedor'].append('')
    excelObj['RFC'].append('')
    excelObj['Folio Fiscal'].append('')
    excelObj['# Comprobante'].append('')
    excelObj['Concepto facturado'].append('')
    excelObj['Moneda'].append('')
    excelObj['Tipo de Cambio'].append('')
    excelObj['Importe'].append("$ " + str(importe))
    excelObj['0%'].append("$ " + str(ceroporcentaje))
    excelObj['IVA'].append("$ " + str(iva))
    excelObj['IVA RETENIDO'].append("$ " + str(ivaretenido))
    excelObj['IEPS'].append("$ " + str(ieps))
    excelObj['Total'].append("$ " + str(total))
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
    emptyRowsExcel(excelObj)
    numFilas = numFilas + 3

    excelObj['Proveedor'].append('Proveedor')
    excelObj['RFC'].append('RFC')
    excelObj['Folio Fiscal'].append('Folio Fiscal')
    excelObj['# Comprobante'].append('# Comprobante')
    excelObj['Concepto facturado'].append('Concepto facturado')
    excelObj['Moneda'].append('')
    excelObj['Tipo de Cambio'].append('')
    excelObj['Importe'].append('')
    excelObj['0%'].append('')
    excelObj['IVA'].append('')
    excelObj['IVA RETENIDO'].append('')
    excelObj['IEPS'].append('')
    excelObj['Total'].append('')
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
    listRows_titles.append(numFilas-1)
    numFilas = numFilas + 1
    for PagoData in listProviders[provedor]["pagosTodos"]:
        if PagoData['Pcoincidencia'] == TRUE:
            excelObj['Proveedor'].append(listProviders[provedor]['nombre'])
            excelObj['RFC'].append(listProviders[provedor]['rfc'])
            excelObj['Folio Fiscal'].append(PagoData['Pffiscal'])
            excelObj['# Comprobante'].append(PagoData['Pfolio'])
            excelObj['Concepto facturado'].append('PAGO')
            excelObj['Moneda'].append('')
            excelObj['Tipo de Cambio'].append('')
            excelObj['Importe'].append('')
            excelObj['0%'].append('')
            excelObj['IVA'].append('')
            excelObj['IVA RETENIDO'].append('')
            excelObj['IEPS'].append('')
            excelObj['Total'].append('')
            excelObj['# Cheque o transacción'].append('')
            excelObj['Fecha cargos'].append('')
            excelObj['Nombre banco'].append('')
            excelObj['Referencia'].append('')
            numFilas = numFilas + 1
    emptyRowsExcel(excelObj)
    numFilas = numFilas + 2

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
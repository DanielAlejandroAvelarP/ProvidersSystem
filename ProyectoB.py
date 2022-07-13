# Biblioteca para Excel
from asyncio.windows_events import NULL
from pickle import FALSE, TRUE
import pandas
import numpy as np
from collections import OrderedDict
from datetime import datetime

# Abrimos el archivo de Excel
excelarchive = pandas.read_excel("C:\\Users\\User\\Desktop\\ProyectoB\\resources\\DEVOLUCION\\FORMATO_DIOT_OCTUBRE_2021_FARMACIA_TEPA.xlsx", sheet_name = 'PAGOS')

# Guardamos la columna 15 (Nombres)
column15 = excelarchive.columns[15]

# Creamos nuestra lista
listaNom = []

#Creamos un For con al info de la columna y le pasamos la info a nuestra lista
for index, row in excelarchive.iterrows():
    listaNom.append(row[column15])

# Ahora que tenemos en nuestra lista todos los nombres eliminamos los repetidos 
listaNomFil= list(OrderedDict.fromkeys(listaNom))

# Creamos una lista y un objeto
# listaProvedores=[]
listaProvedores2={}

#Actualizamos nuestra liista de proovedores y creamos un diccionarios con diccionarios dentro y le agregamos el nombre
for nombre in listaNomFil:
    # listaProvedores2.update({nombre: {'nombre': nombre, 'rfc': [], 'ffiscal': [],'folios': [], 'concepto': [], 'moneda': "MN", 'tipocambio': "1", 'importe': "0", '0%': "0", 'iva': [], 'ivaRetenido': "0", 'total': [], 'cheque': [], 'fecha': [], 'banco': "BANORTE", 'totalMorado': []}})
    listaProvedores2.update({nombre: {'nombre': nombre, 'rfc': NULL,'folios': [], 'pagosTodos': []}})


    # {
    #     nombre: 'qwerty',
    #     rfc: Null,
    #     folios: [],
        # pagosTodos [
            # TotalMorado
            # folio 
            # foliofiscal
            # coincidencia = False 
        # ]
    #     folioftodos; [
    #         {
    #             folio : 'jjjjm',
    #              iva: '',
    #               ...
    #         }
    #     ]
    # }

# foliosarray[
#     {
#         folio: '1'
#     },
#     {
#         folio: '2abc'
#     },
#     {
#         folio: '3'
#     }
# ]

# foliosdiccionario: {
#     '1': {
#         folio: '1'
#     },
#     '2abc': {
#         folio: '2'
#     },
#      '3': {
#         folio: '3'
#     },
# }

# foliosdiccionario['2abc']
# foliosarray

# Columnas Archivo Morado
# Guardamos la columna 20 (Folios)
column20 = excelarchive.columns[20]
# Columna 0 (Cheque)
column0 = excelarchive.columns[0]
# Columna 14 (Fech}a)
column14 = excelarchive.columns[14]
# Columna 11 (Morado)
column11 = excelarchive.columns[11]


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


# For que itera todo el excel y añadimos a cada proovedor sus folios
for index, row in excelarchive.iterrows():
    # Folios
    listaProvedores2[row[column15]]['folios'].append({
        'folio': row[column20],
        'fecha': row[column14],
        'cheque': row[column0],
        })
    pagosTodosObj = createPagosTodosObj(row[column11])
    listaProvedores2[row[column15]]["pagosTodos"].append(pagosTodosObj)
    # Cheque
    # listaProvedores2[row[column15]]['cheque'].append(row[column0])
    # Fecha
    # listaProvedores2[row[column15]]['fecha'].append(row[column14])
    # Total Morado
    # listaProvedores2[row[column15]]['totalMorado'].append(row[column11])


for proveedorNombre in listaProvedores2:
    #print(proveedorNombre) #// cada proveedor 
    Numfolios = len(listaProvedores2[proveedorNombre]["folios"])
    listaProvedores2[proveedorNombre]["foliosData"] = []
    for j in listaProvedores2[proveedorNombre]["folios"]:
        #print(j) // cada folio del proveedor[i]
        listFolios = str(j['folio']).split("-")
        if(listFolios[0]== "F"):
            del listFolios[0]

        # Añadimos a FoliosTodos nuestro folio sin "-" y ademas le agrgamos los campos necesarios que dependen del folio
        for folio in listFolios:
            #listaProvedores2[proveedorNombre]["foliosData"].append({
            # 'folio': folio, 
            # 'ffiscal': NULL, 
            # 'concepto': NULL, 
            # 'moneda': "MN", 
            # 'tipocambio': "1", 
            # 'importe': "0", 
            # '0%': "0", 
            # 'iva': NULL, 
            # 'ivaretendio': "0", 
            # 'total': NULL, 
            # 'cheque': NULL,
            #  'fecha': j['fecha'], 
            # 'banco': "BANORTE", '
            # totalmorado': NULL})
            folioObj = createFolioObj(folio, j['fecha'], j['cheque'])
            listaProvedores2[proveedorNombre]["foliosData"].append(folioObj)

        # listaProvedores2[proveedorNombre]["foliosTodos"] = np.concatenate((listaProvedores2[proveedorNombre]["foliosTodos"], listFolios))
    # Por ultimo eliminamos la lista de folios del provedor que ya no nos sirve 
    del listaProvedores2[proveedorNombre]["folios"]

# For que itera todo el excel y añadimos a cada proovedor sus folios
# for i, row in excelarchive.iterrows():
#     for j in listaProvedores2[proveedorNombre]["foliosData"]:
            


# print(listaProvedores2['BEBIDAS PURIFICADAS S DE RL DE CV'])
# for i in listaProvedores2:
#     print("\n")
#     print(listaProvedores2[i]["foliosTodos"])

# Creamos un For para abrir los 10 archivos de Excel con los que vamos a obtener la info que nos falta 
#listaFiles = ["1.xls", "2.xls","3.xls","4.xlsx","5.xlsx","6.xlsx","7.xlsx","8.xlsx","9.xlsx","10.xlsx" ]
# listaFiles = ["4.xlsx","5.xlsx","6.xlsx","7.xlsx","8.xlsx","9.xlsx","10.xlsx" ]
# Creamos una lista de objetos, el cual contendra el nombre del archivo, la pagina en donde se encuentra la info y las columnas que necesitamos
listaFilesInfo = [
    # {
    #     # Archivo 4
    #     "filename": "4.xlsx",
    #     "sheetName": "I",
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnRFC": 0,
    #     "columnFFiscal": 9,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 21,
    #     "columnTOTAL": 8,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 1,
    #     "columnPFolio": 7,
    #     "columnPTOTAL": 15,
    #     "columnPFFiscal": 16,
    # },
    # {
    #     # Archivo 5
    #     "filename": "5.xlsx",
    #     "sheetName": "I",
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnRFC": 0,
    #     "columnFFiscal": 9,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 19,
    #     "columnTOTAL": 8,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 8,
    #     "columnPFFiscal": 9,
    # },
    # {
    #     # Archivo 6
    #     "filename": "6.xlsx",
    #     "sheetName": "I",
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnRFC": 0,
    #     "columnFFiscal": 8,
    #     "columnConceptos": 11,
    #     "columnIVA": 12,
    #     "columnIVARetenido": 18,
    #     "columnTOTAL": 7,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 7,
    #     "columnPFFiscal": 8,
    # },
    # {
    #     # Archivo 7
    #     "filename": "7.xlsx",
    #     "sheetName": "I",
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnRFC": 0,
    #     "columnFFiscal": 9,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 21,
    #     "columnTOTAL": 8,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 3,
    #     "columnPTOTAL": 5,
    #     "columnPFFiscal": 6,
    # },
    {
        # Archivo 8
        "filename": "8.xlsx",
        "sheetName": "I",
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnRFC": 0,
        "columnFFiscal": 8,
        "columnConceptos": 12,
        "columnIVA": 13,
        "columnIVARetenido": 21,
        "columnTOTAL": 7,
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {  
        # Archivo 9
        "filename": "9.xlsx",
        "sheetName": "I",
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnRFC": 0,
        "columnFFiscal": 8,
        "columnConceptos": 11,
        "columnIVA": 13,
        "columnIVARetenido": 21,
        "columnTOTAL": 7,
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5, 
    },
    {
        # Archivo 10
        "filename": "10.xlsx",
        "sheetName": "I",
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnRFC": 0,
        "columnFFiscal": 9,
        "columnConceptos": 13,
        "columnIVA": 14,
        "columnIVARetenido": 21,
        "columnTOTAL": 8,
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 5,
        "columnPFFiscal": 6,
    }
]


#listaFiles = []
for fileData in listaFilesInfo:
    print("Analizando el archivo: " + fileData['filename'])
    # Creamos el nombre del archivo que vamos a buscar 
    nombreArchivo = "C:\\Users\\User\\Desktop\\ProyectoB\\resources\\DEVOLUCION\\XML\\" + fileData['filename']
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

    # For que itera todo el excel 
    for j, row in excelarchiveNew.iterrows():
        #print(row[columnProvedor])  nombre de la empresa en los archivos xlsx
        for provedor in listaProvedores2:
            # print(provedor) #// cada proveedor
            if(str(provedor).lower() == str(row[columnProvedor]).lower()):
                # Añadimos el RFC a todos 
                listaProvedores2[provedor]["rfc"] = row[columnRFC]
                for folioData in listaProvedores2[provedor]["foliosData"]:
                    if(folioData["folio"] == row[columnFolio]):
                        # print("Provedor:" + provedor )
                        # print("Folio del DIOT:" + folioData["folio"] )
                        # print("Folio de archivo " + fileData['filename'] + ":" + row[columnFolio] +"\n")
                        # Agregamos los datos cuando encontro una coincidencia de folios
                        folioData["ffiscal"] = row[columnFFiscal]
                        folioData["concepto"] = row[columnConceptos]
                        folioData["iva"] = row[columnIVA]
                        folioData["ivaretenido"] = row[columnIVARetenido]
                        folioData["total"] = row[columnTOTAL]
            



    # print("Analizando el archivo: " + fileData['filename'])
    # Creamos el nombre del archivo que vamos a buscar 
    nombreArchivo = "C:\\Users\\User\\Desktop\\ProyectoB\\resources\\DEVOLUCION\\XML\\" + fileData['filename']
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
        for provedor in listaProvedores2:
            # print(provedor) #// cada proveedor
            # print(listaProvedores2[provedor]["rfc"])
            # print(row[columnPRFC])
            # print(listaProvedores2[provedor]["nombre"])
            # print(listaProvedores2[provedor]["rfc"])
            # if(listaProvedores2[provedor]["rfc"] == "TMM720509PYA"):
            #     print("Folio de archivo " + fileData['filename'] + ":" + row[columnFolio] +"\n")
            #     print(listaProvedores2[provedor]["rfc"].lower())
            #     print(row[columnPRFC].lower())
            #     print(listaProvedores2["3M MEXICO SA DE CV"]["pagosTodos"]["totalmorado"])
            #     print(row[columnPTOTAL])
            if(str(listaProvedores2[provedor]["rfc"]).lower() == row[columnPRFC].lower()):
                for pagosTodosData in listaProvedores2[provedor]["pagosTodos"]:
                    if(pagosTodosData["totalmorado"] - 3 <= row[columnPTOTAL]) and (row[columnPTOTAL] <= pagosTodosData["totalmorado"] + 3):
                    #if(97 <= 99) or (99 <= 103):
                        # print("ENTRO AL IF")
                        pagosTodosData["Pffiscal"] = row[columnPFFiscal]
                        pagosTodosData["Pfolio"] = row[columnPFolio]
                        pagosTodosData['Pcoincidencia'] = TRUE

                            


#Datos para  la parte de PAGGOS


# print(listaProvedores2['BEBIDAS PURIFICADAS S DE RL DE CV'])
#Dividimos los folios que se encuentren unidos por un "-"
# for i in listaProvedores2:
    # print(i)
    # Numfolios = len(listaProvedores2[index]["folios"])
    # for j in listaProvedores2[i]["folios"]:
        # print(str(j).split("-"))
        # print("hola")

    

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
        'Total': [], 
        '# Cheque o transacción': [], 
        'Fecha cargos': [], 
        'Nombre banco': [],
        'Referencia': []
    }



excelObj = createExcelData()
numFilas = 2
for provedor in listaProvedores2:
    excelObj['Proveedor'].append('Proveedor')
    excelObj['RFC'].append('RFC')
    excelObj['Folio Fiscal'].append('Folio Fiscal')
    excelObj['# Comprobante'].append('# Comprobante')
    excelObj['Concepto facturado'].append('Concepto facturado')
    excelObj['Moneda'].append('Moneda')
    excelObj['Tipo de Cambio'].append('Tipo de Cambio')
    excelObj['Importe'].append('Importe')
    excelObj['0%'].append('0%')
    excelObj['IVA'].append('IVA')
    excelObj['IVA RETENIDO'].append('IVA RETENIDO')
    excelObj['Total'].append('Total')
    excelObj['# Cheque o transacción'].append('# Cheque o transacción')
    excelObj['Fecha cargos'].append('Fecha cargos')
    excelObj['Nombre banco'].append('Nombre banco')
    excelObj['Referencia'].append('Referencia')
    numFilas = numFilas + 1
    inicioProveedor = numFilas
    importe = 0
    ceroporcentaje = 0 
    iva = 0
    ivaretenido = 0
    total = 0
    for folioData in listaProvedores2[provedor]["foliosData"]:
        excelObj['Proveedor'].append(listaProvedores2[provedor]['nombre'])
        excelObj['RFC'].append(listaProvedores2[provedor]['rfc'])
        excelObj['Folio Fiscal'].append(folioData['ffiscal'])
        excelObj['# Comprobante'].append(folioData['folio'])
        excelObj['Concepto facturado'].append(folioData['concepto'])
        excelObj['Moneda'].append(folioData['moneda'])
        excelObj['Tipo de Cambio'].append(folioData['tipocambio'])
        #excelObj['Importe'].append(folioData['iva']/0.16)
        excelObj['Importe'].append('=K' + str(numFilas) + '/0.16')
        excelObj['0%'].append('=M' + str(numFilas) + '-I' + str(numFilas) + '-K' + str(numFilas))
        excelObj['IVA'].append(folioData['iva'])
        excelObj['IVA RETENIDO'].append(folioData['ivaretenido'])
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
    excelObj['Importe'].append(importe)
    excelObj['0%'].append(ceroporcentaje)
    excelObj['IVA'].append(iva)
    excelObj['IVA RETENIDO'].append(ivaretenido)
    excelObj['Total'].append(total)
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
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
    excelObj['Total'].append('')
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
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
    excelObj['Total'].append('')
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
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
    excelObj['Total'].append('')
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
    numFilas = numFilas + 1
    for PagoData in listaProvedores2[provedor]["pagosTodos"]:
        # print(PagoData['Pcoincidencia'])
        # print(PagoData['Pffiscal'])
        # print(PagoData['Pfolio'])
        if PagoData['Pcoincidencia'] == TRUE:
            excelObj['Proveedor'].append(listaProvedores2[provedor]['nombre'])
            excelObj['RFC'].append(listaProvedores2[provedor]['rfc'])
            excelObj['Folio Fiscal'].append(PagoData['Pffiscal'])
            excelObj['# Comprobante'].append(PagoData['Pfolio'])
            excelObj['Concepto facturado'].append('PAGO')
            excelObj['Moneda'].append('')
            excelObj['Tipo de Cambio'].append('')
            excelObj['Importe'].append('')
            excelObj['0%'].append('')
            excelObj['IVA'].append('')
            excelObj['IVA RETENIDO'].append('')
            excelObj['Total'].append('')
            excelObj['# Cheque o transacción'].append('')
            excelObj['Fecha cargos'].append('')
            excelObj['Nombre banco'].append('')
            excelObj['Referencia'].append('')
            numFilas = numFilas + 1
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
    excelObj['Total'].append('')
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
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
    excelObj['Total'].append('')
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
    numFilas = numFilas + 2


df = pandas.DataFrame(excelObj)
# crear el objeto ExcelWriter
escrito = pandas.ExcelWriter('Anexo_1_proveedores_OCTUBRE.xlsx')
# escribir el DataFrame en excel
df.to_excel(escrito,'Try')
# guardar el excel
escrito.save()
print("Archivo de Excel creado Exitosamente.")
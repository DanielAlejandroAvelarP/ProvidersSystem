# Biblioteca para Excel
from asyncio.windows_events import NULL
from pickle import FALSE, TRUE
import pandas
import numpy as np
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

# Abrimos el archivo de Excel
base_path = Path(__file__).parent
file_path = (base_path / "resources/DEVOLUCION/FORMATO DIOT MARZO 2022.xlsx").resolve()
#path = os.path.join(base_path, "../resources/DEVOLUCION/FORMATO_DIOT_OCTUBRE_2021_FARMACIA_TEPA.xlsx")
excelarchive = pandas.read_excel(file_path, sheet_name = 'PAGOS')

# Guardamos la columna 15 (Nombres)
column15 = excelarchive.columns[15]

# Creamos nuestra lista
listaNom = []

#Creamos un For con al info de la columna y le pasamos la info a nuestra lista
for index, row in excelarchive.iterrows():
    listaNom.append(row[column15])
listaNom.sort();
# Ahora que tenemos en nuestra lista todos los nombres eliminamos los repetidos 
listaNomFil= list(OrderedDict.fromkeys(listaNom))

# Creamos una lista y un objeto
# listaProvedores=[]
listaProvedores2={}

#Actualizamos nuestra liista de proovedores y creamos un diccionarios con diccionarios dentro y le agregamos el nombre
for nombre in listaNomFil:
    # listaProvedores2.update({nombre: {'nombre': nombre, 'rfc': [], 'ffiscal': [],'folios': [], 'concepto': [], 'moneda': "MN", 'tipocambio': "1", 'importe': "0", '0%': "0", 'iva': [], 'ivaRetenido': "0", 'total': [], 'cheque': [], 'fecha': [], 'banco': "BANORTE", 'totalMorado': []}})
    listaProvedores2.update({nombre: {'nombre': nombre, 'rfc': NULL,'folios': [], 'pagosTodos': []}})

# Columnas Archivo Morado
# Guardamos la columna 20 (Folios)
column20 = excelarchive.columns[20]
# Columna 0 (Cheque)
column0 = excelarchive.columns[0]
# Columna 14 (Fecha)
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

# Creamos una lista de objetos, el cual contendra el nombre del archivo, la pagina en donde se encuentra la info y las columnas que necesitamos
listaFilesInfo = [
    # {
    #     # Archivo 05 2020
    #     "filename": "05 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 19,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 06 2022
    #     "filename": "06-2022.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 16,
    #     "columnIVARetenido": 22,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 3,
    #     "columnPTOTAL": 5,
    #     "columnPFFiscal": 6,
    # },
    # {
    #     # Archivo 05 2020
    #     "filename": "05 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 19,
    #     "IEPScolumns": [13, 15, 16, 17, 18],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 06 2020
    #     "filename": "06 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 20,
    #     "IEPScolumns": [13, 15, 16, 17, 18],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 07 2020
    #     "filename": "07 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 13,
    #     "columnIVARetenido": 20,
    #     "IEPScolumns": [14, 15, 16, 17, 18],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 08 2020
    #     "filename": "08 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 10,
    #     "columnFFiscal": 11,
    #     "columnConceptos": 13,
    #     "columnIVA": 15,
    #     "columnIVARetenido": 21,
    #     "IEPScolumns": [14, 16, 17, 18, 19],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 09 2020
    #     "filename": "09 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 10,
    #     "columnFFiscal": 11,
    #     "columnConceptos": 13,
    #     "columnIVA": 15,
    #     "columnIVARetenido": 20,
    #     "IEPScolumns": [14, 16, 17, 18, 19],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 10 2020
    #     "filename": "10 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 10,
    #     "columnFFiscal": 11,
    #     "columnConceptos": 13,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 21,
    #     "IEPScolumns": [15, 16, 17, 18, 19],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 11 2020
    #     "filename": "11 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 10,
    #     "columnFFiscal": 11,
    #     "columnConceptos": 13,
    #     "columnIVA": 15,
    #     "columnIVARetenido": 22,
    #     "IEPScolumns": [14, 16, 17, 18, 19, 20],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 12 2020
    #     "filename": "12 2020.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 21,
    #     "IEPScolumns": [13, 15, 16, 17, 18, 19],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 1 2021
    #     "filename": "01.xls",
    #     "sheetName": "I",
    #     "columnProvedor": 2,
    #     "columnFolio": 7,
    #     "columnRFC": 1,
    #     "columnFFiscal": 16,
    #     "columnConceptos": 27,
    #     "columnIVA": 28,
    #     "IEPScolumns": [29, 30, 31, 32, 33, 34],
    #     "columnIVARetenido": 36,
    #     "columnTOTAL": 15,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 3,
    #     "columnPTOTAL": 5,
    #     "columnPFFiscal": 6,
    # },
    # {
    #     # Archivo 2 2021
    #     "filename": "02.xls",
    #     "sheetName": "I",
    #     "columnProvedor": 2,
    #     "columnFolio": 4,
    #     "columnRFC": 1,
    #     "columnFFiscal": 11,
    #     "columnConceptos": 13,
    #     "columnIVA": 15,
    #     "IEPScolumns": [14, 16, 17, 18, 19, 20],
    #     "columnIVARetenido": 22,
    #     "columnTOTAL": 10,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 4,
    #     "columnPTOTAL": 6,
    #     "columnPFFiscal": 7,
    # },
    # {
    #     # Archivo 3 2021
    #     "filename": "03.xls",
    #     "sheetName": "I",
    #     "columnProvedor": 2,
    #     "columnFolio": 4,
    #     "columnRFC": 1,
    #     "columnFFiscal": 11,
    #     "columnConceptos": 14,
    #     "columnIVA": 16,
    #     "IEPScolumns": [15, 17, 18, 19],
    #     "columnIVARetenido": 21,
    #     "columnTOTAL": 10,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 1,
    #     "columnPFolio": 5,
    #     "columnPTOTAL": 8,
    #     "columnPFFiscal": 9,
    # },
    # {
    #     # Archivo 4 2021
    #     "filename": "04.xlsx",
    #     "sheetName": "I",
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnRFC": 0,
    #     "columnFFiscal": 7,
    #     "columnConceptos": 10,
    #     "columnIVA": 12,
    #     "IEPScolumns": [11, 13, 14, 15, 16, 17],
    #     "columnIVARetenido": 19,
    #     "columnTOTAL": 6,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 1,
    #     "columnPFolio": 5,
    #     "columnPTOTAL": 7,
    #     "columnPFFiscal": 8,
    # },
    # {
    #     # Archivo 5 2021
    #     "filename": "05.xlsx",
    #     "sheetName": "I",
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnRFC": 0,
    #     "columnFFiscal": 9,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "IEPScolumns": [13, 15, 16, 17, 18],
    #     "columnIVARetenido": 19,
    #     "columnTOTAL": 8,
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 3,
    #     "columnPTOTAL": 5,
    #     "columnPFFiscal": 6,
    # },
    {
        # Archivo 01 2022
        "filename": "01 2022.xls",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 8,
        "columnFFiscal": 9,
        "columnConceptos": 10,
        "columnIVA": 12,
        "columnIVARetenido": 19,
        "IEPScolumns": [11, 13, 14, 15, 16, 17],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 5,
        "columnPFFiscal": 6,
    },
    {
        # Archivo 02 2022
        "filename": "02 2022.xls",
        "sheetName": "Ingresos",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 13,
        "columnIVARetenido": 21,
        "IEPScolumns": [14, 15, 16, 17, 18, 19],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 03 2022
        "filename": "03 2022.xls",
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
        # Archivo 04 2022
        "filename": "04 2022.xls",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 13,
        "columnIVARetenido": 19,
        "IEPScolumns": [14, 15, 16, 17],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 2,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 05 2022
        "filename": "05 2022.xls",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 13,
        "columnIVARetenido": 17,
        "IEPScolumns": [14, 15],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 5,
        "columnPFFiscal": 6,
    },
    {
        # Archivo 6 2022
        "filename": "06 2022.xls",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 12,
        "columnIVA": 15,
        "columnIVARetenido": 21,
        "IEPScolumns": [14, 16, 17, 18, 19],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 5,
        "columnPFFiscal": 6,
    },
    {
        # Archivo 9
        "filename": "09.xlsx",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 7,
        "columnFFiscal": 8,
        "columnConceptos": 10,
        "columnIVA": 12,
        "columnIVARetenido": 20,
        "IEPScolumns": [11, 13, 14, 15, 16, 17, 18],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 0,
        "columnPFolio": 3,
        "columnPTOTAL": 4,
        "columnPFFiscal": 5,
    },
    {
        # Archivo 12
        "filename": "12.xls",
        "sheetName": "I",
        "columnRFC": 0,
        "columnProvedor": 1,
        "columnFolio": 2,
        "columnTOTAL": 9,
        "columnFFiscal": 10,
        "columnConceptos": 11,
        "columnIVA": 12,
        "columnIVARetenido": 19,
        "IEPScolumns": [13, 14, 15, 16, 17, 18],
        # Pagina P
        "sheetNameP": "P",
        "columnPRFC": 1,
        "columnPFolio": 3,
        "columnPTOTAL": 5,
        "columnPFFiscal": 6,
    },
    # {
    #     # Archivo 7 2022
    #     "filename": "07 2022.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 8,
    #     "columnFFiscal": 9,
    #     "columnConceptos": 11,
    #     "columnIVA": 12,
    #     "columnIVARetenido": 17,
    #     "IEPScolumns": [13, 14, 15],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 8 2022
    #     "filename": "08 2022.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 13,
    #     "columnIVARetenido": 18,
    #     "IEPScolumns": [14, 15, 16],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 9 2022
    #     "filename": "09 2022.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 13,
    #     "columnIVARetenido": 17,
    #     "IEPScolumns": [14, 15],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 10 2022
    #     "filename": "10 2022.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 13,
    #     "columnIVARetenido": 17,
    #     "IEPScolumns": [14, 15],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 11 2022
    #     "filename": "11 2022.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 13,
    #     "columnIVARetenido": 19,
    #     "IEPScolumns": [14, 15, 16, 17],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 12 2022
    #     "filename": "12 2022.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 18,
    #     "IEPScolumns": [13, 15, 16, 17],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 1 2023
    #     "filename": "01 2023.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 16,
    #     "IEPScolumns": [13, 15],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 2 2023
    #     "filename": "02 2023.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 19,
    #     "IEPScolumns": [13, 15, 16, 17],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 3 2023
    #     "filename": "03 2023.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 14,
    #     "columnIVARetenido": 20,
    #     "IEPScolumns": [13, 15, 16, 17, 18],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # },
    # {
    #     # Archivo 4 2023
    #     "filename": "04 2023.xls",
    #     "sheetName": "I",
    #     "columnRFC": 0,
    #     "columnProvedor": 1,
    #     "columnFolio": 2,
    #     "columnTOTAL": 9,
    #     "columnFFiscal": 10,
    #     "columnConceptos": 12,
    #     "columnIVA": 13,
    #     "columnIVARetenido": 19,
    #     "IEPScolumns": [14, 15, 16, 17],
    #     # Pagina P
    #     "sheetNameP": "P",
    #     "columnPRFC": 0,
    #     "columnPFolio": 2,
    #     "columnPTOTAL": 4,
    #     "columnPFFiscal": 5,
    # }
]


#listaFiles = []
for fileData in listaFilesInfo:
    print("Analizando el archivo: " + fileData['filename'])
    # Creamos el nombre del archivo que vamos a buscar 
    nombreArchivo = (base_path / "resources/DEVOLUCION/XML" / fileData['filename']).resolve()
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
        for provedor in listaProvedores2:
            # print(provedor) #// cada proveedor
            if(str(provedor).lower() == str(row[columnProvedor]).lower()):
                # Añadimos el RFC a todos 
                listaProvedores2[provedor]["rfc"] = row[columnRFC]
                for folioData in listaProvedores2[provedor]["foliosData"]:
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
    nombreArchivo = (base_path / "resources/DEVOLUCION/XML" / fileData['filename']).resolve()
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
for provedor in listaProvedores2:
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
    listRows_titles.append(numFilas-1)
    numFilas = numFilas + 1
    inicioProveedor = numFilas
    importe = 0
    ceroporcentaje = 0 
    iva = 0
    ivaretenido = 0
    ieps = 0
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
            excelObj['IEPS'].append('')
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
    excelObj['IEPS'].append('')
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
    excelObj['IEPS'].append('')
    excelObj['Total'].append('')
    excelObj['# Cheque o transacción'].append('')
    excelObj['Fecha cargos'].append('')
    excelObj['Nombre banco'].append('')
    excelObj['Referencia'].append('')
    numFilas = numFilas + 2

df = pandas.DataFrame(excelObj)
# crear el objeto ExcelWriter
escrito = pandas.ExcelWriter('Anexo_1_proveedores.xlsx', engine='xlsxwriter')
# escribir el DataFrame en excel
#df.to_excel(escrito,'Try', startrow=0, startcol=0, header=False, index=False)
df.to_excel(escrito,'Try', startrow=0, startcol=1, header=False, index=False)
worksheet = escrito.sheets['Try']
cell_format = escrito.book.add_format({
    'align':    'center',
    'valign':   'vcenter',
    'bg_color':   '#92D050'})
cell_format_yellow = escrito.book.add_format()
cell_format_yellow.set_bg_color('#FFFF00')
currency_format = escrito.book.add_format({'num_format': '_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-'})
# worksheet.conditional_format('A1:Q8000',{ 
#     'type': 'cell', 'criteria': '==', 'value': '"RFC"', 
#     'format': cell_format
#     })
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
    
# guardar el excel
escrito.close()
print("Archivo de Excel creado Exitosamente.")
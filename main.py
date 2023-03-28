import pandas as pd
import fileexplorer as fe
import sharepoint as sp
import time
from threading import Thread
import tkinter
import asyncio
tik = time.perf_counter()
try:
    excelPath = fe.getFilePath("Select Excel File")
except:
    print("Didn't select a file. Exiting program.")
    exit()
if(fe.isExcelFile(excelPath)):
    try:
        excelDoc = pd.ExcelFile(excelPath)
        print(excelDoc.sheet_names)
        sheetname = str(input("Select Sheet: "))
        excelDoc = pd.read_excel(excelPath, sheet_name=sheetname)
        
    except:
        print("Sheet does not exist. Exiting Program.")
        exit()
    site =sp.getSite("https://cernerprod.sharepoint.com/sites/centralizeddesktopsupport/", "RA1774524@cerner.net", "Ccarma50")
    list = site.List("Testlist", exclude_hidden_fields=False) 
    sp.addHeadersConditional(excelSheet=excelDoc, list=list)
    sp.addDataToList(list=list, excelSheet=excelDoc, threaded=True, sortThreads=10)
    tok = time.perf_counter()
    print("Process took: " + str((tok - tik)/60)+ " mins.")
else:
    print("Not an excel file. Exiting program.")

  

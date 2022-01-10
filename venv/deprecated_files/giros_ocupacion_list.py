import pandas as pd
import openpyxl as opyxl
from giros import giros_list

giros_ocupacion_list=[]

def loadDummys():
    try:
        #Getting the dummy data
        wb = opyxl.load_workbook("../dummydata_ocupacion.xlsx")
        #Updating my template with the new Giros
        ws = wb.worksheets[0]
        for i in range(0,ws.max_row):
            giros_ocupacion_list.append([ws.cell(i+1,0+1).value, ws.cell(i+1,1+1).value])
        wb.close()
        return giros_ocupacion_list
    except Exception as ex:
        return []

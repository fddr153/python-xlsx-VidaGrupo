import pandas as pd
import openpyxl as opyxl
from giros import giros_list
from giros_ocupacion_list import loadDummys

def updateTemplate():
    try:
        #Getting the master template
        wb = opyxl.load_workbook("../master_template.xlsx")
        #Updating my template with the new Giros
        ws = wb.worksheets[4]
        ws2=wb.worksheets[5]
        i=0
        for giro in giros_list:
            ws.cell(row=i+2, column=1).value = giros_list[i]
            i+=1
        #Update the Template with the Ocupacion fields
        ocupacion_list= loadDummys()
        i=0
        for row in ocupacion_list:
            ws2.cell(row=i+1, column=1).value = ocupacion_list[i][0]
            ws2.cell(row=i+1, column=2).value = ocupacion_list[i][1]
            i+=1
        #salv el archivo final
        wb.save("testing.xlsx")
        return jsonify({"message":"succes"})
    except Exception as ex:
        return ex


updateTemplate()
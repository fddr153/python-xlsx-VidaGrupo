from flask import Flask, jsonify, request, Response, make_response
import numpy as np
import pandas as pd
import openpyxl as opyxl
from openpyxl.comments import Comment
from openpyxl.utils import units
from getlists import getBothLists
from validations import doCellValidations
app = Flask(__name__)



# Testing Route
@app.route('/ping', methods=['GET'])
def ping():
    return jsonify({'response': 'pong!'})

# Create Data Routes
@app.route('/formulario-usuario', methods=['POST'])
def getfile():
    try:
        #Getting the custom xlsx uploaded by the user
        file = request.files['xlsx_file']
        foo = file.filename
        unparsedFile = file.read()
        dframe = pd.ExcelFile(unparsedFile)
        #Getting the sheets from uploaded file
        file_sheets = []
        for sheet in dframe.sheet_names:
            file_sheets.append(dframe.parse(sheet))
        # Uploading my template and inserting vales of Giros and Ocupacion
        # Getting the master template
        wb = opyxl.load_workbook("master_template.xlsx")
        # Gettin' my template sheets with the new Giros and Ocupacion
        ws = wb.worksheets[4]
        ws2 = wb.worksheets[5]
        giros_list, ocupacion_list = getBothLists()
        # Update the Template with the Giros fields
        i = 0
        for giro in giros_list:
            ws.cell(row=i + 2, column=1).value = giros_list[i]
            i += 1
        # Update the Template with the Ocupacion fields
        i = 0
        for row in ocupacion_list:
            ws2.cell(row=i + 1, column=1).value = ocupacion_list[i][0]
            ws2.cell(row=i + 1, column=2).value = ocupacion_list[i][1]
            i += 1
        #Hidding my Configuration Sheets
        wb.worksheets[2].sheet_state = wb.worksheets[2].SHEETSTATE_VERYHIDDEN
        wb.worksheets[3].sheet_state = wb.worksheets[3].SHEETSTATE_VERYHIDDEN
        wb.worksheets[4].sheet_state = wb.worksheets[4].SHEETSTATE_VERYHIDDEN
        wb.worksheets[5].sheet_state = wb.worksheets[5].SHEETSTATE_VERYHIDDEN
        #Retrieving the user values to validate
        my_array = file_sheets[0].values
        #Gettin' the type of document
        sheet_number = 0
        sheet = my_array[1][0]
        if sheet == 'VOLUNTARIO':
            sheet_number = 1
        ws = wb.worksheets[sheet_number]
        wb.worksheets[abs(sheet_number-1)].sheet_state = wb.worksheets[abs(sheet_number-1)].SHEETSTATE_VERYHIDDEN
        #And now its finally time to fill it
        myshape = my_array.shape
        for i in range(0, myshape[0]):
            for j in range(0, myshape[1]):
                auxvalue = my_array[i][j]
                ws.cell(row=i+2, column=j+1).value = auxvalue
                validation, errormessage = doCellValidations(sheet, j+1, auxvalue)
                if validation:
                    ws.cell(row=i + 2, column=j + 1).style = Style(fill=PatternFill(patternType='solid', fill_type='solid', fgColor=Color('F3FF00')))
                    comment = Comment(errormessage)
                    comment.width = units.points_to_pixels(300)
                    comment.height = units.points_to_pixels(50)
                    ws.cell(row=i + 2, column=j + 1).comment = comment
        wb.save("revised.xlsx")
        return jsonify({"message": "succes"})
    except Exception as ex:
        return ex

if __name__ == '__main__':
    app.run(debug=True, port=4000)

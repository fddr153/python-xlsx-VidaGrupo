from flask import Flask, jsonify, request, Response, make_response
import numpy as np
import pandas as pd
import openpyxl as opyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import units
from getlists import getBothLists
from validations import doCellValidations
from io import BytesIO


app = Flask(__name__)
app.config['JSON_SORT_KEYS'] = False
context_path = "/api/vida-grupo/"

# Testing Route
@app.route('/ping', methods=['GET'])
def ping():
    return jsonify({'response': 'pong!'})

# Create Data Routes


@app.route(context_path + '/plantillas', methods=['GET'])
def plantillaVacia():
    try:
        # GETTIN' THE PARAMS FROM THE FRONT
        folio = request.args['folio']
        negocio = request.args['negocio']
        # GETTIN MY DATA FROM THE DATABASE ACCORDING TO THE ENTRY PARAMS, FOR NOW ITS JUST DUMMY
        lineas = []
        lineas.append(['TRADICIONAL', 'DIRECTORES'])
        lineas.append(['TRADICIONAL', 'DIRECTORES'])
        lineas.append(['TRADICIONAL', 'DIRECTORES'])
        lineas.append(['TRADICIONAL', 'SUBGERENTES'])
        lineas.append(['TRADICIONAL', 'SUBGERENTES'])
        lineas.append(['TRADICIONAL', 'SUBGERENTES'])
        lineas.append(['TRADICIONAL', 'GERENTES'])
        lineas.append(['TRADICIONAL', 'GERENTES'])
        lineas.append(['TRADICIONAL', 'GERENTES'])
        lineas.append(['TRADICIONAL', 'ADMINISTRATIVOS'])
        lineas.append(['TRADICIONAL', 'ADMINISTRATIVOS'])
        lineas.append(['TRADICIONAL', 'ADMINISTRATIVOS'])
        # NOW THAT I HAVE MY DATA, I LOAD THE MASTER_TEMPLATE AND FILL IT
        wb = opyxl.load_workbook("master_template.xlsx")
        sheet_number = 0
        sheet = lineas[0][0]
        if sheet == 'VOLUNTARIO':
            sheet_number = 1
        mainws = wb.worksheets[sheet_number]
        ws = wb.worksheets[4]
        ws2 = wb.worksheets[5]
        giros_list, ocupacion_list = getBothLists()
        if(len(ocupacion_list) <= 0):
            return make_response(jsonify(message="Error de conexion con la base de datos",error="Listas de parametros de control vacias"), 400)
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
        # NOW I UPDATE MY MAIN WORKSHEET
        ws = wb.worksheets[sheet_number]
        i = 0
        for linea in lineas:
            ws.cell(row=i + 2, column=1).value = linea[0]
            ws.cell(row=i + 2, column=2).value = linea[1]
            i += 1
        # NOW I HIDE THE CONFIGURATION SHEETS AND DELETE THE UNUSED USER SHEET
        wb.worksheets[2].sheet_state = wb.worksheets[2].SHEETSTATE_VERYHIDDEN
        wb.worksheets[3].sheet_state = wb.worksheets[3].SHEETSTATE_VERYHIDDEN
        wb.worksheets[4].sheet_state = wb.worksheets[4].SHEETSTATE_VERYHIDDEN
        wb.worksheets[5].sheet_state = wb.worksheets[5].SHEETSTATE_VERYHIDDEN
        remove_ws = wb.worksheets[abs(sheet_number - 1)]
        wb.remove(remove_ws)
        # NOW MY FILE IS DONE, ITS TIME TO SEND IT
        # wb.save("prueba_init.xlsx") //IF SAVED TO LOCAL FILE ITS ACCURATE
        # NOW TO SEND IT WITH RESPONSE I NEED TO SAVE IT INTO A BytesIO OBJECT
        virtual_wb = BytesIO()
        wb.save(virtual_wb)
        return Response(virtual_wb.getvalue(), mimetype=wb.mime_type,
                        headers={"Content-Disposition": "attachment;filename=plantilla.xlsx"})
    except Exception as ex:
        return make_response(
            jsonify(message="Error en la generacion del archivo xlsx", error="Error en el manejo del archivo template"),
            400)


@app.route(context_path + '/plantillas', methods=['POST'])
def plantillaRevisada():
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
        if (len(ocupacion_list) <= 0):
            return make_response(jsonify(message="Error de conexion con la base de datos",error="Listas de parametros de control vacias"), 400)
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
        #SELECCIONANDO EL WORKSHEET CORRECTO DE LA PLANTILLA
        ws = wb.worksheets[sheet_number]
        #ELMINANDO DE LA PLANTILLA EL WORKSHEET INCORRECTO
        remove_ws = wb.worksheets[abs(sheet_number-1)]
        wb.remove(remove_ws)
        #And now its finally time to fill it with the users data
        myshape = my_array.shape
        yellowFill = PatternFill(fill_type='solid', start_color='F3FF00', end_color='F3FF00')
        for i in range(0, myshape[0]):
            for j in range(0, myshape[1]):
                auxvalue = my_array[i][j]
                ws.cell(row=i+2, column=j+1).value = auxvalue
                #And now, after fillin' the cell, its time to validate the data, each cell at a time
                validation, errormessage = doCellValidations(sheet, j+1, auxvalue)
                if j + 1 == 4:
                    if [ws.cell(row=i+2, column=j).value, auxvalue] not in ocupacion_list:
                        validation = False
                        errormessage = "El valor debe coincidir con los valores de la lista"
                        ws.cell(row=i + 2, column=j).fill = yellowFill
                if not validation:
                    ws.cell(row=i + 2, column=j + 1).fill = yellowFill
                    comment = Comment(errormessage, "Sura")
                    comment.width = units.points_to_pixels(300)
                    comment.height = units.points_to_pixels(50)
                    ws.cell(row=i + 2, column=j + 1).comment = comment
        #POR AHORA GUARDO EL ARCHIVO EN LA MAQUINA LOCAL DE PRUEBAS
        #wb.save("revised.xlsx")
        #SENDING FILE BACK TO USER FOR NOW - TESTING
        virtual_wb = BytesIO()
        wb.save(virtual_wb)
        return Response(virtual_wb.getvalue(), mimetype=wb.mime_type, headers={"Content-Disposition": "attachment;filename=plantilla_revisada.xlsx"})
        #return jsonify({"message": "succes"})
    except Exception as ex:
        return make_response(
            jsonify(message="Error en la generacion del archivo xlsx", error="Error en el manejo del archivo template"),
            400)


if __name__ == '__main__':
    app.run(debug=True, port=4000)

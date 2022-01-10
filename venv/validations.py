import re


def valid_tipo_de_plan(tipo) -> bool:
    return (tipo=="TRADICIONAL" or tipo== "VOLUNTARIO")


def valid_fecha(value) -> bool:

    r = re.compile('.{2}/.{2}/.{4}')
    if len(value) == 10:
        if r.match(value):
            x = value.split("/")
            res = (int(x[0]) > 0 and int(x[0]) < 32) and (int(x[1]) > 0 and int(x[1]) <= 12 )
            return res
    return False


def valid_email(email) -> bool:
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    if re.fullmatch(regex, email):
        return True

    else:
        return False


def valid_salario(value) -> bool:
    try:
        value = value.replace(".", "")
        value.isnumeric()
        x = value.split(",")
        if len(x) == 2 and len(x[1]) == 2:
            return (float(x[0])+float(x[1])) > 0
        return False

    except ValueError:
        return False
def valid_meses_sueldo(value) -> bool:
    try:
        if str.isdigit(value):
            return int(value) > 0
        return False
    except ValueError:
        return False

##MAIN FUNCTION ##NOTA, LOS GIROS Y OCUPACIONES SIEMPRE SON INDEX 3 Y 4, LOS HAGO EN EL PRINCIPAL


def doCellValidations(plan, index, value):
    res = False
    message = ""
    if index == 1:
        res = valid_tipo_de_plan(value)
        message = "El tipo de plan debe ser TRADICIONAL o VOLUNTARIO"
    if index == 8:
        res = valid_fecha(value)
        message = "La fecha debe tener formato dd/mm/yyyy"
    if index == 11:
        res = valid_salario(value)
        message = "El salario debe ser un número válido con dos decimales"
    if index == 12:
        res = valid_meses_sueldo(value)
        message = "El valor debe ser un número entero mayor a cero"
    return res, message

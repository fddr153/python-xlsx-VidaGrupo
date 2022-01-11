import cx_Oracle
import sys


def getBothLists():
    giros_list = []
    giros_ocupacion_list = []

    host = "localhost"
    port = 1521
    service_name = 'ORCLCDB.localdomain'
    user = 'felix'
    password = 'daza'
    sid = cx_Oracle.makedsn(host, port, service_name=service_name)
    try:
        connection = cx_Oracle.connect(f"{user}/{password}@{host}:{port}/{service_name}")
        try:
            cursor = connection.cursor()
            cursor.execute('SELECT * from GIRO')
            for row in cursor:
                giros_list.append(row[0])
            cursor.execute('SELECT * from OCUPACION')
            for row in cursor:
                giros_ocupacion_list.append([row[1], row[0]])

        except:
            raise RuntimeError('Fallo en conexión con la base de datos.') from exc
        else:
            return giros_list, giros_ocupacion_list
        finally:
            cursor.close()
    except:
        raise RuntimeError('Fallo en conexión con la base de datos.') from exc
    else:
        if connection is not None:
            connection.close()
    finally:
        return [],[]

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
            cursor.execute('SELECT * from GIROS')
            for row in cursor:
                giros_list.append([row[0],row[1]])
            cursor.execute('SELECT g.ID , g.NOMBRE, o.ID_OCUPACION, o.NOMBRE FROM GIROS g INNER JOIN OCUPACIONES o ON g.ID =o.ID_GIROS ORDER BY g.ID ASC ,o.ID_OCUPACION ASC')
            for row in cursor:
                giros_ocupacion_list.append([row[0], row[1],row[2],row[3]])
            cursor.close()
            connection.close()
            return giros_list, giros_ocupacion_list
        except:
            cursor.close()
            connection.close()
            return [],[]
    except:
        return [],[]


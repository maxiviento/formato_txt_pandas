from openpyxl import load_workbook
import pandas as pd

def dar_formato(txt, total, tipo):
    if len(txt) + 1 <= total:
        resto = total - len(txt)
        if tipo == 'A':
            txt = txt + ' ' * resto
        elif tipo == 'N':
            txt = '0' * resto + txt

    if len(txt) > total:
        txt = txt[:total]

    return txt

wb = load_workbook("formato.xlsx")
ws_formato = wb.active
index = 1
columnas = []

for row in ws_formato:
    try:
        if int(row[2].value) > 0:
            tupla_columnas = (index, row[1].value, int(row[2].value), row[0].value)
            columnas.append(tupla_columnas)
            index = index + 1
    except:
        pass

df_cuotas = pd.read_csv('cuotas.csv')

print(df_cuotas)

txt = ""
for index, row in df_cuotas.iterrows():
    txt += "002501202129012021"
    cuit = str(row['CUIT'])
    cuit = cuit[:-2]
    txt += dar_formato(cuit, 11, 'N')
    txt += "000"
    txt += dar_formato(str(row['NRO_DOCUMENTO']), 8, 'N')
    txt += "0020931"
    apellido_nombre = str(row['APELLIDOS']) + " " + str(row['NOMBRES'])
    apellido_nombre = apellido_nombre.title().upper()
    txt += dar_formato(apellido_nombre, 27, 'A')
    txt += "1"
    txt += dar_formato(str(row['NRO_DOCUMENTO']), 8, 'N')
    txt += "0400000000000"
    txt += dar_formato("", 27, 'A')
    txt += "000000000040010000002000000"
    sum = 0
    for i in range(23,98):
        sum  += columnas[i][2]
    sum += 2
    txt += dar_formato("", sum, 'N')
    txt += "101210000"
    txt += dar_formato("", 36, 'N')
    txt += "GOBIERNO DE LA PROVINCIA DE CORDOBA   MIN DE PROMOCIÓN DEL EMPLEO Y ECONOMIAPROGRAMA ASIGNACION ESTIMULO          APODERADO DE                          "
    txt += dar_formato(apellido_nombre, 38, 'A')
    txt += dar_formato("", 114, 'A')
    txt += "100000000000000000000000000 000000000000000000000000000020000006481"
    txt += dar_formato("", 38, 'A')
    txt += "\n"

f = open("formateado final definitivo 1 link megaupload.txt", "a")
f.write(txt)
f.close()
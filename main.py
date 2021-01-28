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

df_cuotas = pd.read_csv('SEGUNDAS CUOTAS A PAGAR VIDA DIGNA TOTAL - NOMINA DE BENEFICIOS .csv')
for column in df_cuotas:
    print(column)
df_xlsx = df_cuotas[['N_DEPARTAMENTO_GU', 'LOCALIDAD', 'COD SUCURSAL', 'NOMBRE SUCURSAL', 'NOMBRES', 'APELLIDOS', 'CUIT', 'FECHA DESDE', 'FECHA HASTA']]
try:
    df_xlsx['COD SUCURSAL'] = df_xlsx['COD SUCURSAL'].astype(str).str.slice(0, -2, 1)
    df_xlsx['CUIT'] = df_xlsx['CUIT'].astype(str).astype(str).str.slice(0, -2, 1)
except:
    pass
print(df_xlsx)
df_xlsx.to_excel("Test.xlsx", index=False)
#for index,row in df_xlsx.iterrows():


print(df_cuotas)

txt = "00012021000093350186700000000000000000000000000001022021050220210205541" + dar_formato("", 953, 'A')
for index, row in df_cuotas.iterrows():
    txt += "\n"
    #GRUPO DE PAGO	FECHA DESDE	FECHA HASTA
    txt += "000102202105022021"
    cuit = str(row['CUIT'])
    cuit = cuit[:-2]
    txt += dar_formato(cuit, 11, 'N')
    #EXCAJA	 TIPO BENEFICIARIO
    txt += "000"
    txt += dar_formato(str(row['NRO_DOCUMENTO']), 8, 'N')
    #BCO: FIJO= 020	SUCURSAL
    txt += "0020931"
    apellido_nombre = str(row['APELLIDOS']) + " " + str(row['NOMBRES'])
    apellido_nombre = apellido_nombre.title().upper()
    txt += dar_formato(apellido_nombre, 27, 'A')
    #TIPO DE DOC: DNI=1
    txt += "1"
    txt += dar_formato(str(row['NRO_DOCUMENTO']), 8, 'N')
    #PCIA EMISIÓN: CBA=04	CUIL APODERADO
    txt += "0400000000000"
    #APELLIDO Y NOMBRE DE APODERADO
    txt += dar_formato("", 27, 'A')
    #TIPO DOC. APODERADO NRO DOC APODERADO PCIA EMISIÓN: CBA=04	CÓDIGO CONCEPTO 1: FIJO 001	SUBCÓDIGO DEL CONCEPTO 1: FIJO 000	IMPORTE DEL CONCEPTO 1
    txt += "000000000040010000002000000"
    #columna desde la 23 a la 97
    sum = 0
    for i in range(23,98):
        sum  += columnas[i][2]
    sum += 2
    txt += dar_formato("", sum, 'N')
    #IMPORTE MAYOR A $ 9.999,00.-: 0=NO ; 1=SI	PERIODO LIQUIDACIÓN: MMAA	TIPO DE PAGO: VER ANEXO	FORMA DE PAGO: VENTANILLA= 0	TIPO DE CUENTA: VALOR FIJO 0
    txt += "101210000"
    #NRO CUENTA VALOR FIJO 0	FECHA DESDE PROX PAGO: DDMMAAA	FECHA HASTA PROX PAGO: DDMMAAAA
    txt += dar_formato("", 36, 'N')
    #LEYENDA 1	LEYENDA 2	LEYENDA 3	LEYENDA 4
    txt += "GOBIERNO DE LA PROVINCIA DE CORDOBA   MIN DE PROMOCIÓN DEL EMPLEO Y ECONOMIAPROGRAMA ASIGNACION ESTIMULO          APODERADO DE                          "
    #LEYENDA 5
    txt += dar_formato(apellido_nombre, 38, 'A')
    #LEYENDA 6	LEYENDA 7	LEYENDA 8
    txt += dar_formato("", 114, 'A')
    #CÓDIGO PAGO-IMPAGO: LA EMPRESA DEBE PONER SIEMPRE 1= IMPAGO. ELBANCO EN LA RENDICIÓN DEVUELVE 0=PAGO O 1=IMPAGO	FECHA PAGO:COMPLETA EL BANCO LA EMPRESA DEBE PONER 0	PAGO CON TARJETA VALOR FIJO = 0	MOTIVO IMPAGO VALOR FIJO PARA LA EMPRESA=0- LUEGO COMPLETA EL BANCO - VER ANEXOS	NÚMERO DE COMPROBANTE: COMPLETA EL BANCO	ÚLTIMO MOV CUENTA: COMPLETA EL BANCO	RETENCIÓN DE TARJETA: 0=PAGO NORMAL; 1=TARJ.RETENIDA CAJERO	COMISIÓN: VALOR FIJO= ESPACIO	UR ASIGNADA: VALOR FIJO=000	IMPORTE  MORATORIA AFIP: VALOR FIJO=0000000000	IMP. RETROACTIVO MOR. AFIP: VALOR FIJO=0000000000	IMPORTE NETO A COBRAR 	CÓDIGO DE EMPRESA
    txt += "100000000000000000000000000 000000000000000000000000000020000005541"
    #USO FUTURO: VALOR FIJO= ESPACIOS
    txt += dar_formato("", 38, 'A')

    

f = open("formateado.txt", "a")
f.write(txt)
f.close()
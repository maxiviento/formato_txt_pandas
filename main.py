from openpyxl import load_workbook
import pandas as pd

codigo_empresa = '5533'

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

col_list = ['MONTO', 'N_DEPARTAMENTO_GU', 'localidad', 'NUMERO DE SUCURSAL', 'NOMBRE SUCURSAL', 'NOMBRES', 'APELLIDOS', 'CUIT', 'FECHA_DESDE', 'FECHA_HASTA', 'NRO_DOCUMENTO']
df_cuotas = pd.read_excel(codigo_empresa+'.xlsx', usecols=col_list)

print(list(df_cuotas.columns))

try:
    df_cuotas['FECHA_HASTA'] = '0' + df_cuotas['FECHA_HASTA'].astype(str)
    df_cuotas['FECHA_DESDE'] = '0' + df_cuotas['FECHA_DESDE'].astype(str)
    df_cuotas['NUMERO DE SUCURSAL'] = df_cuotas['NUMERO DE SUCURSAL'].astype(str).str.slice(0, -2, 1)
    df_cuotas['CUIT'] = df_cuotas['CUIT'].astype(str).astype(str).str.slice(0, -2, 1)
except:
    pass
print(df_cuotas['FECHA_HASTA'])

print(df_cuotas)
df_cuotas.to_excel(codigo_empresa+" formateado.xlsx", index=False)

print(df_cuotas)

calc_monto_gral = df_cuotas['MONTO'].sum()
monto_gral = dar_formato(str(calc_monto_gral), 10, 'N')
monto_gral = monto_gral + '00'
cantidad_registros = dar_formato(str(df_cuotas.shape[0]), 8, 'N')
txt = "00022021"+cantidad_registros+monto_gral+"000000000000000000000103202105032021020"+ codigo_empresa + dar_formato("", 953, 'A')
for index, row in df_cuotas.iterrows():
    txt += "\n"
    txt += "00"
    #GRUPO DE PAGO	FECHA DESDE	FECHA HASTA
    txt += str(row['FECHA_DESDE']) + str(row['FECHA_HASTA'])
    cuit = str(row['CUIT'])
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
    txt += "00000000004001000000" + str(row['MONTO']) + '00'
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
    if codigo_empresa == '5533':
        ministerio = 'MINISTERIO DE PROMOCION DEL EMPLEO Y E'
    elif codigo_empresa == '5541':
        ministerio = 'MINISTERIO DE DESARROLLO SOCIAL       '

    programa = 'PROGRAMA VIDA DIGNA                   '
    txt += "GOBIERNO DE LA PROVINCIA DE CORDOBA   "+ministerio+programa+"APODERADO DE                          "
    #LEYENDA 5
    txt += dar_formato(apellido_nombre, 38, 'A')
    #LEYENDA 6	LEYENDA 7	LEYENDA 8
    txt += dar_formato("", 114, 'A')
    #CÓDIGO PAGO-IMPAGO: LA EMPRESA DEBE PONER SIEMPRE 1= IMPAGO. ELBANCO EN LA RENDICIÓN DEVUELVE 0=PAGO O 1=IMPAGO	FECHA PAGO:COMPLETA EL BANCO LA EMPRESA DEBE PONER 0	PAGO CON TARJETA VALOR FIJO = 0	MOTIVO IMPAGO VALOR FIJO PARA LA EMPRESA=0- LUEGO COMPLETA EL BANCO - VER ANEXOS	NÚMERO DE COMPROBANTE: COMPLETA EL BANCO	ÚLTIMO MOV CUENTA: COMPLETA EL BANCO	RETENCIÓN DE TARJETA: 0=PAGO NORMAL; 1=TARJ.RETENIDA CAJERO	COMISIÓN: VALOR FIJO= ESPACIO	UR ASIGNADA: VALOR FIJO=000	IMPORTE  MORATORIA AFIP: VALOR FIJO=0000000000	IMP. RETROACTIVO MOR. AFIP: VALOR FIJO=0000000000	IMPORTE NETO A COBRAR 	CÓDIGO DE EMPRESA
    txt += "100000000000000000000000000 0000000000000000000000000000" + str(row['MONTO']) + '00' + codigo_empresa
    #USO FUTURO: VALOR FIJO= ESPACIOS
    txt += dar_formato("", 38, 'A')

f = open(codigo_empresa+".txt", "w")
f.write(txt)
f.close()
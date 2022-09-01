#from openpyxl import load_workbook
import pandas as pd

#cargamos informe como dataframe
df = pd.read_excel('Informe Tranformadores MLU.xlsx')
df = df[['Fecha', 'TERRITORIO', 'CIRCUITO', 'ID_BDI','CODELEME', 'CODIGO_BDI_CIRCUITO', 'PLACA MT COLOCADA', 'Equipo Ruta Id', 'PLACA MT ANTERIOR',
'Matricula MT anterior', 'POTENCIA', 'POTENCIA NOMINAL MODIFICADA', 'FABRICANTE', 'MARCA MODIFICADA', 'OTRA MARCA', 'RELTRANS', 'TENSIÓN SECUNDARIA MODIFICADA (V)',
'TIPINS', 'CONFIRME TIPO DE INSTALACIÓN', 'TIPOFASE','CONFIRME TIPO DE TRANSFORMADOR', 'ESTADO DEL TRANSFORMADOR', 'Longitud del equipo padre', 'Latitud del equipo padre',
'FOTO VERIFICACIÓN FRENTE - LADO 1','FOTO VERIFICACIÓN FRENTE - LADO 2', 'Foto matricula MT','FOTO MT COLOCADA', 'SOPORTE 01','SOPORTE 02', 'SOPORTE FOTOGRÁFICO 03', 'SOPORTE FOTOGRÁFICO 04', 'Nombre Equipo padre',
'Nombre del Usuario', 'Longitud', 'Latitud', "Estado", "¿TRANSFORMADORES EN BANCO?"]]


#Eliminar duplicados equipo ruta id
df = df.drop_duplicates(subset=['Equipo Ruta Id']) 

#rellenar vacias por ceros
df.fillna(0, inplace=True)

#Eliminar usuario MARIA JOSE BLANCO OCHOA
indexNames = df[df['Nombre del Usuario'] == 'MARIA JOSÉ BLANCO OCHOA' ].index
df.drop(indexNames , inplace=True)

#Eliminar ESTADO RE_PUBL Y RE_ELIM
indexNames = df[df['Estado'] == 'RE_ELIMIN' ].index
df.drop(indexNames , inplace=True)
indexNames = df[df['Estado'] == 'RE_PUBL' ].index
df.drop(indexNames , inplace=True)
del df['Estado'] #ELIMINAR COLUMNA

#Eliminar placas mt colocadas en cero
indexNames = df[df['PLACA MT COLOCADA'] == 0 ].index
df.drop(indexNames , inplace=True)

#Generacion de codigo id_bdi, codeleme o nuevo
for i in df.index:
    if df["CODELEME"][i] == 0:
        df["CODELEME"][i] = str(10) + str(df["Equipo Ruta Id"][i])
    if df["ID_BDI"][i] == 0:
        df["ID_BDI"][i] = df["CODELEME"][i]
df.rename(columns={'ID_BDI': 'CODIGO'}, inplace=True) #RENOMBRANDO COLUMNA
del df['CODELEME'] #ELIMINAR COLUMNA CODELEME

#INSTALACION SUPERIOR
df.rename(columns={'CODIGO_BDI_CIRCUITO': 'INSTALACION_SUPERIOR'}, inplace=True) #RENOMBRANDO COLUMNA

#PLACA MT ANTERIOR
for i in df.index:
    if df["Matricula MT anterior"][i] != 0:
        df["PLACA MT ANTERIOR"][i] = df["Matricula MT anterior"][i]
df.rename(columns={'PLACA MT ANTERIOR': 'MATRICULA_ANTIGUA'}, inplace=True) #RENOMBRANDO COLUMNA
del df["Matricula MT anterior"]  #ELIMINAR COLUMNA

#POTENCIA
for i in df.index:
    if df["POTENCIA NOMINAL MODIFICADA"][i] != 0:
        df["POTENCIA"][i] = df["POTENCIA NOMINAL MODIFICADA"][i]
df.rename(columns={'POTENCIA': 'Potencia_Nominal_(kVA)'}, inplace=True) #RENOMBRANDO COLUMNA
del df["POTENCIA NOMINAL MODIFICADA"] #ELIMINAR COLUMNA

#MARCA
for i in df.index:
    if df["OTRA MARCA"][i] != 0:
        df["MARCA MODIFICADA"][i] = df["OTRA MARCA"][i]
    if df["MARCA MODIFICADA"][i] != 0:
        df["FABRICANTE"][i] = df["MARCA MODIFICADA"][i]
df.rename(columns={'FABRICANTE': 'Marca'}, inplace=True) #RENOMBRANDO COLUMNA
del df["OTRA MARCA"] #ELIMINAR COLUMNA
del df["MARCA MODIFICADA"] #ELIMINAR COLUMNA

#TENSION SECUNDARIA
for i in df.index:
    if df["TENSIÓN SECUNDARIA MODIFICADA (V)"][i] == 0:
        df["TENSIÓN SECUNDARIA MODIFICADA (V)"][i] = df["RELTRANS"][i]
    df["RELTRANS"][i] = 0
df.rename(columns={'TENSIÓN SECUNDARIA MODIFICADA (V)': 'Tension_Secundaria_(V)',
                   'RELTRANS': 'TENS_SEC'                                         }, inplace=True)  #RENOMBRANDO COLUMNA

#ubicacion-tipo de instalacion
for i in df.index:
    if df["CONFIRME TIPO DE INSTALACIÓN"][i] == 0:
        df["CONFIRME TIPO DE INSTALACIÓN"][i] = df["TIPINS"][i]
    if df["CONFIRME TIPO DE INSTALACIÓN"][i] == 'AÉREO':
        df["CONFIRME TIPO DE INSTALACIÓN"][i] = 'AEREO'
    if df["CONFIRME TIPO DE INSTALACIÓN"][i] == 'SUBTERRÁNEO':
        df["CONFIRME TIPO DE INSTALACIÓN"][i] = 'SUBTERRANEO'
    df["TIPINS"][i] = 0
df.rename(columns={'TIPINS':"UBICACION_TRAFO",
                   'CONFIRME TIPO DE INSTALACIÓN':'Ubicacion_Trafo' }, inplace=True)  #RENOMBRANDO COLUMNA

#TIPO DE TX
for i in df.index:
    if df['CONFIRME TIPO DE TRANSFORMADOR'][i] == 0:
        df['CONFIRME TIPO DE TRANSFORMADOR'][i] = df['TIPOFASE'][i]
    if df['CONFIRME TIPO DE TRANSFORMADOR'][i] == 'MONOFÁSICO BIFILAR':
        df['CONFIRME TIPO DE TRANSFORMADOR'][i] = 'MONOFASICO BIFILAR'
    if df['CONFIRME TIPO DE TRANSFORMADOR'][i] == 'MONOFÁSICO':
        df['CONFIRME TIPO DE TRANSFORMADOR'][i] = 'MONOFASICO'
    if df['CONFIRME TIPO DE TRANSFORMADOR'][i] == 'TRIFÁSICO':
        df['CONFIRME TIPO DE TRANSFORMADOR'][i] = 'TRIFASICO'   
    df['TIPOFASE'][i] = 0
df.rename(columns={'TIPOFASE':"Tipo_de_Conexion",
                   'CONFIRME TIPO DE TRANSFORMADOR':'TIPOFASE' }, inplace=True)  #RENOMBRANDO COLUMNA

#estado del elemento
df.insert(16,"Estado_Elemento",0)

#COD,OBSERVACIONES, ORIGEN DE LOS DATOS, TIPO AREA
df.insert(18,"COD",0)
df.insert(19,"Observaciones",0)
df.insert(20,"Origen_de_los_datos",17)
df.insert(21,"Origen de los datos","CENSO II")
df.insert(22,"UC_R015",0)
df.insert(23,"TIPO_AREA",0)
df.insert(24,"TIPO DE AREA",0)

#LONGITUD Y LATITUD
for i in df.index:
    if df['Longitud del equipo padre'][i] == 0:
        df['Longitud del equipo padre'][i] = df['Longitud'][i]
    if df['Latitud del equipo padre'][i] == 0:
        df['Latitud del equipo padre'][i] = df['Latitud'][i]

df.rename(columns={'Longitud del equipo padre':"Longitud (WGS84)",
                   'Latitud del equipo padre':'Latitud (WGS84)' }, inplace=True)  #RENOMBRANDO COLUMNA
del df["Latitud"] #ELIMINAR COLUMNA
del df["Longitud"] #ELIMINAR COLUMNA
df.insert(37,"ESTADO COORDENADAS",0)

# COMPARAMOS LOS CODIGOS CON LA BD MT

#cargamos BD MT como dataframe
df_mt = pd.read_excel('BD MT - CODIGO.xlsx')

#INSERTAMOS COLUMNAS PARA LA VERIFICACION Y CRUCE DE CODIGO
df.insert(4,"CODIGO MT",0)
df.insert(5,"VERIFICACION CODIGO",0)

#se cruzan los df por matricula antigua 
df = df.merge(df_mt, how='left', on='MATRICULA_ANTIGUA')

#se copian los datos, se organizan y se elimina la columna sobrante
for i in df.index:
    df['CODIGO MT'][i] = df['CODIGO_TRANSFORMADOR'][i]
del df['CODIGO_TRANSFORMADOR']

#imprimimos el df en un excel
df.to_excel('COORDENADAS_xx_SEPTIEMBRE2022_APG.xlsx', sheet_name='CC')
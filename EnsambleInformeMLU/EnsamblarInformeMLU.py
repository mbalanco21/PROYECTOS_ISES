import dask
import pandas as pd
import dask.dataframe as dd
import numpy as np
from pandas import ExcelWriter
# from fConsolidarClientes import DistanciaCoord
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from xlwt import Workbook
import xlwt
import time
import sys
from normalize import normalize

from RemStrDuplicated import unique_list
import warnings
warnings.filterwarnings('ignore')
warnings.warn('DelftStack')
warnings.warn('Do not show this message')
print("No Warning Shown")


# =============================================================================
# 1. Se Carga el archivo MLU
# =============================================================================
# path = 'MLU_2.xlsx'

#VENTANA DE ABRIR DOCUMENTO
file_path = filedialog.askopenfilename(defaultextension=".xls")
if file_path is None:
    # quit()
    sys.exit()

# MLU = pd.read_excel(path, header=0,usecols = ["TERRITORIO","SUBESTACIÓN","CIRCUITO",\
# "Nombre Ruta","Nombre del Usuario","Equipo Ruta Id","Nombre Equipo","Equipo Padre",\
# "Nombre Equipo Padre","Nombre Trabajo","rutaid","Fecha","Hora",\
# "Longitud equipo padre","Latitud equipo padre","Estado del sticker",\
#      "Digite número del sticker","Estado del número del medidor",\
#          "Digite número del medidor","Marca del medidor","¿Dirección disponible?",\
#              "Tipo Vía","Nombre vía","Duplicador","Número puerta",\
#                  "Longitud","Latitud"])
# shutil.copy("Plantillas/Reporte_bin.xlsb", file_path[:len(file_path)-4] + 'xlsb')

# La pestaña TRAMOS se utiliza para sacar los elementos TRAMOS y Puentes BT
# La pestaña AP se utiliza para sacar los demás
MLU = pd.read_excel(file_path, header=0,sheet_name = 'AP')
MLU_Tramo = pd.read_excel(file_path, header=0,sheet_name = 'TRAMOS')


Glosario_Apoyo = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'APOYO',usecols = 'K:N')
Glosario_Apoyo.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True)

    
Glosario_Tramo = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'TRAMO',usecols = 'K:N')
Glosario_Tramo.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True)

Glosario_Caja = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'CAJA DE ABONADO',usecols = 'K:N')
Glosario_Caja.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True)
    
Glosario_Puente = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'PUENTE',usecols = 'K:N')    
Glosario_Puente.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True) 
    
Glosario_TVCABLE = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'TV CABLE',usecols = 'K:N')    
Glosario_TVCABLE.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True) 
    
Glosario_Camara = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'CAMARA',usecols = 'K:N')    
Glosario_Camara.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True)
    
Glosario_Antena = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'ANTENA',usecols = 'K:N')    
Glosario_Antena.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True) 
     
Glosario_Ficticio = pd.read_excel('GLOSARIO NUEVO_Ref_Codigos.xlsb', header=0,\
                               sheet_name = 'Ent_1.1',usecols = 'K:N')    
Glosario_Ficticio.rename(columns={'Unnamed: 10':'Editor','Unnamed: 11':'Codigo',\
                        'Unnamed: 12':'Descripcion','Unnamed: 13':'Observacion'},inplace=True)  
        
UUCC = pd.read_excel('unidades constructivas_UUCC_CREG.xlsx', header=0,\
                               sheet_name = 'UC N1 Postes',usecols = 'C:D')    
UUCC = UUCC.iloc[5:,:] #seleccion de fila 5
UUCC.reset_index(inplace=True) #resetar index
UUCC = UUCC.drop(columns = 'index') #elimina columna index
UUCC.rename(columns={'ANÁLISIS DE PRECIOS UNITARIOS ':'UC','Unnamed: 3':'DESCRIPCION'\
                        },inplace=True) #renombrando columnas
UUCC.DESCRIPCION = UUCC.loc[:,'DESCRIPCION'].str.lower().apply(str.replace,args=(' ', '')) #se coloca la columna en minusculas


UUCC_Tramo = pd.read_excel('unidades constructivas_UUCC_CREG.xlsx', header=0,\
                               sheet_name = 'UC N1 Conductores',usecols = 'B:C')
    
UUCC_Tramo = UUCC_Tramo.iloc[4:,:] #seleccion de fila 4
UUCC_Tramo.reset_index(inplace=True) #resetar index
UUCC_Tramo = UUCC_Tramo.drop(columns = 'index') #elimina columna index
UUCC_Tramo.rename(columns={'ANÁLISIS DE PRECIOS UNITARIOS ':'UC','Unnamed: 2':'DESCRIPCION'\
                        },inplace=True) #renombrando columnas
    
UUCC_Tramo1 = pd.read_excel('unidades constructivas_UUCC_CREG.xlsx', header=0,\
                               sheet_name = 'UC N1 Conductores',usecols = 'R:S')
UUCC_Tramo1 = UUCC_Tramo1.iloc[4:,:]  
UUCC_Tramo1.rename(columns={'Unnamed: 17':'UC','Unnamed: 18':'DESCRIPCION'\
                        },inplace=True)
UUCC_Tramo1.reset_index(inplace=True,drop=True)

UUCC_Tramo = pd.concat([UUCC_Tramo,UUCC_Tramo1])
UUCC_Tramo = UUCC_Tramo.drop(UUCC_Tramo.loc[UUCC_Tramo.loc[:,'UC'].isna(),:].index) #eliminar campos vacios 
UUCC_Tramo.reset_index(inplace=True,drop=True)

UUCC_Tramo.DESCRIPCION = UUCC_Tramo.loc[:,'DESCRIPCION'].astype(str) #convertir columna a tipo str
UUCC_Tramo.UC = UUCC_Tramo.loc[:,'UC'].astype(str) #convertir columna a tipo str
UUCC_Tramo.DESCRIPCION = UUCC_Tramo.loc[:,'DESCRIPCION'].str.lower().apply(str.replace,args=(' ', ''))
UUCC_Tramo = UUCC_Tramo.loc[~UUCC_Tramo.UC.str.contains('nan'),:]
UUCC_Tramo = UUCC_Tramo.loc[~UUCC_Tramo.UC.str.contains('UC'),:]
UUCC_Tramo.reset_index(inplace=True)
UUCC_Tramo = UUCC_Tramo.drop(columns = 'index')

# Sel = UUCC_Tramo.DESCRIPCION.str.contains('<')
# A = UUCC_Tramo.loc[Sel,:]
# A.DESCRIPCION.str.rfind('<', start=0, end=None)
# A.loc[0,'DESCRIPCION'][:66]
# df_repeated = pd.concat([A]*4, ignore_index=True)
# df_repeated['sort'] = df_repeated['UC']#.str.extract('(\d+)', expand=False).astype(int)
# df_repeated.sort_values('sort',inplace=True, ascending=False)
# df_repeated = df_repeated.drop('sort', axis=1)
# df_repeated.reset_index(inplace=True,drop=True)
# idx = df_repeated.DESCRIPCION.str.rfind('<', start=0, end=None)
# df_repeated.DESCRIPCION.apply(lambda x: x[:idx[i]])

# DataList = []
# for i in range(0,len(A.loc[:,'DESCRIPCION']),1):
#     B = pd.DataFrame(np.repeat(A.loc[i,:].values, 4, axis=0))
#     DataList.append(B) 
# DataListF = pd.concat(DataList)    
# newdf.columns = df.columns
# A.DESCRIPCION.apply(lambda x: x.join([char*(7-len(x)) for char in '0']) + x)
# Se valida si hay duplicados en QR y/o No medidor

#se filtran los contains
Sel = MLU.loc[:,'Tipo de Apoyo'].str.contains('Baja Tensión|Media y Baja Tensión|Solo retención', regex=True)
MLU = MLU.loc[Sel,:]
Apoyo = MLU
Apoyo = Apoyo.loc[Apoyo.loc[:,'Tipo de Apoyo'].str.len() != 0,:]
Apoyo.loc[:,'Codigo'] = '11' + Apoyo.loc[:,'Equipo Ruta Id'].astype(str)
# Apoyo.loc[:,'Codigo'] = '11' + Apoyo.loc[:,'Equipo Ruta Id'].astype(str)
Apoyo.loc[:,'Instalacion_Superior'] = Apoyo.loc[:,'BDI'].astype(str)
Sel1 = Apoyo.loc[:,'Estructura'].astype(str).str.contains('ALINEACIÓN', regex=True)
Apoyo.loc[Sel1,'Funcion_de_Aislador'] = 'S'
Sel2 = Apoyo.loc[:,'Estructura'].astype(str).str.contains('FIN DE LÍNEA|ÁNGULO', regex=True)
Apoyo.loc[Sel2,'Funcion_de_Aislador'] = 'R'
Apoyo.loc[Sel1,'DESC_Funcion_de_Aislador'] = 'Suspension'
Apoyo.loc[Sel2,'DESC_Funcion_de_Aislador'] = 'Retencion'
Apoyo.loc[:,'Tipo_de_apoyo'] = Apoyo.loc[:,'Tipo de Apoyo']
Apoyo.loc[:,'Tipo_de_apoyo_'] = Apoyo.loc[:,'Tipo_de_apoyo']
Apoyo.loc[:,'Tipo_de_apoyo_'] = Apoyo.loc[:,'Tipo_de_apoyo_'].astype(str).apply(\
                str.replace,args=('Baja Tensión', 'BT'))
Apoyo.loc[:,'Tipo_de_apoyo_'] = Apoyo.loc[:,'Tipo_de_apoyo_'].astype(str).apply(\
                str.replace,args=('Media Tensión', 'MT'))    
Apoyo.loc[:,'Tipo_de_apoyo_'] = Apoyo.loc[:,'Tipo_de_apoyo_'].astype(str).apply(\
                str.replace,args=('Caja BT', 'CA'))     
Apoyo.loc[:,'Tipo_de_apoyo_'] = Apoyo.loc[:,'Tipo_de_apoyo_'].astype(str).apply(\
                str.replace,args=('Media y BT', 'MT-BT'))
Apoyo.loc[:,'Tipo_de_apoyo_'] = Apoyo.loc[:,'Tipo_de_apoyo_'].astype(str).apply(\
                str.replace,args=('Media y Baja Tensión', 'Media y Baja Tension'))
Apoyo.reset_index(inplace=True,drop=True)    
Apoyo.loc[:,'Tipo_de_apoyo'] = [normalize(str(Apoyo.loc[i,'Tipo_de_apoyo'])) \
                            for i in range(0, len(Apoyo))]    
Sel1 = Apoyo.loc[:,'Material del apoyo'].astype(str).str.contains('CONCRETO', regex=True)
Apoyo.loc[Sel1,'Material_Apoyo'] = 'C'
Apoyo.loc[Sel1,'DESC_Material_Apoyo'] = 'CONCRETO'
Sel1 = Apoyo.loc[:,'Material del apoyo'].astype(str).str.contains('METALICO|METÁLICO', regex=True)
Apoyo.loc[Sel1,'Material_Apoyo'] = 'PM'
Apoyo.loc[Sel1,'DESC_Material_Apoyo'] = 'METALICO'
Apoyo.loc[Sel1,'Material del apoyo'] = 'POSTE METALICO'
Sel1 = Apoyo.loc[:,'Material del apoyo'].astype(str).str.contains('MADERA', regex=True)
Apoyo.loc[Sel1,'Material_Apoyo'] = 'M'
Apoyo.loc[Sel1,'DESC_Material_Apoyo'] = 'MADERA'
Sel1 = Apoyo.loc[:,'Material del apoyo'].astype(str).str.contains('FIBRA', regex=True)
Apoyo.loc[Sel1,'Material del apoyo'] = 'POSTE DE FIBRA DE VIDRIO'
Apoyo.loc[Sel1,'Material_Apoyo'] = 'F'
Apoyo.loc[Sel1,'DESC_Material_Apoyo'] = 'FIBRA'
Sel1 = Apoyo.loc[:,'Material del apoyo'].astype(str).str.contains('PRFV', regex=True)
Apoyo.loc[Sel1,'Material_Apoyo'] = 'PR'
Apoyo.loc[Sel1,'DESC_Material_Apoyo'] = 'PRFV'

Apoyo.loc[:,'Carga Rotura'] = Apoyo.loc[:,'Carga Rotura'].astype(str)
Apoyo.loc[Apoyo.loc[:,'Carga Rotura'].str.contains('nan'),'Carga Rotura'] = '0'
# Apoyo.loc[:,'Carga Rotura'] = Apoyo.loc[:,'Carga Rotura'].fillna(0)
try:
    Apoyo.loc[:,'Carga Rotura'] = Apoyo.loc[:,'Carga Rotura'].astype(float)
    Apoyo.loc[:,'Carga Rotura'] = Apoyo.loc[:,'Carga Rotura'].astype(int)
    Apoyo.loc[:,'Carga Rotura'] = Apoyo.loc[:,'Carga Rotura'].astype(str)    
except:
    print("La columna Carga Rotura contiene texto.")

Apoyo.loc[Apoyo.loc[:,'Carga Rotura'] == '0','Carga Rotura'] = ''
CargaRotura = Glosario_Apoyo.loc[Glosario_Apoyo.Editor == 'APPE',:]
CargaRotura.Descripcion = CargaRotura.Descripcion.astype(str)
TESC = Glosario_Apoyo.loc[Glosario_Apoyo.Editor == 'TESC',:]
CRotura = pd.merge(left=Apoyo, right=CargaRotura,how='left', \
        left_on='Carga Rotura',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
# Apoyo.loc[:'Altura del apoyo'] = 
Apoyo.loc[:,'Estado de la carga de rotura'] = Apoyo.loc[:,'Estado de la carga de rotura'].fillna('')
Sel = Apoyo.loc[:,'Estado de la carga de rotura'].str.contains('ILEGIBLE|INEXISTENTE')
Apoyo.loc[:,'Carga_Rotura'] = CRotura.Codigo_y
Apoyo.loc[Sel,'Carga_Rotura'] =  Apoyo.loc[Sel,'Estado de la carga de rotura']
Apoyo.loc[:,'DESC_Carga_Rotura'] = CRotura.Descripcion
Apoyo.loc[Sel,'DESC_Carga_Rotura'] =  Apoyo.loc[Sel,'Estado de la carga de rotura']
Apoyo.loc[:,'Estructura BT_00'] = Apoyo.loc[:,'Estructura BT_00'].fillna('')
Apoyo.loc[:,'Estructura BT_00'] = Apoyo.loc[:,'Estructura BT_00'].apply(lambda x:normalize(x))
TESC.loc[:,'Descripcion1'] = TESC.Descripcion.apply(lambda x:normalize(x))
CTesc = pd.merge(left=Apoyo, right=TESC,how='left', \
        left_on='Estructura BT_00',right_on='Descripcion1',suffixes=(None, '_y'),indicator=True)
# Apoyo.loc[:'Altura del apoyo'] = 
Apoyo.loc[:,'Estructura_Apoyo_BT'] = CTesc.Codigo_y
Apoyo.loc[:,'Estructura_Apoyo_BT_'] = CTesc.Descripcion
Apoyo.loc[:,'DESC_ Elementos_Telematicos'] = Apoyo.loc[:,'¿El apoyo sostiene tramo de telecomunicaciones?'].str.upper() #mayuscula
Apoyo.loc[:,'Elementos_Telematicos'] = Apoyo.loc[:,'¿El apoyo sostiene tramo de telecomunicaciones?']
Apoyo.loc[:,'Elementos_Telematicos'] = Apoyo.replace({'Elementos_Telematicos': {'Si': 'S',\
                    'No': 'N'}})
A = Apoyo.loc[:,'Altura del apoyo'].astype(int)
B = Apoyo.loc[:,'Tipo de red'].str.lower()
Sel1 = B.str.contains('trenzada')
# Sel1 = B.str.contains('trenzada')
B.loc[Sel1] = 'red trenzada'
B.loc[~Sel1] = 'red común'
# Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'Material del apoyo'].str.lower() + '-' + \
#     A.astype(str).str.lower() + ' m-' + \
#     'urbano-' + Apoyo.loc[:,'DESC_Funcion_de_Aislador'].str.lower() + '-' + \
#         B
Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'Material del apoyo'].str.lower() + '-' + \
    A.astype(str).str.lower() + ' m- ' + Apoyo.loc[:,'Desc Tipo de Área'].str.lower() + \
    '-' + Apoyo.loc[:,'DESC_Funcion_de_Aislador'].str.lower() + '-' + \
        B
Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'UUCC'].astype(str).apply(str.replace,args=(' ', ''))
Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'UUCC'].astype(str).apply(str.replace,args=('7m','8m'))
Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'UUCC'].astype(str).apply(str.replace,args=('9m','10m'))
Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'UUCC'].astype(str).apply(str.replace,args=('11m','12m'))
Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'UUCC'].astype(str).apply(str.replace,args=('14m','12m'))
Apoyo.loc[:,'UUCC'] = Apoyo.loc[:,'UUCC'].astype(str).apply(str.replace,args=('16m','12m'))
Apoyo.loc[:,'UUCC'] = [normalize(str(Apoyo.loc[i,'UUCC'])) \
                            for i in range(0, len(Apoyo))]
UUCC.loc[:,'DESCRIPCION'] = [normalize(str(UUCC.loc[i,'DESCRIPCION'])) \
                            for i in range(0, len(UUCC))] 
uucc = pd.Series(Apoyo.loc[:,'UUCC']) 
DataList=[]
In = pd.DataFrame(uucc, columns=["UUCC"])
In.reset_index(inplace=True)
In = In.drop(['index'],axis=1)
for i in range(0,len(Apoyo.loc[:,'UUCC']),1):
    In2 = pd.DataFrame([''],columns=['UUCC'])
    if i < len(Apoyo.loc[:,'UUCC'])-1:
        In1 = pd.concat([In.loc[i:i+1,:],In2])
        Sel = [True,False,True]
        In1 = In1[Sel]
        In1.reset_index(inplace=True)
        In1 = In1.drop(['index'],axis=1)
    else:
        In1 = pd.concat([In.loc[i-1:i,:],In2])
        Sel = [False,True,True]
        In1 = In1[Sel]
        In1.reset_index(inplace=True)
        In1 = In1.drop(['index'],axis=1)        
    DataCruce = pd.merge(left=In1,\
                         right=UUCC,how='left', left_on='UUCC',right_on='DESCRIPCION',suffixes=(None, '_y'),\
                          indicator=True) 
    DataCruce = DataCruce.iloc[:-1,:]
    # DataCruce = DataCruce.loc[DataCruce.loc[:,'_merge'] == 'both',:]
    if np.sum(DataCruce._merge == 'both') > 0:
        DataCruce = DataCruce.groupby(['_merge'])['UC'].transform(\
                            lambda x : '|'.join(x))
        DataCruce = DataCruce.drop_duplicates()
    else:
        DataCruce = pd.Series(np.array([' ']))
    DataList.append(DataCruce) 
DataListF = pd.concat(DataList)
DataListF.reset_index(inplace=True, drop=True)
DataListF = DataListF.to_frame(name="UUCC")
# DataListF = pd.DataFrame(DataListF, columns=["UUCC"])
# DataListF.reset_index(inplace=True)
# DataListF = DataListF.drop(['index'],axis=1)
Apoyo.loc[:,'UC_R015'] = DataListF
# Preguntar para los casos de altura de 9m pq no estan cruzando.
Apoyo.loc[:,'ORIGEN'] = '17'
Apoyo.loc[:,'Origen_de_los_datos'] = 'CENSO II'
Apoyo.loc[:,'Tipo_de_Conexion'] = ''
SelV = (Apoyo.loc[:,'Tipo_de_apoyo_'] == 'BT') & (Apoyo.loc[:,'Altura del apoyo'] >= 11)
SelV1= (Apoyo.loc[:,'Tipo_de_apoyo_'] == 'MT-BT') & (Apoyo.loc[:,'Altura del apoyo'] <= 10)
Apoyo.loc[SelV,'Observación Apoyo BT AE'] == Apoyo.loc[SelV,'Observación Apoyo BT AE'] + '| Altura Vs BT Validada'
Apoyo.loc[SelV1,'Observación Apoyo BT AE'] == Apoyo.loc[SelV1,'Observación Apoyo BT AE'] + '| Altura Vs MT-BT Validada'
Apoyo = Apoyo.fillna('')

Apoyo = Apoyo.loc[:,['Equipo Ruta Id','Nombre Equipo','Nombre Ruta','Tipo_de_Conexion',\
'Codigo','Instalacion_Superior',\
'Funcion_de_Aislador','DESC_Funcion_de_Aislador','Tipo_de_apoyo_','Tipo_de_apoyo',\
'Material_Apoyo','DESC_Material_Apoyo','Altura del apoyo','Carga_Rotura','DESC_Carga_Rotura',\
'Estructura_Apoyo_BT','Estructura_Apoyo_BT_','Elementos_Telematicos',\
'DESC_ Elementos_Telematicos','UC_R015','Observación Apoyo BT AE','ORIGEN',\
'Origen_de_los_datos','Longitud','Latitud','Foto Apoyo BT AE 01',\
'Tipo de Área', 'Desc Tipo de Área']]

Apoyo = Apoyo.fillna('')

# =============================================================================
# # Tramo
# =============================================================================
Tramo = MLU_Tramo
Tramo = Tramo.loc[(Tramo.loc[:,'Tipo de tramo'].str.len() != 0) & \
                  (~Tramo.loc[:,'Tipo de tramo'].isna()),:]
Tramo.loc[:,'Codigo'] = '12' + Tramo.loc[:,'Equipo Ruta Id'].astype(str)
Tramo.loc[:,'Instalacion_origen'] = Tramo.loc[:,'BDI'].astype(str)
Tramo.loc[:,'APOYO_FIN'] = Tramo.loc[:,'Apoyo Final'].astype(str)
Tramo.loc[:,'APOYO_INI'] = Tramo.loc[:,'Apoyo Inicial'].astype(str)
Tramo.loc[:,'Longitud_(mts)'] = Tramo.loc[:,'LONGITUD (MTS)'].astype(str)
Tramo.loc[:,'measuredlength'] = Tramo.loc[:,'Digite la longitud fija'].astype(str)
Tramo.loc[:,'assettype'] = Tramo.loc[:,'Tipo de tramo'].astype(str)
Tramo.loc[:,'assettype'] = Tramo.loc[:,'assettype'].str.upper()
Tramo.reset_index(inplace=True)
Tramo = Tramo.drop(['index'],axis=1)
Tramo.loc[:,'assettype'] = [normalize(str(Tramo.loc[i,'assettype'])) \
                            for i in range(0, len(Tramo))]
Tramo.loc[:,'Tipo_de_tramo'] = Tramo.loc[:,'assettype']
TipoTramo = Glosario_Tramo.loc[Glosario_Tramo.Editor == 'TTAT',:]
CTesc = pd.merge(left=Tramo, right=TipoTramo,how='left', \
        left_on='assettype',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
# Apoyo.loc[:'Altura del apoyo'] = 
Tramo.loc[:,'assettype'] = CTesc.Codigo_y

EstadoTramo = Glosario_Tramo.loc[Glosario_Tramo.Editor == 'ETRF',:]
CTesc = pd.merge(left=Tramo, right=EstadoTramo,how='left', \
        left_on='Estado Tramo',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
# Apoyo.loc[:'Altura del apoyo'] = 
Tramo.loc[:,'lifecyclestatus'] = CTesc.Codigo_y

Tramo.loc[:,'Estado_Tramo'] = CTesc.Descripcion

TRSTramo = Glosario_Tramo.loc[Glosario_Tramo.Editor == 'TRSC',:]
CTesc = pd.merge(left=Tramo, right=TRSTramo,how='left', \
        left_on='Tipo de red secundaria',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'TIP_RED_SEC'] = CTesc.Codigo_y
Tramo.loc[:,'Tipo_Red_Secundaria'] = CTesc.Descripcion

ConductorTramo = Glosario_Tramo.loc[Glosario_Tramo.Editor == 'TPCO',:]
CTesc = pd.merge(left=Tramo, right=ConductorTramo,how='left', \
        left_on='Tipo de Conductor',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'tip_cond'] = CTesc.Codigo_y
Tramo.loc[:,'Tipo_de_conductor'] = CTesc.Descripcion

Tramo.loc[:,'Cantidad_de_conductores'] = Tramo.loc[:,'Cantidad de conductores']

Tramo.loc[:,'Material Conductor 01'] = Tramo.loc[:,'Material Conductor 01'].str.upper()
Tramo.reset_index(inplace=True)
Tramo = Tramo.drop(['index'],axis=1)

Tramo.loc[:,'Calibre Conductor 01'] = Tramo.loc[:,'Calibre Conductor 01'].astype(str)
Tramo.loc[:,'Calibre Conductor 02'] = Tramo.loc[:,'Calibre Conductor 02'].astype(str)
Tramo.loc[:,'Calibre Conductor 03'] = Tramo.loc[:,'Calibre Conductor 03'].astype(str)
Tramo.loc[:,'Calibre Conductor Neutro'] = Tramo.loc[:,'Calibre Conductor Neutro'].astype(str)

# Si el tipo de conductor es trenzado y el material es aluminio con calibre 3/0 
# se debe redondear a 4/0
Sel = Tramo.loc[:,'Tipo de Conductor'].str.contains('TRENZADO') & \
    Tramo.loc[:,'Material Conductor 01'].str.contains('ALUMINIO') & \
     Tramo.loc[:,'Calibre Conductor 01'].str.contains('3/0')
Tramo.loc[Sel,'Calibre Conductor 01'] = '4/0'
    

Tramo.loc[:,'Material Conductor 01'] = [normalize(str(Tramo.loc[i,'Material Conductor 01'])) \
                            for i in range(0, len(Tramo))]

MatConductorTramo = Glosario_Tramo.loc[Glosario_Tramo.Editor == 'TRCM',:]
MatConductorTramo.loc[:,'Descripcion1'] = MatConductorTramo.loc[:,'Descripcion'].str.upper()
MatConductorTramo.reset_index(inplace=True)
MatConductorTramo = MatConductorTramo.drop(['index'],axis=1)
MatConductorTramo.loc[:,'Descripcion1'] = [normalize(str(MatConductorTramo.loc[i,'Descripcion1'])) \
                            for i in range(0, len(MatConductorTramo))]

CTesc = pd.merge(left=Tramo, right=MatConductorTramo,how='left', \
        left_on='Material Conductor 01',right_on='Descripcion1',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'commonconductortype'] = CTesc.Codigo_y
Tramo.loc[:,'Material_Fase_1'] = CTesc.Descripcion

Tramo.loc[:,'Material Conductor 02'] = Tramo.loc[:,'Material Conductor 02'].fillna('')
CTesc = pd.merge(left=Tramo, right=MatConductorTramo,how='left', \
        left_on='Material Conductor 02',right_on='Descripcion1',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'MATERIAL_FASE_2'] = CTesc.Codigo_y
Tramo.loc[:,'Material_Fase_2'] = CTesc.Descripcion

Tramo.loc[:,'Material Conductor 03'] = Tramo.loc[:,'Material Conductor 03'].fillna('')
CTesc = pd.merge(left=Tramo, right=MatConductorTramo,how='left', \
        left_on='Material Conductor 03',right_on='Descripcion1',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'MATERIAL_FASE_3'] = CTesc.Codigo_y
Tramo.loc[:,'Material_Fase_3'] = CTesc.Descripcion


Tramo.loc[:,'Calibre Conductor 01'] = Tramo.loc[:,'Calibre Conductor 01'].fillna('')
Tramo.loc[:,'Calibre Conductor 02'] = Tramo.loc[:,'Calibre Conductor 02'].fillna('')
Tramo.loc[:,'Calibre Conductor 03'] = Tramo.loc[:,'Calibre Conductor 03'].fillna('')
CalConductorTramo = Glosario_Tramo.loc[Glosario_Tramo.Editor == 'TCON',:]
CalConductorTramo.Descripcion = CalConductorTramo.Descripcion.astype(str)
CalConductorTramo.reset_index(inplace=True, drop=True)
CTesc = pd.merge(left=Tramo, right=CalConductorTramo,how='left', \
        left_on='Calibre Conductor 01',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'TIP_COND_C1'] = CTesc.Codigo_y
# Tramo.loc[:,'Calibre_del_conductor_1'] = CTesc.loc[:,'Calibre Conductor 01']
Tramo.loc[:,'Calibre_del_conductor_1'] = CTesc.loc[:,'Descripcion']

CTesc = pd.merge(left=Tramo, right=CalConductorTramo,how='left', \
        left_on='Calibre Conductor 02',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'TIP_COND_C2'] = CTesc.Codigo_y
# Tramo.loc[:,'Calibre_del_conductor_2'] = CTesc.loc[:,'Calibre Conductor 02']
Tramo.loc[:,'Calibre_del_conductor_2'] = CTesc.loc[:,'Descripcion']

Tramo.loc[:,'Calibre Conductor 03'] = Tramo.loc[:,'Calibre Conductor 03'].astype(str)
CTesc = pd.merge(left=Tramo, right=CalConductorTramo,how='left', \
        left_on='Calibre Conductor 03',right_on='Descripcion',suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'TIP_COND_C3'] = CTesc.Codigo_y
# Tramo.loc[:,'Calibre_del_conductor_3'] = CTesc.loc[:,'Calibre Conductor 03']
Tramo.loc[:,'Calibre_del_conductor_3'] = CTesc.loc[:,'Descripcion']

Tramo.loc[:,'Tipo_de_Conexion'] = ''

# =============================================================================
# Preguntar
# =============================================================================
Tramo.loc[:,'Material Conductor 01'] = Tramo.loc[:,'Material Conductor 01'].fillna('')
Tramo.loc[:,'Material Conductor 02'] = Tramo.loc[:,'Material Conductor 02'].fillna('')
Tramo.loc[:,'Material Conductor 03'] = Tramo.loc[:,'Material Conductor 03'].fillna('')
Tramo.loc[:,'Calibre Conductor 01'] = Tramo.loc[:,'Calibre Conductor 01'].fillna('')
Tramo.loc[:,'Calibre Conductor 02'] = Tramo.loc[:,'Calibre Conductor 02'].fillna('')
Tramo.loc[:,'Calibre Conductor 03'] = Tramo.loc[:,'Calibre Conductor 03'].fillna('')
Sel_NEmp = (Tramo.loc[:,'Material Conductor 01'].str.len() == 0) & \
    (Tramo.loc[:,'Material Conductor 02'].str.len() != 0) & \
    (Tramo.loc[:,'Calibre Conductor 01'].str.len() != 0) & \
    (Tramo.loc[:,'Calibre Conductor 02'].str.len() != 0)
    
Sel_ = (Tramo.loc[:,'Material Conductor 01'] != Tramo.loc[:,'Material Conductor 02']) |\
    (Tramo.loc[:,'Calibre Conductor 01'] != Tramo.loc[:,'Calibre Conductor 02'])
    
Tramo.loc[:,'Cantidad_de_Conductores_1'] = Tramo.loc[:,'Cantidad de conductores']
Tramo.loc[:,'Cantidad_de_Conductores_2'] = ''
Tramo.loc[:,'Mat_Cal_Cond1'] = Tramo.loc[:,'Material Conductor 01'].str.lower() + \
        Tramo.loc[:,'Calibre Conductor 01'].astype(str)
Tramo.loc[:,'Mat_Cal_Cond2'] = Tramo.loc[:,'Material Conductor 02'].str.lower() + \
        Tramo.loc[:,'Calibre Conductor 02'].astype(str)
Tramo.loc[:,'Mat_Cal_Cond3'] = Tramo.loc[:,'Material Conductor 03'].str.lower() + \
        Tramo.loc[:,'Calibre Conductor 03'].astype(str)
Tramo.loc[:,'Mat_Cal_Cond1'] = Tramo.loc[:,'Mat_Cal_Cond1'].fillna('')
Tramo.loc[:,'Mat_Cal_Cond2'] = Tramo.loc[:,'Mat_Cal_Cond2'].fillna('')
Tramo.loc[:,'Mat_Cal_Cond3'] = Tramo.loc[:,'Mat_Cal_Cond3'].fillna('')
Tramo.loc[:,'Mat_Cal_Cond1'] = Tramo.loc[:,'Mat_Cal_Cond1'].replace('nan','')
Tramo.loc[:,'Mat_Cal_Cond2'] = Tramo.loc[:,'Mat_Cal_Cond2'].replace('nan','')
Tramo.loc[:,'Mat_Cal_Cond3'] = Tramo.loc[:,'Mat_Cal_Cond3'].replace('nan','')
      
A = Tramo.loc[:,'Cantidad_de_Conductores_1'].fillna(0)
A = A.astype(int)
Sel = A.astype(str) == '3'
Sel_3 = (Tramo.loc[:,'Mat_Cal_Cond1'].str.len() != 0) & (Tramo.loc[:,'Mat_Cal_Cond2'].str.len() == 0) & \
    (Tramo.loc[:,'Mat_Cal_Cond2'].str.len() == 0)
Tramo.loc[Sel & Sel_3,'Cantidad_de_Conductores_1']  = '3'
Tramo.loc[Sel & Sel_3,'Cantidad_de_Conductores_2'] = '0'
Tramo.loc[Sel & Sel_3,'Cantidad_de_Conductores_3'] = '0'

Sel = A.astype(str) == '2'
Tramo.loc[Sel,'Cantidad_de_Conductores_1']  = '2'
Tramo.loc[Sel,'Cantidad_de_Conductores_2'] = '0'
Tramo.loc[Sel,'Cantidad_de_Conductores_3'] = '0'

Sel_3 = (Tramo.loc[:,'Mat_Cal_Cond1'].str.len() != 0) & (Tramo.loc[:,'Mat_Cal_Cond2'].str.len() != 0)
Sel_4 = Tramo.loc[:,'Mat_Cal_Cond1'] != Tramo.loc[:,'Mat_Cal_Cond2']
Tramo.loc[Sel & Sel_3 & Sel_4,'Cantidad_de_Conductores_1']  = '1'
Tramo.loc[Sel & Sel_3 & Sel_4,'Cantidad_de_Conductores_2'] = '1'
Tramo.loc[Sel & Sel_3 & Sel_4,'Cantidad_de_Conductores_3'] = '0'

Tramo.loc[Sel_NEmp & Sel_,'Cantidad_de_Conductores_1']  = '1'
Tramo.loc[Sel_NEmp & Sel_,'Cantidad_de_Conductores_2'] = '1'
Tramo.loc[Sel_NEmp & Sel_,'Cantidad_de_Conductores_3'] = '0'
# Aplicar la regla para cuando sean los tres calibres 3 (si viene el 3 siempre hay 1 y 2)
Sel = (Tramo.loc[:,'Cantidad_de_Conductores_1'].str.len() != 0) & \
 (Tramo.loc[:,'Cantidad_de_Conductores_2'].str.len() == 0) & \
     (~Tramo.loc[:,'Cantidad_de_Conductores_1'].isna())
Tramo.loc[Sel,'Cantidad_de_Conductores_2'] = '0'
# l 3 siempre hay 1 y 2)
# Sel = (Tramo.loc[:,'Cantidad_de_Conductores_1'].str.len() != 0) & \
#  (Tramo.loc[:,'Cantidad_de_Conductores_3'].str.len() == 0) & \
#   (~Tramo.loc[:,'Cantidad_de_Conductores_1'].isna())
Tramo.loc[Sel,'Cantidad_de_Conductores_3'] = '0'
# Tramo.loc[:,'Cantidad_de_Conductores_3'] = ''


CTesc = pd.merge(left=Tramo, right=MatConductorTramo,how='left', \
        left_on='Material conductor Neutro',right_on='Descripcion1',\
            suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'MAT_NEU'] = CTesc.Codigo_y
Tramo.loc[:,'Material_Neutro'] = CTesc.loc[:,'Descripcion']

CTesc = pd.merge(left=Tramo, right=MatConductorTramo,how='left', \
        left_on='Material conductor Neutro',right_on='Descripcion1',\
            suffixes=(None, '_y'),indicator=True)
Tramo.loc[:,'MAT_NEU'] = CTesc.Codigo_y
Tramo.loc[:,'Material_Neutro'] = CTesc.loc[:,'Descripcion']


CTesc = pd.merge(left=Tramo, right=CalConductorTramo,how='left', \
        left_on='Calibre Conductor Neutro',right_on='Descripcion',\
            suffixes=(None, '_y'),indicator=True)

Tramo.loc[:,'NEUTRO'] = CTesc.Codigo_y
Tramo.loc[:,'Calibre_del_conductor_del_neutro'] = CTesc.loc[:,'Descripcion']

Tramo.loc[:,'UUCC_1'] = 'km de conductor/fase ' + Tramo.loc[:,'Tipo de tramo'].str.lower() + \
    ' ' + Tramo.loc[:,'Desc Tipo de Área'].str.lower() + \
    '-' + Tramo.loc[:,'Tipo de Conductor'].str.lower() + '-' \
        + Tramo.loc[:,'Material Conductor 01'].str.lower() + '-' + 'calibre ' + \
        Tramo.loc[:,'Calibre Conductor 01'].astype(str)
Tramo.loc[:,'UUCC_2'] = 'km de conductor/fase ' + Tramo.loc[:,'Tipo de tramo'].str.lower() + \
        ' ' + Tramo.loc[:,'Desc Tipo de Área'].str.lower() + \
    '-' + Tramo.loc[:,'Tipo de Conductor'].str.lower() + '-' \
        + Tramo.loc[:,'Material Conductor 02'].str.lower() + '-' + 'calibre ' + \
        Tramo.loc[:,'Calibre Conductor 02'].astype(str)            
# Tramo.loc[:,'UUCC_2'] = 'km de conductor/fase ' + Tramo.loc[:,'Tipo de tramo'].str.lower() + \
#     ' urbano-' + Tramo.loc[:,'Tipo de Conductor'].str.lower() + '-' \
#         + Tramo.loc[:,'Material Conductor 02'].str.lower() + '-' + 'calibre ' + \
        # Tramo.loc[:,'Calibre Conductor 02'].astype(str)
Tramo.loc[:,'UUCC_3'] = 'km de conductor/fase ' + Tramo.loc[:,'Tipo de tramo'].str.lower() + \
        ' ' +  Tramo.loc[:,'Desc Tipo de Área'].str.lower() + \
    '-' + Tramo.loc[:,'Tipo de Conductor'].str.lower() + '-' \
        + Tramo.loc[:,'Material Conductor 03'].str.lower() + '-' + 'calibre ' + \
        Tramo.loc[:,'Calibre Conductor 03'].astype(str)        
# Tramo.loc[:,'UUCC_3'] = 'km de conductor/fase ' + Tramo.loc[:,'Tipo de tramo'].str.lower() + \
#     ' urbano-' + Tramo.loc[:,'Tipo de Conductor'].str.lower() + '-' \
#         + Tramo.loc[:,'Material Conductor 03'].str.lower() + '-' + 'calibre ' + \
#         Tramo.loc[:,'Calibre Conductor 03'].astype(str)
Tramo.loc[:,'UUCC_4'] = 'km de conductor/fase ' + Tramo.loc[:,'Tipo de tramo'].str.lower() + \
        ' ' +  Tramo.loc[:,'Desc Tipo de Área'].str.lower() + \
    '-' + Tramo.loc[:,'Tipo de Conductor'].str.lower() + '-' \
        + Tramo.loc[:,'Material conductor Neutro'].str.lower() + '-' + 'calibre ' + \
        Tramo.loc[:,'Calibre Conductor Neutro'].astype(str) 
# Tramo.loc[:,'UUCC_4'] = 'km de conductor/fase ' + Tramo.loc[:,'Tipo de tramo'].str.lower() + \
#     ' urbano-' + Tramo.loc[:,'Tipo de Conductor'].str.lower() + '-' \
#         + Tramo.loc[:,'Material conductor Neutro'].str.lower() + '-' + 'calibre ' + \
#         Tramo.loc[:,'Calibre Conductor Neutro'].astype(str)       

Tramo.loc[:,'UUCC_1'] = Tramo.loc[:,'UUCC_1'].astype(str).apply(str.replace,args=(' ', ''))
Tramo.loc[:,'UUCC_2'] = Tramo.loc[:,'UUCC_2'].astype(str).apply(str.replace,args=(' ', ''))
Tramo.loc[:,'UUCC_3'] = Tramo.loc[:,'UUCC_3'].astype(str).apply(str.replace,args=(' ', ''))
Tramo.loc[:,'UUCC_4'] = Tramo.loc[:,'UUCC_4'].astype(str).apply(str.replace,args=(' ', ''))
B = []
Tramo.loc[:,'Unidad_constructiva_R015']  = ''
Tramo.reset_index(inplace=True, drop=True)
for j in range(1,5,1):
    uucc = pd.Series(Tramo.loc[:,'UUCC_' + str(j)])
    DataList=[]
    In = pd.DataFrame(uucc, columns=['UUCC_' + str(j)])
    In.reset_index(inplace=True)
    In = In.drop(['index'],axis=1)    
    for i in range(0,len(Tramo.loc[:,'UUCC_' + str(j)]),1):
        In2 = pd.DataFrame([''],columns=['UUCC_' + str(j)])
        if i < len(Tramo.loc[:,'UUCC_' + str(j)])-1:
            In1 = pd.concat([In.loc[i:i+1,:],In2])
            Sel = [True,False,True]
            In1 = In1[Sel]
            In1.reset_index(inplace=True)
            In1 = In1.drop(['index'],axis=1)
        else:
            In1 = pd.concat([In.loc[i-1:i,:],In2])
            Sel = [False,True,True]
            In1 = In1[Sel]
            In1.reset_index(inplace=True)
            In1 = In1.drop(['index'],axis=1)        
        DataCruce = pd.merge(left=In1,\
                             right=UUCC_Tramo,how='left', left_on='UUCC_' + str(j),right_on='DESCRIPCION',suffixes=(None, '_y'),\
                              indicator=True) 
        DataCruce = DataCruce.iloc[:-1,:]
        # DataCruce = DataCruce.loc[DataCruce.loc[:,'_merge'] == 'both',:]
        if np.sum(DataCruce._merge == 'both') > 0:
            DataCruce = DataCruce.groupby(['_merge'])['UC'].transform(\
                                lambda x : '|'.join(x))
            DataCruce = DataCruce.drop_duplicates()
        else:
            DataCruce = pd.Series(np.array([' ']))
        DataList.append(DataCruce) 
    DataListF = pd.concat(DataList)
    DataListF.reset_index(inplace=True, drop=True)
    DataListF = DataListF.to_frame(name="UUCC")
    Tramo.loc[:,'Unidad_constructiva_R015'] = Tramo.loc[:,'Unidad_constructiva_R015'].str.cat(DataListF, sep =" ")
Tramo.loc[:,'Unidad_constructiva_R015'] = Tramo.loc[:,'Unidad_constructiva_R015'].apply(lambda x: ' '.join(unique_list(x.split())))
Tramo.loc[:,'Unidad_constructiva_R015'] = Tramo.loc[:,'Unidad_constructiva_R015'].apply(str.replace,args=(' ', '|'))
# DataListF = pd.DataFrame(DataListF, columns=["UUCC"])
# DataListF.reset_index(inplace=True)
# DataListF = DataListF.drop(['index'],axis=1)
# Tramo.loc[:,'Unidad_constructiva_R015'] = DataListF
Tramo.loc[:,'Observaciones'] = Tramo.loc[:,'Observación Tramo BT']
Tramo.loc[:,'ORIGEN'] = '17'
Tramo.loc[:,'Origen_de_los_datos'] = 'CENSO II'
Tramo.loc[:,'Longitud_(mts)'] = ''
Tramo.loc[:,'Longitud_Fija_(mts)'] = ''
Tramo.loc[:,'Apoyo_inicial'] = ''
Tramo.loc[:,'Apoyo_final'] = ''

Tramo.rename(columns={'BDI':'cod. BDIv10','Longitud':'Longitud_final',\
                             'Latitud':'Latitud_final'},inplace=True) 
Tramo.loc[:,'Calibre_del_conductor_1'] = Tramo.loc[:,'Calibre_del_conductor_1'].astype(str)
A = Tramo.loc[:,'Calibre_del_conductor_1'].str.split("/", n=2, expand=True)
A.iloc[:,0] = A.iloc[:,0].str.replace(' ','0')
A = A.rename(columns={0:'Cal',1:'Cal1'})
A.iloc[:,0] = A.iloc[:,0].str.replace('<Null>','0')
A.loc[A.loc[:,'Cal'].isna(),'Cal'] = 0
A.loc[(A.loc[:,'Cal'].str.len() == 0),'Cal'] = 0
A.loc[A.loc[:,'Cal'].isna(),'Cal'] = 0
A.loc[A.loc[:,'Cal'].str.contains('nan'),'Cal'] = 0
A.loc[:,'Cal'] = A.loc[:,'Cal'].astype(int)
    
SelV = (Tramo.loc[:,'Tipo_de_tramo'] == 'AEREO') & \
    (Tramo.loc[:,'Tipo_de_conductor'] == 'AISLADO') & \
        (Tramo.loc[:,'Material_Fase_1'] == 'Aluminio') & \
           (A.iloc[:,0] > 6)
SelV = SelV | (Tramo.loc[:,'Tipo_de_tramo'] == 'SUBTERRANEO') & \
    (Tramo.loc[:,'Tipo_de_conductor'] == 'AISLADO') & \
        (Tramo.loc[:,'Material_Fase_1'] == 'Aluminio') & \
           (A.iloc[:,0] > 6)           
Tramo.loc[SelV,'Observaciones'] = Tramo.loc[SelV,'Observaciones'] + 'Calibre validado'         
    
Tramo = Tramo.loc[:,['Equipo Ruta Id','Nombre Equipo','Equipo padre vinculado',\
'Nombre Equipo Padre','Rutaid',\
'Nombre Ruta','cod. BDIv10','Tipo_de_Conexion','Tipo de Apoyo','Material del apoyo',\
'Tipo de red','Estructura','Foto Apoyo BT AE 01','Foto Tramo BT 01',\
'Foto Tramo BT 02','Codigo','Instalacion_origen','Apoyo_final','Apoyo_inicial',\
'Longitud_(mts)','Longitud_Fija_(mts)','assettype','Tipo_de_tramo',\
'lifecyclestatus','Estado_Tramo','TIP_RED_SEC','Tipo_Red_Secundaria','tip_cond',\
'Tipo_de_conductor','Cantidad_de_conductores','commonconductortype',\
'Material_Fase_1','MATERIAL_FASE_2','Material_Fase_2','MATERIAL_FASE_3',\
'Material_Fase_3','TIP_COND_C1','Calibre_del_conductor_1','TIP_COND_C2',\
'Calibre_del_conductor_2','TIP_COND_C3','Calibre_del_conductor_3',\
'Cantidad_de_Conductores_1','Cantidad_de_Conductores_2',\
'Cantidad_de_Conductores_3','MAT_NEU','Material_Neutro','NEUTRO',\
'Calibre_del_conductor_del_neutro','Unidad_constructiva_R015',\
'Observaciones','ORIGEN','Origen_de_los_datos','Longitud_final','Latitud_final',\
'Longitud Equipo Padre','Latitud Equipo Padre','Tipo de Área', 'Desc Tipo de Área']]
Tramo = Tramo.fillna('')
Tramo.Tipo_de_tramo = Tramo.Tipo_de_tramo.replace('nan','')
Tramo.Calibre_del_conductor_1 = Tramo.Calibre_del_conductor_1.replace('nan','')
# =============================================================================
# # Caja de abonados
# =============================================================================
TCAJ = Glosario_Caja.loc[Glosario_Caja.Editor == 'TCAJ',:]
TCAJ.loc[:,'Descripcion1'] = TCAJ.loc[:,'Descripcion'].str.upper()
TCAJ.reset_index(inplace=True)
TCAJ = TCAJ.drop(['index'],axis=1)
TCAJ.loc[:,'Descripcion1'] = [normalize(str(TCAJ.loc[i,'Descripcion1'])) \
                            for i in range(0, len(TCAJ))]
CAbonados = MLU
# CAbonados = CAbonados.loc[(CAbonados.loc[:,'¿Hay Caja?'].str.len() != 0) & \
#                   (~CAbonados.loc[:,'¿Hay Caja?'].isna()),:]
CAbonados.loc[:,'¿Hay Caja?'] = CAbonados.loc[:,'¿Hay Caja?'].fillna('')
CAbonados = CAbonados.loc[(CAbonados.loc[:,'¿Hay Caja?'].str.contains('Si')),:]# | \
                           # (CAbonados.loc[:,'¿Hay Caja?'].str.len() != 0),:] 
CAbonados.loc[:,'Cantidad de caja'] = CAbonados.loc[:,'Cantidad de caja'].fillna('')
CAbonados.reset_index(inplace=True, drop = True)
# Cuando el campo "Cantidad de caja" sea mayor a 1 se debe duplicar el mismo registros
# la cantidad de veces que hay indique y se va añadiendo el código como se indica:
# Si son 3 el código sería 41 + equiporutaid, 42 + equiporutaid y 43 + equiporutaid.
CantidadCaja = CAbonados.loc[:,'Cantidad de caja'].replace('','0')
CantidadCaja = CantidadCaja.astype(int)
CantidadCaja.reset_index(inplace=True,drop=True)
CAbonados = CAbonados.loc[CAbonados.index.repeat(CantidadCaja)]
Sel = CAbonados.index.duplicated(keep=False)
Filter = CAbonados.loc[CAbonados.loc[:,'Equipo Ruta Id'].duplicated(keep='first'),'Equipo Ruta Id']
Filter.reset_index(inplace=True,drop=True)
CantidadCaja = CAbonados.loc[:,'Cantidad de caja']
for i in range(0,len(Filter),1):
    Sel = CAbonados.loc[:,'Equipo Ruta Id'] == Filter[i]
    CantidadCaja[Sel] = np.linspace(1,np.sum(Sel),np.sum(Sel))
# 
CantidadCaja = CAbonados.loc[:,'Cantidad de caja'].astype(int)
CAbonados.loc[:,'Codigo'] = '4' + \
   CantidadCaja.astype(str) + CAbonados.loc[:,'Equipo Ruta Id'].astype(str)
# CAbonados.loc[:,'Codigo'] = '4' + \
#    CAbonados.loc[:,'Cantidad de caja'].astype(str) + CAbonados.loc[:,'Equipo Ruta Id'].astype(str)
CAbonados.loc[:,'Instalacion_Superior'] = CAbonados.loc[:,'BDI']
CAbonados.loc[:,'Ubicación caja'] = CAbonados.loc[:,'Ubicación caja'].fillna('')
CAbonados.loc[:,'Ubicacion_Caja_Abonado'] = CAbonados.loc[:,'Ubicación caja'].apply(str.replace,args=('Apoyo','Poste'))
CAbonados.loc[:,'desc_Ubicacion_Caja_Abonado'] = CAbonados.loc[:,'Ubicacion_Caja_Abonado']
CAbonados.loc[:,'Tipo_de_caja'] = CAbonados.loc[:,'Tipo de caja']
CAbonados.loc[:,'Tipo_de_caja'] = CAbonados.loc[:,'Tipo_de_caja'].replace('Derivación (Caja de abonado)','Derivacion')
CAbonados.loc[:,'Tipo_de_caja'] = CAbonados.loc[:,'Tipo_de_caja'].replace('Derivación con medida centralizada','Derivacion con medida')
CAbonados.loc[:,'Tipo_de_caja'] = CAbonados.loc[:,'Tipo_de_caja'].str.upper()
CAbonados.reset_index(inplace=True,drop=True)
# CAbonados = CAbonados.drop(['index'],axis=1)
CAbonados.loc[:,'Tipo_de_caja'] = [normalize(str(CAbonados.loc[i,'Tipo_de_caja'])) \
                            for i in range(0, len(CAbonados.Tipo_de_caja))]
Tesc = pd.merge(left=CAbonados, right=TCAJ,how='left', \
        left_on='Tipo_de_caja',right_on='Descripcion1',\
            suffixes=(None, '_y'),indicator=True)

CAbonados.loc[:,'Tipo_de_caja'] = Tesc.Codigo_y
CAbonados.loc[:,'desc_Tipo_de_caja'] = Tesc.loc[:,'Descripcion']
# Sel1 = [CAbonados.loc[i,'Tipo_de_caja'].contains())]) \
#                             for i in range(0, len(CAbonados))]

# CAbonados.loc[:,'desc_Tipo_de_caja'] = ''
NFASE = Glosario_Caja.loc[Glosario_Caja.Editor == 'NFASE',:]
NFASE.loc[:,'Descripcion1'] = NFASE.loc[:,'Descripcion'].str.upper()
NFASE.reset_index(inplace=True)
NFASE = NFASE.drop(['index'],axis=1)
NFASE.loc[:,'Descripcion1'] = [normalize(str(NFASE.loc[i,'Descripcion1'])) \
                            for i in range(0, len(NFASE))]  
CAbonados.loc[:,'NFASESCAJA'] = CAbonados.loc[:,'Número de fases caja 01'].str.upper()
CAbonados.loc[:,'NFASESCAJA'] = [normalize(str(CAbonados.loc[i,'NFASESCAJA'])) \
                            for i in range(0, len(CAbonados))]     
Tesc = pd.merge(left=CAbonados, right=NFASE,how='left', \
        left_on='NFASESCAJA',right_on='Descripcion1',\
            suffixes=(None, '_y'),indicator=True)

CAbonados.loc[:,'NUMERO_FASES'] = Tesc.Codigo_y
CAbonados.loc[:,'desc_Numero_de_Fases'] = Tesc.loc[:,'Descripcion']
CAbonados.loc[:,'UC_R015'] = CAbonados.loc[:,'Unidad constructiva R015 Caja de abonado']
CAbonados.loc[:,'Observaciones'] = '' 
# Sel = CAbonados.loc[:,'Ubicacion_Caja_Abonado'].str.len() == 0
# CAbonados.loc[Sel,'Ubicacion_Caja_Abonado'] = 'Apoyo'
# CAbonados.loc[Sel,'desc_Ubicacion_Caja_Abonado'] = 'Apoyo'

CAbonados.loc[:,'ORIGEN'] = '17'
CAbonados.loc[:,'Origen_de_los_datos'] = 'CENSO II'

CAbonados = CAbonados.fillna('')
CAbonados = CAbonados.loc[~CAbonados.loc[:,'Tipo de caja'].str.contains('nan'),:]
CAbonados.reset_index(inplace=True)
CAbonados = CAbonados.drop(columns = 'index')

# Cuando desc_Tipo_de_caja sea Derivación con Medida y Ubicacion_Caja_Abonado este 
# en blanco se debe asignar Apoyo
Sel = (CAbonados.loc[:,'desc_Tipo_de_caja'] == 'Derivación con Medida') & \
    (CAbonados.loc[:,'Ubicacion_Caja_Abonado'].str.len() == 0)
CAbonados.loc[Sel,'Ubicacion_Caja_Abonado'] = 'Poste'
CAbonados.loc[Sel,'desc_Ubicacion_Caja_Abonado'] = 'Poste' 
CAbonados = CAbonados.loc[:,['Codigo','Instalacion_Superior',\
                            'Ubicacion_Caja_Abonado','desc_Ubicacion_Caja_Abonado',\
'Tipo_de_caja','desc_Tipo_de_caja','NUMERO_FASES','desc_Numero_de_Fases',\
'UC_R015','Observaciones','ORIGEN','Origen_de_los_datos','Longitud','Latitud',\
    'Nombre Ruta','Foto Caja de abonado']]
CAbonados.loc[:,'Tipo_de_caja'] = CAbonados.Tipo_de_caja.replace('nan','')
# =============================================================================
# # Puente BT
# =============================================================================
PUENTEBT = Glosario_Puente.loc[Glosario_Puente.Editor == 'ETRF',:]
PUENTEBT.loc[:,'Descripcion1'] = PUENTEBT.loc[:,'Descripcion'].str.upper()
PUENTEBT.reset_index(inplace=True)
PUENTEBT = PUENTEBT.drop(['index'],axis=1)
PUENTEBT.loc[:,'Descripcion1'] = [normalize(str(PUENTEBT.loc[i,'Descripcion1'])) \
                            for i in range(0, len(PUENTEBT))]
PuenteBT = MLU_Tramo
# PuenteBT = PuenteBT.loc[(PuenteBT.loc[:,'¿Hay puente BT?'].str.len() != 0) & \
                  # (~PuenteBT.loc[:,'¿Hay puente BT?'].isna()),:]
PuenteBT.loc[:,'¿Hay puente BT?'] = PuenteBT.loc[:,'¿Hay puente BT?'].fillna('')                  
PuenteBT = PuenteBT.loc[PuenteBT.loc[:,'¿Hay puente BT?'].str.contains('Si'),:]    
PuenteBT.loc[:,'Codigo'] =  '17' + PuenteBT.loc[:,'Equipo Ruta Id'].astype(str)
PuenteBT.loc[:,'Instalacion_origen'] = PuenteBT.loc[:,'BDI']

# PuenteBT = PuenteBT.drop(['level_0'],axis=1)
PuenteBT.reset_index(inplace=True)
PuenteBT = PuenteBT.drop(['index'],axis=1)
if PuenteBT.shape[0] > 0:
    PuenteBT.loc[:,'Estado del puente'] = PuenteBT.loc[:,'Estado del puente'].str.upper()
    Tesc = pd.merge(left=PuenteBT, right=PUENTEBT,how='left', \
            left_on='Estado del puente',right_on='Descripcion1',\
                suffixes=(None, '_y'),indicator=True)
    
    PuenteBT.loc[:,'Estado_Elemento'] = Tesc.Codigo_y
    PuenteBT.loc[:,'DESC_Estado_Elemento'] = Tesc.loc[:,'Descripcion']
    
    
    PuenteBT.loc[:,'ORIGEN'] = '17'
    PuenteBT.loc[:,'DESC_Origen_de_los_datos'] = 'CENSO II'
    
    PuenteBT = PuenteBT.fillna('')
    PuenteBT = PuenteBT.loc[:,['Codigo','Instalacion_origen','Estado_Elemento','DESC_Estado_Elemento',\
    'ORIGEN','DESC_Origen_de_los_datos','Longitud','Latitud','Nombre Ruta',\
        'Soporte Puente BT']]
else:
    PuenteBT = MLU_Tramo.loc[:1,:]
    PuenteBT.loc[:,'Codigo'] = ''
    PuenteBT.loc[:,'Instalacion_origen'] = ''
    PuenteBT.loc[:,'Estado_Elemento'] = ''
    PuenteBT.loc[:,'DESC_Estado_Elemento'] = ''
    
    PuenteBT.loc[:,'ORIGEN'] = ''
    PuenteBT.loc[:,'DESC_Origen_de_los_datos'] = ''
    PuenteBT.loc[:,'Nombre Ruta'] = '' 
    PuenteBT = PuenteBT.loc[:,['Codigo','Instalacion_origen','Estado_Elemento','DESC_Estado_Elemento',\
    'ORIGEN','DESC_Origen_de_los_datos','Longitud','Latitud','Nombre Ruta',\
        'Soporte Puente BT']]
    PuenteBT.Longitud = ''
    PuenteBT.Latitud = '' 


# =============================================================================
# # Transicion
# =============================================================================
# Trans = Glosario_Puente.loc[Glosario_Puente.Editor == 'ETRF',:]
# PUENTEBT.loc[:,'Descripcion1'] = PUENTEBT.loc[:,'Descripcion'].str.upper()
# PUENTEBT.reset_index(inplace=True)
# PUENTEBT = PUENTEBT.drop(['index'],axis=1)
# PUENTEBT.loc[:,'Descripcion1'] = [normalize(str(PUENTEBT.loc[i,'Descripcion1'])) \
#                             for i in range(0, len(PUENTEBT))]
Transicion = MLU
# Transicion = Transicion.loc[(Transicion.loc[:,'¿Paso Aéreo/Subterráneo (Transición)?'].str.len() != 0) & \
#                   (~Transicion.loc[:,'¿Paso Aéreo/Subterráneo (Transición)?'].isna()),:]
Transicion.loc[:,'¿Paso Aéreo/Subterráneo (Transición)?'] = Transicion.loc[:,'¿Paso Aéreo/Subterráneo (Transición)?'].fillna('')
Transicion = Transicion.loc[Transicion.loc[:,'¿Paso Aéreo/Subterráneo (Transición)?'].str.contains('Si'),:] 
Transicion.loc[:,'Codigo'] =  '14' + Transicion.loc[:,'Equipo Ruta Id'].astype(str)
Transicion.loc[:,'Instalacion_origen'] = Transicion.loc[:,'BDI']

# Transicion = Transicion.drop(['level_0'],axis=1)
Transicion.reset_index(inplace=True)
Transicion = Transicion.drop(['index'],axis=1)

# Se condicina el campo longitud_fija_(mts) dependiendo del campo Altura del apoyo
# Si Altura del apoyo >= 9 la longitud fija será igual a 8, si Alturadel apoyo 
# = 8 la longitud fija será igual a 7.
Transicion.loc[:,'Longitud_Fija_(mts)'] = Transicion.loc[:,'Longitud de la transición (Metros)']
Transicion.loc[Transicion.loc[:,'Altura del apoyo'] >= 9,'Longitud_Fija_(mts)'] = 8
Transicion.loc[Transicion.loc[:,'Altura del apoyo'] <= 8,'Longitud_Fija_(mts)'] = 7
#-----------------------------------
if Transicion.shape[0] > 0:
    Transicion.loc[:,'UC_R015'] = ''
    Transicion.loc[:,'Observaciones'] = ''
    Transicion.loc[:,'Origen_de_los_Datos'] = '17'
    Transicion.loc[:,'Desc_Origen_de_los_Datos'] = 'CENSO II'
    
    Transicion = Transicion.fillna('')
    Transicion = Transicion.loc[:,['Codigo','Instalacion_origen','Longitud_Fija_(mts)','UC_R015','Observaciones',\
    'Origen_de_los_Datos','Desc_Origen_de_los_Datos','Longitud','Latitud',\
        'Nombre Ruta','Soporte Transición']]
else:
    Transicion = MLU.loc[:1,:]
    Transicion.loc[:,'Codigo'] = ''
    Transicion.loc[:,'Instalacion_origen'] = ''
    Transicion.loc[:,'Longitud_Fija_(mts)'] = ''
    Transicion.loc[:,'UC_R015'] = ''
    
    Transicion.loc[:,'Observaciones'] = ''
    Transicion.loc[:,'Origen_de_los_Datos'] = ''
    Transicion.loc[:,'Desc_Origen_de_los_Datos'] = ''
    # Transicion.loc[:,'Longitud'] = ''
    
    # Transicion.loc[:,'Latitud'] = ''
    Transicion.loc[:,'Nombre Ruta'] = ''

    Transicion = Transicion.loc[:,['Codigo','Instalacion_origen','Longitud_Fija_(mts)','UC_R015','Observaciones',\
    'Origen_de_los_Datos','Desc_Origen_de_los_Datos','Longitud','Latitud',\
        'Nombre Ruta','Soporte Transición']]
    Transicion.Longitud = ''
    Transicion.Latitud = ''
    Transicion.loc[:,'Nombre Ruta'] = ''    
# =============================================================================
# # Puesta a tierra
# =============================================================================
# Trans = Glosario_Puente.loc[Glosario_Puente.Editor == 'ETRF',:]
# PUENTEBT.loc[:,'Descripcion1'] = PUENTEBT.loc[:,'Descripcion'].str.upper()
# PUENTEBT.reset_index(inplace=True)
# PUENTEBT = PUENTEBT.drop(['index'],axis=1)
# PUENTEBT.loc[:,'Descripcion1'] = [normalize(str(PUENTEBT.loc[i,'Descripcion1'])) \
#                             for i in range(0, len(PUENTEBT))]
PuestaTierra = MLU
# PuestaTierra = PuestaTierra.loc[(PuestaTierra.loc[:,'¿Puesta tierra?'].str.len() != 0) & \
#                   (~PuestaTierra.loc[:,'¿Puesta tierra?'].isna()),:]
PuestaTierra.loc[:,'¿Puesta tierra?'] = PuestaTierra.loc[:,'¿Puesta tierra?'].fillna('')
PuestaTierra = PuestaTierra.loc[PuestaTierra.loc[:,'¿Puesta tierra?'].str.contains('Si'),:]
PuestaTierra.loc[:,'Codigo'] = '19' + PuestaTierra.loc[:,'Equipo Ruta Id'].astype(str)
PuestaTierra.loc[:,'Instalacion_superior'] = PuestaTierra.loc[:,'BDI']

# PuestaTierra = PuestaTierra.drop(['level_0'],axis=1)
PuestaTierra.reset_index(inplace=True)
PuestaTierra = PuestaTierra.drop(['index'],axis=1)

PuestaTierra.loc[:,'UC_R015'] = PuestaTierra.loc[:,'Unidad constructiva R015 Puesta a tierra']
PuestaTierra.loc[:,'Observaciones'] = PuestaTierra.loc[:,'Código Puesta a tierra']
PuestaTierra.loc[:,'ORIGEN'] = '17'
PuestaTierra.loc[:,'Desc_Origen_de_los_datos'] = 'CENSO II'

PuestaTierra = PuestaTierra.fillna('')
PuestaTierra = PuestaTierra.loc[:,['Codigo','Instalacion_superior','UC_R015','Observaciones','ORIGEN',\
'Desc_Origen_de_los_datos','Longitud','Latitud','Nombre Ruta',\
    'Foto Puesta a tierra']]

# =============================================================================
# # TV Cable
# =============================================================================
TV = Glosario_TVCABLE.loc[Glosario_TVCABLE.Editor == 'SINO',:]
TV.loc[:,'Descripcion1'] = TV.loc[:,'Descripcion'].str.upper()
TV.reset_index(inplace=True)
TV = TV.drop(['index'],axis=1)

TVCable = MLU
# TVCable.loc[:,'Status Tv cable'] = ''
TVCable.loc[:,'Status Tv cable'] = TVCable.loc[:,'Status Tv cable'].astype(str)
TVCable.loc[TVCable.loc[:,'Status Tv cable'].str.contains('nan'),'Status Tv cable'] = ''
TVCable = TVCable.loc[(TVCable.loc[:,'Status Tv cable'].str.len() != 0) & \
                  (~TVCable.loc[:,'Status Tv cable'].isna()),:]
TVCable.loc[:,'Codigo'] = '31' + TVCable.loc[:,'Equipo Ruta Id'].astype(str)
TVCable.loc[:,'Instalacion superior (Trafo)'] = TVCable.loc[:,'BDI']

# TVCable = TVCable.drop(['level_0'],axis=1)
TVCable.reset_index(inplace=True)
TVCable = TVCable.drop(['index'],axis=1)

if TVCable.shape[0] > 0:
    TVCable.loc[:,'TIP_CARGA'] = 'Tpp03'
    TVCable.loc[:,'Tipo de carga'] = 'TV CABLE'
    
    TVCable.loc[:,'ORIGEN'] = '17'
    TVCable.loc[:,'Origen de los datos'] = 'CENSO II'
    
    TVCable.loc[:,'¿Marquilla TV Cable?'] = TVCable.loc[:,'¿Marquilla TV Cable?'].str.upper()
    Tesc = pd.merge(left=TVCable, right=TV,how='left', \
            left_on='¿Marquilla TV Cable?',right_on='Descripcion1',\
                suffixes=(None, '_y'),indicator=True)
    
    TVCable.loc[:,'Marquilla_'] = Tesc.loc[:,'Descripcion']
    TVCable.loc[:,'Marquilla'] = Tesc.Codigo_y
    
    TV = Glosario_TVCABLE.loc[Glosario_TVCABLE.Editor == 'TPC',:]
    TV.loc[:,'Descripcion1'] = TV.loc[:,'Descripcion'].str.upper()
    TV.reset_index(inplace=True)
    TV = TV.drop(['index'],axis=1)
    
    TVCable.loc[:,'Tipo de Tv cable'] = TVCable.loc[:,'Tipo de Tv cable'].str.upper()
    Tesc = pd.merge(left=TVCable, right=TV,how='left', \
            left_on='Tipo de Tv cable',right_on='Descripcion1',\
                suffixes=(None, '_y'),indicator=True)
    
    TVCable.loc[:,'TIP_TV_CABLE'] = Tesc.Codigo_y
    TVCable.loc[:,'Tipo TV Cable'] = Tesc.loc[:,'Descripcion1']
    
    TV = Glosario_TVCABLE.loc[Glosario_TVCABLE.Editor == 'OPTE',:]
    TV.loc[:,'Descripcion1'] = TV.loc[:,'Descripcion'].str.upper()
    TV.reset_index(inplace=True)
    TV = TV.drop(['index'],axis=1)
    
    TVCable.loc[:,'Operador Tv cable'] = TVCable.loc[:,'Operador Tv cable'].fillna('')
    TVCable.loc[:,'Operador Tv cable'] = TVCable.loc[:,'Operador Tv cable'].apply(str.replace,args=('Otros', 'Otro'))    
    TVCable.loc[:,'Operador Tv cable'] = TVCable.loc[:,'Operador Tv cable'].apply(str.replace,args=('Otro', 'Otros'))
    TVCable.loc[:,'Operador Tv cable'] = TVCable.loc[:,'Operador Tv cable'].apply(str.replace,args=('Desconocido', 'Otros'))
    TVCable.loc[:,'Operador Tv cable'] = TVCable.loc[:,'Operador Tv cable'].str.upper()
    Tesc = pd.merge(left=TVCable, right=TV,how='left', \
            left_on='Operador Tv cable',right_on='Descripcion1',\
                suffixes=(None, '_y'),indicator=True)
    
    TVCable.loc[:,'OPERADOR'] = Tesc.Codigo_y
    TVCable.loc[:,'Operador de Telecomunicaciones'] = Tesc.loc[:,'Descripcion']
    
    TV = Glosario_TVCABLE.loc[Glosario_TVCABLE.Editor == 'TMED',:]
    TV.loc[:,'Descripcion1'] = TV.loc[:,'Descripcion'].str.upper()
    
    
    TV.reset_index(inplace=True)
    TV = TV.drop(['index'],axis=1)
    
    TVCable.loc[:,'Tipo de medida Tv cable'] = TVCable.loc[:,'Tipo de medida Tv cable'].str.upper()
    Tesc = pd.merge(left=TVCable, right=TV,how='left', \
            left_on='Tipo de medida Tv cable',right_on='Descripcion1',\
                suffixes=(None, '_y'),indicator=True)
    
    TVCable.loc[:,'TIP_MEDIDA'] = Tesc.Codigo_y
    TVCable.loc[:,'Tipo de medida'] = Tesc.loc[:,'Descripcion1']
    
    TVCable.loc[:,'Marca'] = ''
    TVCable.loc[:,'Numero de serie (Num Medidor)'] = ''
    TVCable.loc[:,'OBSERVACIONES'] = TVCable.loc[:,'Observación Tv cable']
    TVCable = TVCable.fillna('')
    TVCable = TVCable.loc[:,['Codigo','Instalacion superior (Trafo)','TIP_CARGA','Tipo de carga','Marquilla',\
    'Marquilla_','TIP_TV_CABLE','Tipo TV Cable','OPERADOR',\
    'Operador de Telecomunicaciones','TIP_MEDIDA','Tipo de medida','Marca',\
    'Numero de serie (Num Medidor)','OBSERVACIONES','ORIGEN',\
    'Origen de los datos','Longitud','Latitud','Nombre Ruta','Soporte Tv cable']]
else:
    TVCable = MLU.loc[:1,:]
    TVCable.loc[:,'Codigo'] = ''
    TVCable.loc[:,'Instalacion superior (Trafo)'] = ''
    TVCable.loc[:,'TIP_CARGA'] = ''
    TVCable.loc[:,'Tipo de carga'] = ''
    TVCable.loc[:,'Marquilla_'] = ''
    TVCable.loc[:,'Marquilla'] = ''
    TVCable.loc[:,'TIP_TV_CABLE'] = ''
    TVCable.loc[:,'Tipo TV Cable'] = ''
    TVCable.loc[:,'OPERADOR'] = ''
    TVCable.loc[:,'Operador de Telecomunicaciones'] = ''
    TVCable.loc[:,'TIP_MEDIDA'] = ''
    TVCable.loc[:,'Tipo de medida'] = ''
    TVCable.loc[:,'Marca'] = ''
    TVCable.loc[:,'Numero de serie (Num Medidor)'] = ''
    TVCable.loc[:,'OBSERVACIONES'] = ''
    TVCable.loc[:,'ORIGEN'] = ''
    TVCable.loc[:,'Origen de los datos'] = ''
    TVCable.loc[:,'Soporte Tv cable'] = ''
    
    TVCable = TVCable.loc[:,['Codigo','Instalacion superior (Trafo)','TIP_CARGA','Tipo de carga','Marquilla',\
    'Marquilla','TIP_TV_CABLE','Tipo TV Cable','OPERADOR',\
    'Operador de Telecomunicaciones','TIP_MEDIDA','Tipo de medida','Marca',\
    'Numero de serie (Num Medidor)','OBSERVACIONES','ORIGEN',\
    'Origen de los datos','Longitud','Latitud','Nombre Ruta','Soporte Tv cable']]
    TVCable.Longitud = ''
    TVCable.Latitud = ''
    TVCable.loc[:,'Nombre Ruta'] = ''

# =============================================================================
# # Cámara y/o Sensor
# =============================================================================
Camara = Glosario_Camara.loc[Glosario_Camara.Editor == 'SINO',:]
Camara.loc[:,'Descripcion1'] = Camara.loc[:,'Descripcion'].str.upper()
Camara.reset_index(inplace=True)
Camara = Camara.drop(['index'],axis=1)

CAMARA = MLU
CAMARA.loc[:,'Tipo de cámara'] = CAMARA.loc[:,'Tipo de cámara'].fillna('')
CAMARA = CAMARA.loc[(CAMARA.loc[:,'Tipo de cámara'].str.len() != 0) & \
                  (~CAMARA.loc[:,'Tipo de cámara'].isna()),:]
# CAMARA = CAMARA.loc[CAMARA.loc[:,'Foto Cámara'] == 'Foto',:]
CAMARA.loc[:,'CODIGO'] = '34' + CAMARA.loc[:,'Equipo Ruta Id'].astype(str)
CAMARA.loc[:,'INSTALACION_ORIGEN_V10'] = CAMARA.loc[:,'BDI']

# CAMARA = CAMARA.drop(['level_0'],axis=1)
CAMARA.reset_index(inplace=True)
CAMARA = CAMARA.drop(['index'],axis=1)

if CAMARA.shape[0] > 0:
    CAMARA.loc[:,'TIP_CARGA'] = 'Tpp05'
    CAMARA.loc[:,'Tipo de carga'] = 'CAMARAS Y/O SENSORES'
    
    CAMARA.loc[:,'ORIGEN'] = '17'
    CAMARA.loc[:,'Origen de los datos'] = 'CENSO II'
    
    CAMARA.loc[:,'¿Marquilla Cámara?'] = CAMARA.loc[:,'¿Marquilla Cámara?'].str.upper()
    Tesc = pd.merge(left=CAMARA, right=Camara,how='left', \
            left_on='¿Marquilla Cámara?',right_on='Descripcion1',indicator=True)
    # Tesc.rename(columns={'Codigo': Codigo_y'},inplace=True) 
    CAMARA.loc[:,'Marquilla'] = Tesc.Codigo
    CAMARA.loc[:,'Marquilla_'] = Tesc.loc[:,'Descripcion1']
    
    CAMARA.loc[:,'TIP_CAMARA'] = 'TC01'
    CAMARA.loc[:,'Tipo de Camara'] = 'SEGURIDAD'
    
    Camara = Glosario_Camara.loc[Glosario_Camara.Editor == 'TMED',:]
    Camara.loc[:,'Descripcion1'] = Camara.loc[:,'Descripcion'].str.upper()
    Camara.reset_index(inplace=True)
    Camara = Camara.drop(['index'],axis=1)
    
    # CAMARA.loc[:,'Tipo de medida Tv cable'] = CAMARA.loc[:,'Tipo de medida Tv cable'].str.upper()
    # Tesc = pd.merge(left=CAMARA, right=Camara,how='left', \
    #         left_on='Tipo de medida Tv cable',right_on='Descripcion1',\
    #             suffixes=(None, '_y'),indicator=True)
    
    # CAMARA.loc[:,'TIP_MEDIDA'] = Tesc.Codigo_y
    # CAMARA.loc[:,'Tipo de medida'] = Tesc.loc[:,'Descripcion1']
    CAMARA.loc[:,'TIP_MEDIDA'] = 'TM01'
    CAMARA.loc[:,'Tipo de medida'] = 'DIRECTO'
    
    CAMARA.loc[:,'MODELO'] = ''
    CAMARA.loc[:,'Numero de serie (Num Medidor)'] = ''
    CAMARA.loc[:,'OBSERVACION'] = CAMARA.loc[:,'Observación cámara']
    CAMARA = CAMARA.fillna('')
    CAMARA = CAMARA.loc[:,['CODIGO','INSTALACION_ORIGEN_V10','TIP_CARGA','Tipo de carga','Marquilla',\
    'Marquilla_','TIP_CAMARA','Tipo de Camara','TIP_MEDIDA','Tipo de medida',\
    'MODELO','Numero de serie (Num Medidor)','OBSERVACION','ORIGEN','Origen de los datos',\
    'Longitud','Latitud','Nombre Ruta','Foto Cámara']]    
else:
    CAMARA = MLU.loc[:1,:]
    CAMARA.loc[:,'CODIGO'] = ''
    CAMARA.loc[:,'INSTALACION_ORIGEN_V10'] = ''
    CAMARA.loc[:,'TIP_CARGA'] = ''
    CAMARA.loc[:,'Tipo de carga'] = ''
    
    CAMARA.loc[:,'ORIGEN'] = ''
    CAMARA.loc[:,'Origen de los datos'] = ''
    CAMARA.loc[:,'Marquilla'] = ''
    CAMARA.loc[:,'Marquilla_'] = ''
    
    CAMARA.loc[:,'TIP_CAMARA'] = ''
    CAMARA.loc[:,'Tipo de Camara'] = ''
    
    CAMARA.loc[:,'TIP_MEDIDA'] = ''
    CAMARA.loc[:,'Tipo de medida'] = ''
    
    CAMARA.loc[:,'MODELO'] = ''
    CAMARA.loc[:,'Numero de serie (Num Medidor)'] = ''
    CAMARA.loc[:,'OBSERVACION'] = '' 
    CAMARA.loc[:,'Foto Cámara'] = ''
    CAMARA = CAMARA.loc[:,['CODIGO','INSTALACION_ORIGEN_V10','TIP_CARGA','Tipo de carga','Marquilla',\
    'Marquilla','TIP_CAMARA','Tipo de Camara','TIP_MEDIDA','Tipo de medida',\
    'MODELO','Numero de serie (Num Medidor)','OBSERVACION','ORIGEN','Origen de los datos',\
    'Longitud','Latitud','Nombre Ruta','Foto Cámara']]
    CAMARA.Longitud = ''
    CAMARA.Latitud = ''
    CAMARA.loc[:,'Nombre Ruta'] = ''
    
# =============================================================================
# # Antena
# =============================================================================
# Antena = Glosario_Antena.loc[Glosario_Antena.Editor == 'SINO',:]
# Antena.loc[:,'Descripcion1'] = Antena.loc[:,'Descripcion'].str.upper()
# Antena.reset_index(inplace=True)
# Antena = Antena.drop(['index'],axis=1)

# ANTENA = MLU
# ANTENA = ANTENA.loc[ANTENA.loc[:,'Foto Cámara'] == 'Foto',:]
# ANTENA.loc[:,'CODIGO'] = '32' + ANTENA.loc[:,'Equipo Ruta Id'].astype(str)
# ANTENA.loc[:,'Instalacion superior (Trafo)'] = ANTENA.loc[:,'BDI']

# # ANTENA = ANTENA.drop(['level_0'],axis=1)	
# ANTENA.reset_index(inplace=True)
# ANTENA = ANTENA.drop(['index'],axis=1)

# ANTENA.loc[:,'TIP_CARGA'] = 'Tpp01'
# ANTENA.loc[:,'Tipo de carga'] = 'ANTENAS DE TELECOMUNICACIONES' 
 

# ANTENA.loc[:,'ORIGEN'] = '17'
# ANTENA.loc[:,'Origen de los datos'] = 'CENSO II'

# ANTENA.loc[:,'¿Marquilla Cámara?'] = ANTENA.loc[:,'¿Marquilla Cámara?'].str.upper()
# Tesc = pd.merge(left=ANTENA, right=Antena,how='left', \
#         left_on='¿Marquilla Cámara?',right_on='Descripcion1',indicator=True)

# ANTENA.loc[:,'MARQUILLA'] = Tesc.Codigo
# ANTENA.loc[:,'Marquilla'] = Tesc.loc[:,'Descripcion1']

# ANTENA.loc[:,'TIP_ANT'] = 'TC01'
# ANTENA.loc[:,'Tipo de apoyo Antena'] = 'SEGURIDAD'

# ANT = Glosario_Antena.loc[Glosario_Antena.Editor == 'TMED',:]
# ANT.loc[:,'Descripcion1'] = ANT.loc[:,'Descripcion'].str.upper()
# ANT.reset_index(inplace=True)
# ANT = ANT.drop(['index'],axis=1)

# # CAMARA.loc[:,'Tipo de medida Tv cable'] = CAMARA.loc[:,'Tipo de medida Tv cable'].str.upper()
# # Tesc = pd.merge(left=CAMARA, right=Camara,how='left', \
# #         left_on='Tipo de medida Tv cable',right_on='Descripcion1',\
# #             suffixes=(None, '_y'),indicator=True)

# # CAMARA.loc[:,'TIP_MEDIDA'] = Tesc.Codigo_y
# # CAMARA.loc[:,'Tipo de medida'] = Tesc.loc[:,'Descripcion1']
# ANTENA.loc[:,'TIP_MEDIDA'] = 'TM01'
# ANTENA.loc[:,'Tipo de medida'] = 'DIRECTO'

# ANT = Glosario_Antena.loc[Glosario_Antena.Editor == 'OPTE',:]
# ANT.loc[:,'Descripcion1'] = ANT.loc[:,'Descripcion'].str.upper()
# ANT.reset_index(inplace=True)
# ANT = ANT.drop(['index'],axis=1)

# ANTENA.loc[:,'Operador Tv cable'] = ANTENA.loc[:,'Operador Tv cable'].fillna('')
# ANTENA.loc[:,'Operador Tv cable'] = ANTENA.loc[:,'Operador Tv cable'].apply(str.replace,args=('Otro', 'Otros'))
# ANTENA.loc[:,'Operador Tv cable'] = ANTENA.loc[:,'Operador Tv cable'].str.upper()
# Tesc = pd.merge(left=ANTENA, right=ANT,how='left', \
#         left_on='Operador Tv cable',right_on='Descripcion1',indicator=True)

# ANTENA.loc[:,'OPERADOR'] = Tesc.Codigo
# ANTENA.loc[:,'Operador de Telecomunicaciones'] = Tesc.loc[:,'Descripcion1']

# ANTENA.loc[:,'Marca'] = ''
# ANTENA.loc[:,'Numero de serie (Num Medidor)'] = ''
# ANTENA.loc[:,'OBSERVACIONES'] = ANTENA.loc[:,'Observación cámara']
# ANTENA.loc[:,'Longitud equipo padre'] = ANTENA.loc[:,'Longitud Equipo Padre']
# ANTENA.loc[:,'Latitud equipo padre'] = ANTENA.loc[:,'Latitud Equipo Padre']
# ANTENA = ANTENA.fillna('')

# ANTENA = ANTENA.loc[:,['CODIGO','Instalacion superior (Trafo)','TIP_CARGA','Tipo de carga',\
# 'MARQUILLA','Marquilla','TIP_ANT','Tipo de apoyo Antena','OPERADOR',\
# 'Operador de Telecomunicaciones','TIP_MEDIDA','Tipo de medida','Marca',\
# 'Numero de serie (Num Medidor)','OBSERVACIONES','ORIGEN',\
# 'Origen de los datos','Longitud','Latitud','Longitud equipo padre',\
# 'Latitud equipo padre','Nombre Ruta']]
# ANTENA.rename(columns = {'CODIGO':'Codigo'},inplace=True)

# Generar hojas en blanco Acometida, Semáforo y 
ANTENA = MLU
ANTENA.loc[:,'CODIGO'] = ''
ANTENA.loc[:,'Instalacion superior (Trafo)'] = ''
ANTENA.loc[:,'TIP_CARGA'] = ''
ANTENA.loc[:,'Tipo de carga'] = ''
ANTENA.loc[:,'MARQUILLA'] = ''
ANTENA.loc[:,'Marquilla'] = ''
ANTENA.loc[:,'TIP_ANT'] = ''
ANTENA.loc[:,'Tipo de apoyo Antena'] = ''
ANTENA.loc[:,'OPERADOR'] = ''
ANTENA.loc[:,'Operador de Telecomunicaciones'] = ''
ANTENA.loc[:,'TIP_MEDIDA'] = ''
ANTENA.loc[:,'Tipo de medida'] = ''
ANTENA.loc[:,'Marca'] = ''
ANTENA.loc[:,'Numero de serie (Num Medidor)'] = ''
ANTENA.loc[:,'OBSERVACIONES'] = ''
ANTENA.loc[:,'ORIGEN'] = ''
ANTENA.loc[:,'Origen de los datos'] = ''
ANTENA.loc[:,'Longitud'] = ''
ANTENA.loc[:,'Latitud'] = ''
ANTENA.loc[:,'Longitud equipo padre'] = ''
ANTENA.loc[:,'Latitud equipo padre'] = ''
ANTENA.loc[:,'Nombre Ruta'] = ''

ANTENA = ANTENA.loc[:,['CODIGO','Instalacion superior (Trafo)','TIP_CARGA','Tipo de carga',\
'MARQUILLA','Marquilla','TIP_ANT','Tipo de apoyo Antena','OPERADOR',\
'Operador de Telecomunicaciones','TIP_MEDIDA','Tipo de medida','Marca',\
'Numero de serie (Num Medidor)','OBSERVACIONES','ORIGEN',\
'Origen de los datos','Longitud','Latitud','Longitud equipo padre',\
'Latitud equipo padre','Nombre Ruta']]
    
Semaforo = MLU
Semaforo.loc[:,'Codigo'] = ''
Semaforo.loc[:,'Instalacion superior (Trafo)'] = ''
Semaforo.loc[:,'TIP_CARGA'] = ''
Semaforo.loc[:,'Tipo de carga'] = ''
Semaforo.loc[:,'MARQUILLA'] = ''
Semaforo.loc[:,'Marquilla'] = ''
Semaforo.loc[:,'TIP_MEDIDA'] = ''
Semaforo.loc[:,'Tipo de medida'] = ''
Semaforo.loc[:,'Marca'] = ''
Semaforo.loc[:,'Numero de serie (Num Medidor)'] = ''
Semaforo.loc[:,'OBSERVACIONES'] = ''
Semaforo.loc[:,'ORIGEN'] = ''
Semaforo.loc[:,'Origen de los datos'] = ''
Semaforo.loc[:,'Longitud'] = ''
Semaforo.loc[:,'Latitud'] = ''
Semaforo.loc[:,'Longitud equipo padre'] = ''
Semaforo.loc[:,'Latitud equipo padre'] = ''
Semaforo.loc[:,'Nombre Ruta'] = ''

Semaforo = Semaforo.loc[:,['Codigo','Instalacion superior (Trafo)','TIP_CARGA',\
        'Tipo de carga','MARQUILLA','Marquilla','TIP_MEDIDA','Tipo de medida',\
            'Marca','Numero de serie (Num Medidor)','OBSERVACIONES','ORIGEN',\
    'Origen de los datos','Longitud','Latitud','Longitud equipo padre',\
        'Latitud equipo padre','Nombre Ruta']]

Acometida = MLU
Acometida.loc[:,'Codigo'] = ''
Acometida.loc[:,'Instalacion Superior'] = ''
Acometida.loc[:,'campo1'] = ''
Acometida.loc[:,'Acometida Normalizada'] = ''
Acometida.loc[:,'Numero Identificacion Finca'] = ''
Acometida.loc[:,'Numero de Identificacion del Suministro'] = ''
Acometida.loc[:,'Numero de Identificacion del Contrato'] = ''
Acometida.loc[:,'TIP_ACOME'] = ''
Acometida.loc[:,'Tipo de Acometida'] = ''
Acometida.loc[:,'Suministro'] = ''
Acometida.loc[:,'Observaciones'] = ''
Acometida.loc[:,'ORIGEN'] = ''
Acometida.loc[:,'Origen de los datos'] = ''
Acometida.loc[:,'TERRITORIO'] = ''
Acometida.loc[:,'CIRCUITO'] = ''
Acometida.loc[:,'FECHA'] = ''
Acometida.loc[:,'Longitud'] = ''
Acometida.loc[:,'Latitud'] = ''
Acometida.loc[:,'Longitud equipo padre'] = ''
Acometida.loc[:,'Latitud equipo padre'] = ''


Acometida = Acometida.loc[:,['Codigo','Instalacion Superior','campo1','Acometida Normalizada',\
'Numero Identificacion Finca','Numero de Identificacion del Suministro',\
'Numero de Identificacion del Contrato','TIP_ACOME','Tipo de Acometida',\
'Suministro','Observaciones','ORIGEN','Origen de los datos','TERRITORIO','CIRCUITO',\
'FECHA','Longitud','Latitud','Longitud equipo padre','Latitud equipo padre']]
    
    
VallaPublicitaria = MLU
VallaPublicitaria.loc[:,'Codigo'] = ''
VallaPublicitaria.loc[:,'Instalacion superior (Trafo)'] = ''
VallaPublicitaria.loc[:,'TIP_CARGA'] = ''
VallaPublicitaria.loc[:,'Tipo de carga'] = ''
VallaPublicitaria.loc[:,'MARQUILLA'] = ''
VallaPublicitaria.loc[:,'Marquilla'] = ''
VallaPublicitaria.loc[:,'TIP_PUBLI'] = ''
VallaPublicitaria.loc[:,'Tipo Publicidad'] = ''
VallaPublicitaria.loc[:,'Numero de Caras'] = ''
VallaPublicitaria.loc[:,'TIP_MEDIDA'] = ''
VallaPublicitaria.loc[:,'Tipo de medida'] = ''
VallaPublicitaria.loc[:,'Marca'] = ''
VallaPublicitaria.loc[:,'Numero de serie (Num Medidor)'] = ''
VallaPublicitaria.loc[:,'OBSERVACIONES'] = ''
VallaPublicitaria.loc[:,'ORIGEN'] = ''
VallaPublicitaria.loc[:,'Origen de los datos'] = ''
VallaPublicitaria.loc[:,'Longitud'] = ''
VallaPublicitaria.loc[:,'Latitud'] = ''
VallaPublicitaria.loc[:,'Longitud equipo padre'] = ''
VallaPublicitaria.loc[:,'Latitud equipo padre'] = ''
VallaPublicitaria.loc[:,'Nombre Ruta'] = ''

VallaPublicitaria = VallaPublicitaria.loc[:,['Codigo','Instalacion superior (Trafo)','TIP_CARGA',\
    'Tipo de carga','MARQUILLA','Marquilla','TIP_PUBLI','Tipo Publicidad',\
    'Numero de Caras','TIP_MEDIDA','Tipo de medida','Marca',\
   'Numero de serie (Num Medidor)','OBSERVACIONES','ORIGEN','Origen de los datos',\
    'Longitud','Latitud','Longitud equipo padre','Latitud equipo padre','Nombre Ruta']]    

# =============================================================================
# # Ficticio
# =============================================================================
Fict = Glosario_Ficticio.loc[Glosario_Ficticio.Editor == 'CLAS',:]
Fict.loc[:,'Descripcion1'] = Fict.loc[:,'Descripcion'].str.lower()
Fict.reset_index(inplace=True)
Fict = Fict.drop(['index'],axis=1)
Fict.loc[:,'Descripcion1'] = Fict.loc[:,'Descripcion1'].astype(str).apply(str.replace,args=(' ', ''))
Fict.loc[:,'Descripcion1'] = [normalize(str(Fict.loc[i,'Descripcion1'])) \
                            for i in range(0, len(Fict))]
   
Ficticio = MLU_Tramo
Ficticio = Ficticio.loc[Ficticio.loc[:,'Tipo de Apoyo'] == 'Ficticio',:]
Ficticio.reset_index(inplace=True)

Ficticio.loc[:,'Clasificacion Apoyo Ficticio'] = Ficticio.loc[:,'Clasificacion Apoyo Ficticio'].astype(str).apply(str.replace,args=(' ', ''))
Ficticio.loc[:,'Clasificacion Apoyo Ficticio'] = Ficticio.loc[:,'Clasificacion Apoyo Ficticio'].str.lower()
Ficticio.loc[:,'Clasificacion Apoyo Ficticio'] = [normalize(str(Ficticio.loc[i,'Clasificacion Apoyo Ficticio'])) \
                            for i in range(0, len(Ficticio))]
    
Tesc = pd.merge(left=Ficticio, right=Fict,how='left', \
        left_on='Clasificacion Apoyo Ficticio',right_on='Descripcion1',indicator=True)
Tesc.loc[Tesc.loc[:,'Codigo'].isna(),'Codigo'] = ''    
A = Tesc.loc[:,'Codigo'].str.split("|", n=2, expand=True)  
# Tramo.loc[:,'Clasificacion Apoyo Ficticio']    
Ficticio.loc[:,'CODIGO'] = A.iloc[:,1].astype(str).replace('None','') + \
    Ficticio.loc[:,'Equipo Ruta Id'].astype(str)
Ficticio.loc[:,'INSTALACION_ORIGEN_V10'] = Ficticio.loc[:,'BDI']
Ficticio.loc[:,'CLASIFICACIÓN'] = A.iloc[:,0]
Ficticio.loc[:,'CLASIFICACIÓN_'] = Tesc.loc[:,'Descripcion']
# Tesc.rename(columns={'Codigo': Codigo_y'},inplace=True) 
# Ficticio.loc[:,'CLASIFICACIÓN'] = Ficticio.loc[:,'CLASIFICACIÓN_']
Ficticio.loc[:,'Longitud_final'] = Ficticio.loc[:,'Longitud']
Ficticio.loc[:,'Latitud_final'] = Ficticio.loc[:,'Latitud']
Ficticio.loc[:,'Observaciones'] = Ficticio.loc[:,'Descripcion del apoyo NO normalizado']
Ficticio.loc[:,'ORIGEN'] = '17'
Ficticio.loc[:,'Origen de los datos'] = 'CENSO II'
# Ficticio.loc[:,'Foto del fitcicio'] 
# CODIGO	INSTALACION_ORIGEN_V10	CLASIFICACIÓN	CLASIFICACIÓN	
# Observaciones	ORIGEN	Origen de los datos	Longitud_final	Latitud_final	
# Foto del fitcicio
# CAMARA = CAMARA.drop(['level_0'],axis=1)
Ficticio.reset_index(inplace=True)
Ficticio = Ficticio.drop(['index'],axis=1)
Ficticio = Ficticio.loc[:,['CODIGO','INSTALACION_ORIGEN_V10','CLASIFICACIÓN','CLASIFICACIÓN_',\
    'Observaciones','ORIGEN','Origen de los datos','Longitud_final','Latitud_final',\
    'Foto del fitcicio']]    
# =============================================================================
# #. Generar reporte MLU
# =============================================================================
root = tk.Tk()
root.withdraw()

# file_path = filedialog.askopenfilename()
file_path = filedialog.asksaveasfile(mode='w', defaultextension=".xlsx")
if file_path is None:
  a = 0
else:
    # Turn off the default header and skip one row to allow us to insert a
    # user defined header.
    # df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False)
    
    # # Get the xlsxwriter workbook and worksheet objects.
    # workbook  = writer.book
    # worksheet = writer.sheets['Sheet1']    
    # Add a header format. 
        
    # book = Workbook()
    # sheet1 = book.add_sheet('Acometida')
    # book.save(file_path.name)     
    t = time.time()
    with pd.ExcelWriter(file_path.name) as writer: 
        # shutil.copy("Plantillas/Reporte_bin.xlsb", file_path.name[:len(file_path.name)-4] + '.xlsb')      
        # MLU.to_excel(writer, sheet_name='MLU', na_rep='',float_format=None, columns=None, header=True,index=False)          
        Apoyo.to_excel(writer, sheet_name='APOYO', na_rep='',float_format=None, columns=None, header=True,index=False)
        # writer = pd.ExcelWriter(file_path.name,
        #                         engine='xlsxwriter')        
        workbook  = writer.book
        worksheet = writer.sheets['APOYO']
        for col_num, value in enumerate(Apoyo.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [4,5,7,9,11,12,14,16,18,19,20,22]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [0,1,2,3,23,24,25]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})                                      
            elif any(item == col_num for item in [6,8,10,13,15,17,21]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})            
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)         
        Tramo.to_excel(writer, sheet_name='TRAMO', na_rep='',float_format=None, columns=None, header=True,index=False)    
        workbook  = writer.book
        worksheet = writer.sheets['TRAMO']
        for col_num, value in enumerate(Tramo.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if col_num <= 14:
                # st.pattern.pattern_fore_colour = 22 # Gris np.linspace(7,8,2)
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [53,54]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})                
            elif any(item == col_num for item in [15,16,17,18,19,20,22,24,26,28,29,31,33,35,37,39,\
                    41,42,43,44,46,48,49,50,52,55,56]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})                                     
            elif any(item == col_num for item in [21,23,25,27,30,32,34,36,38,40,45,47,51]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)            
        CAbonados.to_excel(writer, sheet_name='CAJA DE ABONADO', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['CAJA DE ABONADO']
        for col_num, value in enumerate(CAbonados.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,5,7,8,9,11]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [12,13,14,15]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,4,6,10]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)
        PuenteBT.to_excel(writer, sheet_name='PUENTE BT', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['PUENTE BT']
        for col_num, value in enumerate(PuenteBT.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,5]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [6,7,8,9]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,4]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)  
        Transicion.to_excel(writer, sheet_name='TRANSICION', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['TRANSICION']
        for col_num, value in enumerate(Transicion.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,2,3,4,6]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [7,8,9,10]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [5]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)    
        PuestaTierra.to_excel(writer, sheet_name='PUESTA TIERRA', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['PUESTA TIERRA']
        for col_num, value in enumerate(PuestaTierra.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,2,3,5]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [6,7,8,9]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [4]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)
        TVCable.to_excel(writer, sheet_name='TV Cable', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['TV Cable']
        for col_num, value in enumerate(TVCable.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,5,7,9,11,12,13,14,16]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [17,18,19]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,4,6,8,10,15]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)  
        CAMARA.to_excel(writer, sheet_name='Cámara y o Sensor', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['Cámara y o Sensor']
        for col_num, value in enumerate(CAMARA.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,5,7,9,10,11,12,14]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [15,16,17]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,4,6,8,13]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)     
        ANTENA.to_excel(writer, sheet_name='Antena', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['Antena']
        for col_num, value in enumerate(ANTENA.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,5,7,9,11,12,13,14,16]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [17,18,19,20,21]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,4,6,8,15]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)
        Semaforo.to_excel(writer, sheet_name='Semaforo', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['Semaforo']
        for col_num, value in enumerate(Semaforo.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,5,7,8,9,10,12]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [13,14,15,16,17]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,4,6,11]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)  
        Acometida.to_excel(writer, sheet_name='Acometida', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['Acometida']
        for col_num, value in enumerate(Acometida.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,4,5,6,8,9,11,12,13,14]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [15,16,17,18]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,7,10]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)    
        VallaPublicitaria.to_excel(writer, sheet_name='Valla Publicitaria', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['Valla Publicitaria']
        for col_num, value in enumerate(VallaPublicitaria.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,5,7,8,9,10,11,12,13,15]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [16,17,18,19,20]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,4,6,9,14]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)   
        Ficticio.to_excel(writer, sheet_name='Apoyo NN & Cruce Fic', na_rep='',float_format=None, columns=None, header=True,index=False)        
        workbook  = writer.book
        worksheet = writer.sheets['Apoyo NN & Cruce Fic']
        for col_num, value in enumerate(Ficticio.columns.values):
            # worksheet.write(0, col_num + 1, value, header_format)        
        # for i in range(0, Acometida.shape[1]):
            # st = xlwt.easyxf('pattern: pattern solid;')
            if any(item == col_num for item in [0,1,3,4,6,9]):
                # st.pattern.pattern_fore_colour = 17 # Verde
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#51ab42',
                    'border': 1})              
            elif any(item == col_num for item in [7,8]):
                # st.pattern.pattern_fore_colour = 22 # Azul
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#0000ff',
                    'border': 1})            
            elif any(item == col_num for item in [2,5,10,11]):
                # st.pattern.pattern_fore_colour = 5 # Amarillo
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': False,
                    'valign': 'top',
                    'fg_color': '#ecff01',
                    'border': 1})              
            # sheet1.write(i % 24, i // 24,'',st)
            # sheet1.write(0,i,'', st)
            worksheet.write(0, col_num, value, header_format)               
file_path.close()
elapsed_time = time.time() - t
elapsed_time = '{0:.2f}'.format(elapsed_time)
print('Reporte generado en ' + elapsed_time)     
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

import warnings

warnings.filterwarnings('ignore')
warnings.warn('DelftStack')
warnings.warn('Do not show this message')

# cargamos informe como dataframe
file_path = filedialog.askopenfilename(defaultextension=".xls")
if file_path is None:
    # quit()
    sys.exit()

df = pd.read_excel(file_path, header=0)

# df = pd.read_excel('Informe Tranformadores MLU.xlsx')
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

df.reset_index(inplace=True) #resetar index
df = df.drop(columns = 'index') #elimina columna index

#Generacion de codigo id_bdi, codeleme o nuevo
Sel = df[df['CODELEME'] == 0].index  #filtramos codeleme en cero
df.loc[Sel,"CODELEME"] = str(10) + df.loc[:,"Equipo Ruta Id"].astype(str) #colocamos el 10 + el ruta id a los codeleme en cero
Sel = df[df['ID_BDI'] == 0].index #filtramos ID_BDI en cero
df.loc[Sel,"ID_BDI"] =  df.loc[:,"CODELEME"] #les asignamos el codeleme
df.rename(columns={'ID_BDI': 'CODIGO'}, inplace=True) #RENOMBRANDO COLUMNA
del df['CODELEME'] #ELIMINAR COLUMNA CODELEME 


#INSTALACION SUPERIOR
df.rename(columns={'CODIGO_BDI_CIRCUITO': 'INSTALACION_SUPERIOR'}, inplace=True) #RENOMBRANDO COLUMNA

#PLACA MT ANTERIOR
Sel = df[df['Matricula MT anterior'] != 0 ].index #filtramos Matricula MT anterior diferentes de cero
df.loc[Sel,"PLACA MT ANTERIOR"] = df.loc[:,"Matricula MT anterior"] # se lo asignamos a MATRICULA ANTIGUA
df.rename(columns={'PLACA MT ANTERIOR': 'MATRICULA_ANTIGUA'}, inplace=True) #RENOMBRANDO COLUMNA
del df["Matricula MT anterior"]  #ELIMINAR COLUMNA

#POTENCIA
Sel = df[df['POTENCIA NOMINAL MODIFICADA'] != 0 ].index #filtramos 
df.loc[Sel,"POTENCIA"] = df.loc[:,"POTENCIA NOMINAL MODIFICADA"] #asignamos 
df.rename(columns={'POTENCIA': 'Potencia_Nominal_(kVA)'}, inplace=True) #RENOMBRANDO COLUMNA
del df["POTENCIA NOMINAL MODIFICADA"] #ELIMINAR COLUMNA

#MARCA
Sel = df[df['OTRA MARCA'] != 0].index #filtramos 
df.loc[Sel,"MARCA MODIFICADA"] = df.loc[:,"OTRA MARCA"] #asignamos
Sel = df[df['MARCA MODIFICADA'] != 0].index #filtramos
df.loc[Sel,"FABRICANTE"] = df.loc[:,"MARCA MODIFICADA"] #asignamos
df.rename(columns={'FABRICANTE': 'Marca'}, inplace=True) #RENOMBRANDO COLUMNA
del df["OTRA MARCA"] #ELIMINAR COLUMNA
del df["MARCA MODIFICADA"] #ELIMINAR COLUMNA

#TENSION SECUNDARIA
Sel = df[df["TENSIÓN SECUNDARIA MODIFICADA (V)"] == 0].index #filtramos
df.loc[Sel,"TENSIÓN SECUNDARIA MODIFICADA (V)"] = df.loc[:,"RELTRANS"] #asignamos
df["RELTRANS"] = 0 #se coloca columna en cero
df.rename(columns={'TENSIÓN SECUNDARIA MODIFICADA (V)': 'Tension_Secundaria_(V)',
                    'RELTRANS': 'TENS_SEC'    }, inplace=True)  #RENOMBRANDO COLUMNA

#ubicacion-tipo de instalacion
Sel = df[df["CONFIRME TIPO DE INSTALACIÓN"] == 0 ].index #filtramos
df.loc[Sel,"CONFIRME TIPO DE INSTALACIÓN"] = df.loc[:,"TIPINS"] #asignamos
df.loc[:,"CONFIRME TIPO DE INSTALACIÓN"] = df.loc[:,"CONFIRME TIPO DE INSTALACIÓN"].astype(str).apply(\
                str.replace,args=('AÉREO', 'AEREO')) #remplazamos
df.loc[:,"CONFIRME TIPO DE INSTALACIÓN"] = df.loc[:,"CONFIRME TIPO DE INSTALACIÓN"].astype(str).apply(\
                str.replace,args=('SUBTERRÁNEO', 'SUBTERRANEO')) #remplazamos
df["TIPINS"] = 0 #se coloca columna en cero
df.rename(columns={'TIPINS':"UBICACION_TRAFO",
                    'CONFIRME TIPO DE INSTALACIÓN':'Ubicacion_Trafo' }, inplace=True)  #RENOMBRANDO COLUMNA

#TIPO DE TX
Sel = df[df['CONFIRME TIPO DE TRANSFORMADOR'] == 0 ].index #filtramos
df.loc[Sel,'CONFIRME TIPO DE TRANSFORMADOR'] = df.loc[:,"TIPOFASE"] #asignamos
df.loc[:,'CONFIRME TIPO DE TRANSFORMADOR'] = df.loc[:,"CONFIRME TIPO DE TRANSFORMADOR"].astype(str).apply(\
                str.replace,args=('Á', 'A')) #remplazamos
df['TIPOFASE'] = 0
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
Sel = df[df['Longitud del equipo padre'] == 0 ].index #filtramos
df.loc[Sel,'Longitud del equipo padre'] = df.loc[:,"Longitud"] #asignamos
Sel = df[df['Latitud del equipo padre'] == 0 ].index #filtramos
df.loc[Sel,'Latitud del equipo padre'] = df.loc[:,"Latitud"] #asignamos
df.rename(columns={'Longitud del equipo padre':"Longitud (WGS84)",
                    'Latitud del equipo padre':'Latitud (WGS84)' }, inplace=True)  #RENOMBRANDO COLUMNA
del df["Latitud"] #ELIMINAR COLUMNA
del df["Longitud"] #ELIMINAR COLUMNA

df.insert(37,"ESTADO COORDENADAS",0)

# =============================================================================
#  COMPARAMOS LOS CODIGOS CON LA BD MT
# =============================================================================

#cargamos BD MT como dataframe
df_mt = pd.read_excel('BD MT - CODIGO.xlsx')

#INSERTAMOS COLUMNAS PARA LA VERIFICACION Y CRUCE DE CODIGO
df.insert(4,"CODIGO MT",0)
df.insert(5,"VERIFICACION CODIGO",0)

#se cruzan los df por matricula antigua 
df = df.merge(df_mt, how='left', on='MATRICULA_ANTIGUA')

df['CODIGO MT'] = df['CODIGO_TRANSFORMADOR']
del df['CODIGO_TRANSFORMADOR']

#imprimimos el df en un excel

df.to_excel('COORDENADAS_xx_SEPTIEMBRE2022_APG.xlsx', sheet_name='CC')

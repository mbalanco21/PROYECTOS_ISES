# LIBRERIAS
from ftplib import FTP
import os
import pandas as pd

#LISTAS USADAS EN TODO EL PROCESO

#lista para descargar 
# lista_descarga = ['761_20589350','761_20589349','761_20589337']
lista_ruta_servidor = []
lista_descarga = []
lista_id_tx = []
lista_id_ap = []
lista_ruta_temp = []
lista_ruta_doc = []
lista_Error_ap =[]
lista_ok_ap = []
lista_Error_tx =[]
lista_ok_tx = []

# #RECORDATORIOS DOC EXCEL:
# =============================================================================
# TILDAR LA A DE ATLANTICO DEBE QUEDAR ATLÁNTICO
# CAMBIAR ESPACIOS POR _
# CAMBIAR LOS NOMBRES DE LOS TERRITORIOS 01_ 02_ ETC
# =============================================================================
# =============================================================================
#  IndexError: list index out of range = 
#  SIGIFICA QUE NO HAY CARPETA MATRICULACION DENTRO DEL CIRCUITO
#
#  error_perm: 550 CWD failed. "/MLU AIR-E/03_MAGDALENA/PLATO/PLATO__I": directory not found. =
#  ALGO DEBER ESTAR MAL ESCRITO EN EL EXCEL
# =============================================================================
# "/ImagenesFormsMap/ImagenesCampo/MLU AIR-E/
df = pd.read_excel('descarga.xlsx') # SE CARGA EL ARCHIVO EXCEL QUE ESTA DENTRO DE LA CARPETA

#SE CREA RUTA INCIAL A PARTIR DE ARCHIVO EXCEL Y SE EMPAQUETA EN LSITAS
for i in df.index:
    ruta_servidor ="/ImagenesFormsMap/ImagenesCampo/MLU AIR-E/"+df["TERRITORIO"][i]+"/"+df["SUB_ESTACION"][i]+"/"+df["CIRCUITO"][i]
    lista_ruta_servidor.append(ruta_servidor)
    
    #SE EXTRAEN LOS RUTA ID TX Y SE LISTAN
    n_id_tx =str(df["RUTA ID"][i])
    lista_descarga.append(n_id_tx)
    lista_id_tx.append(n_id_tx)
    
    #SE EXTRAEN LOS RUTA ID AP Y SE LISTAN
    n_id_ap = str(df["RUTA ID APOYO"][i])
    lista_descarga.append(n_id_ap)
    lista_id_ap.append(n_id_ap)
    
    #SE EXTRAEN LOS NOMBRE RUTA DEL EXCEL Y SE LISTAN
    lista_ruta_doc.append(df["RUTA"][i]) 
    
# print(lista_ruta_servidor)
print(len(lista_ruta_servidor))

# print(lista_descarga)
# print(len(lista_descarga))

#INICIAMOS SESION EN EL SERVIDOR 
ftp = FTP()
ftp.set_pasv(False)                                     #modo activo
ftp.connect('formap.co', 21, timeout= 60 )              # servidor, puerto y tiempo de espera
ftp.login('lcabrera', '123456')                        #credenciales
ftp.encoding = "UTF-8"
print("conexion correcta") 
# print(ftp.getwelcome())                             #mensaje de bienvenida
print(" ")

#SE AÑADE LA CARPETA MATRICULACIO_XXXX EN LA RUTA SERVIDOR
for ruta in lista_ruta_servidor[0:len(lista_ruta_servidor)]:
    try:
        ftp.cwd(ruta)
        # print(ftp.cwd(ruta))
        carpetas = ftp.nlst()
        # print(carpetas)
        if 'LEVANTAMIENTO' in carpetas:
            # print ('existen archivos temporales')
            carpetas.remove('LEVANTAMIENTO')            #ELIMINO CARPETA LEVANTAMIENTO DE LAS RUTAS POSIBLES
            # print(carpetas)
           
        else:
            print(carpetas)
        
    
        # print(ruta)
        ruta_temp = ruta+"/"+carpetas[0]
        # print(ruta_temp)
        # print(" ")
        lista_ruta_temp.append(ruta_temp)
    except Exception as e:                  #SE CAPTRAN LOS ERRORES COMO: NO ENCONTRADO, MAL ESCRITO
        ruta_temp = ruta+"/error"
        lista_ruta_temp.append(ruta_temp)
        #print(ruta)
        #lista_ruta_servidor.remove(ruta)
#print(lista_ruta_temp)

#SE AÑADE RUTA EN LA RUTA SERVIDOR Y SE LISTA
n=0
lista_ruta_temp1= []
for ruta in lista_ruta_temp[0:len(lista_ruta_temp)]:
    ruta_temp1 = ruta+"/"+lista_ruta_doc[n]
    n = n+1
    lista_ruta_temp1.append(ruta_temp1)
    
#SE CREAN LAS POSIBLES RUTAS FINALES
n=0
lista_ruta_final_tx = []
lista_ruta_final_ap = []
#RUTA FINAL TX
for ruta in lista_ruta_temp1[0:len(lista_ruta_temp)]:
    ruta_final = ruta+"/"+"749_"+lista_id_tx[n]
    n = n+1
    lista_ruta_final_tx.append(ruta_final)
#RUTA FINAL AP
n=0
for ruta in lista_ruta_temp1[0:len(lista_ruta_temp)]:
    ruta_final = ruta+"/"+"761_"+lista_id_ap[n]
    n = n+1
    lista_ruta_final_ap.append(ruta_final)
# print(lista_ruta_final_tx)
# print(" ")
# print(lista_ruta_final_ap)
print(len(lista_ruta_final_ap))

#SE VALIDAN LAS RUTASDINALES AP Y SE LISTAN LAS OK Y ERRORES
n=0
for ruta in lista_ruta_final_ap[0:len(lista_ruta_final_ap)]:
    id_ap = lista_id_ap[n]
    try:
        
        ftp.cwd(ruta)
        n = n+1
        lista_ok_ap.append(id_ap)
    except Exception as e:
        lista_Error_ap.append(id_ap)
        # print(ruta)
        lista_ruta_final_ap.remove(ruta)
        n = n+1

print("SE ENCONTRARON " + str(len(lista_ruta_final_ap)) + " AP")
# print(len(lista_ruta_final_ap))
print(" ")

#SE VALIDAN LAS RUTASDINALES TX Y SE LISTAN LAS OK Y ERRORES
print(len(lista_ruta_final_tx))
n=0
for ruta in lista_ruta_final_tx[0:len(lista_ruta_final_tx)]:
    id_tx = lista_id_tx[n]
    try:
        ftp.cwd(ruta)
        n = n+1
        lista_ok_tx.append(id_tx)
    except Exception as e:
        lista_Error_tx.append(id_tx)
        # print(ruta)
        lista_ruta_final_tx.remove(ruta)
        n = n+1
    
print("SE ENCONTRARON " + str(len(lista_ruta_final_tx)) + " TX")
# print(len(lista_ruta_final_tx))

#AUTORIZACION PARA DESCARGAR
confirma_descarga = "n"
print(" ")
confirma_descarga = input("¿desea comenzar la descargar?  si / no  : ")
print(" ")

#CODIGO DE DESCARGA
if confirma_descarga == "si":    
    n=0
    for nombre_carpeta in lista_ok_tx[0:len(lista_ok_tx)]:  #SE CREAN LAS CARPETAS LOCALES
        os.mkdir("749_"+nombre_carpeta)
    for nombre_carpeta in lista_ok_tx[0:len(lista_ok_tx)]:  #SE CARGAN LAS LAS CARPETAS LOCALES
        os.chdir('C:/Users/P568/Desktop/AUTO_FTP/FTP_DESCARGA/'+"749_"+nombre_carpeta)
        print("749_"+nombre_carpeta+" "+str(n))
        ftp.cwd(lista_ruta_final_tx[n])         #SE BUSCA LA RUTA EN EL SERVIDOR
        n = n+1
        # ftp.dir() 
        archivos = ftp.nlst()
        # print(archivos)
        if 'Temp' in archivos:              #SE ELIMINA LA CARPETA TEMP DE LO QUE SE DESCARGARA
            # print ('existen archivos temporales')
            archivos.remove('Temp')
            # print(archivos)
            
        else:
            print ('NO existen archivos temporales')
            #print(archivos)
            
        for archivo in archivos[0:len(archivos)]:   #SE PROCEDE A DESCARGAR
            abrir = open(archivo, 'wb')
            ftp.retrbinary("RETR "+ archivo, abrir.write)
            print("descargando")
            
    print("------- termenino descarga tx-------")
    
    os.chdir('C:/Users/P568/Desktop/AUTO_FTP/FTP_DESCARGA')
    
    n=0
    for nombre_carpeta in lista_ok_ap[0:len(lista_ok_ap)]: #SE CREAN LAS CARPETAS LOCALES
        try:
            os.mkdir("761_"+nombre_carpeta)
        except FileExistsError:
            print("apoyo duplicado / carpeta ya creada: 761_"+str(nombre_carpeta))
        
    for nombre_carpeta in lista_ok_ap[0:len(lista_ok_ap)]: #SE CARGAN LAS LAS CARPETAS LOCALES
        os.chdir('C:/Users/P568/Desktop/AUTO_FTP/FTP_DESCARGA/'+"761_"+nombre_carpeta)
        print("761_"+nombre_carpeta+" "+str(n))
        ftp.cwd(lista_ruta_final_ap[n])              #SE BUSCA LA RUTA EN EL SERVIDOR
        n = n+1
        # ftp.dir() 
        archivos = ftp.nlst()
        # print(archivos)
        if 'Temp' in archivos:              #SE ELIMINA LA CARPETA TEMP DE LO QUE SE DESCARGARA
            # print ('existen archivos temporales')
            archivos.remove('Temp')
            # print(archivos)
            
        else:
            print ('NO existen archivos temporales')
            #print(archivos)
            
        for archivo in archivos[0:len(archivos)]:   #SE PROCEDE A DESCARGAR
            abrir = open(archivo, 'wb')
            ftp.retrbinary("RETR "+ archivo, abrir.write)
            print("descargando")
            
    print("------- termenino descarga ap-------")

else:
    print("no se descargo nada")    #EN CASO DE NO COLOCAR si 

ftp.close()
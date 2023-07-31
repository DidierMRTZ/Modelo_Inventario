from Librerias_SAP import SAP_GUI, Funtions
import win32com.client
import pandas as pd
from time import sleep
import json

"""----------------------------Inciar session----------------------------------------------------"""
# Insert User name and password
Keys=pd.read_excel("C:\\Users\\practicante.picking\\OneDrive - Prebel S.A\\Escritorio\\bot_picking\\Archivos_CSV\\Keys.xlsx")
user=Keys["User"][0]
password=Keys["Password"][0]
# Initialize session
session=SAP_GUI.SessionSAP(user,password)

"""----------------- Search ZD110 -------------------------------------------------------"""
Transsaccion_ZSD110="zsd110"
Varian_FIRME="PEN-FIRME"
Varian_MERCADEO="PEN-MERCADEO"
Varian_DISNAL_TOTAL="PENDIENTES"
provision_ZSD110="/REVISIONPEN"

"""-----------------------------------BUSCAR y DESCARGAR TRANSSACCION ZD110---------------------------------------------------------------- """
# Pendiente Firme
Name_ZSD110_FIRME="Pendiente_Firme"  # Revisar suele no guardar
# Ruta_Name_ZSD110="C:\\Users\\practicante.picking\\OneDrive - Prebel S.A\\Escritorio\\Bot_Disnal_Pendientes_2\\Archivos_CSV\\"
SAP_GUI.Search_ZSD110(Transsaccion_ZSD110,Varian_FIRME,provision_ZSD110,session)
SAP_GUI.Export_TXT2(Name_ZSD110_FIRME,session)

# Pendiente Mercadeo
Name_ZSD110_MERCADEO="Pendiente_Mercadeo"  # Revisar suele no guardar
SAP_GUI.Search_ZSD110(Transsaccion_ZSD110,Varian_MERCADEO,provision_ZSD110,session)
SAP_GUI.Export_TXT2(Name_ZSD110_MERCADEO,session)

# Pendiente Total 
Name_ZSD110_TOTAL="Pendiente_Total"  # Revisar suele no guardar
tabla_Total=SAP_GUI.Search_ZSD110(Transsaccion_ZSD110,Varian_DISNAL_TOTAL,provision_ZSD110,session)
SAP_GUI.Export_TXT2(Name_ZSD110_TOTAL,session)

# Descargar archivos Csv de chanales Pendiente y Entrega
Download_channel=["01"]
List_Channels=SAP_GUI.Download_ZSD110_Channels(tabla_Total,Download_channel,session)

# Descargar json de los chanales
ruta_archivo_json="C:\\Users\\practicante.picking\\OneDrive - Prebel S.A\\Escritorio\\bot_picking\\Archivos_CSV\\Channels.json"
Funtions.list_to_json(List_Channels,ruta_archivo_json)
from dash import Dash, dash_table, dcc, html
from dash.dependencies import Input, Output
import pandas as pd
from flask import Flask
from Librerias_SAP import SAP_GUI, Funtions
import pandas as pd
import re
from datetime import datetime,timedelta
import numpy as np
import win32com.client

"""----------------------------Inciar session----------------------------------------------------"""
# Insert User name and password

Keys=pd.read_excel("C:\\Users\\practicante.picking\\OneDrive - Prebel S.A\\Escritorio\\bot_picking\\Archivos_CSV\\Keys.xlsx")
user=Keys["User"][0]
password=Keys["Password"][0]
# user= "jespinosap"
# password= "Ela2023*"
# Initialize session
session=SAP_GUI.SessionSAP(user,password)

Defaul_Column_Pedidos_dia=['Documento', 'GTr', 'Denomin.', 'ClVt', 'Denominación', 'Solic.',
       'Creado el', 'Fecha doc.', 'Pedido', 'Func.', 'Responsab', 'Creado',
       'OrgVt', 'CDis', 'Se', 'OfVta', 'GVen', 'Mon.', 'Valor neto']

# PEDIDOSEXITO
Transsaccion='va05n'
provision='PEDIDOSEXTCEN'
variant='JESPINOSAP'

SAP_GUI.Search_VA05N(Transsaccion,session,provision,variant)

Name_VA05N="Pedidos_diarios"
# Ruta_VA05N="C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Archivos_CSV\\"
# Ruta_VA05N="C:\\Users\\practicante.picking\\OneDrive - Prebel S.A\\Escritorio\\Archivos_CSV\\"
Ruta_VA05N="C:\\Users\\practicante.picking\\OneDrive - Prebel S.A\\Escritorio\\bot_picking\\Archivos_CSV\\"
SAP_GUI.Export_TXT2(Name_VA05N,session,Ruta_VA05N)

Pedidos_VN05N=pd.read_csv(Ruta_VA05N+Name_VA05N+".txt",delimiter="\t",skiprows=1)
Pedidos_VN05N=Funtions.Clean_Columns(Pedidos_VN05N)
Pedidos_VN05N=Funtions.default_column(Defaul_Column_Pedidos_dia,Pedidos_VN05N)
#Elimino pedidos con valores nulos
Pedidos_VN05N=Pedidos_VN05N[Pedidos_VN05N['Pedido'].notnull()]

# Estandarizo

Agenda=["85","20","146","149","50","138","45"]

Pedidos_VN05N['Pedido']=Funtions.Estandarizo_Pedidos(Pedidos_VN05N['Pedido'])
Pedidos_VN05N['Pedido']=Funtions.complete_pedidos(Pedidos_VN05N['Pedido'],Agenda)

#Dia actual
now=datetime.now().date()

if now.strftime("%A")=='Monday':
    #Clientes lunes
    Lunes_Cliente_Exito=["0085","0045"]  #"0085"  Funza, Surtimayoristas
    Lunes_Cliente_Cencosub=["93","122","127","95"] #"93-","122-","127-" Medellin, Barranquilla, Bucaramanga y cali
    buscar_exito=Funtions.Search_Agenda_Exito(Pedidos_VN05N['Pedido'],Lunes_Cliente_Exito)
    buscar_cencosub=Funtions.Search_agenda_Cencosub(Pedidos_VN05N['Pedido'],Lunes_Cliente_Cencosub)
    filtro_exito_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_exito)]
    filtro_cencosub_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_cencosub)]
elif now.strftime("%A")=='Tuesday':
    #Clientes Martes
    Martes_Cliente_Exito=["0020","0045"]  #"0020"  VEGAS, Surtimayoristas
    Martes_Cliente_Cencosub=["93","122","127"] #"93-","122-","127-" Medellin, Barranquilla y Bucaramanga
    buscar_exito=Funtions.Search_Agenda_Exito(Pedidos_VN05N['Pedido'],Martes_Cliente_Exito)
    buscar_cencosub=Funtions.Search_agenda_Cencosub(Pedidos_VN05N['Pedido'],Martes_Cliente_Cencosub)
    filtro_exito_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_exito)]
    filtro_cencosub_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_cencosub)]
elif now.strftime("%A")=='Wednesday':
    #Clientes Miercoles
    Miecoles_Cliente_Exito=["0085","0045"]  #"0085"  Funza
    Martes_Cliente_Cencosub=["Sin programa"]
    buscar_exito=Funtions.Search_Agenda_Exito(Pedidos_VN05N['Pedido'],Miecoles_Cliente_Exito)
    buscar_cencosub=Funtions.Search_agenda_Cencosub(Pedidos_VN05N['Pedido'],Martes_Cliente_Cencosub)
    filtro_exito_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_exito)]
    filtro_cencosub_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_cencosub)]
elif now.strftime("%A")=='Thursday':   
    #Clientes Jueves
    Jueves_Cliente_Exito=["0020","0146","0149","0045"]  #"0020"  VEGAS, Barranquilla, Bucaramanga, Surtimayoristas
    Jueves_Cliente_Cencosub=["93","122","127"] #"93-","122-","127-" Medellin, Barranquilla y Bucaramanga
    buscar_exito=Funtions.Search_Agenda_Exito(Pedidos_VN05N['Pedido'],Jueves_Cliente_Exito)
    buscar_cencosub=Funtions.Search_agenda_Cencosub(Pedidos_VN05N['Pedido'],Jueves_Cliente_Cencosub)
    filtro_exito_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_exito)]
    filtro_cencosub_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_cencosub)]
elif now.strftime("%A")=='Friday': 
    #Clientes Jueves
    Viernes_Cliente_Exito=["0050","0138","0045"]  #"0020"  Cali, Pereira, Surtimayoristas 
    Viernes_Cliente_Cencosub=["60"] #"60" Bogota
    buscar_exito=Funtions.Search_Agenda_Exito(Pedidos_VN05N['Pedido'],Viernes_Cliente_Exito)
    buscar_cencosub=Funtions.Search_agenda_Cencosub(Pedidos_VN05N['Pedido'],Viernes_Cliente_Cencosub)
    filtro_exito_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_exito)]
    filtro_cencosub_dia=Pedidos_VN05N[Pedidos_VN05N['Pedido'].isin(buscar_cencosub)]


filtro_exito_cencosub_dia=pd.concat([filtro_exito_dia['Pedido'],filtro_cencosub_dia['Pedido']])

#Busco los que tengan Pedidos, Entregados, Despachados y facturados. 
Tabla_ZSD79=SAP_GUI.Search_ZSD79('zsd79',filtro_exito_cencosub_dia,session)

def Search_Table_ZSD79(table,session):
    """
    -table: Tabla ZSD79
    -session: session
    (Busca filtro especial de colores)
    """
    #Columna Pedido concluido 
    table.SelectColumn("LFGSK")
    #Columna entregas concluidas
    table.SelectColumn("WBSTK")
    session.findById("wnd[0]/mbar/menu[1]/menu[3]").select() 
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "@0A@"  #Rojo Pedido
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN002-LOW").text = "@08@"  #Verde Entrega
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    Row=table.RowCount
    dic={"Nº Pedido cliente":[],"Clase Orden":[],"PrimFecEnt":[],"ÚltEntrega":[]}
    for i in range(0,Row):
        #Nº Pedido cliente Col "BSTNK"
        #Clase Orden Col "AUART"
        #PrimFecEnt Col "AUDAT"
        #ÚltEntrega Col "VDATU"
        Pedido_cliente,Clase_Orden,PrimFecEnt,ÚltEntrega=table.GetCellValue(i,"BSTNK"),table.GetCellValue(i,"AUART"),table.GetCellValue(i,"AUDAT"),table.GetCellValue(i,"VDATU") 
        dic["Nº Pedido cliente"].append(Pedido_cliente)
        dic["Clase Orden"].append(Clase_Orden)
        dic["PrimFecEnt"].append(PrimFecEnt)
        dic["ÚltEntrega"].append(ÚltEntrega)
    return(pd.DataFrame(dic))

#Buscar los pedidos diarios de clientes
Tabla_Consolidado_Diaria=Search_Table_ZSD79(Tabla_ZSD79,session)

"""Send email"""

correos="practicante.picking@prebel.com.co"

def send_emails(*args,emails="",htmlbody="",subject=""):
    email=emails
    outlook=win32com.client.Dispatch("outlook.application")
    mail=outlook.CreateItem(0)
    mail.Subject=subject+" "+datetime.now().strftime('%#d %b %Y %H:%M')
    mail.To=email
    mail.HTMLBody=htmlbody.format(*args)
    mail.Send()


def style_df(df):
    """
    -df: Dataframe
    -column: Nombre de la columna en Str
    -value_left:Rango de valor izquierdo
    -value_right: Rango de valor derecho
    """
    return df.style \
        .set_table_styles([{'selector': "table,tr,th,td", 'props': [("border", "1px solid"), ('color', '#000'),("text-align","center")]}]) \


html="""
    <h2 style="text-align: center">INFORME DE PEDIDOS DIARIOS</h2>
    <p> Por medio del presente informe se evidencia las permanencia de pedidos diarios pendientes de entrega</p>

    <div">{0}</div>

    <p> Anticipo sinceros agradecimientos. </p>
 """

#Tabla=style_df(Tabla)     #Style between LI and LS

send_emails(Tabla_Consolidado_Diaria.to_html(),emails=correos,htmlbody=html,subject="INFORME DE PEDIDOS DIARIOS")

# Tabla_Consolidado_Diaria.to_csv("C:\\Users\\prac.ingindustrial2\\OneDrive - Prebel S.A\\Escritorio\\SAP\\Archivos_CSV\\Tabla_Consolidado_Diaria.txt",sep='\t',index=False)

Tabla_Consolidado_Diaria.to_csv("C:\\Users\\practicante.picking\\OneDrive - Prebel S.A\\Escritorio\\bot_picking\\Archivos_CSV\\Tabla_Consolidado_Diaria.txt",sep='\t',index=False)

#!/usr/bin/env python
# coding: utf-8

# In[47]:


# IMPORTAR MODULOS
import re #Este módulo proporciona operaciones de coincidencia de expresiones regulares similares.
import os #permiten manipular la estructura de directorios (para leer y escribir archivos)
from datetime import datetime #manejo formato fechas
import pdfplumber #módulo para manejo PDFs, permite la extracción del texto
import pandas as pd #módulo que facilita el análisis de datos, proporciona unas estructuras de datos flexibles y que permiten trabajar con ellos de forma muy eficiente
from collections import namedtuple #módulo que permite crear subclases de tuplas con campos con nombre
import psycopg2
import psycopg2.extras as extras
import numpy as np

DB_NAME="facturasmedicas"
DB_USER="postgres"
DB_PASS="S1c2020"
DB_HOST="10.41.0.110"
DB_PORT="5432"

try:
    conn=psycopg2.connect(database=DB_NAME,user=DB_USER,password=DB_PASS,host=DB_HOST,port=DB_PORT)
    
    print("SISAS")
except:
    
    print("PAILA")

Line = namedtuple('Line', 'No_Factura CUM Cantidad Valor_Unitario Valor_Total Contrato Fecha Numero')  #ENCABEZADO DE LAS COLUMNAS DEL EXCEL

cur=conn.cursor()
cur.execute("""

CREATE TABLE Prueba1
(
NO_FACTURA INT NOT NULL,
CUM TEXT NOT NULL,
CANTIDAD INT NOT NULL,
VALOR_UNITARIO DOUBLE PRECISION NOT NULL,
VALOR_TOTAL DOUBLE PRECISION NOT NULL,
CONTRATO TEXT NOT NULL,
FECHA TEXT NOT NULL,
NUMERO TEXT NOT NULL
)

""")

conn.commit()
print("dale menol")
#CAMBIO DE ESCRITORIO

Directorio=r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\UNIDROGAS\FACTURAS' #DIRECTORIO CON LAS FACTURAS
os.chdir(Directorio) #Cambio de escritorio
facturas=os.listdir(Directorio) #listado de facturas

#############DEFINICION DE LAS CADENAS DE CARACTERES CON LA INFORMACION QUE VA A SER BUSCADA#########################################


text_Contrato='^\d+.\d+.\d*.?\d* (\D+$)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NOMBRE DEL CLIENTE EN LAS FACTURAS  
text_Fecha='^(\d{4}.\d{2}.\d{2})$' #CADENAS DE CARACTERES QUE CORRESPONDEN A LA FECHA EN LAS FACTURAS
text_NoFactura='^\D{2}-(\d+)$' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NUMERO DE LAS FACTURAS
text_CUM='.*CUM:(\d+-\d*)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL CUM DE LOS MEDICAMENTOS EN LAS FACTURAS
text_Line='.*\s+(\d+)\s+(\d+)\s+\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+\s+(\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+)\s+(\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+)\s+\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+ \d{0,3}.?\d{0,3}.?\d{0,3}.?\d+' #CADENAS DE CARACTERES QUE CORRESPONDEN A LAS LINEAS DE LOS MEDICAMENTOS EN LAS FACTURAS



######################GUARDA LA INFORMACION QUE VA A SER BUSCADA EN LAS FACTURAS###############################################

contrato_re = re.compile(text_Contrato)
Fecha_re=re.compile(text_Fecha)
Numero_re=re.compile(text_NoFactura)
CUM_re=re.compile(text_CUM)
Line_re=re.compile(text_Line)

#########################################VARIABLES AUXILIARES#############################################################################
No_FacturaDF,CUMDF,CantidadDF,Valor_UnitarioDF,Valor_TotalDF,ContratoDF,FechaDF,NumeroDF=0,0,0,0,0,0,0,0
data=[]
sql_data=[]
Numero=0
aux1=0
aux2=0
#n=facturas[2]

for n in (facturas):
    print(n)
    Numero=Numero+1
    aux1=0
    aux2=0
    with pdfplumber.open(n) as pdf:
        pages = pdf.pages
        print (pages)

############################SE RECORREN TODAS LAS PAGINAS DEL PDF Y SE EXTRAE EL TEXTO############################################################
        for page in pdf.pages:
            text = page.extract_text()

############################SE RECORREN TODAS LAS LINEAS SEPARADAS POR LOS ESPACIOS DEL PDF BUSCANDO LA INFORMACIÓN DEFINIDA AL INICIO###############
            for line in text.split('\n'):
    
############################   BUSCA LA INFORMACIÓN DEL ENCABEZADO   ##################################################################
              
                cont = contrato_re.search(line)
                fecha=Fecha_re.search(line)
                num=Numero_re.search(line)
                if cont:
                    contra = cont.group(1) 
                elif fecha:
                    if aux1==0:
                        fechaf=fecha.group(1)
                        aux1=1
                elif num:
                    if aux2==0:
                        nume=num.group(1)
                        aux2=1

###########################   BUSCA LA INFORMACIÓN DE LAS TABLAS  #######################################################################

                if Line_re.search(line):
                    lin=Line_re.search(line)
                    valorU, cant, valorT = float(lin.group(3).replace(',', '')),lin.group(1), float(lin.group(4).replace(',', ''))
                elif CUM_re.search(line):
                        CUM=CUM_re.search(line)
                        v_CUM=CUM.group(1)
    
##########################  ARREGLA LOS DATOS EN CASO DE QUE SEA NECESARIO  #############################################################
                        fechaf=fechaf.replace('-', '/')
                        fecha_F=datetime.strptime(fechaf, '%Y/%m/%d')
                        fecha_F=datetime.strftime(fecha_F, '%d/%m/%Y')
                        data.append(Line(Numero, v_CUM, cant, valorU, valorT, contra, fecha_F, nume))
                        
###################################CONSOLIDA LA INFORMACION OBTENIDA EN UN DATAFRAME############################################                    

df = pd.DataFrame(data)

try:
    conn=psycopg2.connect(database=DB_NAME,user=DB_USER,password=DB_PASS,host=DB_HOST,port=DB_PORT)
    
    print("SISAS")
except:
    
    print("PAILA")

cur=conn.cursor()
nombre_Tabla='Prueba1'

# Create a list of tupples from the dataframe values
tuples = [tuple(x) for x in df.to_numpy()]
# Comma-separated dataframe columns
cols = ','.join(list(df.columns))
# SQL quert to execute
query  = "INSERT INTO %s(%s) VALUES %%s" % (nombre_Tabla, cols)
cursor = cur
try:
    extras.execute_values(cursor, query, tuples)
    conn.commit()
except (Exception, psycopg2.DatabaseError) as error:
    print("Error: %s" % error)
    conn.rollback()
    cursor.close()
print("execute_values() done")
cursor.close()


#for indice,fila in df.iterrows():
#    print("hola")
#    No_FacturaDF,CantidadDF,Valor_UnitarioDF,Valor_TotalDF,ContratoDF,FechaDF,NumeroDF=fila['No_Factura'],fila['Cantidad'], fila['Valor_Unitario'],fila['Valor_Total'],fila['Contrato'],fila['Fecha'],fila['Numero']
#    CUMDF=fila['CUM']
#    sql_data=No_FacturaDF,CUMDF,CantidadDF,Valor_UnitarioDF,Valor_TotalDF,ContratoDF,FechaDF,NumeroDF
#    sql="INSERT INTO Prueba1 (NO_FACTURA,CUM,CANTIDAD,VALOR_UNITARIO,VALOR_TOTAL,CONTRATO,FECHA,NUMERO) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
#    cur.execute(sql,sql_data) 
#    conn.commit
conn.close
####################################CREACIÓN DEL ARCHIVO EN EXCEL##################################################################
os.chdir(r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\UNIDROGAS')
#df_excel=df.to_excel("baseD_UD.xlsx")


# In[ ]:





# In[ ]:





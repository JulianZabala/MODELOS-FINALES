#!/usr/bin/env python
# coding: utf-8

# In[2]:


# IMPORTAR MODULOS
import re #Este módulo proporciona operaciones de coincidencia de expresiones regulares similares.
import os #permiten manipular la estructura de directorios (para leer y escribir archivos)
from datetime import datetime #manejo formato fechas
import pdfplumber #módulo para manejo PDFs, permite la extracción del texto
import pandas as pd #módulo que facilita el análisis de datos, proporciona unas estructuras de datos flexibles y que permiten trabajar con ellos de forma muy eficiente
from collections import namedtuple #módulo que permite crear subclases de tuplas con campos con nombre

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

Line = namedtuple('Line', 'No_Factura CUM Cantidad Valor_Unitario Valor_Total Contrato Fecha Numero')  #ENCABEZADO DE LAS COLUMNAS DEL EXCEL
contrato_re = re.compile(text_Contrato)
Fecha_re=re.compile(text_Fecha)
Numero_re=re.compile(text_NoFactura)
CUM_re=re.compile(text_CUM)
Line_re=re.compile(text_Line)

#########################################VARIABLES AUXILIARES#############################################################################

data=[]
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
                print(line)
    
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


####################################CREACIÓN DEL ARCHIVO EN EXCEL##################################################################
os.chdir(r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\UNIDROGAS')
df_excel=df.to_excel("baseD_UD.xlsx")


# In[ ]:





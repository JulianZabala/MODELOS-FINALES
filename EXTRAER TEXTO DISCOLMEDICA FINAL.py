#!/usr/bin/env python
# coding: utf-8

# In[15]:


# IMPORTAR PAQUETES

import re #Este módulo proporciona operaciones de coincidencia de expresiones regulares similares.
import os #permiten manipular la estructura de directorios (para leer y escribir archivos)
from datetime import datetime #manejo formato fechas
import pdfplumber #módulo para manejo PDFs, permite la extracción del texto
import pandas as pd #módulo que facilita el análisis de datos, proporciona unas estructuras de datos flexibles y que permiten trabajar con ellos de forma muy eficiente
from collections import namedtuple #módulo que permite crear subclases de tuplas con campos con nombre


#CAMBIO DE ESCRITORIO

Directorio=r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\DISCOLMEDICA\FACTURAS' #DIRECTORIO CON LAS FACTURAS
os.chdir(Directorio) #Cambio de escritorio
facturas=os.listdir(Directorio) #listado de facturas

#############  DEFINICION DE LAS CADENAS DE CARACTERES CON LA INFORMACION QUE VA A SER BUSCADA   #########################################

text_Contrato='(Cliente) (.*) (NIT)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NOMBRE DEL CLIENTE EN LAS FACTURAS
text_Fecha='(F.EXP:).* (\d{2}\-\d{2}\-\d{4}).*' #CADENAS DE CARACTERES QUE CORRESPONDEN A LA FECHA EN LAS FACTURAS
text_NoFactura='^\s*[a-zA-Z]{3}(\d+$)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NUMERO DE LAS FACTURAS
text_CUM='.*CUM:(\d+-\d+)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL CUM DE LOS MEDICAMENTOS EN LAS FACTURAS
text_Line='.*\w \d{2}\/\d{2}\/\d{4} (\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+)' #CADENAS DE CARACTERES QUE CORRESPONDEN A LAS LINEAS DE LOS MEDICAMENTOS EN LAS FACTURAS


######################   GUARDA LA INFORMACION QUE VA A SER BUSCADA EN LAS FACTURAS  ###############################################

Line = namedtuple('Line', 'No_Factura CUM Cantidad Valor_Unitario Valor_Total Contrato Fecha Numero')
contrato_re = re.compile(text_Contrato)
Fecha_re=re.compile(text_Fecha)
Numero_re=re.compile(text_NoFactura)
CUM_re=re.compile(text_CUM)
Line_re=re.compile(text_Line)

#########################################VARIABLES AUXILIARES#############################################################################

data=[]
Numero=0
aux1=0
n=facturas[0]

for n in (facturas):
    print(n)
    Numero=Numero+1
    with pdfplumber.open(n) as pdf:
            pages = pdf.pages
            
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
                        contra = cont.group(2) 
                    elif fecha:
                        fechaf=fecha.group(2)
                    elif num:
                        nume=num.group(1)

###########################   BUSCA LA INFORMACIÓN DE LAS TABLAS  #######################################################################                            

                    if Line_re.search(line):
                        lin=Line_re.search(line)
                        v_u=lin.group(3).replace('.', '')
                        v_t=lin.group(5).replace('.', '')
                        valorU, cant, valorT = float(v_u.replace(',', '.')),lin.group(1), float(v_t.replace(',', '.'))
                    elif CUM_re.search(line):
                        CUM=CUM_re.search(line)
                        v_CUM=CUM.group(1)
                        fechaf=fechaf.replace('-', '/')
                        data.append(Line(Numero, v_CUM, cant, valorU, valorT, contra, fechaf, nume))
                        
###################################CONSOLIDA LA INFORMACION OBTENIDA EN UN DATAFRAME############################################                    

df = pd.DataFrame(data)

####################################CREACIÓN DEL ARCHIVO EN EXCEL##################################################################

os.chdir(r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\DISCOLMEDICA')
df_excel=df.to_excel("baseD_DM.xlsx")


# In[ ]:





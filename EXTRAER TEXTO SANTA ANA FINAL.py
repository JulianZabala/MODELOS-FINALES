#!/usr/bin/env python
# coding: utf-8

# In[1]:


# IMPORTAR PAQUETES
import re #Este módulo proporciona operaciones de coincidencia de expresiones regulares similares.
import os #permiten manipular la estructura de directorios (para leer y escribir archivos)
from datetime import datetime #manejo formato fechas
import pdfplumber #módulo para manejo PDFs, permite la extracción del texto
import pandas as pd #módulo que facilita el análisis de datos, proporciona unas estructuras de datos flexibles y que permiten trabajar con ellos de forma muy eficiente
from collections import namedtuple #módulo que permite crear subclases de tuplas con campos con nombre


#CAMBIO DE ESCRITORIO
Directorio=r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\SANTA ANA\FACTURAS' #DIRECTORIO CON LAS FACTURAS
os.chdir(Directorio) #Cambio de escritorio
facturas=os.listdir(Directorio) #listado de facturas


#############  DEFINICION DE LAS CADENAS DE CARACTERES CON LA INFORMACION QUE VA A SER BUSCADA  #########################################

text_Contrato='(SEÑORES) (.*) (NIT)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NOMBRE DEL CLIENTE EN LAS FACTURAS
text_Fecha='(Fecha:) (\d{2}/\d{2}/\d{4})' #CADENAS DE CARACTERES QUE CORRESPONDEN A LA FECHA EN LAS FACTURAS
text_NoFactura='(Número:) \D{4}(\d+)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NUMERO DE LAS FACTURAS
text_CUM='(\d{8}\-\d{2}|^.?\w{3,10}\-?\d*?\s)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL CUM DE LOS MEDICAMENTOS EN LAS FACTURAS
text_Line='(\d{8}\-\d{2}|^.?\w{3,10}\-?\d*?\s)\-?.*? (\d{2}\/\d{2}\/\d{4}) (\d+.\d+.?\d*?.?\d*?) (\d+) (\d+\.\d+) (\d+\.\d+) (\d+\.\d+) (\d+.\d+\.?\d*?.?\d*?)' #CADENAS DE CARACTERES QUE CORRESPONDEN A LAS LINEAS DE LOS MEDICAMENTOS EN LAS FACTURAS
#CADENAS DE CARACTERES QUE CORRESPONDEN A LA CATEGORIA QUE DEBE SER BUSCADA
text_MEDICAMENTOS='^MEDICAMENTO NO POS|^MEDICAMENTOS'
text_MEDICAMENTOS_F='SUBTOTAL MEDICAMENTO NO POS|SUBTOTAL MEDICAMENTOS'

###################### GUARDA LA INFORMACION QUE VA A SER BUSCADA EN LAS FACTURAS ###############################################

Line = namedtuple('Line', 'No_Factura CUM Cantidad Valor_Unitario Valor_Total Contrato Fecha Numero') #ENCABEZADO DE LAS COLUMNAS DEL EXCEL
contrato_re = re.compile(text_Contrato)
Fecha_re=re.compile(text_Fecha)
Numero_re=re.compile(text_NoFactura)
CUM_re=re.compile(text_CUM)
MEDICAMENTOS_re=re.compile(text_MEDICAMENTOS)
MEDICAMENTOS_F_re=re.compile(text_MEDICAMENTOS_F)
Line_re=re.compile(text_Line)

#########################################VARIABLES AUXILIARES#############################################################################

data=[]
Numero=0


for n in (facturas):
    aux=0
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
                            nume=num.group(2)
                
############################   BUSCA LA INFORMACIÓN DE LA CATEGORÍA QUE DEBE SER ANALIZADA   ##################################################################
    
                        inicio=MEDICAMENTOS_re.search(line)
                        fin=MEDICAMENTOS_F_re.search(line)
                        if inicio:
                            aux=1
                        if fin:
                            aux=0
                            
                        if aux==1:                            
###########################   BUSCA LA INFORMACIÓN DE LAS TABLAS  #######################################################################                            
                            if Line_re.search(line):
                                lin=Line_re.search(line)
                                CUM, valorU, cant, valorT = lin.group(1),float(lin.group(3).replace(',', '')), lin.group(4), float(lin.group(8).replace(',', ''))
                                data.append(Line(Numero, CUM, cant, valorU, valorT, contra, fechaf, nume))

###################################CONSOLIDA LA INFORMACION OBTENIDA EN UN DATAFRAME############################################                    

df = pd.DataFrame(data)

####################################CREACIÓN DEL ARCHIVO EN EXCEL##################################################################

os.chdir(r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\SANTA ANA')
df_excel=df.to_excel("baseD_SA.xlsx")


# In[ ]:





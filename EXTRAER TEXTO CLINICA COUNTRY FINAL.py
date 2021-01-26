#!/usr/bin/env python
# coding: utf-8

# In[27]:


# IMPORTAR PAQUETES

import re #Este módulo proporciona operaciones de coincidencia de expresiones regulares similares.
import os #permiten manipular la estructura de directorios (para leer y escribir archivos)
from datetime import datetime #manejo formato fechas
import pdfplumber #módulo para manejo PDFs, permite la extracción del texto
import pandas as pd #módulo que facilita el análisis de datos, proporciona unas estructuras de datos flexibles y que permiten trabajar con ellos de forma muy eficiente
from collections import namedtuple #módulo que permite crear subclases de tuplas con campos con nombre

#CAMBIO DE ESCRITORIO

Directorio=r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\CLINICA COUNTRY\FACTURAS' #DIRECTORIO CON LAS FACTURAS
os.chdir(Directorio) #Cambio de escritorio
facturas=os.listdir(Directorio) #listado de facturas

#############  DEFINICION DE LAS CADENAS DE CARACTERES CON LA INFORMACION QUE VA A SER BUSCADA   #########################################

text_Contrato='(SEÑORES) (.*) (FACTURA DE VENTA:)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NOMBRE DEL CLIENTE EN LAS FACTURAS
text_Fecha='(FECHA FACTURA:).* (\d{2}\/\d{2}\/\d{4}).* (FECHA VENCIMIENTO:)' #CADENAS DE CARACTERES QUE CORRESPONDEN A LA FECHA EN LAS FACTURAS
text_NoFactura='.*(No:) \D{5}(\d+)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NUMERO DE LAS FACTURAS
text_CUM1='.*CUM-(\d{4,}-\d+)\)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL PRIMER TIPO DE CUM DE LOS MEDICAMENTOS EN LAS FACTURAS
#CADENAS DE CARACTERES QUE CORRESPONDEN AL SEGUNDO TIPO DE CUM DE LOS MEDICAMENTOS EN LAS FACTURAS
text_CUM2='.*CUM-(\d{4,}-?)$' 
text_Line1_CUM2='^\s{0,15}?(-?\d{2})\).*'
#CADENAS DE CARACTERES QUE CORRESPONDEN AL TERCER TIPO DE CUM DE LOS MEDICAMENTOS EN LAS FACTURAS
text_CUM3='.*CUM-$' 
text_Line1_CUM3='^\s{0,15}?(-?\d{2})\).*'
#CADENAS DE CARACTERES QUE CORRESPONDEN A LAS LINEAS DE LOS MEDICAMENTOS EN LAS FACTURAS
text_Line='.*\w \d{2}\/\d{2}\/\d{4} (\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+) (\d{0,3}.?\d{0,3}.?\d{0,3}.?\d+)' 
text_Line_Vacia='^\s+$|^\s*?(\d{2}\/\d{2}\/\d{4})'

######################   GUARDA LA INFORMACION QUE VA A SER BUSCADA EN LAS FACTURAS  ###############################################

Line = namedtuple('Line', 'No_Factura CUM Cantidad Valor_Unitario Valor_Total Contrato Fecha Numero')

contrato_re = re.compile(text_Contrato)
Fecha_re=re.compile(text_Fecha)
Numero_re=re.compile(text_NoFactura)
CUM1_re=re.compile(text_CUM1)
CUM2_re=re.compile(text_CUM2)
Line1_CUM2_re=re.compile(text_Line1_CUM2)
CUM3_re=re.compile(text_CUM3)
Line1_CUM3_re=re.compile(text_Line1_CUM3)
Line2_re=re.compile(text_Line)
Line_Vacia_re=re.compile(text_Line_Vacia)


#########################################VARIABLES AUXILIARES#############################################################################

data=[]
Numero=0
aux1=0
aux2=0
aux3=0


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
                        nume=num.group(2)
                        
###########################   BUSCA LA INFORMACIÓN DE LAS TABLAS  #######################################################################                            
    
    ########   BUSCA LA INFORMACIÓN PARA EL PRIMER TIPO DE CUM ##############
    
                    if CUM1_re.search(line):
                        CUM=CUM1_re.search(line)
                        if CUM:
                            aux1=1
                            v_CUM=CUM.group(1)
                    if aux1==1:
                        if CUM1_re.search(line):
                            aux1=1
                            aux2=0
                            aux3=0
                        elif Line_Vacia_re.search(line):
                            aux1=1
                        elif Line2_re.search(line):
                            lin=Line2_re.search(line)
                            CUM=v_CUM
                            valorU, cant, valorT = float(lin.group(1).replace('.', '')), lin.group(2), float(lin.group(5).replace('.', ''))
                            data.append(Line(Numero, CUM, cant, valorU, valorT, contra, fechaf, nume))
                        else:
                            aux1=0
                            aux2=0
                            aux3=0
                            
     ########   BUSCA LA INFORMACIÓN PARA EL SEGUNDO TIPO DE CUM ##############
    
                    if CUM2_re.search(line):
                        CUM=CUM2_re.search(line)
                        if CUM:
                            aux2=1
                            v_CUM=CUM.group(1)
                    if aux2==1:
                        if CUM2_re.search(line):
                            CUM=CUM2_re.search(line)
                            v_CUM=CUM.group(1)
                            aux2=1
                            aux1=0
                            aux3=0
                        elif Line_Vacia_re.search(line):
                            aux2=1
                        elif Line1_CUM2_re.search(line):
                            lin=Line1_CUM2_re.search(line)
                            v_CUM2=lin.group(1)
                            CUM=v_CUM + v_CUM2 
                        elif Line2_re.search(line):
                            lin=Line2_re.search(line)
                            valorU, cant, valorT = float(lin.group(1).replace('.', '')), lin.group(2), float(lin.group(5).replace('.', ''))
                            data.append(Line(Numero, CUM, cant, valorU, valorT, contra, fechaf, nume))
                        else:
                            aux1=0
                            aux2=0
                            aux3=0
                            
    ########   BUSCA LA INFORMACIÓN PARA EL TERCER TIPO DE CUM ##############
    
                    if CUM3_re.search(line):
                        CUM=CUM3_re.search(line)
                        if CUM:
                            aux3=1
                            aux2=0
                            aux1=0
                    if aux3==1:
                        if CUM3_re.search(line):
                            aux3=1
                        elif Line_Vacia_re.search(line):
                            aux3=1
                        elif Line1_CUM3_re.search(line):
                            lin=Line1_CUM3_re.search(line)
                            v_CUM2=lin.group(1)
                            CUM= v_CUM2 
                        elif Line2_re.search(line):
                            lin=Line2_re.search(line)
                            valorU, cant, valorT = float(lin.group(1).replace('.', '')), lin.group(2), float(lin.group(5).replace('.', ''))
                            data.append(Line(Numero, CUM, cant, valorU, valorT, contra, fechaf, nume))
                        else:
                            aux1=0
                            aux2=0
                            aux3=0
                            
                            
###################################CONSOLIDA LA INFORMACION OBTENIDA EN UN DATAFRAME############################################                    
    
df = pd.DataFrame(data)

####################################CREACIÓN DEL ARCHIVO EN EXCEL##################################################################

os.chdir(r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\CLINICA COUNTRY')
df_excel=df.to_excel("baseD_CC.xlsx")


# In[ ]:





# In[ ]:





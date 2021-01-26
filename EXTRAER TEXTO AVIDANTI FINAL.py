#!/usr/bin/env python
# coding: utf-8

# In[10]:


# IMPORTAR PAQUETES

import re #Este módulo proporciona operaciones de coincidencia de expresiones regulares similares.
import os #permiten manipular la estructura de directorios (para leer y escribir archivos)
from datetime import datetime #manejo formato fechas
import pdfplumber #módulo para manejo PDFs, permite la extracción del texto
import pandas as pd #módulo que facilita el análisis de datos, proporciona unas estructuras de datos flexibles y que permiten trabajar con ellos de forma muy eficiente
from collections import namedtuple #módulo que permite crear subclases de tuplas con campos con nombre

#CAMBIO DE ESCRITORIO

Directorio=r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\AVIDANTI\FACTURAS' #DIRECTORIO CON LAS FACTURAS
os.chdir(Directorio) #Cambio de escritorio
facturas=os.listdir(Directorio) #listado de facturas

#############  DEFINICION DE LAS CADENAS DE CARACTERES CON LA INFORMACION QUE VA A SER BUSCADA   #########################################

text_Contrato='(Señores:) (.*) (NIT)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NOMBRE DEL CLIENTE EN LAS FACTURAS
text_Fecha='(Plan:).* (Fecha)  (\d{2}\/[a-zA-Z]+\.+\/\d{4})' #CADENAS DE CARACTERES QUE CORRESPONDEN A LA FECHA EN LAS FACTURAS
text_NoFactura='Contrato:.+(No.) \D{4}(\d+)' #CADENAS DE CARACTERES QUE CORRESPONDEN AL NUMERO DE LAS FACTURAS
text_CUM='(\d{5,}-\d{0,2})\-?.*?' #CADENAS DE CARACTERES QUE CORRESPONDEN AL CUM DE LOS MEDICAMENTOS EN LAS FACTURAS

#CADENAS DE CARACTERES QUE CORRESPONDEN A LA CATEGORIA QUE DEBE SER BUSCADA

text_MEDICAMENTOS='^\d+ - MEDICAMENTOS NO POS|^\d+ - MEDICAMENTOS POS')'
text_MEDICAMENTOS_F='^Total \d+ - MEDICAMENTOS NO POS|^Total \d+ - MEDICAMENTOS POS'
text_MEDICAMENTOSr='^\d+ - MEDICAMENTOS REGULADOS POS|^\d+ - MEDICAMENTOS REGULADOS NO POS'
text_MEDICAMENTOSr_F='^Total \d+ - MEDICAMENTOS REGULADOS POS|^Total \d+ - MEDICAMENTOS REGULADOS NO POS'
text_MEDICAMENTOSPBS='^\d+ - MEDICAMENTOS PBS|^\d+ - MEDICAMENTOS REGULADOS PBS'
text_MEDICAMENTOSPBS_F='^Total \d+ - MEDICAMENTOS PBS|^Total \d+ - MEDICAMENTOS REGULADOS PBS'

#CADENAS DE CARACTERES QUE CORRESPONDEN AL PRIMER TIPO DE LINEAS EN FUNCIÓN DEL CUM DE LOS MEDICAMENTOS EN LAS FACTURAS
text_Line1='(\d+)\-?.*? (\d?.\d+) (\d?.\d+) (\d?.\d+) (\d+.\d+.?[\d{2}\d{3}]?.?\d{0,2})' 
text_Line11='^(\d{2})'
text_Line12='(^\d+)\s* (\d+)\s.* (\d?.\d+) (\d?.\d+) (\d?.\d+) (\d+.\d+.?[\d{2}\d{3}]?.?\d{0,2})'

#CADENAS DE CARACTERES QUE CORRESPONDEN AL SEGUNDO TIPO DE LINEAS DE LOS MEDICAMENTOS EN LAS FACTURAS
text_Line2='(\d+)\-?.*? (\d{5,}\-?\d{0,2}|[a-zA-Z]+\d+).*? (\d?.\d+) (\d?.\d+) (\d?.\d+) (\d+.\d+.?[\d{2}\d{3}]?.?\d{0,2})' 

#CADENAS DE CARACTERES QUE CORRESPONDEN AL TERCER TIPO DE LINEAS DE LOS MEDICAMENTOS EN LAS FACTURAS
text_Line3='(\d+)\-?.*? (\d+) (\d{5,}\-?\d{0,2}|[a-zA-Z]+\d+).* (\d?.\d+) (\d?.\d+) (\d?.\d+) (\d+.\d+.?[\d{2}\d{3}]?.?\d{0,2})'

######################   GUARDA LA INFORMACION QUE VA A SER BUSCADA EN LAS FACTURAS  ###############################################


Line = namedtuple('Line', 'No_Factura CUM Cantidad Valor_Unitario Valor_Total Contrato Fecha Numero')
contrato_re = re.compile(text_Contrato)
Fecha_re=re.compile(text_Fecha)
Numero_re=re.compile(text_NoFactura)
CUM_re=re.compile(text_CUM)
MEDICAMENTOS_re=re.compile(text_MEDICAMENTOS)
MEDICAMENTOS_F_re=re.compile(text_MEDICAMENTOS_F)
MEDICAMENTOSr_re=re.compile(text_MEDICAMENTOSr)
MEDICAMENTOSr_F_re=re.compile(text_MEDICAMENTOSr_F)
MEDICAMENTOSPBS_re=re.compile(text_MEDICAMENTOSPBS)  
MEDICAMENTOSPBS_F_re=re.compile(text_MEDICAMENTOSPBS_F)
Line1_re=re.compile(text_Line1)
Line11_re=re.compile(text_Line11)
Line12_re=re.compile(text_Line12)
Line2_re=re.compile(text_Line2)
Line3_re=re.compile(text_Line3)


#########################################VARIABLES AUXILIARES#############################################################################

data=[]
Numero=0
aux=0
aux2=0
aux3=0


for n in (facturas):
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
                            fechaf=fecha.group(3)
                            fechaf=fechaf.replace('.', '')
                        elif num:
                            nume=num.group(2)
                            
############################   BUSCA LA INFORMACIÓN DE LA CATEGORÍA QUE DEBE SER ANALIZADA   ##################################################################

                        inicio=MEDICAMENTOS_re.search(line)
                        fin=MEDICAMENTOS_F_re.search(line)
                        inicior=MEDICAMENTOSr_re.search(line)
                        finr=MEDICAMENTOSr_F_re.search(line)
                        inicioPBS=MEDICAMENTOSPBS_re.search(line)
                        finPBS=MEDICAMENTOSPBS_F_re.search(line)
                        if inicio or inicior or inicioPBS:
                            aux=1
                        if fin or finr or finPBS:
                            aux=0
###########################   BUSCA LA INFORMACIÓN DE LAS TABLAS  #######################################################################                            
                        if aux==1:
        
########   BUSCA LA INFORMACIÓN PARA EL TERCER TIPO DE LÍNEA ##############

                            if Line3_re.search(line):
                                lin=Line3_re.search(line)
                                v_t=lin.group(7).replace('.', '')
                                CUM, cant, valorT = lin.group(3),float(lin.group(2)), float(v_t.replace(',', '.'))
                                valorU=float(valorT/cant)
                                data.append(Line(Numero, CUM, cant, valorU, valorT, contra, fechaf, nume))
                            
########   BUSCA LA INFORMACIÓN PARA EL SEGUNDO TIPO DE LÍNEA ##############

                            elif Line2_re.search(line):

                                lin=Line2_re.search(line)
                                v_t=lin.group(6).replace('.', '')
                                CUM, cant, valorT = lin.group(2), float(lin.group(1)), float(v_t.replace(',', '.'))
                                valorU=float(valorT/cant)
                                data.append(Line(Numero, CUM, cant, valorU, valorT, contra, fechaf, nume))
                            
########   BUSCA LA INFORMACIÓN EN FUNCIÓN DEL CUM ##############
                            elif CUM_re.search(line):
                                CUM=CUM_re.search(line)
                                if CUM:
                                    aux2=1
                            if aux2==1:
                                CUM=CUM_re.search(line)
                                if CUM: 
                                    v_CUM=CUM.group(1)
                                    
########   BUSCA LA INFORMACIÓN PARA EL SEGUNDO TIPO DE LÍNEA EN FUNCION DEL CUM ##############

                                elif Line12_re.search(line):
                                    if Line12_re.search(line):
                                        lin=Line12_re.search(line)
                                        v_t=lin.group(6).replace('.', '')
                                        cant,valorT=float(lin.group(2)),float(v_t.replace(',', '.'))
                                        valorU=float(valorT/cant)
                                        aux3=1
                                    
########   BUSCA LA INFORMACIÓN PARA EL PRIMER TIPO DE LÍNEA EN FUNCION DEL CUM ##############

                                elif Line1_re.search(line):
                                        lin=Line1_re.search(line)
                                        v_t=lin.group(5).replace('.', '')
                                        cant,valorT=float(lin.group(1)),float(v_t.replace(',', '.'))
                                        valorU=float(valorT/cant)
                                        aux3=1 
                            
########   BUSCA LA INFORMACIÓN PARA EL TERCER TIPO DE LÍNEA EN FUNCION DEL CUM ##############

                                elif Line11_re.search(line) and aux3==1:
                                    lin=Line11_re.search(line)
                                    v_CUM2=lin.group(1)
                                    CUM=v_CUM + v_CUM2
                                    data.append(Line(Numero, CUM, cant, valorU, valorT, contra, fechaf, nume))
                                    aux2=0 
                                    aux3=0
                                
###################################CONSOLIDA LA INFORMACION OBTENIDA EN UN DATAFRAME############################################                    
                                
df = pd.DataFrame(data)

####################################CREACIÓN DEL ARCHIVO EN EXCEL##################################################################

os.chdir(r'C:\Users\C.JZABALA\Documents\PROGRAMACION FACTURAS\AVIDANTI')
df_excel=df.to_excel("baseD_AV.xlsx")


# In[ ]:





# In[ ]:





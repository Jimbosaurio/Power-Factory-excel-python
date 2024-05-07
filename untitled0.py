# -*- coding: utf-8 -*-
#El_JIMBOSAURIO
#ESTAS LINEAS VAN EN CUALQUIER SCRIPT 
import os 
os.environ["PATH"]=r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP2"+os.environ["PATH"]
#vinculacion con pf 
import sys 
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.8")
#importar app 
import powerfactory as pf 
app=pf.GetApplication()
#abro pf en modo engine
app.Show()
#ACTIVAR EL PROYECTO 
user=app.GetCurrentUser()
project=app.ActivateProject('04 BOCETO (43) Extraccion')
prj=app.GetActiveProject()
#ACCEDER A OBJETOS 


ldf=app.GetFromStudyCase('ComLdf')
ldf.Execute()


#DICCIONARIO DE LINEAS
lines_dict={}

#EXTRAEMOS DEL DICK
line = app.GetCalcRelevantObjects('*.ElmLne')

#CARGAMOS MEDIANTE EL NOMBRE
for i in line:
    lines_dict[i.loc_name]=i
    
#GENERA VECTOR LINEAS R1 y X1
loading_Vector=[]
Name_Vector=[]
R1_Vector=[]
X1_Vector=[]
for i in line:
    R1=lines_dict[i.loc_name].GetAttribute('R1')
    X1=lines_dict[i.loc_name].GetAttribute('X1')
    cargab=lines_dict[i.loc_name].GetAttribute('c:loading')
    print(i.loc_name)
    #print(R1)
    R1_Vector.append(R1)
    Name_Vector.append(i.loc_name)
    X1_Vector.append(X1)
    loading_Vector.append(cargab)
    




ldf=app.GetFromStudyCase('ComShc')
ldf.iopt_mde=2
ldf.iopt_shc='3psc'
ldf.Execute()

#dickcionario
bus_dict={}
#EXTRAEMOS DEL DICK
buses = app.GetCalcRelevantObjects('*.ElmTerm')


for x in buses:
    bus_dict[x.loc_name]=x 
    
    
Terminal_Vector=[]
ISYM_Vector=[]
for x in bus_dict:
    ISYM=bus_dict[x].GetAttribute('m:Isym_m')
    print(x)
    print(ISYM)
    Terminal_Vector.append(x)
    ISYM_Vector.append(ISYM)






#COMENZAMOS CON GUARDADO AUTOMATICO

from win32com import client
excel=client.Dispatch('excel.Application')
excel.visible= True

wb = excel.Workbooks.Add()
wb.Worksheets[0].Name = 'Registrar'

# Seleccionar la hoja de cálculo "Registrar"
ws = wb.Worksheets['Registrar']

# Escribir encabezados en las celdas A1 y B1
ws.Cells(1, 1).Value = "Nombre de Linea"
ws.Cells(1, 2).Value = "R1"
ws.Cells(1, 3).Value = "X1"
ws.Cells(1, 4).Value = "loading %"
ws.Cells(1, 5).Value = "TERMINALES"
ws.Cells(1, 6).Value = "FALLA 3PSC_Isym_m"
# Escribir los nombres en la columna A a partir de la fila 2
for i, name in enumerate(Name_Vector):
    ws.Cells(i + 2, 1).Value = name
    
for i, name in enumerate(R1_Vector):
    ws.Cells(i + 2, 2).Value = name
    
for i, name in enumerate(X1_Vector):
    ws.Cells(i + 2, 3).Value = name
    
for i, name in enumerate(loading_Vector):
    ws.Cells(i + 2, 4).Value = name    
    
for i, name in enumerate(Terminal_Vector):
    ws.Cells(i + 2, 5).Value = name   
    
for i, name in enumerate(ISYM_Vector):
    ws.Cells(i + 2, 6).Value = name   
ruta_archivo = r"C:\Users\Admin\Desktop\PROTECCIONES\VERSIONES_DIG\phyton\excel_autom/archivo.xlsx"    
# Guardar el libro de trabajo en un archivo Excel
wb.SaveAs(ruta_archivo)


 

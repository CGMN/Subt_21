# -*- coding: utf-8 -*-
#python36
#guardar excel como to_csv



import xlrd
import csv
import time
import pandas as pd
from pandas.api.types import CategoricalDtype
import numpy as np
import tkinter.filedialog, re
import openpyxl
from openpyxl.styles import Alignment


import traceback

f= open("Archivo_de_error.txt","w")

var=0

try:
	#Permite seleccionar el archivo, abre el explorador y guarda la seleccion en la variable file_path
	#______________________________________________________________________________
	root = tkinter.Tk()
	root.withdraw()
	file_path = tkinter.filedialog.askopenfilename()

	#Definiciones___________________________________________________________________
	filtros=["Asig. Urgencia ","Asig. Urgencia (incremento)","Ley 19.536", "Horas extraordinarias",
		"Asignación de turno", "Bonificación compensatoria", "Viáticos", "Función crítica",
		"Asignación de responsabilidad", "Asignación de estímulo", "Experiencia Calificada",
		"Dedicación Exclusiva", "Suplencias y reemplazos"]


	print ("cantidad de tablas solicitadas: " +str(len(filtros)))

	experimentales=["50 Hospital Padre Alberto Hurtado","51 Centro de Referencia de Salud Maipu",
		"52 Centro de Referencia de Salud Penalolen Cordillera Oriente"]


	pos_tablas=[]
	espacio_de_tablas=260

	for i in range(0,len(filtros)+6):
		if i !=0 and i !=1 and i !=2 and i !=3 and i!=4 and i!=5:
			t=i*espacio_de_tablas
			pos_tablas.append(t)

	#Creacion de tablas_______________________________________________________________

	print ("leyendo base")

	starttime =  time.time()

	df = pd.read_csv(str(file_path), encoding="latin1", low_memory = False)


	#Para ordenar los meses al mostrar la tabla
	diccio_meses={"enero":1,"febrero":2, "marzo":3, "abril":4, "mayo":5,"junio":6,"julio":7,"agosto":8,
		"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}

	df.loc[0,"MES_NUMERO"]=""

	df['MES_NUMERO'].update(df['mes'].map(diccio_meses) )

	print (time.time()-starttime)

	print ("base leida")
	print (len(df))
	#Para tener una lista con las cabeceras, _____________________________________

	#Para tener un listado con las clasificaciones y despues usarlas para que no muera el programa si falta alguna
	clasificaciones_todas=[]
	for i in range(0, len(df)):
		clasificaciones_todas.append(df.CLASIFICACION[i])


	clasificaciones=list(set(clasificaciones_todas))

	print("cantidad listado clasificaciones", len(clasificaciones))


	#Quitamos de filtros lo que no está en el archivo______________________________
	quitar_de_filtros=[]
	for i in range(0,len(filtros)):
		if filtros[i] not in clasificaciones:
			quitar_de_filtros.append(filtros[i])


	for i in range(0,len(quitar_de_filtros)):
		a=filtros.index(quitar_de_filtros[i])
		filtros.insert(a,"0")
		filtros.remove(quitar_de_filtros[i])

	print ("cantidad de tablas a crear: "+ str(len(filtros)))

	if len(quitar_de_filtros)>0:
		print ("En el archivo sigfe no se encuentra: ")
		for i in range(0,len(quitar_de_filtros)):
			print (" ",quitar_de_filtros[i])

	print ("creando tablas")
	#creacion de tablas especificas, las 4 primeras______________________________________________________________

	df["CLASIFICACION"]=df["CLASIFICACION"].fillna(value="0")

	tablas=[]

	#TABLA SUBT 21
	tabla4=pd.pivot_table(df[df.SubTitulo=="21 GASTOS EN PERSONAL"], index=["Institucion","CodEstablecimiento SIRH"],
			columns=["MES_NUMERO"],	values=["Devengado"],aggfunc=[np.sum], fill_value=0, margins=True, margins_name='Total')

	#TABLA HONORARIOS SIN EXPERIMENTALES
	if "Honorarios asim. Ley 18.834" in clasificaciones:
		table1=pd.pivot_table(df[df.CLASIFICACION=="Honorarios asim. Ley 18.834"], index=["Institucion","CodEstablecimiento SIRH"],
			columns=["MES_NUMERO"],	values=["Devengado"],aggfunc=[np.sum], fill_value=0, margins=True, margins_name='Total')

	#TABLA HONORARIOS LEY MEDICA SIN EXPERIMENTALES
	if "Honorarios asim. Ley Médica" in clasificaciones:
		table2=pd.pivot_table(df[df.CLASIFICACION=="Honorarios asim. Ley Médica"], index=["Institucion","CodEstablecimiento SIRH"],
			columns=["MES_NUMERO"],	values=["Devengado"],aggfunc=[np.sum], fill_value=0, margins=True, margins_name='Total')

	#TABLA HONORARIOS 18834 Y LEY MEDICA SOLO EXPERIMENTALES
	if "Honorarios asim. Ley Médica" in clasificaciones and "Honorarios asim. Ley 18.834" in clasificaciones:
		table3=pd.pivot_table(df[(df.CLASIFICACION=="Honorarios asim. Ley Médica") | (df.CLASIFICACION=="Honorarios asim. Ley 18.834")],
			index=["Institucion","CodEstablecimiento SIRH"], columns=["MES_NUMERO"],	values=["Devengado"],aggfunc=[np.sum], fill_value=0, margins=True, margins_name='Total')

	#consultores de llamada
	table5=pd.pivot_table(df[(df.SubTitulo=="21 GASTOS EN PERSONAL")&(df.Item=="03 Otras Remuneraciones")&(df.Asignación=="001 Honorarios a Suma Alzada - Personas Naturales")
		&(df.SubAsignación=="001 Honorarios A Suma Alzada Personas Naturales")&(df.Específico=="01 Hon Conv Tratantes O Consult Llamadas Art 24 L 19664")],
			index=["Institucion","CodEstablecimiento SIRH"], columns=["MES_NUMERO"],values=["Devengado"],aggfunc=[np.sum], fill_value=0, margins=True, margins_name='Total')

	#33 mil horas
	table6=pd.pivot_table(df[(df.SubTitulo=="21 GASTOS EN PERSONAL")&(df.Item=="03 Otras Remuneraciones")&(df.Asignación=="001 Honorarios a Suma Alzada - Personas Naturales")
		&(df.SubAsignación=="001 Honorarios A Suma Alzada Personas Naturales")&(df.Específico=="06 Personal Médico Programa Cierre De Brechas")],
			index=["Institucion","CodEstablecimiento SIRH"], columns=["MES_NUMERO"],values=["Devengado"],aggfunc=[np.sum], fill_value=0, margins=True, margins_name='Total')


	#creacion de tablas generales_________________________________________________________________________________
	for i in range(0,len(filtros)):#index=filas

		table = pd.pivot_table(df[df.CLASIFICACION==filtros[i]],index=["Institucion","CodEstablecimiento SIRH"], columns=["MES_NUMERO"],
			values=["Devengado"],aggfunc=[np.sum],fill_value=0,margins=True,margins_name='Total')
		tablas.append(table)

	#Crea el archivo excel________________________________________________________
	writer = pd.ExcelWriter('Resumen_glosas.xlsx')

	#para posicionar las tablas en el archivo______________________________________


	tabla4.to_excel(writer,"Hoja1", startcol=1, startrow=0)

	if "Honorarios asim. Ley 18.834" in clasificaciones:
		table1.to_excel(writer,"Hoja1", startcol=1, startrow=espacio_de_tablas)

	if "Honorarios asim. Ley Médica" in clasificaciones:
		table2.to_excel(writer,"Hoja1", startcol=1, startrow=espacio_de_tablas*2)

	if "Honorarios asim. Ley Médica" in clasificaciones and "Honorarios asim. Ley 18.834" in clasificaciones:
		table3.to_excel(writer,"Hoja1", startcol=1, startrow=espacio_de_tablas*3)

	table5.to_excel(writer,"Hoja1",startcol=1, startrow=espacio_de_tablas*4)

	table6.to_excel(writer,"Hoja1",startcol=1,startrow=espacio_de_tablas*5)

	for i in range(0,len(tablas)):
		tablas[i].to_excel(writer,"Hoja1", startcol=1, startrow=pos_tablas[i])

	#Graba el archivo_____________________________________________________________
	writer.save()


	print ("Resumen_glosas creado")

	#Para darle formato al archivo________________________________________________________
	doc = openpyxl.load_workbook('Resumen_glosas.xlsx')

	H=doc.get_sheet_names() #la variable H guarda un listado con las hojas del excel
	hoja1 = doc.get_sheet_by_name(str(H[0])) #con H[0], le decimos que use la primera hoja del excel

	#Ancho de columna___________________________________________________________________________
	hoja1.column_dimensions['B'].width = 60


	#alineacion a la izquierda______________________________________________________________
	for row in hoja1['B1':'B'+str(pos_tablas[-1]+espacio_de_tablas)]:
		for cell in row:
		 	cell.alignment = Alignment(horizontal="left")

	#alineacion vertical al centro______________________________________________________________
	for row in hoja1['B1':'B'+str(pos_tablas[-1]+espacio_de_tablas)]:
		for cell in row:
			cell.alignment = Alignment(vertical="center")


	#cabeceras__________________________________________________________

	hoja1["B1"]="21 GASTOS EN PERSONAL"

	if "Honorarios asim. Ley 18.834" in clasificaciones:
		hoja1["B"+str(espacio_de_tablas*1+1)]="Honorarios asim. Ley 18.834 sin experimentales"

	if "Honorarios asim. Ley Médica" in clasificaciones:
		hoja1["B"+str(espacio_de_tablas*2+1)]="Honorarios asim. Ley Médica sin experimentales"

	if "Honorarios asim. Ley Médica" in clasificaciones and "Honorarios asim. Ley 18.834" in clasificaciones:
		hoja1["B"+str(espacio_de_tablas*3+1)]="Honorarios Ley 18.834 y Ley 19.664 solo experimentales"

	hoja1["B"+str(espacio_de_tablas*4+1)]="Consultores de Llamada"

	hoja1["B"+str(espacio_de_tablas*5+1)]="33.000 horas"

	for i in range(0,len(filtros)):
		hoja1["B"+str(pos_tablas[i]+1)]=filtros[i]


	#OK Para borrar los datos de las tablas en 0
	#for i in range(1,pos_tablas[-1]+espacio_de_tablas):
	for i in range(1,5000):
		if hoja1["B"+str(i)].value=="0":
			for j in hoja1["B"+str(i+4):"O"+str(i+245)]:
				for k in j:
					k.value=""

	#OK Para borrar los totales de las tablas ubicadas en segunda, tercera y cuarta ubicacion
	for i in range(450,1050):
		if hoja1["B"+str(i)].value=="Total":
			for j in hoja1["D"+str(i):"P"+str(i)]:
				for k in j:
					k.value=""

	#OK Para borrar los experimentales de la tablas ubicadas en segunda y tercera posicion del archivo
	for i in range(261,775):
		if hoja1["C"+str(i)].value==1314 or hoja1["C"+str(i)].value==1320 or hoja1["C"+str(i)].value==1394:
			for j in hoja1["D"+str(i):"P"+str(i)]:
				for k in j:
					k.value=""

	#OK borrar los otros servicios y dejando los experimentales
	todos_los_servicios=[101,103,125,127,130,201,203,211,212,216,217,221,301,306,307,316,317,318,401,406,407,
							408,411,416,417,418,419,420,501,502,503,506,507,511,516,517,525,527,530,531,540,541,
							543,544,545,546,547,548,550,555,556,560,565,566,567,568,569,570,601,603,606,611,616,617,618,
							619,620,621,622,623,624,640,642,647,701,703,704,706,711,716,717,718,719,720,721,722,723,724,
							801,802,803,806,817,819,820,821,824,825,827,829,830,831,834,836,837,838,840,843,845,846,847,848,
							850,852,865,866,867,875,880,890,891,892,893,895,896,901,906,907,911,916,917,919,920,950,953,962,
							963,968,971,972,973,974,990,991,992,993,996,998,1001,1003,1005,1006,1016,1017,1018,1019,1020,1021,
							1025,1027,1040,1041,1042,1043,1044,1050,1051,1055,1059,1060,1061,1065,1066,1067,1068,1069,1070,1071,
							1072,1073,1074,1090,1101,1102,1106,1116,1117,1118,1119,1120,1121,1201,1202,1206,1216,1217,1301,1303,
							1305,1306,1307,1308,1309,1313,1315,1316,1317,1318,1319,1325,1330,1332,1334,1338,1339,1340,1341,1345,
							1346,1349,1351,1352,1357,1358,1360,1362,1364,1365,1368,1369,1372,1374,1379,1390,1392,1396,1397,1398]


	for i in range(785,1035):
		for k in range(0,len(todos_los_servicios)):
			if hoja1["C"+str(i)].value==todos_los_servicios[k]:
				for j in hoja1["D"+str(i):"P"+str(i)]:
					for k in j:
						k.value=""

	#graba archivo
	doc.save("Resumen_glosas_establecimientos.xlsx")

	print ("Resumen_glosas modificado")

except:
	var = traceback.format_exc()
	print (var)
if var !=0:
	f.write(var)
	f.close()
	print ("")
	print (" X X X X X X X X X X X X X X X X X X ")
	print ("")
	print ("El programa arroja un error")
	print ("Se ha creado un archivo llamado 'Archivo_de_error' con el detalle")
	print ("")
	print (" X X X X X X X X X X X X X X X X X X ")

f.close()
print ("")
print ("proceso terminado")

time.sleep(6)

import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
from openpyxl import load_workbook
#wb = Workbook()
#ws = wb.active()
listPromL = []
listPromM = []
listPromC = []

wb = load_workbook(filename = 'prueba.xlsx')
sheet_ranges = wb['Sheet1']
#print(sheet_ranges['J2'].value)

def lectura(list,y): #Se le pasa una lista y una columna
	for x in range(2,17):
		asd = str(y)+ str(x)
		#print(asd)
		list.append(sheet_ranges[asd].value)
lectura(listPromL,'J')
print(listPromL)		
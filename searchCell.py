from openpyxl import Workbook
import openpyxl
import re
import os
import matplotlib.pyplot as plt
from itertools import cycle
#wb = load_workbook('datos.xlsx', data_only=True)

# fecha descripcion sucursal valor saldo
class EstadoDia:
  def __init__(self, fecha, descripcion, sucursal,valor, saldo, anio):
    self.fecha = fecha
    self.anio = anio
    self.descripcion = descripcion
    self.sucursal = sucursal
    self.valor = valor
    self.saldo = saldo
  def __init__(self, fecha, descripcion, sucursal,valor, saldo):
    self.fecha = fecha
    self.descripcion = descripcion
    self.sucursal = sucursal
    self.valor = valor
    self.saldo = saldo
  def __init__(self, fecha, descripcion, sucursal="",valor=0, saldo=0, anio=1996):
    self.fecha = fecha
    self.anio = anio
    self.descripcion = descripcion
    self.sucursal = sucursal
    self.valor = valor
    self.saldo = saldo

  def __repr__(self):
      return " fecha :"+self.fecha+" descripcion :"+self.descripcion+" valor :"+self.valor+" saldo "+self.saldo
  def __str__(self):
      return " fecha :"+self.fecha+" descripcion :"+self.descripcion+" valor :"+self.valor+" saldo "+self.saldo
  def mostrarFechaValor(self):
      print(" fecha : "+self.fecha+" valor "+self.valor)

#p1 = EstadoDia("John", 36)
def extraerAnio():
    nombreCelda = "DESDE"
    for row in ws.iter_rows(min_row=6,  
                            max_row=8,  
                            min_col=0,  
                            max_col=1):
        for cell in row:
            if cell.value == nombreCelda:
                ##print(ws.cell(row=(cell.row)+1, column=1).value) #change column number for any cell value you want
                anio = (ws.cell(row=(cell.row)+1, column=1).value)[0:4]
    if anio is None:
        print("no se encontro nada") 
    return anio
def extraerFechas():
    arrFechas = []
    nombreCelda = "FECHA"
    fechaEncontrada = False
    for row in ws.iter_rows(min_row=6,  
                            max_row=1000,  
                            min_col=0,  
                            max_col=6):
        for cell in row:
            if cell.value == nombreCelda:
                fechaEncontrada = True
                ##print(ws.cell(row=(cell.row)+1, column=1).value) #change column number for any cell value you want
            fecha = (ws.cell(row=(cell.row), column=1).value)
            if fechaEncontrada and fecha != None:
                #print(fecha)
                indice = 0
                #print(' cell row '+str(cell.row))
                if has_numbers(fecha): 
                    #print(fecha)
                    descripcion = (ws.cell(row=(cell.row), column=2).value)
                    sucursal = ""
                    if  ws.cell(row=(cell.row), column=3).value != None:
                        sucursal = ws.cell(row=(cell.row), column=3).value
                    valor = ws.cell(row=(cell.row), column=5).value
                    if has_numbers(valor):
                        #print('valor '+valor)
                        valor = valor.split(".")[0]
                        #print('valor '+valor)
                        valor = valor.replace(",","")
                        #print('valor '+valor)
                        if len(valor) > 1 :
                            valor = str(round(float(valor)))
                            saldo = ws.cell(row=(cell.row), column=6).value
                            fecha = fecha+"/"+anio
                            estadoDia = EstadoDia(fecha=fecha,descripcion=descripcion,sucursal=sucursal,valor=valor,saldo=saldo)
                            agregarEstado(arrEstados=arrFechas,estado=estadoDia)
                            fecha = (ws.cell(row=(cell.row), column=1).value)
                            fecha = fecha+"/"+anio
                            indice = indice +1
                else :
                    continue
                #print(fecha+' salida ')
                
    return arrFechas
def agregarEstado( arrEstados,estado ):
    yaExiste = False
    for elEstado in arrEstados:
        if elEstado.fecha == estado.fecha and elEstado.descripcion == estado.descripcion and elEstado.valor == estado.valor:
            yaExiste = True
    if yaExiste == False:
        arrEstados.append(estado)
    return arrEstados
def greet():
    print('Hello World!')

def has_numbers(inputString):
    return bool(re.search(r'\d', inputString))

##greet()
files = [f for f in os.listdir() if os.path.isfile(f)]
print(files)
#file = "datos.xlsx"#"enter_path_to_file_here"
reporteTotal = []
reporteGeneral = []
contador = 0
ejeX = []
ejeY = []
for file in files:
    if file.endswith('.xlsx'):
        wb = openpyxl.load_workbook(file, read_only=True)
        ws = wb.active
        anio = extraerAnio()
        print(anio+' extraido')
        reporteArchivo = extraerFechas()
        print( ' fechas extraidas ')
        for dato in reporteArchivo:
            dato.mostrarFechaValor()
        print (' lontitud arreglo archivo '+str(len(reporteArchivo)))
        reporteTotal.append(reporteArchivo)
        print (' lontitud arreglo reporte total '+str(len(reporteTotal)))
        contador= contador +1
        break

for rep in reporteTotal:
    for subRep in rep:
        reporteGeneral.append(rep)
        ejeX.append(subRep.fecha)
        ejeY.append(subRep.valor)

print (' conteo reporte general '+str(len(reporteGeneral)))
#creamos el grafico
colors = cycle(["aqua", "black", "blue", "fuchsia", "gray", "green", "lime", "maroon", "navy", "olive", "purple", "red", "silver", "teal", "yellow"])
plt.xlabel('tiempo')
plt.ylabel('intensidad')
for rep in reporteGeneral:
    estado = rep[0]
    plt.plot(ejeX, ejeY, label="Data " + str(estado.fecha),       color=next(colors))
#plt.plot([1, 2, 3], [1, 4, 9], 'rs', label='line 2')
plt.legend(loc='upper left', fontsize='small')
plt.grid(True)
plt.xlim(0,70)
plt.ylim(0,70)
plt.title('Grafica tiempo/intensidad')
plt.show()

from openpyxl import Workbook
import openpyxl
import re
import os
import matplotlib.pyplot as plt
from itertools import cycle
from dateutil import parser 
import datetime
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
      print(" fecha : "+str(self.fecha)+" valor "+self.valor)

#p1 = EstadoDia("John", 36)
def cambiarFormatoAnio( strEntrada ):
    #format = '%b %d %Y %I:%M%p'
    #datetime_str = datetime.datetime.strptime(date_time, format) 
    #DT = parser.parse("Jun 23 2022 07:31PM") 
    #print(' formato entrada '+strEntrada)
    DT = parser.parse(strEntrada) 
    #datetime.datetime.strptime("21/12/2008", "%d/%m/%Y").strftime("%Y-%m-%d")
    formatoCambiado = datetime.datetime.strptime(strEntrada, "%d/%m/%Y").strftime("%Y-%m-%d")
    #print(' formato cambiado '+formatoCambiado) 
    #print(DT)
    return str(formatoCambiado)
def extraerAnio():
    nombreCelda = "DESDE"
    celdaHasta = "HASTA"
    anioHasta = 0
    mes = 0
    for row in ws.iter_rows(min_row=6,  
                            max_row=8,  
                            min_col=0,  
                            max_col=2):
        for cell in row:
            if cell.value == nombreCelda:
                ##print(ws.cell(row=(cell.row)+1, column=1).value) #change column number for any cell value you want
                mes = (ws.cell(row=(cell.row)+1, column=1).value)[5:7]
                ##print(' mes para extraccion '+mes)
                anio = (ws.cell(row=(cell.row)+1, column=1).value)[0:4]
               
            if cell.value == celdaHasta:
                anioHasta = (ws.cell(row=(cell.row)+1, column=2).value)[0:4]
                print(' anio hasta '+str(anioHasta))
    print(' mes para extraccion '+mes)        
    if( mes == '12' ):
        anio = anioHasta
    if anio is None:
        print("no se encontro nada") 
    return anio
def isDiciembre(fecha):
    #print(' fecha para mes '+fecha)
    mes = fecha.split("/")[1]
    #print(' mes '+mes)
    if( mes == 12 ):
        return True
    return False

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
            if cell.value :    
                fecha = (ws.cell(row=(cell.row), column=1).value)
            if fechaEncontrada and fecha != None and cell.value :
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
                            isDiciembre(fecha)
                            fecha = fecha+"/"+anio
                            fecha = cambiarFormatoAnio(fecha)
                            fecha = fecha[0:10]
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
cambiarFormatoAnio('29/03/2023')
for file in files:
    # se cambia a veces nombre de archivo para hacer pruebas pequeñas
    if file.endswith('.xlsx') or file.endswith('datos.xlsmr'):
        print(' leyendo archivo '+file)
        ##### break para pruebas cortas
        #break
        wb = openpyxl.load_workbook(file, read_only=True)
        ws = wb.active
        anio = extraerAnio()
        print(anio+' extraido')
        ##### break para pruebas cortas
        #break
        reporteArchivo = extraerFechas()
        print( ' fechas extraidas ')
        for dato in reporteArchivo:
            algo = 1
            dato.mostrarFechaValor()
        print (' lontitud arreglo archivo '+str(len(reporteArchivo)))
        reporteTotal.append(reporteArchivo)
        print (' lontitud arreglo reporte total '+str(len(reporteTotal)))
        contador= contador +1
        #break para solo procesar un archivo
        #break

for rep in reporteTotal:
    for subRep in rep:
        reporteGeneral.append(subRep)
        #ejeX.append(subRep.fecha)
        #ejeY.append(int(subRep.valor))

reporteGeneral.sort(key=lambda x: x.fecha, reverse=False)
menoresA10mil = 0
mayoresA10mil = 0
for rep in reporteGeneral:
        ejeX.append(rep.fecha)
        valorNum = int(rep.valor)
        if valorNum < 10000:
            menoresA10mil = menoresA10mil+1
            mayoresA10mil = mayoresA10mil+1
        ejeY.append(valorNum)
print (' conteo reporte general '+str(len(reporteGeneral)))
print(' menores a 10k '+str(menoresA10mil))
print(' mayores a 10k '+str(mayoresA10mil))
#creamos el grafico
colors = cycle(["aqua", "black", "blue", "fuchsia", "gray", "green", "lime", "maroon", "navy", "olive", "purple", "red", "silver", "teal", "yellow"])
plt.xlabel('tiempo')
plt.ylabel('valor')
#subX = ejeX[0:5]
#subY = ejeY[0:5]
subX = ejeX
subY = ejeY
#print(subX)
#print(subY)
print(' min suby '+str(min(subY)))
print(' max suby '+str(max(subY)))
for rep in reporteGeneral:
    estado = rep
    #plt.plot(ejeX, ejeY, label="Data " + str(estado.fecha),       color=next(colors))
    #plt.plot(ejeX, ejeY, 'rs')
    plt.plot(subX, subY, 'rs')
    plt.ylim(min(subY), max(subY))
    #plt.xlim(min(subX), max(subX))
    #plt.plot([1, 2, 3,4], [1, 4, -19,23], 'rs')
plt.legend(loc='upper left', fontsize='small')
plt.grid(True)
#plt.xlim(0,70)
#plt.ylim(0,70)
plt.title('Grafica tiempo/valor')
plt.show()

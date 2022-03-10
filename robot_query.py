"""
made by Oscar Enrique Estrada García
oenriqueg@gmail.com
September 2021

QMF Query Automation Script for invoices

"""

#import modules
#módulos a importar
import pyautogui as robot
import time
import datetime
from win32api import GetKeyState 
from win32con import VK_NUMLOCK 

#variables and constants declaration
#declaración de variables y constantes

#constants
#constantes
robot.FAILSAFE = True
robot.PAUSE = 1

#variables for keyboard
#variables del teclado
numlock_status = GetKeyState(VK_NUMLOCK)
type_speed = 0.050
shorcut_speed = 0.050

#requesting data to open and run QMF app
#variables a solicitar para ejecutar el query en QMF
username = robot.prompt(text="", title="Enter your username: ", default="user")
password = robot.password(text="", title="Enter your password: ", default="password", mask="*")
exchange_rate = robot.prompt(text="", title="Enter exchange rate for this month (MXN to USD): ", default="20.08")
start_date = robot.prompt(text="", title="Enter a start date for data: ", default="20210901")
end_date = robot.prompt(text="", title="Enter a end date for data: ", default="20210930")
start_date_rejected = robot.prompt(text="", title="Ingresa la fecha de 3 meses antes para facturas rechazadas", default="20210601")
today = datetime.datetime.now()
saved_query_path = "C:\\path\\{:}".format(today.strftime("%Y\\%B\\raw\\filename_%d_%B_%Y.xlsx"))
saved_query_rejected_path = "C:\\path\\{:}".format(today.strftime("%Y\\%B\\raw\\filename_%d_%B_%Y_rejected.xlsx"))
template_report_path = "C:\\path\\template.xlsx"
saved_final_report = "C:\\path\\{:}".format(today.strftime("%Y\\%B\\final\\filename_%d_%B_%Y.xlsx"))
qmf_app_path = "C:\\path_app\\qmfdev.exe"


#función para mover el puntero y dar un click
def click(pos,click=1):
    robot.moveTo(pos)
    robot.click(clicks=click)

#inicializar el reloj para medir el tiempo de ejecución al final
runtime_clock_start = time.time()

#Open QMF App
#Abrir QMF
robot.hotkey("winleft","r")
time.sleep(shorcut_speed)
robot.typewrite(qmf_app_path, interval=type_speed)
robot.press("enter")
time.sleep(25)

#Ejecutar Query de facturas
robot.hotkey("altleft","a")
robot.press("enter")
time.sleep(shorcut_speed)
robot.press("1")
time.sleep(0.15)

robot.hotkey("ctrlleft","a")
robot.typewrite(username,interval=type_speed)
robot.hotkey("tab")
time.sleep(0.15)
robot.hotkey("ctrlleft","a")
robot.typewrite(password, interval=type_speed)
robot.hotkey("enter")
time.sleep(5)

robot.hotkey("ctrlleft","r")
time.sleep(1)

robot.hotkey("ctrlleft","a")
robot.typewrite("'" + start_date + "'")
robot.hotkey("tab")
robot.typewrite("'" + end_date + "'")
robot.hotkey("enter")
time.sleep(12)

robot.hotkey("altleft","r")
time.sleep(0.3)
robot.press("e")
robot.hotkey("tab")
robot.hotkey("tab")
robot.hotkey("tab")
robot.hotkey("ctrlleft","a")
robot.press("backspace")
robot.typewrite(saved_query_path, interval=type_speed)
robot.hotkey("enter")
time.sleep(4)

#Comienza query de facturas rechazadas
robot.hotkey("altleft","n")
robot.hotkey("enter")
robot.press("down", presses=19)
robot.hotkey("enter")
robot.hotkey("altleft","n")
robot.hotkey("enter")
robot.press("down", presses=19)
robot.hotkey("enter")
robot.sleep(4)
robot.hotkey("altleft","n")
robot.hotkey("enter")
robot.press("down", presses=19)
robot.hotkey("enter")

robot.press("down", presses=53)
robot.press("right", presses=3)
robot.press("backspace", presses=3)

robot.hotkey("ctrlleft","r")
time.sleep(1)

robot.hotkey("ctrlleft","a")
robot.typewrite("'" + start_date_rejected + "'")
robot.hotkey("tab")
robot.hotkey("ctrlleft","a")
robot.typewrite("'" + end_date + "'")
robot.hotkey("enter")
time.sleep(10)

robot.hotkey("altleft","r")
time.sleep(0.3)
robot.press("e")
robot.hotkey("tab")
robot.hotkey("tab")
robot.hotkey("tab")
robot.hotkey("ctrlleft","a")
robot.press("backspace")
robot.typewrite(saved_query_rejected_path, interval=type_speed)
robot.hotkey("enter")
time.sleep(4)

#Abrimos el primer reporte de facturación y copiamos los valores
robot.hotkey("winleft","r")
time.sleep(shorcut_speed)
robot.hotkey("ctrlleft","a")
robot.press("del")
time.sleep(shorcut_speed)
robot.typewrite(saved_query_path, interval=type_speed)
robot.press("enter")
time.sleep(6)
robot.press("F5")
robot.typewrite("A2")
robot.press("enter")

if numlock_status == 1:
    robot.press("numlock")
else:
    exit

robot.hotkey("ctrlleft","shiftleft","right")
robot.hotkey("ctrlleft","shiftleft","down")
robot.hotkey("ctrlleft","c")

#Abrimos el reporte final, pegamos valores y eliminamos duplicados
robot.hotkey("winleft","r")
time.sleep(shorcut_speed)
robot.hotkey("ctrlleft","a")
robot.press("del")
time.sleep(shorcut_speed)
robot.typewrite(template_report_path, interval=type_speed)
robot.press("enter")
time.sleep(8)
robot.press("F5")
robot.typewrite("C5")
robot.press("enter")
robot.hotkey("ctrlleft","alt","v")
robot.press("down", presses=2)
robot.press("enter")
robot.press("up")
robot.press("down")
robot.hotkey("alt","d","q", interval=shorcut_speed)
robot.press("right")
robot.press("enter")
robot.press("tab", presses=3)
time.sleep(0.1)
robot.press("down", presses=2)
time.sleep(0.1)
robot.hotkey("space")
robot.press("enter")
time.sleep(1)
robot.press("enter")

#Abrtimos el reporte de facturas rechazadas y copiamos valores
robot.hotkey("winleft","r")
time.sleep(shorcut_speed)
robot.hotkey("ctrlleft","a")
robot.press("del")
time.sleep(shorcut_speed)
robot.typewrite(saved_query_rejected_path, interval=type_speed)
robot.press("enter")
time.sleep(6)
robot.press("F5")
robot.typewrite("A2")
robot.press("enter")

"""
if numlock_status == 1:
    robot.press("numlock")
else:
    exit
"""

robot.hotkey("ctrlleft","shiftleft","right")
robot.hotkey("ctrlleft","shiftleft","down")
robot.hotkey("ctrlleft","c")

#Pegamos los valores en el reporte final ya abierto y eliminamos duplicados
robot.hotkey("alt","w","q","2", interval=shorcut_speed)
robot.hotkey("ctrlleft","pagedown")
robot.press("F5")
robot.typewrite("C5")
robot.press("enter")
robot.hotkey("ctrlleft","alt","v")
robot.press("down", presses=2)
robot.press("enter")
robot.press("up")
robot.press("down")
robot.hotkey("alt","d","q", interval=shorcut_speed)
robot.press("right")
robot.press("enter")
robot.press("tab", presses=3)
time.sleep(0.1)
robot.press("down", presses=2)
time.sleep(0.1)
robot.hotkey("space")
robot.press("enter")
time.sleep(1)
robot.press("enter")
robot.press("F5")
robot.typewrite("B2")
robot.press("enter")
robot.typewrite(exchange_rate, interval=type_speed)
robot.press("enter")
robot.hotkey("ctrlleft","pageup")
robot.press("F5")
robot.typewrite("B2")
robot.press("enter")
robot.typewrite(exchange_rate, interval=type_speed)
robot.press("enter")
robot.press("F5")
robot.typewrite("B1")
robot.press("enter")
robot.typewrite(today.strftime("%d/%m/%Y"))
robot.press("enter")
robot.hotkey("ctrlleft","pageup")

#Se actualizan las gráficas del overview
robot.hotkey("alt","d","k","t", interval=shorcut_speed)
robot.hotkey("alt","a","v", interval=shorcut_speed)
time.sleep(1)
robot.hotkey("o")
time.sleep(2)
robot.typewrite(saved_final_report, interval=type_speed)
robot.press("enter")
time.sleep(2)
robot.hotkey("alt","F4")
robot.hotkey("alt","F4")
robot.hotkey("alt","F4")

#Termina la medición del tiempo de ejecución
runtime_clock_end = time.time()
total_runtime = (runtime_clock_end - runtime_clock_start)
print(total_runtime)
robot.press("numlock")
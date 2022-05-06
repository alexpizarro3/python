import win32com.client
import sys
import subprocess
import time
import win32api
import datetime
import calendar
from keyboard import press
import pygetwindow as gw
import win32gui
import os
import pygsheets
import openpyxl
import pandas as pd
import win32com.client as win32

def fechainicio():
    hoy=datetime.datetime.now()
    fechain="01.%s.%s" % (hoy.month, hoy.year)
    print(fechain)
    return(fechain)

def fechafinal():
    hoy=datetime.datetime.now()
    fechafin="%s.%s.%s" % (calendar.monthrange(hoy.year, hoy.month)[1], hoy.month, hoy.year)
    print(fechafin)
    return(fechafin)

class SapGui(): #Crea una clase para Abrir la ruta de SAP Logon y asignar la instancia de la sesion de SAP a self
	def __init__(self):
		self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe" #Ruta del ejecutable de saplogon
		subprocess.Popen(self.path)
		time.sleep(5)

		self.SapGuiAuto = win32com.client.GetObject("SAPGUI") #Instanciar el objeto SAPGUI
		if not type(self.SapGuiAuto)== win32com.client.CDispatch:
			return

		application = self.SapGuiAuto.GetScriptingEngine
		self.connection = application.OpenConnection("1. Grupo Nutresa_ERP_PRD", True) #Conexión asociada
		time.sleep(3)
		self.session = self.connection.Children(0)
		self.session.findById("wnd[0]").maximize

	def sapCostoTot(self):
		try:
				self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
				self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nS_ALR_87013127" #Digita Transacción
				self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
				self.session.findById("wnd[0]").sendVKey(17) #Abre disposiciones Shift F5
				self.session.findById("wnd[1]/usr/txtENAME-LOW").text = "pzadpizarro"
				self.session.findById("wnd[0]").sendVKey(8) #Presiona F8 para ejecutar
				infecha = fechainicio() #Asigna primer dia del mes
				finFecha = fechafinal() #Asigna ultimo día del mes
				self.session.findById("wnd[0]/usr/ctxtRT_GLTRP-LOW").setFocus()
				self.session.findById("wnd[0]/usr/ctxtRT_GLTRP-LOW").text = infecha
				self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
				self.session.findById("wnd[0]/usr/ctxtRT_GLTRP-HIGH").setFocus()
				self.session.findById("wnd[0]/usr/ctxtRT_GLTRP-HIGH").text = finFecha
				self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
				self.session.findById("wnd[0]").sendVKey(8) #Presiona Enter
				self.session.findById("wnd[0]/tbar[1]/btn[45]").press()
				self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
				self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = "\\\Tarcoles\Sim\Ing. Procesos\Archivos Power BI\SAP"
				self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "CostoTotal.txt"
				self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
				self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
				print('Hola')
		    
		except:
				print(sys.exc_info()[0])

SapGui().sapCostoTot() #Ejecuta la clase SapGui y el Metodo sapCostoTot

#Cierra SAP

#subprocess.call(["taskkill","/F","/IM","saplogon.exe"]) #Mata proceso SAPLOGON

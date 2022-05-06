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



def fechainicio(): #Metodo para obtener primer dia del mes actual
		hoy = datetime.datetime.now()
		fechain = "01.%s.%s" % (hoy.month-1, hoy.year)
		print(fechain)
		return (fechain)	

def fechafinal(): #Metodo para obtener ultimo dia del mes
		hoy = datetime.datetime.now()
		fechafin = "%s.%s.%s" % (calendar.monthrange(hoy.year, hoy.month)[1], hoy.month, hoy.year)
		print(fechafin)
		return (fechafin)

	
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

	def sapCooispi(self):
		try:
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nZMM_QRY_1656" #Digita Transacción
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			self.session.findById("wnd[0]").sendVKey(17) #Abre disposiciones
			self.session.findById("wnd[0]").sendVKey(8) #Presiona F8 para ejecutar
			infecha = fechainicio() #Asigna primer dia del mes
			finFecha = fechafinal() #Asigna ultimo día del mes
			self.session.findById("wnd[0]/usr/ctxtFECHACON-LOW").text = infecha
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			self.session.findById("wnd[0]/usr/ctxtFECHACON-HIGH").text = finFecha
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			self.session.findById("wnd[0]").sendVKey(8) #Presiona F8
			self.session.findById("wnd[0]").maximize()  #Maximiza ventana
			self.session.findById("wnd[0]/tbar[1]/btn[45]").press() #Selecciona el boton de exportar como Fichero
			self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()  #Selecciona 4ta opcion clipboard
			self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus() ##Selecciona 4ta opcion clipboard
			self.session.findById("wnd[1]/tbar[0]/btn[0]").press() #Presiona boton check
			print('Hola')
		


		except:
			print(sys.exc_info()[0])


SapGui().sapCooispi() #Ejecuta la clase SapGui y el Metodo sapCooispi

#Abre excel y pega datos

xl=win32.gencache.EnsureDispatch('Excel.Application') #Asigna la aplicación de excel a una variable(Instancia)
wb = xl.Workbooks.Open(r"\\Tarcoles\Sim\Ing. Procesos\Archivos Power BI\SAP\Entradas y Precios de MEMPREP.xlsm") #Abre el libro de excel desde la ruta
time.sleep(3) #Espera 3 segundos
xl.Visible=True #Pone visible el libro
xl.Run("Módulo2.BorrayPega") #Ejecuta macro BorrayPega del libro
time.sleep(3) #Espera 3 Segundos
xl.Run("Módulo2.Ordena")  #Ejecuta macro Ordena del libro
wb.Close(True) #Cierra el libro de excel

#Cierra SAP

subprocess.call(["taskkill","/F","/IM","saplogon.exe"]) #Mata proceso SAPLOGON



			






		
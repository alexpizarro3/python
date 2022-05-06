import win32com.client
import sys
import subprocess
import time
import win32api
import datetime
import calendar
from keyboard import press


def fechainicio(): #Metodo para obtener primer dia del mes actual
		hoy = datetime.datetime.now()
		fechain = "01.%s.%s" % (hoy.month, hoy.year)
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
			self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nCooispi" #Digita cooispi
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			self.session.findById("wnd[0]").sendVKey(17) #Abre disposiciones
			self.session.findById("wnd[0]").sendVKey(8) #Presiona F8 para ejecutar
			self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").pressToolbarButton("&FIND") #Click en boton buscar
			self.session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "ordenessem" 
			self.session.findById("wnd[2]").sendVKey(0) #Presiona enter
			self.session.findById("wnd[2]").sendVKey(12) #Presiona F12 para cerrar ventana
			numfila = self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow #Asigna indice de la fila encontrada
			self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = numfila  #selecciona la fila
			self.session.findById("wnd[1]").sendVKey(2) #Presiona F2 para seleccionar
			self.session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/ordenes_pw" #Digita la disposición
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			infecha = fechainicio() #Asigna primer dia del mes
			finFecha = fechafinal() #Asigna ultimo día del mes
			self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKEN-LOW").text = infecha #Digita el inicio de la fecha
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			self.session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_ECKEN-HIGH").text = finFecha #Digita el fin de fecha
			self.session.findById("wnd[0]").sendVKey(0) #Presiona Enter
			self.session.findById("wnd[0]").sendVKey(8) #Presiona F8
			self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
			self.session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem("&PC")
			self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
			self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
			self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
			self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Prueba1.xls"
			self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
			self.session.findById("wnd[1]/tbar[0]/btn[11]").press()
			


		except:
			print(sys.exc_info()[0])


SapGui().sapCooispi() #Ejecuta la clase SapGui y el Metodo sapCooispi







		
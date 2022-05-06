import win32com.client as win32
import os
import pygsheets
import pandas as pd
import time
#from keyboard import press


#Actualiza consulta de base de datos en excel
xl=win32.gencache.EnsureDispatch('Excel.Application')
path =  os.getcwd().replace('\'','\\') + '\\'
wb = xl.Workbooks.Open(path+'Control de Temperaturas de Horno.xlsm')
#xl.Visible = "true"
time.sleep(10)
#press('enter')
time.sleep(5)
wb.RefreshAll()
time.sleep(10)
wb.Save()
wb.Close()
#xl.Quit()


#autorización
gc = pygsheets.authorize(service_file='C://lotus-notes-cal-sync-144314-02b782326c3b.json')
# Crea un Dataframe en blanco
df = pd.DataFrame()
# Carga los datos del excel al dataframe
df=pd.read_excel('Control de Temperaturas de Horno.xlsm')
#Abre la hoja de google
sh = gc.open('Control temperaturas')
#Selecciona la primer hoja
wks = sh[0]
#Carga los datos del dataframe a la hoja de google 
wks.set_dataframe(df,(1,1))
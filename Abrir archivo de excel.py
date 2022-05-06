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
import os, os.path
import pygsheets
import pandas as pd
import win32com.client as win32

xl=win32.gencache.EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(r"\\Tarcoles\Sim\Ing. Procesos\Archivos Power BI\SAP\Entradas y Precios de MEMPREP.xlsm")
time.sleep(3)
ws = wb.Worksheets(3)
xl.Visible=True
ws.Paste(ws.Range('a1'))
xl.Run("Módulo1.BorraFechas")

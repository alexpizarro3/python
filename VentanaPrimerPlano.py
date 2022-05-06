import win32gui

def loadwindowslist(hwnd, topwindows):
	topwindows.append((hwnd, win32gui.GetWindowText(hwnd)))

def findandshowwindow(swinname, bshow, bbreak):
	topwindows = []
	win32gui.EnumWindows(loadwindowslist, topwindows)
	for hwin in topwindows:	
		sappname=str(hwin[1])
		if swinname in sappname.lower():	
			nhwnd=hwin[0]
			print(">>> Found: " + str(nhwnd) + ": " + sappname)
			if(bshow):
				win32gui.ShowWindow(nhwnd,5)
				win32gui.SetForegroundWindow(nhwnd)
			if(bbreak):
				break

findandshowwindow("Guardar como", True, True)
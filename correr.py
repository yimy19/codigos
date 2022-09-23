import win32gui
 
def loadwindowslist(hwnd, topwindows):
	topwindows.append((hwnd, win32gui.GetWindowText(hwnd)))

def showwindowslist():
	topwindows = []
	win32gui.EnumWindows(loadwindowslist, topwindows)
	for hwin in topwindows:	
		sappname=str(hwin[1])
		nhwnd=hwin[0]
		print(str(nhwnd) + ": " + sappname)

showwindowslist()
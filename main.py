import win32gui
import ctypes
import win32con
import win32api
import PySimpleGUI as sg

user32 = ctypes.windll.user32

def windowsizeupdate(hwnd):
	windowrect = win32gui.GetWindowRect(hwnd)
	#hres = int(windowrect[2] - windowrect[0])
	#window.Element('X').update(int((screenwidth - hres) /2))
	window.Element('X').update(int(windowrect[0]))
	window.Element('Y').update(int(windowrect[1]))
	window.Element('HRes').update(int(windowrect[2] - windowrect[0]))
	window.Element('VRes').update(int(windowrect[3] - windowrect[1]))

def winEnumHandler(hwnd, ctx):
	if win32gui.IsWindowVisible(hwnd):
		n = win32gui.GetWindowText(hwnd)
		s = win32api.GetWindowLong(hwnd, win32con.GWL_STYLE)
		x = win32api.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
		if n:
			allvisiblewindows.update({n: {'hwnd': hwnd, 'STYLE': s, 'EXSTYLE': x}})
			
def aspectratioresize(hwnd, x, y):
	VRes = int(screenheight)-1
	HRes = int(screenheight / y * x)-1
	XPos = int((screenwidth - HRes) /2)
	YPos = int(0)
	user32.SetWindowLongW(hwnd, win32con.GWL_STYLE, win32con.WS_VISIBLE | win32con.WS_CLIPCHILDREN)
	user32.SetWindowLongW(hwnd, win32con.GWL_EXSTYLE, 0)
	user32.MoveWindow(hwnd, XPos, YPos, HRes, VRes, True)
	VRes = VRes + 1
	HRes = HRes + 1
	win32gui.SetWindowPos(hwnd, None, XPos, YPos, HRes, VRes, win32con.SWP_FRAMECHANGED | win32con.SWP_NOZORDER |
						  win32con.SWP_NOOWNERZORDER)
	windowsizeupdate(hwnd)

allvisiblewindows = {}

win32gui.EnumWindows(winEnumHandler, None)

screenwidth = win32api.GetSystemMetrics(0)
screenheight = win32api.GetSystemMetrics(1)

sg.theme('DarkGrey11')
layout = [[sg.Text('Window Selection: ')],
		  [sg.Combo(list(allvisiblewindows.keys()), enable_events=True, key='combo'),sg.Button('Refresh')],
		  [sg.Text("X Pos:"),sg.InputText("840", size=(7,7),key='X'),sg.Text("Y Pos:"),sg.InputText("0", size=(7,7),key='Y')],
		  [sg.Text("H Res:"),sg.InputText("3440", size=(7,7),key='HRes'),sg.Text("V Res:"),sg.InputText("1440", size=(7,7),key='VRes')],
		  [sg.Button('16:9'), sg.Button('21:9'), sg.Button('Borderless Window'), sg.Button('Revert Changes'), sg.Button('Resize/Position', key="Resize"), sg.Exit(key='Cancel')]
		  ]

window = sg.Window('Borderless Window Utility',layout)

while True:
	event, values = window.read()
	if event == sg.WIN_CLOSED or event == 'Cancel':
		break

	if event == 'Borderless Window':
		hwnd = allvisiblewindows.get(values['combo'],{}).get('hwnd')
		try:
			win32gui.GetWindowRect(hwnd)
		except:
			print('Window not found, reloading list')
			allvisiblewindows = {}
			win32gui.EnumWindows(winEnumHandler, None)
			window.Element('combo').update(values=list(allvisiblewindows.keys()))
		else:
			XPos = int(values["X"])
			YPos = int(values["Y"])
			user32.SetWindowLongW(hwnd, win32con.GWL_STYLE, win32con.WS_VISIBLE | win32con.WS_CLIPCHILDREN)
			user32.SetWindowLongW(hwnd, win32con.GWL_EXSTYLE, 0)
			HRes = int(values["HRes"])-1
			VRes = int(values["VRes"])-1
			user32.MoveWindow(hwnd, XPos, YPos, HRes, VRes, True)
			VRes = VRes + 1
			HRes = HRes + 1
			win32gui.SetWindowPos(hwnd, None, XPos, YPos, HRes, VRes, win32con.SWP_FRAMECHANGED | win32con.SWP_NOZORDER |
								  win32con.SWP_NOOWNERZORDER)
			windowsizeupdate(hwnd)
			
	if event == '16:9':
		hwnd = allvisiblewindows.get(values['combo'],{}).get('hwnd')
		try:
			win32gui.GetWindowRect(hwnd)
		except:
			print('Window not found, reloading list')
			allvisiblewindows = {}
			win32gui.EnumWindows(winEnumHandler, None)
			window.Element('combo').update(values=list(allvisiblewindows.keys()))
		else:
			aspectratioresize(hwnd, 16, 9)
			
	if event == '21:9':
		hwnd = allvisiblewindows.get(values['combo'],{}).get('hwnd')
		try:
			win32gui.GetWindowRect(hwnd)
		except:
			print('Window not found, reloading list')
			allvisiblewindows = {}
			win32gui.EnumWindows(winEnumHandler, None)
			window.Element('combo').update(values=list(allvisiblewindows.keys()))
		else:
			aspectratioresize(hwnd, 16, 9)

	if event == 'Revert Changes':
		hwnd = allvisiblewindows.get(values['combo'],{}).get('hwnd')
		try:
			win32gui.GetWindowRect(hwnd)
		except:
			print('Window not found, reloading list')
			allvisiblewindows = {}
			win32gui.EnumWindows(winEnumHandler, None)
			window.Element('combo').update(values=list(allvisiblewindows.keys()))
		else:
			style = allvisiblewindows.get(values['combo'],{}).get('STYLE')
			exstyle = allvisiblewindows.get(values['combo'], {}).get('EXSTYLE')
			hwnd = allvisiblewindows.get(values['combo'], {}).get('hwnd')
			rect = win32gui.GetWindowRect(hwnd)
			user32.SetWindowLongW(hwnd, win32con.GWL_STYLE, style)
			user32.SetWindowLongW(hwnd, win32con.GWL_EXSTYLE, exstyle)
			win32gui.SetWindowPos(hwnd, None, 0, 0, 0, 0, win32con.SWP_FRAMECHANGED | win32con.SWP_NOMOVE |
								  win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOOWNERZORDER)
			windowsizeupdate(hwnd)

	if event == 'Resize':
		hwnd = allvisiblewindows.get(values['combo'],{}).get('hwnd')
		try:
			win32gui.GetWindowRect(hwnd)
		except:
			print('Window not found, reloading list')
			allvisiblewindows = {}
			win32gui.EnumWindows(winEnumHandler, None)
			window.Element('combo').update(values=list(allvisiblewindows.keys()))
		else:
			XPos = int(values["X"])
			YPos = int(values["Y"])
			HRes = int(values["HRes"])
			VRes = int(values["VRes"])
			user32.MoveWindow(hwnd, XPos,YPos, HRes, VRes, True)
			windowsizeupdate(hwnd)

	if event == 'Refresh':
		allvisiblewindows = {}
		win32gui.EnumWindows(winEnumHandler, None)
		window.Element('combo').update(values=list(allvisiblewindows.keys()))

	if event == 'combo':
		hwnd = allvisiblewindows.get(values['combo'], {}).get('hwnd')
		try:
			win32gui.GetWindowRect(hwnd)
		except:
			print('Window not found, reloading list')
			allvisiblewindows = {}
			win32gui.EnumWindows(winEnumHandler, None)
			window.Element('combo').update(values=list(allvisiblewindows.keys()))
		else:
			windowsizeupdate(hwnd)
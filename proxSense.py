import serial
import io
import time
import winsound
import wx
import thread
import warnings
import json
import os.path
warnings.simplefilter('ignore')

### Setup config parameter defaults
serialPort="COM10"
buadRate=57600
serialTimeout=30
panelBgColor="black"
mapImage="storemap.jpg"
serialDataLength=4
serialLineLength=8
triggerDistance=120
triggerHoldTime=0
displayFullScreen=True
displayWidth=0
displayHeight=0
displayPosX=0
displayPosY=0
## End config parameter defaults	

##### Begin config file loading
try:
	with open(os.environ['ProgramData'] + "\\proxSense\config.json",'r') as fp:
		config=json.load(fp)
	serialPort=config["serialPort"]
	buadRate=config["buadRate"]
	serialTimeout=config["serialTimeout"]
	panelBgColor=config["panelBgColor"]
	mapImage=config["mapImage"]
	serialDataLength=config["serialDataLength"]
	serialLineLength=config["serialLineLength"]
	triggerDistance=config["triggerDistance"]
	triggerHoldTime=config["triggerHoldTime"]
	displayFullScreen=config["displayFullScreen"]
	displayWidth=config["displayWidth"]
	displayHeight=config["displayHeight"]
	displayPosX=config["displayPosX"]
	displayPosY=config["displayPosY"]
except ValueError:
	print "Error parsing config file"
except Exception:
	pass
##### End config file loading

##### Begin win32 API code for hiding taskbar
import ctypes
from ctypes import wintypes

taskBarVis=1


FindWindow = ctypes.windll.user32.FindWindowA
FindWindow.restype = wintypes.HWND
FindWindow.argtypes = [
	wintypes.LPCSTR, #lpClassName
	wintypes.LPCSTR, #lpWindowName
]

SetWindowPos = ctypes.windll.user32.SetWindowPos
SetWindowPos.restype = wintypes.BOOL
SetWindowPos.argtypes = [
	wintypes.HWND, #hWnd
	wintypes.HWND, #hWndInsertAfter
	ctypes.c_int,  #X
	ctypes.c_int,  #Y
	ctypes.c_int,  #cx
	ctypes.c_int,  #cy
	ctypes.c_uint, #uFlags
] 

FindWindowEx = ctypes.windll.user32.FindWindowExA
FindWindowEx.restype = wintypes.HWND
FindWindowEx.argtypes = [
	wintypes.HWND, # hwndParent
	wintypes.HWND, # hwndChildAfter
	wintypes.LPCSTR, # lpszClass
	wintypes.LPCSTR, # lpszWindow
]

start_atom = wintypes.LPCSTR(0xc017)
TOGGLE_HIDEWINDOW = 0x80
TOGGLE_UNHIDEWINDOW = 0x40


def hide_taskbar():
	global taskBarVis
	handleW1 = FindWindow(b"Shell_traywnd", b"")
	SetWindowPos(handleW1, 0, 0, 0, 0, 0, TOGGLE_HIDEWINDOW)
	hStart = FindWindowEx(None, None, start_atom, None)
	SetWindowPos(hStart, 0, 0, 0, 0, 0, TOGGLE_HIDEWINDOW)
	taskBarVis=0

def unhide_taskbar():
	global taskBarVis
	handleW1 = FindWindow(b"Shell_traywnd", b"")
	SetWindowPos(handleW1, 0, 0, 0, 0, 0, TOGGLE_UNHIDEWINDOW)
	hStart = FindWindowEx(None, None, start_atom, None)
	SetWindowPos(hStart, 0, 0, 0, 0, 0, TOGGLE_UNHIDEWINDOW)
	taskBarVis=1

def toggleTaskBar(event):
	if taskBarVis:
		hide_taskbar()
	else:
		unhide_taskbar()
		
	
##### End win32 API code for hiding taskbar





### Begin Main Code


# Serial port runs at 57600 baud, 8N1
# We should figure out a way to auto search to grab the right port
ser=serial.Serial(serialPort, buadRate, timeout=serialTimeout)

# Load the image used for the map.  This will be later put into a UI object and scaled.
mapImg=wx.Image(mapImage,wx.BITMAP_TYPE_ANY)

def adjustLayout():
	frame.SetSizeWH(displayWidth,displayHeight)
	panel.SetSizeWH(panel.GetClientSize().x,panel.GetClientSize().y)
	frame.Move((displayPosX, displayPosY))

def dataProcessing():
	# This is the processing function that will be spun off in a separate thread.
	# This thread is what triggers the UI events
	lastPerson=0
	distance=500
	person=False

	while True:
		dataline=ser.read(serialLineLength)
		# Parse the incoming data.  Format comes in as "R<distance inches> P<1 or 0 is there a person>"
		# Depending on the sensor, there may be 3 or 4 digits printed
		#print len(dataline)
		if len(dataline) >= serialDataLength:
			if dataline[0]=="R" and dataline[1:serialDataLength].isdigit():
					distance=int(dataline[1:serialDataLength])
		#print dataline[1:serialDataLength]	
		
		# If there is a person that was detected on the last loop, have they moved away for long enough to take action?
		# We want to limit it flapping back and forth
		if person:
			if distance > triggerDistance and (time.time() - lastPerson) >= triggerHoldTime:
				winsound.Beep(2000,500)
				print "No more person"
				# Hide map
				wx.CallAfter(frame.Show,False)
				#wx.CallAfter(unhide_taskbar)
				person=False
		else:		
			if distance <= triggerDistance:
				winsound.Beep(2500,500)
				print "PERSON!"
				# Show the map
				wx.CallAfter(frame.Show,True)
				# Fix some buginess with screen refreshes
				wx.CallAfter(frame.Iconize,True)
				wx.CallAfter(frame.Iconize,False)
				
				wx.CallAfter(hide_taskbar)
				
				person=True
				lastPerson=time.time()
			

	
#Begin GUI
app = wx.PySimpleApp()
# Get our screen resolution
screenX = wx.SystemSettings_GetMetric(wx.SYS_SCREEN_X)
screenY = wx.SystemSettings_GetMetric(wx.SYS_SCREEN_Y)

if(displayFullScreen):
	displayWidth=screenX
	displayHeight=screenY
	displayPosX=0
	displayPosY=0
	

# create a window/frame, no parent, -1 is default ID
# Size it based on the screen resolution
frame = wx.Frame(None, -1, "Store Map", size = (displayWidth,displayHeight), style=wx.SYSTEM_MENU|wx.MINIMIZE_BOX|wx.MAXIMIZE_BOX|wx.CLOSE_BOX|wx.STAY_ON_TOP)
# Create a panel inside the frame
panel = wx.Panel(frame,-1)

# Color for any part of the panel that doesn't get covered.  Normally shouldn't see this but black can hide any minor
# sizing inaccuracies
panel.SetBackgroundColour(panelBgColor)

# Scale the image so that it fits our resolution
mapImg=mapImg.Scale(displayWidth,displayHeight)

# Create the image control
mapImgCtrl = wx.StaticBitmap(panel, -1, wx.BitmapFromImage(mapImg))

# Redraw and adjust all our UI elements so they are positioned correctly
adjustLayout()

bgThread = thread.start_new_thread(dataProcessing, ())


# Register global hotkey for hiding/showing the taskbar (ALT-T)
frame.RegisterHotKey(100,wx.MOD_ALT,84)
# Bind hotkey event handler
frame.Bind(wx.EVT_HOTKEY, toggleTaskBar, None, 100)


def appClean(event):
	unhide_taskbar()
	wx.CallAfter(taskBarIcon.Destroy)
	wx.CallAfter(frame.Destroy)
	#frame.Destroy()
	app.Exit()
	#bgThread.exit()
	

def create_menu_item(menu, label, func):
	item = wx.MenuItem(menu, -1, label)
	menu.Bind(wx.EVT_MENU, func, id=item.GetId())
	menu.AppendItem(item)
	return item


class TaskBarIcon(wx.TaskBarIcon):
	def __init__(self):
		super(TaskBarIcon, self).__init__()
		self.set_icon('icon.png')

	def CreatePopupMenu(self):
		menu = wx.Menu()
		create_menu_item(menu, 'Exit', self.on_exit)
		return menu

	def set_icon(self, path):
		icon = wx.IconFromBitmap(wx.Bitmap(path))
		self.SetIcon(icon, "Map Popper")

	def on_exit(self, event):
		#unhide_taskbar()
		wx.CallAfter(appClean,None)
		
taskBarIcon=TaskBarIcon()

frame.Bind(wx.EVT_CLOSE, appClean)


# start the event loop
app.MainLoop()


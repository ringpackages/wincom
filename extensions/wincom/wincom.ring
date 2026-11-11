/*
	WinCOM - Windows COM (Component Object Model) for the Ring Programming Language
	Copyright (c) 2026 Mahmoud Fayed <msfclipper@yahoo.com>

	The com_*() functions (written in C) use COM Automation (IDispatch)
	So we can use ADO, ADOX, Scripting, WScript, WMI, Microsoft Office, ... from Ring
	The COMObject and COMEvents classes provide the same features using Object-Oriented Programming
*/

# Load the extension (ring_wincom.c)
if isWindows()
	LoadLib("ring_wincom.dll")
else
	raise("Error, WinCOM: The COM technology is available on MS-Windows only")
ok

load "globals.ring"
load "functions.ring"
load "comobject.ring"
load "comevents.ring"

# Asynchronous WMI events: know when processes start and stop
# WbemScripting.SWbemSink receives the events, and com_processmessages() delivers them

load "wincom.ring"

nStarted = 0
nStopped = 0

func main
	oWMI = com_getobject("winmgmts:")

	# The sink object raises OnObjectReady for each event
	oStartSink = com_createobject("WbemScripting.SWbemSink")
	oStartEvents = com_connectevents(oStartSink)
	com_onevent(oStartEvents, "OnObjectReady", "onStarted()")

	oStopSink = com_createobject("WbemScripting.SWbemSink")
	oStopEvents = com_connectevents(oStopSink)
	com_onevent(oStopEvents, "OnObjectReady", "onStopped()")

	# WMI checks the processes every second (WITHIN 1)
	cProcess = "TargetInstance ISA 'Win32_Process' AND TargetInstance.Name = 'PING.EXE'"
	com_invoke(oWMI, "ExecNotificationQueryAsync", oStartSink,
		"SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE " + cProcess)
	com_invoke(oWMI, "ExecNotificationQueryAsync", oStopSink,
		"SELECT * FROM __InstanceDeletionEvent WITHIN 1 WHERE " + cProcess)

	# Start some processes (hidden windows), then wait for the events
	? "Starting 3 processes (ping.exe)..."
	oShell = com_createobject("WScript.Shell")
	for nIndex = 1 to 3
		com_invoke(oShell, "Run", "ping -n 3 127.0.0.1", 0, com_bool(0))
		com_processmessages(500)
	next
	# Wait (up to 10 seconds) while the events are delivered
	for nIndex = 1 to 20
		com_processmessages(500)
		if nStopped >= 3
			exit
		ok
	next

	# Stop receiving the events
	com_invoke(oStartSink, "Cancel")
	com_invoke(oStopSink, "Cancel")
	? nl + "Processes started: " + nStarted + ", stopped: " + nStopped

func onStarted
	# OnObjectReady(objWbemObject, objWbemAsyncContext)
	oEvent = com_eventargs()[1]
	? "Started : " + com_get(oEvent, "TargetInstance.Name") + " (Process ID: " +
	  com_get(oEvent, "TargetInstance.ProcessId") + ")"
	nStarted++

func onStopped
	oEvent = com_eventargs()[1]
	? "Stopped : " + com_get(oEvent, "TargetInstance.Name") + " (Process ID: " +
	  com_get(oEvent, "TargetInstance.ProcessId") + ")"
	nStopped++

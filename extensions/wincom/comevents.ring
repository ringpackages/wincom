/*
	WinCOM - COMEvents Class
	Receiving the events of a COM object (Connection Points)

	oEvents = new COMEvents(oConnection)                       # The default events interface
	oEvents = new COMEvents([oWorkbook, "WorkbookEvents"])     # A specific events interface
	oEvents.onEvent("WillExecute", "onWillExecute()")

	The handler is Ring code, it can use com_eventname(), com_eventargs(),
	com_seteventarg() and com_seteventresult()
*/

class COMEvents

	pEvents = NULL

	/*
		vValue is a COM object (C pointer or COMObject), or a list: [oObject, cInterface]
		cInterface is the name of the events interface (like "ConnectionEvents") or its IID ("{...}")
	*/
	func init vValue
		cInterface = ""
		vObject = vValue
		if isList(vValue) and ! com_isobject(vValue)
			if len(vValue) != 2 or ! isString(vValue[2])
				raise("Error, WinCOM: COMEvents expects a COM object or [oObject, cInterface]")
			ok
			vObject = vValue[1]
			cInterface = vValue[2]
		ok
		if isObject(vObject)
			if lower(classname(vObject)) = "comobject"
				vObject = vObject.pointer()
			ok
		ok
		if ! com_isobject(vObject)
			raise("Error, WinCOM: COMEvents expects a COM object or [oObject, cInterface]")
		ok
		if cInterface = ""
			pEvents = com_connectevents(vObject)
		else
			pEvents = com_connectevents(vObject, cInterface)
		ok
		return self

	/*
		Run the Ring code when the event happens ("*" for all the events, "" code removes the handler)
	*/
	func onEvent cEvent, cCode
		com_onevent(pEvents, cEvent, cCode)
		return self

	/*
		The names of the events
	*/
	func names
		return com_eventnames(pEvents)

	/*
		Stop receiving the events
	*/
	func disconnect
		com_disconnectevents(pEvents)

	/*
		The events object (C pointer), used by the com_*() functions
	*/
	func pointer
		return pEvents

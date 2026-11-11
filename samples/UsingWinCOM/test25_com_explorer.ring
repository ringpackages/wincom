# COM Explorer: print the members of a COM object (methods, properties and events)
# Usage: ring test25_com_explorer.ring [ProgID]
# Example: ring test25_com_explorer.ring Scripting.FileSystemObject

load "wincom.ring"

func main
	cProgID = "Scripting.Dictionary"
	aArgs = sysargv
	if len(aArgs) >= 3
		cProgID = aArgs[3]
	ok
	if ! com_progidexists(cProgID)
		? "The ProgID is not registered: " + cProgID
		return
	ok
	oObject = com_createobject(cProgID)
	? "ProgID    : " + cProgID
	? "Interface : " + com_typename(oObject)

	# Group the property get/put members
	aMembers = com_members(oObject)
	aProperties = []
	aMethods = []
	for aMember in aMembers
		switch aMember[2]
			on "method"
				aMethods + aMember
			other
				nPos = find(aProperties, aMember[1], 1)
				if nPos = 0
					aProperties + [aMember[1], [aMember[2]], aMember[4], aMember[5]]
				else
					aProperties[nPos][2] + aMember[2]
				ok
		off
	next

	? nl + "Properties (" + len(aProperties) + "):"
	aProperties = sort(aProperties, 1)
	for aProperty in aProperties
		cLine = "	" + aProperty[1]
		if len(aProperty[4]) > 0
			cLine += "(" + joinList(aProperty[4]) + ")"
		ok
		cLine += "  [" + joinList(aProperty[2]) + "]"
		? cLine
		if aProperty[3] != ""
			? "		" + aProperty[3]
		ok
	next

	? nl + "Methods (" + len(aMethods) + "):"
	aMethods = sort(aMethods, 1)
	for aMethod in aMethods
		? "	" + aMethod[1] + "(" + joinList(aMethod[5]) + ")"
		if aMethod[4] != ""
			? "		" + aMethod[4]
		ok
	next

	# Events (if the object has events)
	try
		oEvents = com_connectevents(oObject)
		aEvents = com_eventnames(oEvents)
		com_disconnectevents(oEvents)
		? nl + "Events (" + len(aEvents) + "):"
		for cEvent in aEvents
			? "	" + cEvent
		next
	catch
		? nl + "Events: none"
	done

func joinList aList
	cOut = ""
	for nIndex = 1 to len(aList)
		cOut += aList[nIndex]
		if nIndex < len(aList)
			cOut += ", "
		ok
	next
	return cOut

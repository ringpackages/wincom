# Using the Windows Shell (Shell.Application): folders, file details and shell verbs

load "wincom.ring"

func main
	oShell = com_createobject("Shell.Application")

	# A folder (Namespace) and its items
	oFolder = com_invoke(oShell, "Namespace", currentdir())
	aItems = com_items(com_invoke(oFolder, "Items"))
	? "Items in the samples folder: " + len(aItems)

	# The columns that Windows Explorer shows (GetDetailsOf(NULL, nColumn) is the column name)
	aColumns = []
	for nColumn = 0 to 3
		aColumns + com_invoke(oFolder, "GetDetailsOf", com_null(), nColumn)
	next
	? "Columns: " + aColumns[1] + ", " + aColumns[2] + ", " + aColumns[3] + ", " + aColumns[4]

	# The details of some files (Name, Size, Type)
	? nl + "Ring files:"
	nShown = 0
	for oItem in aItems
		cName = com_get(oItem, "Name")
		if com_get(oItem, "IsFolder") or (right(lower(com_get(oItem, "Path")), 5) != ".ring")
			loop
		ok
		? "	" + left(cName + copy(" ", 36), 36) + left(com_invoke(oFolder, "GetDetailsOf", oItem, 1) + copy(" ", 12), 12) +
		  com_invoke(oFolder, "GetDetailsOf", oItem, 2)
		nShown++
		if nShown = 5
			exit
		ok
	next

	# One item using ParseName(), and its properties (ExtendedProperty)
	oItem = com_invoke(oFolder, "ParseName", "test01_scripting_dictionary.ring")
	? nl + "test01_scripting_dictionary.ring:"
	? "	Size (bytes) : " + com_invoke(oItem, "ExtendedProperty", "System.Size")
	? "	Item type    : " + com_invoke(oItem, "ExtendedProperty", "System.ItemTypeText")

	# The verbs of the item (Like the menu of the right mouse button in Windows Explorer)
	aVerbs = com_items(com_invoke(oItem, "Verbs"))
	cVerbs = ""
	for oVerb in aVerbs
		cVerb = substr(com_get(oVerb, "Name"), "&", "")
		if cVerb != ""
			cVerbs += cVerb + " | "
		ok
	next
	? "	Verbs        : " + left(cVerbs, len(cVerbs) - 3)

	# A special folder: 36 (ssfWINDOWS) is the Windows folder
	oWindows = com_invoke(oShell, "Namespace", 36)
	? nl + "Windows folder: " + com_get(oWindows, "Self.Path")
	? "Has explorer.exe: " + com_isobject(com_invoke(oWindows, "ParseName", "explorer.exe"))

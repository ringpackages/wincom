# System information using WMI (Windows Management Instrumentation)
# com_getobject() gets the WMI service using the "winmgmts:" moniker

load "wincom.ring"

func main
	oWMI = com_getobject("winmgmts:root\cimv2")

	# ExecQuery() returns a collection of objects
	aSystems = com_items(com_invoke(oWMI, "ExecQuery", "SELECT * FROM Win32_OperatingSystem"))
	for oOS in aSystems
		? "Operating system : " + com_get(oOS, "Caption")
		? "Version          : " + com_get(oOS, "Version")
		? "Architecture     : " + com_get(oOS, "OSArchitecture")
		? "Free memory      : " + floor(com_get(oOS, "FreePhysicalMemory") / 1024) + " MB"
	next

	aProcessors = com_items(com_invoke(oWMI, "ExecQuery", "SELECT Name, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor"))
	for oCPU in aProcessors
		? "Processor        : " + trim(com_get(oCPU, "Name"))
		? "Cores            : " + com_get(oCPU, "NumberOfCores") + " (Logical processors: " + com_get(oCPU, "NumberOfLogicalProcessors") + ")"
	next

	# DriveType = 3 : Local disks
	? nl + "Local disks:"
	aDisks = com_items(com_invoke(oWMI, "ExecQuery", "SELECT DeviceID, FileSystem, Size, FreeSpace FROM Win32_LogicalDisk WHERE DriveType = 3"))
	for oDisk in aDisks
		# Size and FreeSpace are strings (64-bit numbers)
		nSize = 0 + com_get(oDisk, "Size")
		nFree = 0 + com_get(oDisk, "FreeSpace")
		? "	" + com_get(oDisk, "DeviceID") + " " + com_get(oDisk, "FileSystem") + " - Size: " +
		  floor(nSize / pow(1024, 3)) + " GB, Free: " + floor(nFree / pow(1024, 3)) + " GB"
	next

	# The properties of a WMI object are in the Properties_ collection
	? nl + "Properties of Win32_Processor:"
	aProperties = com_items(com_get(aProcessors[1], "Properties_"))
	cNames = ""
	for oProperty in aProperties
		cNames += com_get(oProperty, "Name") + " "
	next
	? cNames

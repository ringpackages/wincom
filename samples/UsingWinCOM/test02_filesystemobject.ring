# Using Scripting.FileSystemObject
# Paths: com_get(oObject, 'GetFolder("C:\").Name') calls GetFolder() then gets Name

load "wincom.ring"

func main
	oFSO = com_createobject("Scripting.FileSystemObject")

	? "Drives:"
	aDrives = com_items(com_get(oFSO, "Drives"))
	for oDrive in aDrives
		cLine = "	" + com_get(oDrive, "DriveLetter") + ": "
		if com_get(oDrive, "IsReady")
			cLine += "Free space: " + floor(com_get(oDrive, "FreeSpace") / (1024*1024*1024)) + " GB"
		else
			cLine += "Not ready"
		ok
		? cLine
	next

	# Create a folder and a Unicode text file
	cFolder = currentdir() + "\FSO Тест"
	if ! com_get(oFSO, "FolderExists", cFolder)
		com_invoke(oFSO, "CreateFolder", cFolder)
	ok
	cFile = com_invoke(oFSO, "BuildPath", cFolder, "Файл.txt")

	# CreateTextFile(FileName, Overwrite, Unicode)
	oStream = com_invoke(oFSO, "CreateTextFile", cFile, com_bool(1), com_bool(1))
	com_invoke(oStream, "WriteLine", "Line 1 - Hello")
	com_invoke(oStream, "WriteLine", "Line 2 - Привет")
	com_invoke(oStream, "WriteLine", "Line 3 - مرحبا")
	com_invoke(oStream, "Close")

	# OpenTextFile(FileName, IOMode (1 = ForReading), Create, Format (-1 = Unicode))
	? nl + "File content:"
	oStream = com_invoke(oFSO, "OpenTextFile", cFile, 1, com_bool(0), -1)
	while ! com_get(oStream, "AtEndOfStream")
		? "	" + com_invoke(oStream, "ReadLine")
	end
	com_invoke(oStream, "Close")

	? nl + "File size : " + com_get(oFSO, 'GetFile("' + cFile + '").Size') + " bytes"
	? "File name : " + com_invoke(oFSO, "GetFileName", cFile)
	? "Extension : " + com_invoke(oFSO, "GetExtensionName", cFile)

	? nl + "Files in the folder:"
	# Store the list in a variable (for ... in evaluates the expression in each iteration)
	aFiles = com_items(com_get(oFSO, 'GetFolder("' + cFolder + '").Files'))
	for oFile in aFiles
		? "	" + com_get(oFile, "Name") + " (" + com_get(oFile, "Type") + ")"
	next

	# DeleteFolder(FolderSpec, Force)
	com_invoke(oFSO, "DeleteFolder", cFolder, com_bool(1))
	lExists = com_get(oFSO, "FolderExists", cFolder)
	? nl + "Folder exists after DeleteFolder(): " + lExists

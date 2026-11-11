# Store files (binary data) in a Microsoft Access table, then read them again
# Binary values are Ring strings: com_bytes() writes them and Field.Value returns them

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_files_blob")
	com_invoke(oConnection, "Execute", "CREATE TABLE Files (ID COUNTER PRIMARY KEY, Name TEXT(255), FileSize LONG, Content LONGBINARY)")

	# The files: some sample files and binary data that has all the byte values (0 to 255)
	cAllBytes = ""
	for nByte = 0 to 255
		cAllBytes += char(nByte)
	next
	aFiles = [
		["accesshelper.ring", read("accesshelper.ring")],
		["test01_scripting_dictionary.ring", read("test01_scripting_dictionary.ring")],
		["allbytes.bin", copy(cAllBytes, 1000)]
	]

	# Add the files (AddNew with the field names and the values)
	oRecordset = com_createobject("ADODB.Recordset")
	com_invoke(oRecordset, "Open", "Files", oConnection, adOpenKeyset, adLockOptimistic, adCmdTable)
	for aFile in aFiles
		com_invoke(oRecordset, "AddNew", ["Name", "FileSize", "Content"],
			   [aFile[1], len(aFile[2]), com_bytes(aFile[2])])
	next
	com_invoke(oRecordset, "Close")

	? "Files in the table:" + nl
	oRecordset = com_invoke(oConnection, "Execute", "SELECT ID, Name, FileSize FROM Files ORDER BY ID")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# Read the files again and compare them with the original data
	? nl + "Read and compare:"
	oRecordset = com_invoke(oConnection, "Execute", "SELECT Name, Content FROM Files ORDER BY ID")
	nIndex = 0
	while ! com_get(oRecordset, "EOF")
		nIndex++
		cName = com_get(oRecordset, 'Fields("Name").Value')
		cContent = com_get(oRecordset, 'Fields("Content").Value')
		cResult = "Different"
		if cContent = aFiles[nIndex][2]
			cResult = "Same"
		ok
		? "	" + cName + " : " + len(cContent) + " bytes, " + cResult
		com_invoke(oRecordset, "MoveNext")
	end
	com_invoke(oRecordset, "Close")

	# Read a part of the data using GetChunk() (useful for large files)
	oRecordset = com_invoke(oConnection, "Execute", "SELECT Content FROM Files WHERE Name = 'allbytes.bin'")
	cPart = com_invoke(oRecordset, "Fields(0).GetChunk", 8)
	cHex = ""
	for nIndex = 1 to len(cPart)
		cHex += right("0" + upper(hex(ascii(cPart[nIndex]))), 2) + " "
	next
	? nl + "The first 8 bytes of allbytes.bin (GetChunk): " + cHex
	com_invoke(oRecordset, "Close")

	# Save a file from the table to the disk using ADODB.Stream
	oRecordset = com_invoke(oConnection, "Execute", "SELECT Content FROM Files WHERE Name = 'allbytes.bin'")
	oStream = com_createobject("ADODB.Stream")
	com_set(oStream, "Type", adTypeBinary)
	com_invoke(oStream, "Open")
	com_invoke(oStream, "Write", com_bytes(com_get(oRecordset, "Fields(0).Value")))
	com_invoke(oStream, "SaveToFile", currentdir() + "\allbytes.bin", adSaveCreateOverWrite)
	com_invoke(oStream, "Close")
	com_invoke(oRecordset, "Close")
	? "allbytes.bin saved: " + getfilesize("allbytes.bin") + " bytes, same as the original: " +
	  (read("allbytes.bin") = aFiles[3][2])
	remove("allbytes.bin")

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_files_blob")

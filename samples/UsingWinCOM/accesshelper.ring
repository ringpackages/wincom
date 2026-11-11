# Helper functions used by the Microsoft Access samples (ADO + ADOX)
#
# The samples need one of these OLE DB providers:
#	Microsoft.ACE.OLEDB.16.0 or Microsoft.ACE.OLEDB.12.0 (.accdb files)
#		Installed with Microsoft Office or the "Microsoft Access Database Engine Redistributable"
#		Ring 64-bit needs the 64-bit engine, Ring 32-bit needs the 32-bit engine
#	Microsoft.Jet.OLEDB.4.0 (.mdb files)
#		Part of MS-Windows, but available for 32-bit programs only (Ring 32-bit)

ACCESS_PROVIDERS = [
	["Microsoft.ACE.OLEDB.16.0", ".accdb"],
	["Microsoft.ACE.OLEDB.12.0", ".accdb"],
	["Microsoft.Jet.OLEDB.4.0",  ".mdb"]
]

# The provider and the file extension selected by accessCreateDatabase()
cAccessProvider  = ""
cAccessExtension = ""

/*
	Create a new database (cName without extension), returns ADODB.Connection (open)
	An old database with the same name is deleted
*/
func accessCreateDatabase cName
	nCount = len(ACCESS_PROVIDERS)
	cErrors = ""
	for nIndex = 1 to nCount
		cProvider = ACCESS_PROVIDERS[nIndex][1]
		if ! com_progidexists(cProvider)
			loop
		ok
		cFile = cName + ACCESS_PROVIDERS[nIndex][2]
		if fexists(cFile)
			remove(cFile)
		ok
		try
			oCatalog = com_createobject("ADOX.Catalog")
			# Catalog.Create() creates the file and returns the open connection
			oConnection = com_invoke(oCatalog, "Create", accessConnectionString(cProvider, cFile))
			cAccessProvider  = cProvider
			cAccessExtension = ACCESS_PROVIDERS[nIndex][2]
			return oConnection
		catch
			cErrors += nl + "	" + cCatchError
		done
	next
	? "Can't create the Microsoft Access database" + cErrors
	? "Install the Microsoft Access Database Engine (Use the 64-bit version with Ring 64-bit)"
	shutdown(1)

/*
	Load the database engine now (Without opening a database)
	Needed with RingQt in Ring versions older than 1.28 (Call it before load "guilib.ring"): the Ring folder
	has an older version of the Visual C++ runtime (msvcp140.dll) that Qt loads, and the Access Database Engine
	(Microsoft Office) can't use it. When the engine is loaded first, it loads the newer version from the system
	Ring 1.28 includes a newer runtime, so the function isn't needed there (but it's harmless)
*/
func accessLoadEngine
	nCount = len(ACCESS_PROVIDERS)
	for nIndex = 1 to nCount
		cProvider = ACCESS_PROVIDERS[nIndex][1]
		if com_progidexists(cProvider)
			oConnection = com_createobject("ADODB.Connection")
			try
				# The file doesn't exist, so we get an error after loading the engine
				com_invoke(oConnection, "Open", accessConnectionString(cProvider, "accesshelper_no_file" + ACCESS_PROVIDERS[nIndex][2]))
			catch
			done
			return
		ok
	next

/*
	Open an existing database (Created by accessCreateDatabase())
*/
func accessOpenDatabase cName
	oConnection = com_createobject("ADODB.Connection")
	com_invoke(oConnection, "Open", accessConnectionString(cAccessProvider, cName + cAccessExtension))
	return oConnection

func accessConnectionString cProvider, cFile
	return "Provider=" + cProvider + ";Data Source=" + cFile

func accessFileName cName
	return cName + cAccessExtension

/*
	Delete the database file
*/
func accessDeleteDatabase cName
	cFile = cName + cAccessExtension
	if fexists(cFile)
		remove(cFile)
	ok

/*
	Create the Products table and add sample data
*/
func accessCreateProducts oConnection
	com_invoke(oConnection, "Execute", "
		CREATE TABLE Products (
			ID       COUNTER PRIMARY KEY,
			Name     TEXT(50) NOT NULL,
			Category TEXT(30),
			Price    CURRENCY,
			Quantity INTEGER,
			Added    DATETIME,
			InStock  YESNO,
			Notes    MEMO
		)
	")
	aProducts = [
		["Product A", "Category 1", 10.5,  100, "2026-01-05", 1, "First product"],
		["Product B", "Category 1", 25,    0,   "2026-01-12", 0, NULL],
		["Product C", "Category 2", 7.25,  42,  "2026-02-03", 1, "Fragile"],
		["Product D", "Category 2", 99.99, 5,   "2026-03-18", 1, NULL],
		["Product E", "Category 3", 3,     250, "2026-04-22", 1, "Sold in packs"]
	]
	# Add the records using Recordset.AddNew(FieldList, Values)
	oRecordset = com_createobject("ADODB.Recordset")
	com_invoke(oRecordset, "Open", "Products", oConnection, adOpenKeyset, adLockOptimistic, adCmdTable)
	aFields = ["Name", "Category", "Price", "Quantity", "Added", "InStock", "Notes"]
	for aProduct in aProducts
		# Dates, Booleans and Null are created using com_date(), com_bool() and com_null()
		aValues = [aProduct[1], aProduct[2], aProduct[3], aProduct[4],
			   com_date(aProduct[5]), com_bool(aProduct[6]), com_null()]
		if aProduct[7] != NULL
			aValues[7] = aProduct[7]
		ok
		com_invoke(oRecordset, "AddNew", aFields, aValues)
	next
	com_invoke(oRecordset, "Close")

/*
	Print the records of an open Recordset as a table
*/
func accessPrintRecordset oRecordset
	nFields = com_get(oRecordset, "Fields.Count")
	# Read the field names and the records as strings
	aHeader = []
	for nField = 0 to nFields - 1
		aHeader + com_get(oRecordset, "Fields(" + nField + ").Name")
	next
	aRows = []
	while ! com_get(oRecordset, "EOF")
		aRow = []
		for nField = 0 to nFields - 1
			vValue = com_get(oRecordset, "Fields(" + nField + ").Value")
			if com_wasnull()
				vValue = "(Null)"
			ok
			aRow + ("" + vValue)
		next
		aRows + aRow
		com_invoke(oRecordset, "MoveNext")
	end
	# The width of each column
	aWidths = []
	for nField = 1 to nFields
		nWidth = accessCharCount(aHeader[nField])
		for aRow in aRows
			nWidth = max(nWidth, accessCharCount(aRow[nField]))
		next
		aWidths + nWidth
	next
	? accessFormatRow(aHeader, aWidths)
	cLine = ""
	for nField = 1 to nFields
		cLine += copy("-", aWidths[nField]) + "-+-"
	next
	? left(cLine, len(cLine) - 3)
	for aRow in aRows
		? accessFormatRow(aRow, aWidths)
	next

func accessFormatRow aRow, aWidths
	cLine = ""
	nCount = len(aRow)
	for nField = 1 to nCount
		if nField < nCount
			cLine += accessPad(aRow[nField], aWidths[nField]) + " | "
		else
			cLine += aRow[nField]
		ok
	next
	return cLine

func accessPad cText, nWidth
	# The width is counted in characters (not bytes), so UTF-8 text is aligned
	nChars = accessCharCount(cText)
	if nChars >= nWidth
		return cText
	ok
	return cText + copy(" ", nWidth - nChars)

func accessCharCount cText
	nCount = 0
	nLen = len(cText)
	for nIndex = 1 to nLen
		# Count the bytes that are not UTF-8 continuation bytes
		nByte = ascii(cText[nIndex])
		if nByte < 128 or nByte >= 192
			nCount++
		ok
	next
	return nCount

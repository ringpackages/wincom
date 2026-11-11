# Read and write Excel workbooks and CSV files using ADO (Without Microsoft Excel)
# The Access Database Engine (ACE) or Jet can use these files like database tables

load "wincom.ring"
load "accesshelper.ring"

func main
	# Select the provider (ACE for Ring 64-bit, ACE or Jet for Ring 32-bit)
	accessDeleteDatabase(accessFileNameOnly())
	cProvider = cAccessProvider
	if cAccessExtension = ".accdb"
		cWorkbookName = "prices.xlsx"
		cExcelVersion = "Excel 12.0 Xml"
	else
		cWorkbookName = "prices.xls"
		cExcelVersion = "Excel 8.0"
	ok
	cWorkbook = currentdir() + "\" + cWorkbookName
	if fexists(cWorkbook)
		remove(cWorkbook)
	ok
	? "Provider : " + cProvider

	# Create a workbook: the table is a worksheet (HDR=YES: the first row is the header)
	oExcel = com_createobject("ADODB.Connection")
	com_invoke(oExcel, "Open", "Provider=" + cProvider + ";Data Source=" + cWorkbook +
		";Extended Properties='" + cExcelVersion + ";HDR=YES'")
	com_invoke(oExcel, "Execute", "CREATE TABLE Prices (Product TEXT, Price DOUBLE, Updated DATETIME)")
	aRows = [["Product A", 12.5, "2026-01-05"], ["Product B", 27.75, "2026-02-10"], ["Продукт", 7.25, "2026-03-15"]]
	for aRow in aRows
		oCommand = com_createobject("ADODB.Command")
		com_setref(oCommand, "ActiveConnection", oExcel)
		com_set(oCommand, "CommandText", "INSERT INTO Prices (Product, Price, Updated) VALUES (?, ?, ?)")
		# Use adDate for dates (Parameters created by ADO from the values may swap the day and the month)
		com_invoke(oCommand, "Parameters.Append", com_invoke(oCommand, "CreateParameter", "p1", adVarWChar, adParamInput, 50, aRow[1]))
		com_invoke(oCommand, "Parameters.Append", com_invoke(oCommand, "CreateParameter", "p2", adDouble, adParamInput, 0, aRow[2]))
		com_invoke(oCommand, "Parameters.Append", com_invoke(oCommand, "CreateParameter", "p3", adDate, adParamInput, 0, com_date(aRow[3])))
		com_invoke(oCommand, "Execute")
	next
	com_invoke(oExcel, "Close")
	? "Created  : " + cWorkbookName + " (" + getfilesize(cWorkbook) + " bytes)"

	# Read the worksheet: [Prices$] is the worksheet, [Prices$A1:B3] is a range
	com_invoke(oExcel, "Open", "Provider=" + cProvider + ";Data Source=" + cWorkbook +
		";Extended Properties='" + cExcelVersion + ";HDR=YES;IMEX=1'")
	? nl + "SELECT * FROM [Prices$]:" + nl
	oRecordset = com_invoke(oExcel, "Execute", "SELECT * FROM [Prices$] ORDER BY Price")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# The worksheets of the workbook
	oSchema = com_invoke(oExcel, "OpenSchema", adSchemaTables)
	? nl + "Worksheets:"
	while ! com_get(oSchema, "EOF")
		? "	" + com_get(oSchema, 'Fields("TABLE_NAME").Value')
		com_invoke(oSchema, "MoveNext")
	end
	com_invoke(oSchema, "Close")
	com_invoke(oExcel, "Close")

	# CSV files: the folder is the database and each file is a table
	cCSV = "Product,Quantity" + nl + "Product A,100" + nl + "Product B,0" + nl + "Product C,42" + nl
	write("quantities.csv", cCSV)
	oText = com_createobject("ADODB.Connection")
	com_invoke(oText, "Open", "Provider=" + cProvider + ";Data Source=" + currentdir() +
		";Extended Properties='text;HDR=YES;FMT=Delimited'")
	? nl + "SELECT * FROM [quantities.csv] WHERE Quantity > 0:" + nl
	oRecordset = com_invoke(oText, "Execute", "SELECT * FROM [quantities.csv] WHERE Quantity > 0")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")
	com_invoke(oText, "Close")

	# Import the prices from the workbook into an Access table
	oConnection = accessCreateDatabase("sample_files")
	accessCreateProducts(oConnection)
	com_invoke(oExcel, "Open", "Provider=" + cProvider + ";Data Source=" + cWorkbook +
		";Extended Properties='" + cExcelVersion + ";HDR=YES;IMEX=1'")
	oRecordset = com_invoke(oExcel, "Execute", "SELECT Product, Price FROM [Prices$]")
	aPrices = com_rowsfromcolumns(com_invoke(oRecordset, "GetRows"))
	com_invoke(oRecordset, "Close")
	com_invoke(oExcel, "Close")
	nUpdated = 0
	for aPrice in aPrices
		oCommand = com_createobject("ADODB.Command")
		com_setref(oCommand, "ActiveConnection", oConnection)
		com_set(oCommand, "CommandText", "UPDATE Products SET Price = ? WHERE Name = ?")
		oAffected = com_byref(0)
		com_invoke(oCommand, "Execute", oAffected, [com_currency(aPrice[2]), aPrice[1]], adExecuteNoRecords)
		nUpdated += com_byrefvalue(oAffected)
	next
	? nl + "Prices imported from the workbook: " + nUpdated + " products updated" + nl
	oRecordset = com_invoke(oConnection, "Execute", "SELECT Name, Price FROM Products ORDER BY Name")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_files")

	remove("quantities.csv")
	remove(cWorkbook)

# The helper selects the provider when it creates a database, so we create and delete one
func accessFileNameOnly
	oConnection = accessCreateDatabase("sample_provider")
	com_invoke(oConnection, "Close")
	return "sample_provider"

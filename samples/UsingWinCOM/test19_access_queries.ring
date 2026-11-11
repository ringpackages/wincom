# Saved queries in Microsoft Access: views (SELECT queries) and procedures (parameter and action queries)

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_queries")
	accessCreateProducts(oConnection)

	# A view (A saved SELECT query)
	com_invoke(oConnection, "Execute", "CREATE VIEW ProductsInStock AS
		SELECT Name, Price, Quantity FROM Products WHERE InStock = True")

	# Procedures (Saved queries with parameters)
	com_invoke(oConnection, "Execute", "CREATE PROCEDURE ProductsByCategory (pCategory TEXT(30)) AS
		SELECT Name, Price FROM Products WHERE Category = pCategory ORDER BY Price")
	com_invoke(oConnection, "Execute", "CREATE PROCEDURE RaisePrices (pPercent DOUBLE) AS
		UPDATE Products SET Price = Price * (1 + pPercent / 100)")

	# Use the view like a table
	? "ProductsInStock:" + nl
	oRecordset = com_invoke(oConnection, "Execute", "SELECT * FROM ProductsInStock ORDER BY Name")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# Call a procedure using a Command (adCmdStoredProc)
	oCommand = com_createobject("ADODB.Command")
	com_setref(oCommand, "ActiveConnection", oConnection)
	com_set(oCommand, "CommandText", "ProductsByCategory")
	com_set(oCommand, "CommandType", adCmdStoredProc)
	com_invoke(oCommand, "Parameters.Append",
		com_invoke(oCommand, "CreateParameter", "pCategory", adVarWChar, adParamInput, 30, "Category 2"))
	oRecordset = com_invoke(oCommand, "Execute")
	? nl + "ProductsByCategory('Category 2'):" + nl
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# An action query: the number of changed records is returned using com_byref()
	oCommand = com_createobject("ADODB.Command")
	com_setref(oCommand, "ActiveConnection", oConnection)
	com_set(oCommand, "CommandText", "RaisePrices")
	com_set(oCommand, "CommandType", adCmdStoredProc)
	com_invoke(oCommand, "Parameters.Append",
		com_invoke(oCommand, "CreateParameter", "pPercent", adDouble, adParamInput, 0, 10))
	oAffected = com_byref(0)
	com_invoke(oCommand, "Execute", oAffected, com_missing(), adExecuteNoRecords)
	? nl + "RaisePrices(10): " + com_byrefvalue(oAffected) + " records changed"

	# The procedures can be executed using SQL too
	oRecordset = com_invoke(oConnection, "Execute", "EXECUTE ProductsByCategory 'Category 1'")
	? nl + "EXECUTE ProductsByCategory 'Category 1':" + nl
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# List the saved queries using ADOX (Views and Procedures)
	oCatalog = com_createobject("ADOX.Catalog")
	com_setref(oCatalog, "ActiveConnection", oConnection)
	? nl + "Views:"
	aViews = com_items(com_get(oCatalog, "Views"))
	for oView in aViews
		? "	" + com_get(oView, "Name") + " : " + trimSQL(com_get(oView, "Command.CommandText"))
	next
	? nl + "Procedures:"
	aProcedures = com_items(com_get(oCatalog, "Procedures"))
	for oProcedure in aProcedures
		? "	" + com_get(oProcedure, "Name") + " : " + trimSQL(com_get(oProcedure, "Command.CommandText"))
	next

	com_invoke(oConnection, "Execute", "DROP PROCEDURE RaisePrices")
	# The catalog keeps the objects that it read, so we read them again
	com_invoke(oCatalog, "Procedures.Refresh")
	? nl + "Procedures after DROP PROCEDURE: " + com_get(oCatalog, "Procedures.Count")

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_queries")

# The SQL in one line
func trimSQL cSQL
	cSQL = substr(substr(cSQL, char(13), " "), char(10), " ")
	cSQL = substr(cSQL, char(9), " ")
	while substr(cSQL, "  ")
		cSQL = substr(cSQL, "  ", " ")
	end
	return trim(cSQL)

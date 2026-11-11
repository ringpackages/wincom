# Using ADODB.Command with parameters
# Parameters keep the values out of the SQL statement (No SQL injection, no quotes problems)

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_parameters")
	accessCreateProducts(oConnection)

	# Search using a parameter
	oCommand = com_createobject("ADODB.Command")
	com_setref(oCommand, "ActiveConnection", oConnection)
	com_set(oCommand, "CommandText", "SELECT Name, Price FROM Products WHERE Price >= ? ORDER BY Price")
	# CreateParameter(Name, Type, Direction, Size, Value)
	oParameter = com_invoke(oCommand, "CreateParameter", "pMinPrice", adCurrency, adParamInput, 0, 10)
	com_invoke(oCommand, "Parameters.Append", oParameter)

	for nMinPrice in [10, 50]
		com_set(oCommand, 'Parameters("pMinPrice").Value', nMinPrice)
		oRecordset = com_invoke(oCommand, "Execute")
		? "Price >= " + nMinPrice + ":"
		while ! com_get(oRecordset, "EOF")
			? "	" + com_get(oRecordset, "Fields(0).Value") + " - " + com_get(oRecordset, "Fields(1).Value")
			com_invoke(oRecordset, "MoveNext")
		end
		com_invoke(oRecordset, "Close")
	next

	# Values with quotes and Unicode are safe
	aNewProducts = [
		["Product 'Quoted'", "Category 4", 12.5, "2026-06-01 09:30:00", "Name with quotes"],
		["Продукт",          "Category 4", 8,    "2026-06-02",          "Russian name"],
		["منتج",             "Category 4", 4.25, "2026-06-03",          "Arabic name"],
		["x'; DROP TABLE Products; --", "Category 4", 1, "2026-06-04",  "Not executed as SQL"]
	]
	for aProduct in aNewProducts
		insertProduct(oConnection, aProduct)
	next

	# Execute(RecordsAffected, Parameters): the parameters can be passed as a list too
	oSearch = com_createobject("ADODB.Command")
	com_setref(oSearch, "ActiveConnection", oConnection)
	com_set(oSearch, "CommandText", "SELECT Name, Price, Added, Notes FROM Products WHERE Category = ? AND Price < ?")
	oRecordset = com_invoke(oSearch, "Execute", com_missing(), ["Category 4", 100])
	? nl + "Category 4:" + nl
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	oRecordset = com_invoke(oConnection, "Execute", "SELECT COUNT(*) FROM Products")
	? nl + "The Products table still exists, records: " + com_get(oRecordset, "Fields(0).Value")
	com_invoke(oRecordset, "Close")

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_parameters")

/*
	Insert a product using a Command with parameters
	Microsoft Access (ACE) note: executing the same INSERT Command again with new values
	may swap the day and the month of dates, or fail with Boolean values
	So we create a new Command for each record (SELECT commands can be executed many times)
*/
func insertProduct oConnection, aProduct
	oInsert = com_createobject("ADODB.Command")
	com_setref(oInsert, "ActiveConnection", oConnection)
	com_set(oInsert, "CommandText", "INSERT INTO Products (Name, Category, Price, Added, Notes) VALUES (?, ?, ?, ?, ?)")
	# CreateParameter(Name, Type, Direction, Size, Value)
	com_invoke(oInsert, "Parameters.Append",
		com_invoke(oInsert, "CreateParameter", "pName", adVarWChar, adParamInput, 50, aProduct[1]))
	com_invoke(oInsert, "Parameters.Append",
		com_invoke(oInsert, "CreateParameter", "pCategory", adVarWChar, adParamInput, 30, aProduct[2]))
	com_invoke(oInsert, "Parameters.Append",
		com_invoke(oInsert, "CreateParameter", "pPrice", adCurrency, adParamInput, 0, aProduct[3]))
	# Use adDate (not adDBTimeStamp) for the dates of Microsoft Access
	com_invoke(oInsert, "Parameters.Append",
		com_invoke(oInsert, "CreateParameter", "pAdded", adDate, adParamInput, 0, com_date(aProduct[4])))
	com_invoke(oInsert, "Parameters.Append",
		com_invoke(oInsert, "CreateParameter", "pNotes", adLongVarWChar, adParamInput, 1000, aProduct[5]))
	com_invoke(oInsert, "Execute", com_missing(), com_missing(), adExecuteNoRecords)

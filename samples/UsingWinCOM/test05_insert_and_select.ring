# Insert, update, delete and select records using ADODB.Connection.Execute()

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_select")
	accessCreateProducts(oConnection)

	# Execute() returns a Recordset for SELECT statements
	? "All the products:" + nl
	oRecordset = com_invoke(oConnection, "Execute", "SELECT * FROM Products ORDER BY ID")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# INSERT (Use parameters for values from users, see test06_parameters.ring)
	com_invoke(oConnection, "Execute", "INSERT INTO Products (Name, Category, Price, Quantity, Added, InStock)
					    VALUES ('Product F', 'Category 3', 15.75, 20, #2026-05-01#, True)")

	# UPDATE and DELETE (RecordsAffected is a ByRef parameter, so we read it using a SELECT)
	com_invoke(oConnection, "Execute", "UPDATE Products SET Price = Price * 1.1 WHERE Category = 'Category 1'")
	com_invoke(oConnection, "Execute", "DELETE FROM Products WHERE Quantity = 0")

	# Read the fields by name
	? nl + "Products in stock:" + nl
	oRecordset = com_invoke(oConnection, "Execute",
		"SELECT Name, Price, Quantity, Added, Notes FROM Products WHERE InStock = True ORDER BY Price DESC")
	while ! com_get(oRecordset, "EOF")
		cName = com_get(oRecordset, 'Fields("Name").Value')
		nPrice = com_get(oRecordset, 'Fields("Price").Value')
		nQuantity = com_get(oRecordset, 'Fields("Quantity").Value')
		cAdded = com_get(oRecordset, 'Fields("Added").Value')
		# Null values are returned as "", com_wasnull() tells us if the last value was Null
		cNotes = com_get(oRecordset, 'Fields("Notes").Value')
		if com_wasnull()
			cNotes = "(No notes)"
		ok
		? cName + " | " + nPrice + " | " + nQuantity + " | " + cAdded + " | " + cNotes
		com_invoke(oRecordset, "MoveNext")
	end
	com_invoke(oRecordset, "Close")

	# Aggregate functions
	oRecordset = com_invoke(oConnection, "Execute",
		"SELECT Category, COUNT(*) AS Products, SUM(Price * Quantity) AS StockValue FROM Products GROUP BY Category")
	? nl + "Products by category:" + nl
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_select")

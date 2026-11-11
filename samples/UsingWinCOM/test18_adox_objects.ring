# Create tables, keys and indexes using the ADOX objects (without SQL)

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_adox")
	oCatalog = com_createobject("ADOX.Catalog")
	com_setref(oCatalog, "ActiveConnection", oConnection)

	# The Customers table
	oTable = com_createobject("ADOX.Table")
	com_set(oTable, "Name", "Customers")
	# Columns.Append(Name, Type, DefinedSize)
	com_invoke(oTable, "Columns.Append", "CustomerID", adInteger)
	com_invoke(oTable, "Columns.Append", "Name", adVarWChar, 50)
	com_invoke(oTable, "Columns.Append", "City", adVarWChar, 30)
	com_invoke(oTable, "Columns.Append", "Notes", adLongVarWChar)
	# The provider properties of a column need the catalog (ParentCatalog)
	com_setref(oTable, 'Columns("CustomerID").ParentCatalog', oCatalog)
	com_set(oTable, 'Columns("CustomerID").Properties("AutoIncrement").Value', com_bool(1))
	com_set(oTable, 'Columns("City").Attributes', adColNullable)
	com_set(oTable, 'Columns("Notes").Attributes', adColNullable)
	# Keys.Append(Name, Type, Column)
	com_invoke(oTable, "Keys.Append", "PK_Customers", adKeyPrimary, "CustomerID")
	com_invoke(oCatalog, "Tables.Append", oTable)

	# The Orders table, with a foreign key: Keys.Append(Name, Type, Column, RelatedTable, RelatedColumn)
	oTable = com_createobject("ADOX.Table")
	com_set(oTable, "Name", "Orders")
	com_invoke(oTable, "Columns.Append", "OrderID", adInteger)
	com_invoke(oTable, "Columns.Append", "CustomerID", adInteger)
	com_invoke(oTable, "Columns.Append", "OrderDate", adDate)
	com_invoke(oTable, "Columns.Append", "Total", adCurrency)
	com_invoke(oTable, "Keys.Append", "PK_Orders", adKeyPrimary, "OrderID")
	com_invoke(oTable, "Keys.Append", "FK_Orders_Customers", adKeyForeign, "CustomerID", "Customers", "CustomerID")
	com_invoke(oCatalog, "Tables.Append", oTable)

	# An index: ADOX.Index
	oIndex = com_createobject("ADOX.Index")
	com_set(oIndex, "Name", "idxCity")
	com_invoke(oIndex, "Columns.Append", "City")
	com_set(oIndex, "Unique", com_bool(0))
	com_invoke(oCatalog, 'Tables("Customers").Indexes.Append', oIndex)

	# Use the tables
	com_invoke(oConnection, "Execute", "INSERT INTO Customers (Name, City) VALUES ('Customer 1', 'City A')")
	com_invoke(oConnection, "Execute", "INSERT INTO Customers (Name, City) VALUES ('Customer 2', 'City B')")
	com_invoke(oConnection, "Execute", "INSERT INTO Orders VALUES (1, 1, #2026-03-01#, 150)")
	com_invoke(oConnection, "Execute", "INSERT INTO Orders VALUES (2, 2, #2026-03-05#, 75.5)")
	try
		# The foreign key doesn't allow orders for customers that don't exist
		com_invoke(oConnection, "Execute", "INSERT INTO Orders VALUES (3, 99, #2026-03-06#, 10)")
	catch
		? "Expected error: " + cCatchError + nl
	done

	oRecordset = com_invoke(oConnection, "Execute",
		"SELECT o.OrderID, c.Name, c.City, o.OrderDate, o.Total FROM Orders o INNER JOIN Customers c ON o.CustomerID = c.CustomerID")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# Read the structure (Refresh() reads the tables again, the catalog keeps the objects that we appended)
	com_invoke(oCatalog, "Tables.Refresh")
	aTables = com_items(com_get(oCatalog, "Tables"))
	for oTable in aTables
		if com_get(oTable, "Type") != "TABLE"
			loop
		ok
		? nl + "Table: " + com_get(oTable, "Name")
		aKeys = com_items(com_get(oTable, "Keys"))
		for oKey in aKeys
			cLine = "	Key: " + com_get(oKey, "Name") + " (Type: " + com_get(oKey, "Type") + ")"
			if com_get(oKey, "Type") = adKeyForeign
				cLine += " --> " + com_get(oKey, "RelatedTable")
			ok
			? cLine
		next
		aIndexes = com_items(com_get(oTable, "Indexes"))
		for oIndex in aIndexes
			? "	Index: " + com_get(oIndex, "Name") + " (Unique: " + com_get(oIndex, "Unique") + ")"
		next
	next

	# Delete a table
	com_invoke(oCatalog, "Tables.Delete", "Orders")
	? nl + "Tables after deleting Orders: " + com_get(oCatalog, "Tables.Count") + " (with the system tables)"

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_adox")

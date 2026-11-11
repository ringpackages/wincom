# Create a Microsoft Access database using ADOX, then create tables using SQL (DDL)
# Then list the tables and the columns using the ADOX Catalog

load "wincom.ring"
load "accesshelper.ring"

func main
	# Create the database file (ADOX.Catalog.Create() returns an open ADODB.Connection)
	oConnection = accessCreateDatabase("sample_create")
	? "Database : " + accessFileName("sample_create")
	? "Provider : " + com_get(oConnection, "Provider")
	? "ADO      : " + com_get(oConnection, "Version")

	# Create tables using SQL
	com_invoke(oConnection, "Execute", "
		CREATE TABLE Customers (
			CustomerID COUNTER PRIMARY KEY,
			Name       TEXT(50) NOT NULL,
			City       TEXT(30),
			Email      TEXT(50)
		)
	")
	com_invoke(oConnection, "Execute", "
		CREATE TABLE Orders (
			OrderID    COUNTER PRIMARY KEY,
			CustomerID INTEGER NOT NULL REFERENCES Customers (CustomerID),
			OrderDate  DATETIME,
			Total      CURRENCY
		)
	")
	com_invoke(oConnection, "Execute", "CREATE INDEX idxOrderDate ON Orders (OrderDate)")

	# Read the structure using ADOX (Catalog.ActiveConnection is an object, so we use com_setref())
	oCatalog = com_createobject("ADOX.Catalog")
	com_setref(oCatalog, "ActiveConnection", oConnection)
	aTables = com_items(com_get(oCatalog, "Tables"))
	for oTable in aTables
		# Skip the system tables (Type: "SYSTEM TABLE", "ACCESS TABLE")
		if com_get(oTable, "Type") != "TABLE"
			loop
		ok
		? nl + "Table: " + com_get(oTable, "Name")
		aColumns = com_items(com_get(oTable, "Columns"))
		for oColumn in aColumns
			? "	" + com_get(oColumn, "Name") + " (Type: " + com_get(oColumn, "Type") +
			  ", Size: " + com_get(oColumn, "DefinedSize") + ")"
		next
		aIndexes = com_items(com_get(oTable, "Indexes"))
		for oIndex in aIndexes
			? "	Index: " + com_get(oIndex, "Name") + " (Primary key: " + com_get(oIndex, "PrimaryKey") + ")"
		next
	next

	# Close the connection (The file is unlocked) then delete the database
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_create")
	? nl + "Done"

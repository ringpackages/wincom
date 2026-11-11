# Read all the records in one call using Recordset.GetRows()
# Then read the database structure using Connection.OpenSchema()

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_getrows")
	accessCreateProducts(oConnection)

	# GetRows() returns a two dimensions array: aColumns[nField][nRow]
	oRecordset = com_invoke(oConnection, "Execute", "SELECT Name, Price, Quantity FROM Products ORDER BY Name")
	aColumns = com_invoke(oRecordset, "GetRows")
	com_invoke(oRecordset, "Close")
	? "Fields: " + len(aColumns) + ", Records: " + len(aColumns[1])

	# com_rowsfromcolumns() converts it to a list of records: aRows[nRow][nField]
	aRows = com_rowsfromcolumns(aColumns)
	? nl + "Records:"
	for aRow in aRows
		? "	" + aRow[1] + " | " + aRow[2] + " | " + aRow[3]
	next

	# Process the data using Ring
	nTotal = 0
	for aRow in aRows
		nTotal += aRow[2] * aRow[3]
	next
	? nl + "Stock value (calculated in Ring): " + nTotal

	# GetRows(Rows, Start, Fields): read some fields only
	oRecordset = com_invoke(oConnection, "Execute", "SELECT * FROM Products ORDER BY ID")
	aColumns = com_invoke(oRecordset, "GetRows", adGetRowsRest, com_missing(), ["Name", "Category"])
	com_invoke(oRecordset, "Close")
	? nl + "Name and Category:"
	aRows = com_rowsfromcolumns(aColumns)
	for aRow in aRows
		? "	" + aRow[1] + " (" + aRow[2] + ")"
	next

	# The tables (OpenSchema returns a Recordset)
	? nl + "Tables:"
	oSchema = com_invoke(oConnection, "OpenSchema", adSchemaTables)
	while ! com_get(oSchema, "EOF")
		if com_get(oSchema, 'Fields("TABLE_TYPE").Value') = "TABLE"
			? "	" + com_get(oSchema, 'Fields("TABLE_NAME").Value')
		ok
		com_invoke(oSchema, "MoveNext")
	end
	com_invoke(oSchema, "Close")

	# The columns of the Products table: OpenSchema(adSchemaColumns, [Catalog, Schema, Table, Column])
	? nl + "Columns of the Products table:"
	oSchema = com_invoke(oConnection, "OpenSchema", adSchemaColumns, [com_empty(), com_empty(), "Products"])
	aColumns = []
	while ! com_get(oSchema, "EOF")
		aColumns + [com_get(oSchema, 'Fields("ORDINAL_POSITION").Value'),
			    com_get(oSchema, 'Fields("COLUMN_NAME").Value'),
			    com_get(oSchema, 'Fields("DATA_TYPE").Value')]
		com_invoke(oSchema, "MoveNext")
	end
	com_invoke(oSchema, "Close")
	aColumns = sort(aColumns, 1)
	for aColumn in aColumns
		? "	" + aColumn[1] + ". " + aColumn[2] + " (Data type: " + aColumn[3] + ")"
	next

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_getrows")

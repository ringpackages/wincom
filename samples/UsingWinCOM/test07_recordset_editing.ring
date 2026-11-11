# Editing records using ADODB.Recordset: AddNew, Update, Delete, Find, Filter and Sort

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_recordset")
	accessCreateProducts(oConnection)

	# Open(Source, ActiveConnection, CursorType, LockType, Options)
	oRecordset = com_createobject("ADODB.Recordset")
	com_set(oRecordset, "CursorLocation", adUseClient)
	com_invoke(oRecordset, "Open", "Products", oConnection, adOpenStatic, adLockOptimistic, adCmdTable)
	? "Records: " + com_get(oRecordset, "RecordCount")

	# Add a new record
	com_invoke(oRecordset, "AddNew")
	com_set(oRecordset, 'Fields("Name").Value', "Product New")
	com_set(oRecordset, 'Fields("Category").Value', "Category 5")
	com_set(oRecordset, 'Fields("Price").Value', 49.5)
	com_set(oRecordset, 'Fields("Quantity").Value', 8)
	com_set(oRecordset, 'Fields("Added").Value', com_date("2026-07-15 14:00:00"))
	com_set(oRecordset, 'Fields("InStock").Value', com_bool(1))
	com_invoke(oRecordset, "Update")
	? "New record ID: " + com_get(oRecordset, 'Fields("ID").Value')

	# AddNew(FieldList, Values) adds a record in one call
	com_invoke(oRecordset, "AddNew", ["Name", "Category", "Price"], ["Product Quick", "Category 5", 1.25])
	? "Records after AddNew: " + com_get(oRecordset, "RecordCount")

	# Find a record and change it
	com_invoke(oRecordset, "MoveFirst")
	com_invoke(oRecordset, "Find", "Name = 'Product C'")
	if ! com_get(oRecordset, "EOF")
		? "Found: " + com_get(oRecordset, 'Fields("Name").Value') + ", Quantity: " + com_get(oRecordset, 'Fields("Quantity").Value')
		com_set(oRecordset, 'Fields("Quantity").Value', 40)
		com_set(oRecordset, 'Fields("Notes").Value', "Quantity updated from Ring")
		com_invoke(oRecordset, "Update")
	ok

	# Delete the current record
	com_invoke(oRecordset, "MoveFirst")
	com_invoke(oRecordset, "Find", "Name = 'Product B'")
	if ! com_get(oRecordset, "EOF")
		com_invoke(oRecordset, "Delete", adAffectCurrent)
		? "Product B deleted"
	ok

	# Filter and Sort (Client side cursor)
	com_set(oRecordset, "Filter", "Category = 'Category 5' OR Price > 20")
	com_set(oRecordset, "Sort", "Price DESC")
	? nl + "Filtered and sorted:" + nl
	com_invoke(oRecordset, "MoveFirst")
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# The changes are saved in the database
	oRecordset = com_invoke(oConnection, "Execute", "SELECT ID, Name, Quantity, Notes FROM Products ORDER BY ID")
	? nl + "Products table:" + nl
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_recordset")

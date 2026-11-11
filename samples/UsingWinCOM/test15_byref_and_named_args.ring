# ByRef (output) arguments: com_byref() and com_byrefvalue()
# Named arguments (Like Name:=Value in Visual Basic): com_named()

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_byref")
	accessCreateProducts(oConnection)

	# Connection.Execute(CommandText, RecordsAffected, Options): RecordsAffected is a ByRef argument
	oAffected = com_byref(0)
	com_invoke(oConnection, "Execute", "UPDATE Products SET Price = Price * 2 WHERE Category = 'Category 2'",
		   oAffected, adExecuteNoRecords)
	? "Updated records: " + com_byrefvalue(oAffected)

	oAffected = com_byref(0)
	com_invoke(oConnection, "Execute", "DELETE FROM Products WHERE Quantity < 10", oAffected, adExecuteNoRecords)
	? "Deleted records: " + com_byrefvalue(oAffected)

	# Named arguments: skip the arguments that we don't need, in any order
	oRecordset = com_createobject("ADODB.Recordset")
	com_invoke(oRecordset, "Open",
		com_named("Source", "SELECT Name, Price FROM Products ORDER BY Price"),
		com_named("ActiveConnection", oConnection),
		com_named("LockType", adLockReadOnly),
		com_named("CursorType", adOpenStatic))
	? nl + "Recordset.Open() using named arguments:" + nl
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# Positional arguments first, then named arguments
	oRecordset = com_invoke(oConnection, "Execute", "SELECT COUNT(*) FROM Products", com_named("Options", adCmdText))
	? nl + "Records: " + com_get(oRecordset, "Fields(0).Value")
	com_invoke(oRecordset, "Close")

	# Using the COMObject class (The arguments are passed in the list)
	oDict = new COMObject("Scripting.Dictionary")
	oDict.invoke("Add", [com_named("Key", "one"), com_named("Item", 1)])
	? nl + "Dictionary item (Added using named arguments): " + oDict.getPropArgs("Item", ["one"])

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_byref")

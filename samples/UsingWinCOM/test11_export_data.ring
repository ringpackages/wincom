# Export the data: Recordset.GetString() (CSV), Recordset.Save() (XML) and ADODB.Stream (UTF-8 file)

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_export")
	accessCreateProducts(oConnection)

	# GetString(StringFormat (2 = adClipString), NumRows, ColumnDelimiter, RowDelimiter, NullExpr)
	oRecordset = com_invoke(oConnection, "Execute", "SELECT ID, Name, Category, Price FROM Products ORDER BY ID")
	cCSV = "ID,Name,Category,Price" + nl
	cCSV += com_invoke(oRecordset, "GetString", 2, adGetRowsRest, ",", nl, "")
	com_invoke(oRecordset, "Close")
	? "CSV:" + nl + cCSV

	# Save the CSV as UTF-8 using ADODB.Stream
	oStream = com_createobject("ADODB.Stream")
	com_set(oStream, "Type", adTypeText)
	com_set(oStream, "Charset", "utf-8")
	com_invoke(oStream, "Open")
	com_invoke(oStream, "WriteText", cCSV)
	com_invoke(oStream, "SaveToFile", currentdir() + "\products.csv", adSaveCreateOverWrite)
	com_invoke(oStream, "Close")
	? "products.csv size: " + len(read("products.csv")) + " bytes"

	# Recordset.Save() writes XML (The file must not exist)
	if fexists("products.xml")
		remove("products.xml")
	ok
	oRecordset = com_createobject("ADODB.Recordset")
	com_set(oRecordset, "CursorLocation", adUseClient)
	com_invoke(oRecordset, "Open", "SELECT ID, Name, Price FROM Products WHERE Price > 20", oConnection, adOpenStatic, adLockReadOnly)
	com_invoke(oRecordset, "Save", currentdir() + "\products.xml", adPersistXML)
	com_invoke(oRecordset, "Close")
	? nl + "products.xml size: " + len(read("products.xml")) + " bytes"

	# Load the XML file again (Without the database)
	oRecordset = com_createobject("ADODB.Recordset")
	com_invoke(oRecordset, "Open", currentdir() + "\products.xml", "Provider=MSPersist")
	? nl + "Records loaded from products.xml:" + nl
	accessPrintRecordset(oRecordset)
	com_invoke(oRecordset, "Close")

	# Binary data: Stream.Read() returns a Ring string
	oStream = com_createobject("ADODB.Stream")
	com_set(oStream, "Type", adTypeBinary)
	com_invoke(oStream, "Open")
	com_invoke(oStream, "LoadFromFile", currentdir() + "\products.csv")
	cBytes = com_invoke(oStream, "Read")
	com_invoke(oStream, "Close")
	? nl + "products.csv bytes: " + len(cBytes)
	# ADODB.Stream writes the UTF-8 BOM (EF BB BF) before the text
	? "The first 3 bytes: " + upper(hex(ascii(cBytes[1]))) + " " + upper(hex(ascii(cBytes[2]))) + " " + upper(hex(ascii(cBytes[3])))

	remove("products.csv")
	remove("products.xml")
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_export")

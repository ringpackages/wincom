# XML using MSXML (MSXML2.DOMDocument.6.0): create, save, load and query using XPath

load "wincom.ring"
load "accesshelper.ring"

func main
	# Create an XML document from the records of an Access table
	oConnection = accessCreateDatabase("sample_xml")
	accessCreateProducts(oConnection)
	oRecordset = com_invoke(oConnection, "Execute", "SELECT ID, Name, Category, Price, Quantity FROM Products ORDER BY ID")
	aRows = com_rowsfromcolumns(com_invoke(oRecordset, "GetRows"))
	com_invoke(oRecordset, "Close")
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_xml")

	oXML = com_createobject("MSXML2.DOMDocument.6.0")
	com_invoke(oXML, "appendChild", com_invoke(oXML, "createProcessingInstruction", "xml", 'version="1.0" encoding="UTF-8"'))
	oRoot = com_invoke(oXML, "createElement", "products")
	com_invoke(oXML, "appendChild", oRoot)
	for aRow in aRows
		oProduct = com_invoke(oXML, "createElement", "product")
		com_invoke(oProduct, "setAttribute", "id", aRow[1])
		com_invoke(oProduct, "setAttribute", "category", aRow[3])
		addElement(oXML, oProduct, "name", aRow[2])
		addElement(oXML, oProduct, "price", aRow[4])
		addElement(oXML, oProduct, "quantity", aRow[5])
		com_invoke(oRoot, "appendChild", oProduct)
	next
	# Unicode text
	oProduct = com_invoke(oXML, "createElement", "product")
	com_invoke(oProduct, "setAttribute", "id", 99)
	com_invoke(oProduct, "setAttribute", "category", "Category 9")
	addElement(oXML, oProduct, "name", "Продукт - منتج")
	addElement(oXML, oProduct, "price", 1)
	addElement(oXML, oProduct, "quantity", 1)
	com_invoke(oRoot, "appendChild", oProduct)

	cFile = currentdir() + "\products.xml"
	com_invoke(oXML, "save", cFile)
	? "Saved: products.xml (" + getfilesize(cFile) + " bytes)"

	# Load the file again
	oDoc = com_createobject("MSXML2.DOMDocument.6.0")
	com_set(oDoc, "async", com_bool(0))
	if ! com_invoke(oDoc, "load", cFile)
		? "Error: " + com_get(oDoc, "parseError.reason")
		return
	ok
	? "Products: " + com_get(oDoc, 'selectNodes("/products/product").length')

	# XPath queries
	? nl + "Products in Category 2:"
	aNodes = com_items(com_invoke(oDoc, "selectNodes", "/products/product[@category='Category 2']"))
	for oNode in aNodes
		? "	" + com_get(oNode, 'selectSingleNode("name").text') + " : " + com_get(oNode, 'selectSingleNode("price").text')
	next

	? nl + "Products with quantity > 40:"
	aNodes = com_items(com_invoke(oDoc, "selectNodes", "//product[quantity > 40]/name"))
	for oNode in aNodes
		? "	" + com_get(oNode, "text")
	next

	? nl + "Product 99: " + com_get(oDoc, `selectSingleNode("//product[@id='99']/name").text`)

	# The XML text
	? nl + "The first product (XML):"
	? com_get(oDoc, "documentElement.firstChild.xml")

	# Parse a string, and get the error of bad XML
	oBad = com_createobject("MSXML2.DOMDocument.6.0")
	if ! com_invoke(oBad, "loadXML", "<a><b></a>")
		cReason = substr(substr(com_get(oBad, "parseError.reason"), char(13), ""), char(10), "")
		? nl + "Bad XML: " + cReason + " (Line " + com_get(oBad, "parseError.line") + ")"
	ok

	remove(cFile)

func addElement oXML, oParent, cName, vValue
	oElement = com_invoke(oXML, "createElement", cName)
	com_set(oElement, "text", "" + vValue)
	com_invoke(oParent, "appendChild", oElement)

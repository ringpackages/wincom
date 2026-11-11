# Unit tests for the WinCOM package
# Values are stored in a Scripting.Dictionary and read again (The dictionary keeps the VARIANT as it is)

load "wincom.ring"
load "accesshelper.ring"

nTests  = 0
nFailed = 0

# Used by the event handlers
aEventLog = []
oEventConnection = NULL
oByRefInEvent = NULL

func main
	testValues()
	testPaths()
	testObjects()
	testCOMObjectClass()
	testErrors()
	testByRefAndNamed()
	testEvents()
	testAccess()
	testAccessEvents()
	? nl + "Tests: " + nTests + ", Failed: " + nFailed
	if nFailed > 0
		shutdown(1)
	ok

func check cName, vValue, vExpected
	nTests++
	if isNumber(vValue) and isNumber(vExpected)
		lOK = fabs(vValue - vExpected) < 0.000001
	elseif isString(vValue) and isString(vExpected)
		lOK = (vValue = vExpected) and (len(vValue) = len(vExpected))
	elseif isList(vValue) and isList(vExpected)
		lOK = listText(vValue) = listText(vExpected)
	else
		lOK = False
	ok
	if lOK
		? "Pass : " + cName
	else
		nFailed++
		? "FAIL : " + cName
		? "	Value    : " + valueText(vValue)
		? "	Expected : " + valueText(vExpected)
	ok

func valueText vValue
	if isList(vValue)
		return listText(vValue)
	ok
	return "" + vValue

func listText aList
	cText = "["
	nCount = len(aList)
	for nIndex = 1 to nCount
		if isList(aList[nIndex])
			cText += listText(aList[nIndex])
		elseif isString(aList[nIndex])
			cText += '"' + aList[nIndex] + '"'
		else
			cText += "" + aList[nIndex]
		ok
		if nIndex < nCount
			cText += ", "
		ok
	next
	return cText + "]"

func checkError cName, cCode, cExpected
	nTests++
	cError = ""
	try
		eval(cCode)
	catch
		cError = cCatchError
	done
	if substr(cError, cExpected) > 0
		? "Pass : " + cName
	else
		nFailed++
		? "FAIL : " + cName + " (Error: " + cError + ")"
	ok

func roundTrip vValue
	oDict = com_createobject("Scripting.Dictionary")
	com_set(oDict, 'Item("v")', vValue)
	return com_get(oDict, 'Item("v")')

func testValues
	? nl + "Values" + nl
	check("Integer", roundTrip(42), 42)
	check("Negative integer", roundTrip(-7), -7)
	check("Integer type is VT_I4", com_lasttype(), VT_I4)
	check("Double", roundTrip(3.25), 3.25)
	check("Double type is VT_R8", com_lasttype(), VT_R8)
	check("Large number", roundTrip(5000000000), 5000000000)
	check("Empty string", roundTrip(""), "")
	check("ASCII string", roundTrip("Hello"), "Hello")
	check("UTF-8 string", roundTrip("Привет مرحبا 你好 😀"), "Привет مرحبا 你好 😀")
	check("String with NUL", roundTrip("a" + char(0) + "b"), "a" + char(0) + "b")
	check("String type is VT_BSTR", com_lasttype(), VT_BSTR)
	check("List", roundTrip([1, "two", 3.5]), [1, "two", 3.5])
	check("Nested list", roundTrip([1, [2, [3, "x"]], []]), [1, [2, [3, "x"]], []])
	check("Empty list", roundTrip([]), [])
	check("Boolean True", roundTrip(com_bool(1)), 1)
	check("Boolean type is VT_BOOL", com_lasttype(), VT_BOOL)
	check("Boolean False", roundTrip(com_bool(0)), 0)
	check("Null", roundTrip(com_null()), "")
	check("com_wasnull()", com_wasnull(), 1)
	check("Empty", roundTrip(com_empty()), "")
	check("com_wasnull() after Empty", com_wasnull(), 0)
	check("Date", roundTrip(com_date("2026-03-04")), "2026-03-04")
	check("Date type is VT_DATE", com_lasttype(), VT_DATE)
	check("Date and time", roundTrip(com_date("2026-12-31 23:59:58")), "2026-12-31 23:59:58")
	check("Date with T", roundTrip(com_date("2026-01-02T03:04:05")), "2026-01-02 03:04:05")
	check("Currency", roundTrip(com_currency(12.3456)), 12.3456)
	check("Currency type is VT_CY", com_lasttype(), VT_CY)
	check("Double from integer", roundTrip(com_double(5)), 5)
	check("com_double() type is VT_R8", com_lasttype(), VT_R8)
	check("com_variant() string to VT_I4", roundTrip(com_variant(VT_I4, "123")), 123)
	check("com_variant() string to VT_R8 (invariant)", roundTrip(com_variant(VT_R8, "1.5")), 1.5)
	check("com_variant() number to VT_BSTR", roundTrip(com_variant(VT_BSTR, 15)), "15")
	cBytes = char(0) + char(1) + char(255) + "ABC"
	check("Byte array", roundTrip(com_bytes(cBytes)), cBytes)
	check("Empty byte array", roundTrip(com_bytes("")), "")
	check("Missing", roundTrip(com_missing()), -2147352572)
	check("VT_ERROR unsigned code", roundTrip(com_variant(VT_ERROR, 2148139465)), -2146827831)
	check("VT_ERROR signed code", roundTrip(com_variant(VT_ERROR, -2146827831)), -2146827831)
	checkError("VT_ERROR code out of range", "com_variant(VT_ERROR, 5000000000)", "32-bit number")
	check("Variant list items", roundTrip([com_null(), com_bool(1), com_date("2026-05-06")]), ["", 1, "2026-05-06"])
	checkError("Bad date", 'com_date("2026-13-01")', "com_variant()")
	checkError("Bad day (February 30)", 'com_date("2026-02-30")', "com_variant()")
	checkError("Bad day (February 29, not leap year)", 'com_date("2026-02-29")', "com_variant()")
	check("February 29 (leap year)", roundTrip(com_date("2028-02-29")), "2028-02-29")
	checkError("Date with extra text", 'com_date("2026-01-02 10:30:00xyz")', "com_variant()")
	check("Date without seconds", roundTrip(com_date("2026-01-02 10:30")), "2026-01-02 10:30:00")
	checkError("Bad number string", 'com_variant(VT_I4, "abc")', "com_variant()")
	check("com_matrix()", roundTrip(com_matrix([[1, 2], [3, 4], [5, 6]])), [[1, 2], [3, 4], [5, 6]])
	check("com_matrix() short rows", roundTrip(com_matrix([[1], [2, "x"]])), [[1, ""], [2, "x"]])
	check("com_matrix() values as rows", roundTrip(com_matrix([1, 2])), [[1], [2]])
	check("com_matrix() empty", roundTrip(com_matrix([])), [])
	check("com_matrix() with variants", roundTrip(com_matrix([[com_null(), com_date("2026-01-02")]])), [["", "2026-01-02"]])
	checkError("com_matrix() not a list", "com_matrix(5)", "Bad parameter type")
	checkError("com_matrix() C pointer", 'com_matrix(com_createobject("Scripting.Dictionary"))', "list of rows")
	check("Typed array VT_I4", roundTrip(com_variant(VT_ARRAY + VT_I4, [1, "2", 3.0])), [1, 2, 3])
	check("Typed array VT_I4 type", com_lasttype(), VT_ARRAY + VT_I4)
	check("Typed array VT_BSTR", roundTrip(com_variant(VT_ARRAY + VT_BSTR, ["a", "Привет", 5])), ["a", "Привет", "5"])
	check("Typed array VT_BSTR type", com_lasttype(), VT_ARRAY + VT_BSTR)
	check("Typed array VT_R8", roundTrip(com_variant(VT_ARRAY + VT_R8, [1.5, "2.25"])), [1.5, 2.25])
	check("Typed array VT_BOOL", roundTrip(com_variant(VT_ARRAY + VT_BOOL, [1, 0])), [1, 0])
	check("Typed array VT_DECIMAL", roundTrip(com_variant(VT_ARRAY + VT_DECIMAL, [12.5])), [12.5])
	check("Typed array empty", roundTrip(com_variant(VT_ARRAY + VT_I2, [])), [])
	checkError("Typed array bad item", 'com_variant(VT_ARRAY + VT_I4, ["abc"])', "Type mismatch")
	checkError("Typed array with a list", "com_variant(VT_ARRAY + VT_I4, [[1]])", "can't contain lists")
	checkError("Typed array unsupported type", "com_variant(VT_ARRAY + 36, [1])", "not supported")
	checkError("Typed array out of range", "com_variant(VT_ARRAY + VT_I2, [100000])", "0x8002000A")
	check("com_version()", com_version(), "1.0.0")

func testPaths
	? nl + "Paths" + nl
	oDict = com_createobject("Scripting.Dictionary")
	com_invoke(oDict, "Add", "a", 1)
	com_invoke(oDict, "Add", "it's", 2)
	com_invoke(oDict, "Add", 'say "hi"', 3)
	com_invoke(oDict, "Add", 10, "ten")
	com_invoke(oDict, "Add", "a.b(c)", 4)
	check("Name and argument", com_get(oDict, "Item", "a"), 1)
	check("Path with double quotes", com_get(oDict, 'Item("a")'), 1)
	check("Path with single quotes", com_get(oDict, "Item('a')"), 1)
	check("Quote inside single quotes", com_get(oDict, "Item('it''s')"), 2)
	check("Quote inside double quotes", com_get(oDict, 'Item("say ""hi""")'), 3)
	check("Number argument", com_get(oDict, "Item(10)"), "ten")
	check("Dots and parentheses inside a string", com_get(oDict, 'Item("a.b(c)")'), 4)
	check("Spaces", com_get(oDict, ' Item ( "a" ) '), 1)
	check("Case insensitive names", com_get(oDict, "COUNT"), 5)
	com_set(oDict, 'Item("a")', 100)
	check("com_set() with path arguments", com_get(oDict, 'Item("a")'), 100)
	com_set(oDict, "Item", "a", 200)
	check("com_set() with arguments", com_get(oDict, 'Item("a")'), 200)
	oFSO = com_createobject("Scripting.FileSystemObject")
	check("Path with a method call", com_get(oFSO, 'GetFolder("' + currentdir() + '").Name') != "", 1)
	check("com_dispatch() COM_GET", com_dispatch(oDict, "Item", COM_GET, ["a"]), 200)
	com_dispatch(oDict, "Item", COM_SET, ["a", 300])
	check("com_dispatch() COM_SET", com_get(oDict, 'Item("a")'), 300)
	check("com_dispatch() COM_INVOKE", com_dispatch(oDict, "Exists", COM_INVOKE, ["a"]), 1)
	checkError("Missing parenthesis", 'com_get(com_createobject("Scripting.Dictionary"), "Item(")', "Bad member name or path")
	checkError("Missing quote", 'com_get(com_createobject("Scripting.Dictionary"), "Item(' + "'a)" + '")', "Bad member name or path")
	checkError("Not an object in the path", 'com_get(com_createobject("Scripting.Dictionary"), "Count.Value")', "not a COM object")
	checkError("Path ending with a dot", 'com_get(com_createobject("Scripting.Dictionary"), "Count.")', "Bad member name or path")
	checkError("Path ending with a dot and spaces", 'com_get(com_createobject("Scripting.Dictionary"), "Count. ")', "Bad member name or path")
	checkError("COM_SET without a value", 'com_dispatch(com_createobject("Scripting.Dictionary"), ' + "'Item(" + '"k")' + "'" + ', COM_SET, [])', "The value is missing")
	checkError("Bad mode", 'com_dispatch(com_createobject("Scripting.Dictionary"), "Count", 9, [])', "Bad mode")

func testObjects
	? nl + "Objects" + nl
	oInner = com_createobject("Scripting.Dictionary")
	oOuter = com_createobject("Scripting.Dictionary")
	check("com_isobject()", com_isobject(oInner), 1)
	check("com_isobject() string", com_isobject("x"), 0)
	check("com_isobject() list", com_isobject([1, 2]), 0)
	check("com_typename()", com_typename(oInner), "IDictionary")
	check("Reference count", com_refcount(oInner), 1)
	com_invoke(oOuter, "Add", "inner", oInner)
	check("Reference count after Add", com_refcount(oInner), 2)
	com_invoke(oInner, "Add", "x", 5)
	check("Object in object", com_get(oOuter, 'Item("inner").Item("x")'), 5)
	oCopy = oInner
	check("Copy shares the object", com_refcount(oInner), 2)
	checkRefCountInFunction(oOuter, oInner)
	check("Reference count after function", com_refcount(oInner), 2)
	aItems = com_items(oOuter)
	check("com_items() keys", aItems, ["inner"])
	aMembers = com_members(oInner)
	lFound = False
	for aMember in aMembers
		if aMember[1] = "Add" and aMember[2] = "method" and aMember[3] = 2
			lFound = True
		ok
	next
	check("com_members() has Add", lFound, True)
	cDescription = ""
	aParameters = []
	for aMember in aMembers
		if aMember[1] = "Add"
			cDescription = aMember[4]
			aParameters = aMember[5]
		ok
	next
	check("com_members() description", substr(cDescription, "Add a new key") > 0, 1)
	check("com_members() parameter names", aParameters, ["Key", "Item"])
	com_release(oCopy)
	check("com_release() releases all the copies", com_isobject(oInner), 0)
	check("Object stays in the dictionary", com_get(oOuter, 'Item("inner").Count'), 1)
	checkError("Using a released object", 'o = com_createobject("Scripting.Dictionary") com_release(o) com_get(o, "Count")', "must be a COM object")
	oFSO = com_createobject("Scripting.FileSystemObject")
	aDrives = com_items(com_get(oFSO, "Drives"))
	check("com_items() returns objects", com_isobject(aDrives[1]), 1)
	check("Progid exists", com_progidexists("Scripting.Dictionary"), 1)
	check("Progid doesn't exist", com_progidexists("NoSuch.Object"), 0)
	oValue = com_createobject("Scripting.Dictionary")
	com_set(oOuter, "Item", "byset", oValue)
	check("com_set() with an object (PUTREF)", com_issame(com_get(oOuter, 'Item("byset")'), oValue), 1)
	check("com_issame() same object", com_issame(oOuter, oOuter), 1)
	check("com_issame() other object", com_issame(oOuter, oValue), 0)
	checkError("com_issame() not an object", 'com_issame(1, 2)', "must be COM objects")
	aConstants = com_constants(com_createobject("ADODB.Connection"))
	check("com_constants() count", len(aConstants) > 400, 1)
	nValue = -1
	cGroup = ""
	for aConstant in aConstants
		if aConstant[2] = "adOpenStatic"
			nValue = aConstant[3]
			cGroup = aConstant[1]
		ok
	next
	check("com_constants() value", nValue, adOpenStatic)
	check("com_constants() group", cGroup, "CursorTypeEnum")
	cADO = sysget("CommonProgramFiles") + "\System\ado\msado15.dll"
	check("com_constants() from a file", len(com_constants(cADO)), len(aConstants))
	checkError("com_constants() bad file", 'com_constants("nofile.dll")', "nofile.dll")
	checkError("com_constants() bad parameter", "com_constants(5)", "type library file name")
	check("com_createobject() with an empty server", com_get(com_createobject("Scripting.Dictionary", ""), "Count"), 0)
	checkError("com_createobject() server: in-process class", 'com_createobject("Scripting.Dictionary", sysget("COMPUTERNAME"))', "0x80040154")
	checkError("com_createobject() server must be a string", 'com_createobject("Scripting.Dictionary", 5)', "Bad parameter type")
	checkError("com_createobject() parameters count", 'com_createobject("Scripting.Dictionary", "", "")', "Bad parameters count")
	oArrayOwner = com_createobject("Scripting.Dictionary")
	com_set(oArrayOwner, "Item", "k", com_variant(VT_ARRAY + VT_DISPATCH, [oOuter, oOuter]))
	aObjects = com_get(oArrayOwner, "Item", "k")
	check("Typed array VT_DISPATCH", com_issame(aObjects[2], oOuter), 1)
	check("CLSID string", com_typename(com_createobject("{EE09B103-97E0-11CF-978F-00A02463E06F}")), "IDictionary")

func checkRefCountInFunction oOuter, oInner
	x = com_get(oOuter, 'Item("inner")')
	aList = [x, oInner, com_invoke(oOuter, "Items")]
	check("Reference count in function", com_refcount(oInner), 4)

func testCOMObjectClass
	? nl + "COMObject class" + nl
	oDict = new COMObject("Scripting.Dictionary")
	oDict.invoke("Add", ["a", 1])
	oDict.setPropArgs("Item", ["b"], 2)
	check("invoke() and getProp()", oDict.getProp("Count"), 2)
	check("getPropArgs()", oDict.getPropArgs("Item", ["b"]), 2)
	oParent = new COMObject("Scripting.Dictionary")
	oParent.invoke("Add", ["child", oDict])
	oChild = oParent.getPropArgs("Item", ["child"])
	check("Returns COMObject", lower(classname(oChild)), "comobject")
	check("Passed COMObject", oChild.getProp('Item("a")'), 1)
	oChild.setProp('Item("a")', 10)
	check("setProp()", oDict.getProp('Item("a")'), 10)
	check("typeName()", oDict.typeName(), "IDictionary")
	check("items()", oDict.items(), ["a", "b"])
	check("isValid()", oDict.isValid(), 1)
	check("pointer()", com_isobject(oDict.pointer()), 1)
	oFSO = new COMObject("Scripting.FileSystemObject")
	aDrives = oFSO.getProp("Drives").items()
	check("items() returns COMObject", lower(classname(aDrives[1])), "comobject")
	oDict.release()
	check("release()", oDict.isValid(), 0)
	oNew = new COMObject(com_createobject("Scripting.Dictionary"))
	check("COMObject from pointer", oNew.getProp("Count"), 0)
	checkError("Bad init value", "o = new COMObject(5)", "COMObject expects")
	checkError("Arguments must be a list", 'o = new COMObject("Scripting.Dictionary") o.invoke("Add", "a")', "must be a list")

func testErrors
	? nl + "Errors" + nl
	checkError("ProgID doesn't exist", 'com_createobject("NoSuch.Object")', "NoSuch.Object")
	checkError("Member doesn't exist", 'com_get(com_createobject("Scripting.Dictionary"), "NoSuchMember")', "NoSuchMember : ")
	checkError("Error from the object", 'o = com_createobject("Scripting.Dictionary") com_invoke(o, "Add", 1, 1) com_invoke(o, "Add", 1, 1)', "Error number 457")
	check("com_lasthresult()", com_lasthresult(), 2148139465)
	checkError("Not an object", 'com_get("x", "Count")', "must be a COM object")
	checkError("Name is not a string", 'com_get(com_createobject("Scripting.Dictionary"), 5)', "must be a string")
	checkError("com_set() without value", 'com_set(com_createobject("Scripting.Dictionary"), "Item")', "Bad parameters count")
	checkError("Ring object as argument", 'o = com_createobject("Scripting.Dictionary") com_invoke(o, "Add", "k", new COMObject("Scripting.Dictionary"))', "Ring objects can't be passed")
	checkError("Wrong arguments count", 'com_invoke(com_createobject("Scripting.Dictionary"), "Add", 1)', "Add : ")
	checkError("com_getobject() bad moniker", 'com_getobject("nosuchmoniker:xyz")', "nosuchmoniker")
	checkError("com_getactiveobject() not running", 'com_getactiveobject("Scripting.Dictionary")', "Scripting.Dictionary")

func testAccess
	? nl + "Microsoft Access" + nl
	oConnection = accessCreateDatabase("unit_tests")
	check("Connection state", com_get(oConnection, "State"), adStateOpen)
	com_invoke(oConnection, "Execute", "CREATE TABLE T (ID COUNTER PRIMARY KEY, S TEXT(50), M MEMO, I INTEGER, L LONG,
		D DOUBLE, C CURRENCY, DT DATETIME, B YESNO, BIN LONGBINARY)")
	oRecordset = com_createobject("ADODB.Recordset")
	com_invoke(oRecordset, "Open", "T", oConnection, adOpenKeyset, adLockOptimistic, adCmdTable)
	cBytes = char(0) + char(10) + char(255)
	com_invoke(oRecordset, "AddNew", ["S", "M", "I", "L", "D", "C", "DT", "B"],
		["Текст", copy("x", 5000), 123, 2000000000, 1.25, 9.99, com_date("2026-08-09 10:11:12"), com_bool(1)])
	com_invoke(oRecordset, "AddNew")
	com_set(oRecordset, 'Fields("BIN").Value', com_bytes(cBytes))
	com_invoke(oRecordset, "Update")
	com_invoke(oRecordset, "Close")
	oRecordset = com_invoke(oConnection, "Execute", "SELECT * FROM T ORDER BY ID")
	check("Text field", com_get(oRecordset, 'Fields("S").Value'), "Текст")
	check("Memo field", len(com_get(oRecordset, 'Fields("M").Value')), 5000)
	check("Integer field", com_get(oRecordset, 'Fields("I").Value'), 123)
	check("Long field", com_get(oRecordset, 'Fields("L").Value'), 2000000000)
	check("Double field", com_get(oRecordset, 'Fields("D").Value'), 1.25)
	check("Currency field", com_get(oRecordset, 'Fields("C").Value'), 9.99)
	check("Date field", com_get(oRecordset, 'Fields("DT").Value'), "2026-08-09 10:11:12")
	check("Yes/No field", com_get(oRecordset, 'Fields("B").Value'), 1)
	com_get(oRecordset, 'Fields("BIN").Value')
	check("Null field", com_wasnull(), 1)
	com_invoke(oRecordset, "MoveNext")
	check("Binary field", com_get(oRecordset, 'Fields("BIN").Value'), cBytes)
	com_get(oRecordset, 'Fields("S").Value')
	check("Null text field", com_wasnull(), 1)
	com_invoke(oRecordset, "MoveFirst")
	aColumns = com_invoke(oRecordset, "GetRows", adGetRowsRest, com_missing(), ["ID", "I"])
	check("GetRows()", aColumns, [[1, 2], [123, ""]])
	check("com_rowsfromcolumns()", com_rowsfromcolumns(aColumns), [[1, 123], [2, ""]])
	com_invoke(oRecordset, "Close")
	cError = ""
	try
		com_invoke(oConnection, "Execute", "SELECT * FROM NoSuchTable")
	catch
		cError = cCatchError
	done
	check("SQL error", substr(cError, "NoSuchTable") > 0, 1)
	com_invoke(oConnection, "Close")
	check("Connection closed", com_get(oConnection, "State"), adStateClosed)
	accessDeleteDatabase("unit_tests")

func testByRefAndNamed
	? nl + "ByRef and named arguments" + nl
	oRef = com_byref(5)
	check("com_byrefvalue() initial value", com_byrefvalue(oRef), 5)
	oDict = com_createobject("Scripting.Dictionary")
	com_invoke(oDict, "Add", "k", oRef)
	check("ByRef passed to a member", com_get(oDict, 'Item("k")'), 5)
	oRef2 = com_byref([1, "a"])
	check("com_byref() list", com_byrefvalue(oRef2), [1, "a"])
	check("ByRef inside a list is copied", roundTrip([oRef]), [5])
	com_invoke(oDict, "Add", com_named("Key", "n1"), com_named("Item", 10))
	check("Named arguments", com_get(oDict, 'Item("n1")'), 10)
	com_invoke(oDict, "Add", com_named("Item", 30), com_named("Key", "n3"))
	check("Named arguments (other order)", com_get(oDict, 'Item("n3")'), 30)
	com_invoke(oDict, "Add", "n2", com_named("Item", 20))
	check("Positional then named", com_get(oDict, 'Item("n2")'), 20)
	check("Named argument (Exists)", com_invoke(oDict, "Exists", com_named("Key", "n2")), 1)
	checkError("Named before positional", 'o = com_createobject("Scripting.Dictionary") com_invoke(o, "Add", com_named("Key", "x"), 1)', "must be after")
	checkError("Unknown argument name", 'o = com_createobject("Scripting.Dictionary") com_invoke(o, "Add", com_named("NoSuchName", "x"), com_named("Item", 1))', "Add : ")
	checkError("Named argument in com_set()", 'o = com_createobject("Scripting.Dictionary") com_set(o, "Item", "x", com_named("Value", 1))', "can't be used to set")
	checkError("Named argument in a list", 'o = com_createobject("Scripting.Dictionary") com_invoke(o, "Add", "x", [com_named("Item", 1)])', "inside lists")
	checkError("Empty argument name", 'com_named("", 1)', "must be a string")
	checkError("com_byrefvalue() bad parameter", 'com_byrefvalue(5)', "created by com_byref()")
	checkError("Released ByRef", 'o = com_byref(1) com_release(o) com_invoke(com_createobject("Scripting.Dictionary"), "Add", "x", o)', "was released")
	check("COMObject with named arguments", testNamedCOMObject(), 7)

func testNamedCOMObject
	oDict = new COMObject("Scripting.Dictionary")
	oDict.invoke("Add", [com_named("Key", "a"), com_named("Item", 7)])
	return oDict.getProp('Item("a")')

func testEvents
	? nl + "Events (without an event)" + nl
	checkError("com_eventargs() outside events", "com_eventargs()", "no active event")
	checkError("com_eventname() outside events", "com_eventname()", "no active event")
	checkError("com_seteventarg() outside events", "com_seteventarg(1, 1)", "no active event")
	checkError("Object without events", 'com_connectevents(com_createobject("Scripting.Dictionary"))', "com_connectevents()")
	checkError("com_onevent() bad parameter", 'com_onevent(1, "x", "y")', "created by com_connectevents()")
	check("com_processmessages()", isNumber(com_processmessages()), 1)
	check("com_processmessages(10)", isNumber(com_processmessages(10)), 1)

func testAccessEvents
	? nl + "Microsoft Access events" + nl
	oConnection = accessCreateDatabase("unit_events")
	com_invoke(oConnection, "Execute", "CREATE TABLE T (ID INTEGER)")
	nRefCount = com_refcount(oConnection)
	runEventTests(oConnection)
	# Ring keeps the temporary values of a function until it returns, so we check after runEventTests()
	check("Reference count after the events", com_refcount(oConnection), nRefCount)
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("unit_events")

func runEventTests oConnection
	oEventConnection = oConnection
	oEvents = com_connectevents(oConnection)
	aNames = com_eventnames(oEvents)
	check("com_eventnames()", find(aNames, "WillExecute") > 0, 1)
	com_onevent(oEvents, "WillExecute", "aEventLog + ('Will:' + com_eventargs()[1])")
	com_onevent(oEvents, "ExecuteComplete", "aEventLog + ('Complete:' + com_eventargs()[1])")
	com_invoke(oConnection, "Execute", "INSERT INTO T VALUES (1)")
	check("Events", aEventLog, ["Will:INSERT INTO T VALUES (1)", "Complete:1"])
	aEventLog = []
	com_onevent(oEvents, "WillExecute", "if substr(com_eventargs()[1], 'DROP') com_seteventarg(5, adStatusCancel) ok")
	cError = ""
	try
		com_invoke(oConnection, "Execute", "DROP TABLE T")
	catch
		cError = cCatchError
	done
	check("Cancel using com_seteventarg()", substr(cError, "0x800A0E80") > 0, 1)
	aEventLog = []
	com_onevent(oEvents, "WillExecute", "")
	com_onevent(oEvents, "ExecuteComplete", "")
	com_onevent(oEvents, "*", "aEventLog + com_eventname()")
	com_invoke(oConnection, "Execute", "SELECT * FROM T")
	check("Handler for all the events", aEventLog, ["WillExecute", "ExecuteComplete"])
	aEventLog = []
	com_disconnectevents(oEvents)
	com_invoke(oConnection, "Execute", "SELECT * FROM T")
	check("No events after disconnect", aEventLog, [])
	checkError("com_onevent() after disconnect", 'o = com_connectevents(oEventConnection) com_disconnectevents(o) com_onevent(o, "WillExecute", "x = 1")', "disconnected")
	# Connect and release many times
	for nIndex = 1 to 50
		oTemp = com_connectevents(oConnection)
		com_onevent(oTemp, "WillExecute", "aEventLog + 1")
		com_release(oTemp)
	next
	com_invoke(oConnection, "Execute", "SELECT * FROM T")
	check("No events after com_release()", aEventLog, [])
	oEvents = new COMEvents(oConnection)
	oEvents.onEvent("ExecuteComplete", "aEventLog + 'class'")
	com_invoke(oConnection, "Execute", "SELECT * FROM T")
	check("COMEvents class", aEventLog, ["class"])
	oEvents.disconnect()
	oEvents = new COMEvents([oConnection, "ConnectionEvents"])
	check("COMEvents with the interface name", len(oEvents.names()), 9)
	oEvents.disconnect()
	checkError("Unknown events interface", 'com_connectevents(oEventConnection, "NoSuchEvents")', "NoSuchEvents")
	oByRefEvents = com_connectevents(oConnection)
	oByRefInEvent = com_byref(0)
	com_onevent(oByRefEvents, "WillExecute", "com_release(oByRefInEvent)")
	com_invoke(oConnection, "Execute", "SELECT * FROM T", oByRefInEvent)
	check("ByRef released during the call", com_isobject(oByRefInEvent), 0)
	com_disconnectevents(oByRefEvents)
	oCloseEvents = com_connectevents(oEventConnection)
	com_onevent(oCloseEvents, "Disconnect", "nZero = 0 x = 1 / nZero")
	cError = ""
	oSecond = accessOpenDatabase("unit_events")
	oSecondEvents = com_connectevents(oSecond)
	com_onevent(oSecondEvents, "Disconnect", "nZero = 0 x = 1 / nZero")
	try
		com_invoke(oSecond, "Close")
	catch
		cError = cCatchError
	done
	check("Error in the Disconnect event is raised by Close", substr(cError, "Event Disconnect") > 0, 1)
	check("No pending error after it was raised", com_get(oConnection, "State"), adStateOpen)
	com_disconnectevents(oCloseEvents)
	com_disconnectevents(oSecondEvents)
	oHandlerEvents = com_connectevents(oConnection)
	com_onevent(oHandlerEvents, "WillExecute", "nZero = 0 x = 1 / nZero")
	cError = ""
	try
		com_invoke(oConnection, "Execute", "SELECT * FROM T")
	catch
		cError = cCatchError
	done
	check("Error in an event handler is raised by the call", substr(cError, "Event WillExecute") > 0, 1)
	com_disconnectevents(oHandlerEvents)
	check("Calls work after an error in a handler", com_get(oConnection, "State"), adStateOpen)
	checkError("com_seteventarg() bad index", 'o = com_connectevents(oEventConnection) com_onevent(o, "WillExecute", "com_seteventarg(99, 1)") com_invoke(oEventConnection, "Execute", "SELECT * FROM T")', "Bad event argument index")
	oEventConnection = NULL

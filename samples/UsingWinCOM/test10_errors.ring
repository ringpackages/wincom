# COM errors are Ring errors, so we can catch them using try/catch
# The error message contains the member name, the error source, the description and the HRESULT code

load "wincom.ring"
load "accesshelper.ring"

func main
	# A ProgID that doesn't exist
	try
		com_createobject("NoSuch.Object")
	catch
		? cCatchError
	done

	# A member that doesn't exist
	oDict = com_createobject("Scripting.Dictionary")
	try
		com_get(oDict, "NoSuchProperty")
	catch
		? cCatchError
	done

	# An error raised by the object (Error number 457: the key already exists)
	com_invoke(oDict, "Add", "key", 1)
	try
		com_invoke(oDict, "Add", "key", 2)
	catch
		? cCatchError
		? "com_lasthresult(): " + upper(hex(com_lasthresult()))
	done

	# Database errors (The description comes from the database engine)
	oConnection = accessCreateDatabase("sample_errors")
	com_invoke(oConnection, "Execute", "CREATE TABLE Items (ID INTEGER PRIMARY KEY, Name TEXT(10) NOT NULL)")
	aStatements = [
		"SELECT * FROM NoSuchTable",
		"SELEC Name FROM Items",
		"INSERT INTO Items (ID) VALUES (1)",
		"INSERT INTO Items (ID, Name) VALUES (1, 'This name is too long')"
	]
	for cSQL in aStatements
		try
			com_invoke(oConnection, "Execute", cSQL)
		catch
			? nl + "SQL   : " + cSQL
			? "Error : " + cCatchError
			# The Errors collection of the connection has the details
			aErrors = com_items(com_get(oConnection, "Errors"))
			for oError in aErrors
				? "State : " + com_get(oError, "SQLState") + ", Native error: " + com_get(oError, "NativeError")
			next
			com_invoke(oConnection, "Errors.Clear")
		done
	next

	# A path that is not valid
	try
		com_get(oConnection, "Properties(")
	catch
		? nl + cCatchError
	done

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_errors")

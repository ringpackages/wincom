# ADO events: log the SQL statements and cancel some of them
# com_connectevents(), com_onevent(), com_eventargs(), com_seteventarg()

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_events")
	accessCreateProducts(oConnection)

	# The default events interface of ADODB.Connection is ConnectionEvents
	oEvents = com_connectevents(oConnection)
	? "Events of ADODB.Connection:"
	aNames = com_eventnames(oEvents)
	for cName in aNames
		? "	" + cName
	next

	# WillExecute(Source, CursorType, LockType, Options, adStatus, pCommand, pRecordset, pConnection)
	com_onevent(oEvents, "WillExecute", "onWillExecute()")
	# ExecuteComplete(RecordsAffected, pError, adStatus, pCommand, pRecordset, pConnection)
	com_onevent(oEvents, "ExecuteComplete", "onExecuteComplete()")

	? nl + "Executing the statements:" + nl
	aStatements = [
		"UPDATE Products SET Price = Price + 1 WHERE Category = 'Category 1'",
		"DELETE FROM Products WHERE Quantity = 0",
		"DROP TABLE Products",
		"SELECT COUNT(*) FROM Products"
	]
	for cSQL in aStatements
		try
			com_invoke(oConnection, "Execute", cSQL)
		catch
			? "	Error: " + cCatchError
		done
	next

	# Transactions events, using one handler for all the events ("*")
	com_onevent(oEvents, "WillExecute", "")
	com_onevent(oEvents, "ExecuteComplete", "")
	com_onevent(oEvents, "*", "? '	Event: ' + com_eventname()")
	? nl + "Transaction:" + nl
	com_invoke(oConnection, "BeginTrans")
	com_invoke(oConnection, "Execute", "UPDATE Products SET Quantity = Quantity + 1")
	com_invoke(oConnection, "CommitTrans")

	# Using the COMEvents class
	com_disconnectevents(oEvents)
	oEvents2 = new COMEvents(oConnection)
	oEvents2.onEvent("Disconnect", "? nl + 'The connection is closed'")
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_events")

func onWillExecute
	aArgs = com_eventargs()
	cSQL = aArgs[1]
	? "	WillExecute     : " + cSQL
	# Don't allow DROP statements (adStatus is the fifth argument, ByRef)
	if substr(upper(cSQL), "DROP ")
		? "	                  Cancelled"
		com_seteventarg(5, adStatusCancel)
	ok

func onExecuteComplete
	aArgs = com_eventargs()
	? "	ExecuteComplete : " + aArgs[1] + " records affected (Status: " + aArgs[3] + ")"

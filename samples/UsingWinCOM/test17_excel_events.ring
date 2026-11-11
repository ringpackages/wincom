# Microsoft Excel events: the Ring code runs when cells are changed
# Needs Microsoft Excel

load "wincom.ring"

# How many times the workbook received BeforeClose
nCloseCount = 0

func main
	if ! com_progidexists("Excel.Application")
		? "This sample needs Microsoft Excel"
		return
	ok
	oExcel = com_createobject("Excel.Application")
	com_set(oExcel, "DisplayAlerts", com_bool(0))
	try
		# The default events interface of Excel.Application is AppEvents
		oAppEvents = new COMEvents(oExcel)
		? "Excel.Application has " + len(oAppEvents.names()) + " events"
		oAppEvents.onEvent("NewWorkbook", "? 'NewWorkbook: ' + com_get(com_eventargs()[1], 'Name')")
		oAppEvents.onEvent("SheetChange", "onSheetChange()")

		oWorkbook = com_invoke(oExcel, "Workbooks.Add")
		oSheet = com_get(oWorkbook, "Worksheets(1)")

		# The events of the workbook (WorkbookEvents interface)
		oBookEvents = new COMEvents([oWorkbook, "WorkbookEvents"])
		oBookEvents.onEvent("BeforeClose", "onBeforeClose()")

		? nl + "Changing cells:"
		com_set(oSheet, 'Range("A1").Value', "Total")
		com_set(oSheet, 'Range("B1:B3").Value', 10)
		com_set(oSheet, 'Range("B4").Formula', "=SUM(B1:B3)")
		? "B4 = " + com_get(oSheet, 'Range("B4").Value')

		# Events from other programs are delivered while we call their methods,
		# or when we process the messages (Like DoEvents in Visual Basic)
		com_processmessages(100)

		# The first time BeforeClose cancels the close
		? nl + "Closing the workbook:"
		com_invoke(oWorkbook, "Close", com_bool(0))
		? "Workbooks: " + com_get(oExcel, "Workbooks.Count")
		com_invoke(oWorkbook, "Close", com_bool(0))
		? "Workbooks: " + com_get(oExcel, "Workbooks.Count")
	catch
		? cCatchError
	done
	com_invoke(oExcel, "Quit")

func onSheetChange
	# SheetChange(Sh As Object, Target As Range)
	aArgs = com_eventargs()
	? "SheetChange: " + com_get(aArgs[1], "Name") + "!" + com_get(aArgs[2], "Address") +
	  " = " + listValue(com_get(aArgs[2], "Value"))

func onBeforeClose
	# BeforeClose(Cancel As Boolean): set Cancel to True to keep the workbook open
	nCloseCount++
	if nCloseCount = 1
		? "BeforeClose: cancelled"
		com_seteventarg(1, com_bool(1))
	else
		? "BeforeClose: allowed"
	ok

func listValue vValue
	if isList(vValue)
		return "(" + len(vValue) + " rows)"
	ok
	return "" + vValue

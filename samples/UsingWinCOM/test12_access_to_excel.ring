# Export the data of a Microsoft Access database to a Microsoft Excel workbook
# Needs Microsoft Excel

load "wincom.ring"
load "accesshelper.ring"

func main
	if ! com_progidexists("Excel.Application")
		? "This sample needs Microsoft Excel"
		return
	ok

	oConnection = accessCreateDatabase("sample_excel")
	accessCreateProducts(oConnection)

	oExcel = com_createobject("Excel.Application")
	com_set(oExcel, "Visible", com_bool(0))
	com_set(oExcel, "DisplayAlerts", com_bool(0))
	try
		oWorkbook = com_invoke(oExcel, "Workbooks.Add")
		oSheet = com_get(oWorkbook, "Worksheets(1)")
		com_set(oSheet, "Name", "Products")

		# The header (Cells(Row, Column))
		oRecordset = com_invoke(oConnection, "Execute", "SELECT Name, Category, Price, Quantity, Added FROM Products ORDER BY ID")
		nFields = com_get(oRecordset, "Fields.Count")
		for nField = 1 to nFields
			com_set(oSheet, "Cells(1, " + nField + ").Value", com_get(oRecordset, "Fields(" + (nField - 1) + ").Name"))
		next
		com_set(oSheet, 'Range("A1:E1").Font.Bold', com_bool(1))

		# Range.CopyFromRecordset() copies all the records in one call
		nRows = com_invoke(oSheet, 'Range("A2").CopyFromRecordset', oRecordset)
		com_invoke(oRecordset, "Close")
		? "Records copied: " + nRows

		# A formula
		nTotalRow = nRows + 2
		com_set(oSheet, "Cells(" + nTotalRow + ", 1).Value", "Total")
		com_set(oSheet, "Cells(" + nTotalRow + ", 4).Formula", "=SUM(D2:D" + (nRows + 1) + ")")
		? "Total quantity (calculated by Excel): " + com_get(oSheet, "Cells(" + nTotalRow + ", 4).Value")

		# Read a range as a two dimensions list
		aValues = com_get(oSheet, 'Range("A2:C4").Value')
		? nl + "Range A2:C4:"
		for aRow in aValues
			? "	" + aRow[1] + " | " + aRow[2] + " | " + aRow[3]
		next

		com_invoke(oSheet, "Columns.AutoFit")
		cFile = currentdir() + "\products.xlsx"
		if fexists(cFile)
			remove(cFile)
		ok
		com_invoke(oWorkbook, "SaveAs", cFile)
		com_invoke(oWorkbook, "Close", com_bool(0))
		? nl + "Saved: products.xlsx (" + len(read(cFile)) + " bytes)"
	catch
		? cCatchError
	done
	com_invoke(oExcel, "Quit")

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_excel")

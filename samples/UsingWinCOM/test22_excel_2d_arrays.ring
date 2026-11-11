# Fast export to Microsoft Excel using two dimensions arrays (com_matrix())
# One call writes all the cells, and one call reads them (Range.Value)
# Needs Microsoft Excel

load "wincom.ring"
load "accesshelper.ring"

func main
	if ! com_progidexists("Excel.Application")
		? "This sample needs Microsoft Excel"
		return
	ok

	# Read the data from the Access database
	oConnection = accessCreateDatabase("sample_excel2d")
	accessCreateProducts(oConnection)
	oRecordset = com_invoke(oConnection, "Execute",
		"SELECT Name, Category, Price, Quantity, Price * Quantity AS StockValue FROM Products ORDER BY Name")
	aRows = com_rowsfromcolumns(com_invoke(oRecordset, "GetRows"))
	com_invoke(oRecordset, "Close")
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_excel2d")

	# The header is the first row
	aData = [["Name", "Category", "Price", "Quantity", "Stock Value"]]
	for aRow in aRows
		aData + aRow
	next
	nRows = len(aData)

	oExcel = com_createobject("Excel.Application")
	com_set(oExcel, "DisplayAlerts", com_bool(0))
	try
		oWorkbook = com_invoke(oExcel, "Workbooks.Add")
		oSheet = com_get(oWorkbook, "Worksheets(1)")
		com_set(oSheet, "Name", "Products")

		# Write all the cells in one call: Range.Value = two dimensions array
		cRange = "A1:E" + nRows
		com_set(oSheet, 'Range("' + cRange + '").Value', com_matrix(aData))
		? "Cells written: " + com_get(oSheet, 'Range("' + cRange + '").Count')

		# Formats and a total row
		com_set(oSheet, 'Range("A1:E1").Font.Bold', com_bool(1))
		com_set(oSheet, 'Range("C2:C' + nRows + '").NumberFormat', "0.00")
		com_set(oSheet, 'Range("E2:E' + (nRows + 1) + '").NumberFormat', "#,##0.00")
		com_set(oSheet, 'Range("A' + (nRows + 1) + '").Value', "Total")
		com_set(oSheet, 'Range("E' + (nRows + 1) + '").Formula', "=SUM(E2:E" + nRows + ")")
		com_invoke(oSheet, "Columns.AutoFit")

		# A chart: Shapes.AddChart2(Style, XlChartType (57 = xlBarClustered))
		oChart = com_invoke(oSheet, "Shapes.AddChart2", 216, 57, 320, 10, 360, 220)
		com_invoke(oChart, "Chart.SetSourceData", com_get(oSheet, 'Range("A1:A' + nRows + ',E1:E' + nRows + '")'))
		com_set(oChart, "Chart.HasTitle", com_bool(1))
		com_set(oChart, "Chart.ChartTitle.Text", "Stock Value")

		# Read all the cells in one call: aValues[nRow][nColumn]
		aValues = com_get(oSheet, 'Range("A1:E' + (nRows + 1) + '").Value')
		? nl + "The cells (read in one call):" + nl
		for aRow in aValues
			cLine = ""
			for vValue in aRow
				cLine += left("" + vValue + copy(" ", 14), 14)
			next
			? cLine
		next

		cFile = currentdir() + "\products_2d.xlsx"
		if fexists(cFile)
			remove(cFile)
		ok
		com_invoke(oWorkbook, "SaveAs", cFile)
		com_invoke(oWorkbook, "Close", com_bool(0))
		? nl + "Saved: products_2d.xlsx"
	catch
		? cCatchError
	done
	com_invoke(oExcel, "Quit")

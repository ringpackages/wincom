# Create a Microsoft Word report from the data of a Microsoft Access database
# Needs Microsoft Word

load "wincom.ring"
load "accesshelper.ring"

# Word constants
wdStyleHeading1  = -2
wdStyleNormal    = -1
wdAlignCenter    = 1
wdAutoFitContent = 1
wdFormatDocument = 16
wdFormatPDF      = 17

func main
	if ! com_progidexists("Word.Application")
		? "This sample needs Microsoft Word"
		return
	ok

	oConnection = accessCreateDatabase("sample_word")
	accessCreateProducts(oConnection)
	oRecordset = com_invoke(oConnection, "Execute", "SELECT Name, Category, Price, Quantity FROM Products ORDER BY Category, Name")
	aRows = com_rowsfromcolumns(com_invoke(oRecordset, "GetRows"))
	com_invoke(oRecordset, "Close")
	oRecordset = com_invoke(oConnection, "Execute", "SELECT SUM(Price * Quantity) FROM Products")
	nTotal = com_get(oRecordset, "Fields(0).Value")
	com_invoke(oRecordset, "Close")
	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_word")

	oWord = com_createobject("Word.Application")
	com_set(oWord, "Visible", com_bool(0))
	try
		oDocument = com_invoke(oWord, "Documents.Add")

		# The title (A paragraph using the Heading 1 style)
		oRange = com_invoke(oDocument, "Range")
		com_set(oRange, "Text", "Products Report")
		com_set(oRange, "Style", wdStyleHeading1)
		com_invoke(oRange, "InsertParagraphAfter")

		oRange = com_invoke(oDocument, "Range", com_get(oDocument, "Content.End") - 1)
		com_set(oRange, "Text", "Products: " + len(aRows) + ", stock value: " + nTotal)
		com_set(oRange, "Style", wdStyleNormal)
		com_invoke(oRange, "InsertParagraphAfter")

		# A table: Tables.Add(Range, Rows, Columns)
		oRange = com_invoke(oDocument, "Range", com_get(oDocument, "Content.End") - 1)
		oTable = com_invoke(oDocument, "Tables.Add", oRange, len(aRows) + 1, 4)
		com_set(oTable, "Borders.Enable", com_bool(1))
		aHeader = ["Name", "Category", "Price", "Quantity"]
		for nColumn = 1 to 4
			com_set(oTable, "Cell(1, " + nColumn + ").Range.Text", aHeader[nColumn])
		next
		com_set(oTable, "Rows(1).Range.Font.Bold", com_bool(1))
		for nRow = 1 to len(aRows)
			for nColumn = 1 to 4
				com_set(oTable, "Cell(" + (nRow + 1) + ", " + nColumn + ").Range.Text", "" + aRows[nRow][nColumn])
			next
		next
		# AutoFitBehavior(wdAutoFitContent)
		com_invoke(oTable, "AutoFitBehavior", wdAutoFitContent)
		com_set(oTable, "Rows.Alignment", wdAlignCenter)

		? "Paragraphs : " + com_get(oDocument, "Paragraphs.Count")
		? "Tables     : " + com_get(oDocument, "Tables.Count")
		? "Words      : " + com_get(oDocument, "Words.Count")

		# Save as .docx and .pdf (SaveAs2(FileName, FileFormat))
		cDocx = currentdir() + "\products_report.docx"
		cPDF = currentdir() + "\products_report.pdf"
		for cFile in [cDocx, cPDF]
			if fexists(cFile)
				remove(cFile)
			ok
		next
		com_invoke(oDocument, "SaveAs2", cDocx, wdFormatDocument)
		com_invoke(oDocument, "SaveAs2", cPDF, wdFormatPDF)
		com_invoke(oDocument, "Close", com_bool(0))
		? nl + "Saved: products_report.docx (" + floor(getfilesize(cDocx) / 1024) + " KB)"
		? "Saved: products_report.pdf (" + floor(getfilesize(cPDF) / 1024) + " KB)"
	catch
		? cCatchError
	done
	com_invoke(oWord, "Quit", com_bool(0))

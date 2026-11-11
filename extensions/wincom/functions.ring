/*
	WinCOM - Functions
	The other com_*() functions are written in C (ring_wincom.c)
*/

/*
	The library version
*/
func com_version
	return WINCOM_VERSION

/*
	Values that don't exist in Ring (Pass them to com_set(), com_invoke(), ...)
*/

# Null (Like Null in Visual Basic, for example to store Null in a database field)
func com_null
	return com_variant(VT_NULL, 0)

# Empty (Like Empty in Visual Basic)
func com_empty
	return com_variant(VT_EMPTY, 0)

# Missing optional parameter (Skip a parameter and pass the next one)
func com_missing
	return com_variant(VT_ERROR, 0)

# Boolean (True/False), lValue is a number (0 = False, other values = True)
func com_bool lValue
	return com_variant(VT_BOOL, lValue)

# Date, cDate is a string: "YYYY-MM-DD" or "YYYY-MM-DD HH:MM:SS"
func com_date cDate
	return com_variant(VT_DATE, cDate)

# Currency (Four digits after the decimal point)
func com_currency nValue
	return com_variant(VT_CY, nValue)

# Double (Numbers without fractions are passed as Long (VT_I4) by default)
func com_double nValue
	return com_variant(VT_R8, nValue)

# Binary data (Byte array), cData is a Ring string
func com_bytes cData
	return com_variant(VT_ARRAY + VT_UI1, cData)

/*
	Information about the last value returned by com_get(), com_invoke() or com_dispatch()
*/

# Returns 1 if the last value was Null (Null is returned as "")
func com_wasnull
	return com_lasttype() = VT_NULL

/*
	Convert ADO Recordset.GetRows() output (aData[nField][nRow]) to a list of rows (aRows[nRow][nField])
*/
func com_rowsfromcolumns aColumns
	aRows = []
	nColumns = len(aColumns)
	if nColumns = 0
		return aRows
	ok
	nRows = len(aColumns[1])
	for nRow = 1 to nRows
		aRow = []
		for nColumn = 1 to nColumns
			aRow + aColumns[nColumn][nRow]
		next
		aRows + aRow
	next
	return aRows

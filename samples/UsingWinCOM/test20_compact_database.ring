# Compact a Microsoft Access database
#	.accdb files: DAO (DAO.DBEngine.120), installed with the Microsoft Access Database Engine (ACE)
#	.mdb files  : JRO (JRO.JetEngine), part of MS-Windows for 32-bit programs

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_compact")
	cFile = currentdir() + "\" + accessFileName("sample_compact")
	cCompacted = currentdir() + "\sample_compacted" + cAccessExtension
	if fexists(cCompacted)
		remove(cCompacted)
	ok
	com_invoke(oConnection, "Execute", "CREATE TABLE Data (ID COUNTER PRIMARY KEY, Content MEMO)")

	# Add then delete many records (The file keeps the space of the deleted records)
	oRecordset = com_createobject("ADODB.Recordset")
	com_invoke(oRecordset, "Open", "Data", oConnection, adOpenKeyset, adLockOptimistic, adCmdTable)
	cContent = copy("Sample data ", 400)
	for nIndex = 1 to 2000
		com_invoke(oRecordset, "AddNew", ["Content"], [cContent])
	next
	com_invoke(oRecordset, "Close")
	com_invoke(oConnection, "Execute", "DELETE FROM Data WHERE ID > 100")
	com_invoke(oConnection, "Close")
	? "Provider               : " + cAccessProvider
	? "Size before compacting : " + floor(getfilesize(cFile) / 1024) + " KB"

	try
		if cAccessExtension = ".accdb"
			compactUsingDAO(cFile, cCompacted)
		else
			compactUsingJRO(cFile, cCompacted)
		ok
		? "Size after compacting  : " + floor(getfilesize(cCompacted) / 1024) + " KB"

		# Check the compacted database
		oConnection = com_createobject("ADODB.Connection")
		com_invoke(oConnection, "Open", accessConnectionString(cAccessProvider, cCompacted))
		oRecordset = com_invoke(oConnection, "Execute", "SELECT COUNT(*) FROM Data")
		? "Records                : " + com_get(oRecordset, "Fields(0).Value")
		com_invoke(oRecordset, "Close")
		com_invoke(oConnection, "Close")
	catch
		? cCatchError
	done

	if fexists(cCompacted)
		remove(cCompacted)
	ok
	accessDeleteDatabase("sample_compact")

# DBEngine.CompactDatabase(SourceFile, DestinationFile)
func compactUsingDAO cFile, cCompacted
	if ! com_progidexists("DAO.DBEngine.120")
		raise("This sample needs DAO (DAO.DBEngine.120), installed with the Microsoft Access Database Engine")
	ok
	oDBEngine = com_createobject("DAO.DBEngine.120")
	com_invoke(oDBEngine, "CompactDatabase", cFile, cCompacted)

# JetEngine.CompactDatabase(SourceConnection, DestinationConnection)
func compactUsingJRO cFile, cCompacted
	if ! com_progidexists("JRO.JetEngine")
		raise("This sample needs JRO (JRO.JetEngine), available for 32-bit programs")
	ok
	oJetEngine = com_createobject("JRO.JetEngine")
	com_invoke(oJetEngine, "CompactDatabase", accessConnectionString(cAccessProvider, cFile),
		   accessConnectionString(cAccessProvider, cCompacted))

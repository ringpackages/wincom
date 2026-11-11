# Transactions: BeginTrans(), CommitTrans() and RollbackTrans()

load "wincom.ring"
load "accesshelper.ring"

func main
	oConnection = accessCreateDatabase("sample_transactions")
	com_invoke(oConnection, "Execute", "CREATE TABLE Accounts (ID INTEGER PRIMARY KEY, Name TEXT(30), Balance CURRENCY)")
	com_invoke(oConnection, "Execute", "INSERT INTO Accounts VALUES (1, 'Account 1', 1000)")
	com_invoke(oConnection, "Execute", "INSERT INTO Accounts VALUES (2, 'Account 2', 500)")
	printBalances(oConnection, "Start")

	# A transfer that succeeds
	transfer(oConnection, 1, 2, 300)
	printBalances(oConnection, "After transfer(1, 2, 300)")

	# A transfer that fails in the middle (Account 3 doesn't exist), so nothing is changed
	transfer(oConnection, 2, 3, 100)
	printBalances(oConnection, "After transfer(2, 3, 100)")

	# A transfer cancelled by a rule (The balance can't be negative)
	transfer(oConnection, 1, 2, 5000)
	printBalances(oConnection, "After transfer(1, 2, 5000)")

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_transactions")

func transfer oConnection, nFrom, nTo, nAmount
	com_invoke(oConnection, "BeginTrans")
	try
		com_invoke(oConnection, "Execute", "UPDATE Accounts SET Balance = Balance - " + nAmount + " WHERE ID = " + nFrom)
		com_invoke(oConnection, "Execute", "UPDATE Accounts SET Balance = Balance + " + nAmount + " WHERE ID = " + nTo)
		if queryCount(oConnection, "SELECT COUNT(*) FROM Accounts WHERE ID = " + nTo) = 0
			raise("Account " + nTo + " doesn't exist")
		ok
		if queryCount(oConnection, "SELECT COUNT(*) FROM Accounts WHERE Balance < 0") > 0
			raise("The balance can't be negative")
		ok
		com_invoke(oConnection, "CommitTrans")
		? "Transfer done"
	catch
		com_invoke(oConnection, "RollbackTrans")
		? "Transfer cancelled: " + cCatchError
	done

func queryCount oConnection, cSQL
	oRecordset = com_invoke(oConnection, "Execute", cSQL)
	nCount = com_get(oRecordset, "Fields(0).Value")
	com_invoke(oRecordset, "Close")
	return nCount

func printBalances oConnection, cTitle
	? nl + cTitle + ":"
	oRecordset = com_invoke(oConnection, "Execute", "SELECT Name, Balance FROM Accounts ORDER BY ID")
	while ! com_get(oRecordset, "EOF")
		? "	" + com_get(oRecordset, "Fields(0).Value") + " : " + com_get(oRecordset, "Fields(1).Value")
		com_invoke(oRecordset, "MoveNext")
	end
	com_invoke(oRecordset, "Close")
	? ""

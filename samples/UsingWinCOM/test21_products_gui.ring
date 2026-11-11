# A GUI application (RingQt) that uses a Microsoft Access database (ADO)
# Search, add, delete and edit products (Double click a Name or a Price to edit it)

load "wincom.ring"
load "accesshelper.ring"

# Load the Access Database Engine before RingQt
# Needed in Ring versions older than 1.28 (See accessLoadEngine() in accesshelper.ring)
accessLoadEngine()

load "guilib.ring"

# The event handlers use global variables
oConnection = NULL
oApp        = NULL
oWin        = NULL
oSearch     = NULL
oTable      = NULL
oStatus     = NULL
lLoading    = False
nNewProduct = 0

func main
	oConnection = accessCreateDatabase("sample_gui")
	accessCreateProducts(oConnection)

	oApp = new QApp {
		oWin = new QWidget() {
			setWindowTitle("WinCOM - Products (Microsoft Access)")
			resize(720, 420)

			oSearch = new QLineEdit(oWin) {
				setPlaceholderText("Search (Name or Category)")
				setTextChangedEvent("loadProducts()")
			}
			oTable = new QTableWidget(oWin) {
				setColumnCount(6)
				setSelectionBehavior(QAbstractItemView_SelectRows)
				setItemChangedEvent("updateProduct()")
				horizontalHeader().setStretchLastSection(True)
			}
			aHeader = ["ID", "Name", "Category", "Price", "Quantity", "Added"]
			for nColumn = 1 to len(aHeader)
				oTable.setHorizontalHeaderItem(nColumn - 1, new QTableWidgetItem(aHeader[nColumn]))
			next
			oAdd = new QPushButton(oWin) {
				setText("Add Product")
				setClickEvent("addProduct()")
			}
			oDelete = new QPushButton(oWin) {
				setText("Delete Product")
				setClickEvent("deleteProduct()")
			}
			oStatus = new QLabel(oWin) { setText("") }

			oButtons = new QHBoxLayout() {
				addWidget(oAdd)
				addWidget(oDelete)
				addWidget(oStatus)
			}
			oLayout = new QVBoxLayout() {
				addWidget(oSearch)
				addWidget(oTable)
				addLayout(oButtons)
			}
			setLayout(oLayout)
			loadProducts()
			show()
		}
		exec()
	}

	com_invoke(oConnection, "Close")
	accessDeleteDatabase("sample_gui")

/*
	Load the products that match the search text (Using a parameter)
*/
func loadProducts
	oCommand = com_createobject("ADODB.Command")
	com_setref(oCommand, "ActiveConnection", oConnection)
	com_set(oCommand, "CommandText", "SELECT ID, Name, Category, Price, Quantity, Added FROM Products
		WHERE Name LIKE ? OR Category LIKE ? ORDER BY ID")
	cPattern = "%" + oSearch.text() + "%"
	oRecordset = com_invoke(oCommand, "Execute", com_missing(), [cPattern, cPattern])
	aRows = []
	if ! com_get(oRecordset, "EOF")
		aRows = com_rowsfromcolumns(com_invoke(oRecordset, "GetRows"))
	ok
	com_invoke(oRecordset, "Close")

	lLoading = True
	oTable.setRowCount(len(aRows))
	for nRow = 1 to len(aRows)
		for nColumn = 1 to 6
			oItem = new QTableWidgetItem("" + aRows[nRow][nColumn])
			# Only Name and Price can be edited
			if nColumn != 2 and nColumn != 4
				oItem.setFlags(Qt_ItemIsSelectable | Qt_ItemIsEnabled)
			ok
			oTable.setItem(nRow - 1, nColumn - 1, oItem)
		next
	next
	oTable.resizeColumnsToContents()
	lLoading = False
	oStatus.setText("Products: " + len(aRows))

/*
	Add a new product (A new Command for each INSERT, see the notes about Microsoft Access)
*/
func addProduct
	nNewProduct++
	oCommand = com_createobject("ADODB.Command")
	com_setref(oCommand, "ActiveConnection", oConnection)
	com_set(oCommand, "CommandText", "INSERT INTO Products (Name, Category, Price, Quantity, Added, InStock) VALUES (?, ?, ?, ?, ?, ?)")
	# date() returns dd/mm/yyyy
	cToday = date()
	cToday = right(cToday, 4) + "-" + substr(cToday, 4, 2) + "-" + left(cToday, 2)
	com_invoke(oCommand, "Execute", com_missing(),
		["New Product " + nNewProduct, "Category 9", 1, 1, com_date(cToday), com_bool(1)])
	loadProducts()

/*
	Delete the selected product
*/
func deleteProduct
	nRow = oTable.currentRow()
	if nRow < 0
		oStatus.setText("Select a product")
		return
	ok
	nID = 0 + oTable.item(nRow, 0).text()
	com_invoke(oConnection, "Execute", "DELETE FROM Products WHERE ID = " + nID)
	loadProducts()

/*
	Save the edited Name or Price
*/
func updateProduct
	if lLoading
		return
	ok
	nRow = oTable.currentRow()
	if nRow < 0
		return
	ok
	nID = 0 + oTable.item(nRow, 0).text()
	cName = oTable.item(nRow, 1).text()
	cPrice = oTable.item(nRow, 3).text()
	if ! isdigit(substr(cPrice, ".", ""))
		oStatus.setText("The price must be a number")
		return
	ok
	oCommand = com_createobject("ADODB.Command")
	com_setref(oCommand, "ActiveConnection", oConnection)
	com_set(oCommand, "CommandText", "UPDATE Products SET Name = ?, Price = ? WHERE ID = ?")
	try
		com_invoke(oCommand, "Execute", com_missing(), [cName, com_currency(0 + cPrice), nID])
		oStatus.setText("Saved: " + cName)
	catch
		oStatus.setText(cCatchError)
	done

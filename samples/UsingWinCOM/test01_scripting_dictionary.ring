# Using a COM object: Scripting.Dictionary
# com_createobject(), com_invoke(), com_get(), com_set() and com_items()

load "wincom.ring"

func main
	oDict = com_createobject("Scripting.Dictionary")

	# Call methods (Arguments come after the method name)
	com_invoke(oDict, "Add", "apple", 5)
	com_invoke(oDict, "Add", "orange", 12)
	com_invoke(oDict, "Add", "Яблоко", "Fruit (Russian)")

	# Get properties
	? "Count       : " + com_get(oDict, "Count")
	? "Exists      : " + com_get(oDict, "Exists", "apple")

	# Properties with arguments: pass the arguments after the name, or write them in the path
	? "Item(apple) : " + com_get(oDict, "Item", "apple")
	? "Item(Яблоко): " + com_get(oDict, 'Item("Яблоко")')

	# Set properties (The value is the last parameter)
	com_set(oDict, "Item", "apple", 7)
	com_set(oDict, 'Item("banana")', 3)
	? "Item(apple) after com_set(): " + com_get(oDict, 'Item("apple")')

	# Arrays are returned as Ring lists
	aKeys = com_invoke(oDict, "Keys")
	aValues = com_invoke(oDict, "Items")
	? nl + "Keys and values:"
	nCount = len(aKeys)
	for nIndex = 1 to nCount
		? "	" + aKeys[nIndex] + " = " + aValues[nIndex]
	next

	# Enumerate the collection (Like For Each in Visual Basic)
	? nl + "com_items():"
	aItems = com_items(oDict)
	for cKey in aItems
		? "	" + cKey
	next

	com_invoke(oDict, "Remove", "orange")
	? nl + "Count after Remove(): " + com_get(oDict, "Count")
	com_invoke(oDict, "RemoveAll")
	? "Count after RemoveAll(): " + com_get(oDict, "Count")

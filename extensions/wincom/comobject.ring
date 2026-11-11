/*
	WinCOM - COMObject Class
	Using COM objects with Object-Oriented Programming

	oConnection = new COMObject("ADODB.Connection")
	oConnection.setProp("ConnectionString", cConnection)
	oConnection.invoke("Open", [])
	oRecordset = oConnection.invoke("Execute", ["SELECT * FROM Products"])
	? oRecordset.getProp('Fields("Name").Value')

	COM objects returned by the methods are COMObject objects too
	Arguments can be Ring values, COMObject objects, or values created by com_null(), com_date(), ...
*/

class COMObject

	pObject = NULL

	/*
		vValue is a ProgID ("ADODB.Connection"), a CLSID ("{...}") or a COM object (C pointer)
	*/
	func init vValue
		if isString(vValue)
			pObject = com_createobject(vValue)
		elseif com_isobject(vValue)
			pObject = vValue
		else
			raise("Error, WinCOM: COMObject expects a ProgID or a COM object")
		ok
		return self

	/*
		Get a property value, cPath can be "Name", "Fields.Count", 'Fields("Name").Value', ...
	*/
	func getProp cPath
		return wrapValue(com_dispatch(pObject, cPath, COM_GET, []))

	/*
		Get a property that has arguments, like Item(nIndex) or Cells(nRow, nColumn)
	*/
	func getPropArgs cPath, aArgs
		return wrapValue(com_dispatch(pObject, cPath, COM_GET, unwrapList(aArgs)))

	/*
		Set a property value
	*/
	func setProp cPath, vValue
		com_dispatch(pObject, cPath, COM_SET, [unwrapValue(vValue)])
		return self

	/*
		Set a property that has arguments, like Item(cKey)
	*/
	func setPropArgs cPath, aArgs, vValue
		aValues = unwrapList(aArgs)
		aValues + unwrapValue(vValue)
		com_dispatch(pObject, cPath, COM_SET, aValues)
		return self

	/*
		Set a property to a COM object (Like Set x.Property = Object in Visual Basic)
	*/
	func setPropRef cPath, oValue
		com_dispatch(pObject, cPath, COM_SETREF, [unwrapValue(oValue)])
		return self

	/*
		Call a method, aArgs is a list of arguments
	*/
	func invoke cMethod, aArgs
		return wrapValue(com_dispatch(pObject, cMethod, COM_INVOKE, unwrapList(aArgs)))

	/*
		The items of a collection (Like For Each in Visual Basic)
	*/
	func items
		aItems = com_items(pObject)
		aOutput = []
		nCount = len(aItems)
		for nIndex = 1 to nCount
			aOutput + wrapValue(aItems[nIndex])
		next
		return aOutput

	/*
		The interface name (like _Connection) and the list of members [cName, cKind, nParameters]
	*/
	func typeName
		return com_typename(pObject)

	func members
		return com_members(pObject)

	/*
		Receive the events of the object (The default events interface), returns a COMEvents object
	*/
	func events
		return new COMEvents(pObject)

	/*
		Returns 1 if oOther is the same COM object (COMObject or C pointer)
	*/
	func isSame oOther
		return com_issame(pObject, unwrapValue(oOther))

	/*
		The COM object (C pointer), used by the com_*() functions
	*/
	func pointer
		return pObject

	/*
		Returns 1 if the object can be used (created and not released)
	*/
	func isValid
		return com_isobject(pObject)

	/*
		Release the COM object now (The Ring garbage collector does that when the object is deleted)
	*/
	func release
		if com_isobject(pObject)
			com_release(pObject)
		ok

	private

	func wrapValue vValue
		if com_isobject(vValue)
			return new COMObject(vValue)
		ok
		return vValue

	func unwrapValue vValue
		if isObject(vValue)
			if lower(classname(vValue)) = "comobject"
				return vValue.pointer()
			ok
		ok
		return vValue

	func unwrapList aArgs
		if ! isList(aArgs)
			raise("Error, WinCOM: The arguments must be a list")
		ok
		aValues = []
		nCount = len(aArgs)
		for nIndex = 1 to nCount
			aValues + unwrapValue(aArgs[nIndex])
		next
		return aValues

# Using the COMObject class (Object-Oriented Programming)
# The methods return COMObject objects for COM objects

load "wincom.ring"

func main
	oShell = new COMObject("WScript.Shell")
	? "Interface : " + oShell.typeName()
	? "Windows   : " + oShell.invoke("ExpandEnvironmentStrings", ["%WINDIR%"])
	? "OS        : " + oShell.invoke("ExpandEnvironmentStrings", ["%OS%"])

	# Environment("PROCESS") returns a collection
	oEnv = oShell.getPropArgs("Environment", ["PROCESS"])
	oEnv.setPropArgs("Item", ["WINCOM_SAMPLE"], "Value from Ring")
	? "Variable  : " + oEnv.getPropArgs("Item", ["WINCOM_SAMPLE"])

	# Run a command and read the output (Exec returns a WshScriptExec object)
	oExec = oShell.invoke("Exec", ["cmd /c echo Hello from cmd"])
	cOutput = oExec.getProp("StdOut.ReadAll")
	cOutput = substr(substr(cOutput, char(13), ""), char(10), "")
	? "Output    : " + cOutput

	# Using braces
	oDict = new COMObject("Scripting.Dictionary")
	oDict {
		invoke("Add", ["one", 1])
		invoke("Add", ["two", 2])
		setPropArgs("Item", ["three"], 3)
		? "Count     : " + getProp("Count")
		aKeys = invoke("Keys", [])
		? "Keys      : " + aKeys[1] + ", " + aKeys[2] + ", " + aKeys[3]
	}

	# COMObject objects can be passed as arguments
	oParent = new COMObject("Scripting.Dictionary")
	oParent.invoke("Add", ["child", oDict])
	oChild = oParent.getPropArgs("Item", ["child"])
	? "Child     : " + oChild.getProp("Count") + " items"

	# The members of the object (From the type information)
	? nl + "Members of " + oDict.typeName() + ":"
	for aMember in oDict.members()
		? "	" + aMember[1] + " (" + aMember[2] + ", parameters: " + aMember[3] + ")"
	next

	# Release the COM object now (The garbage collector does that automatically)
	oDict.release()
	? nl + "oDict.isValid() after release(): " + oDict.isValid()

# Using COM objects from many threads (RingThreads)
# Each thread has its own COM apartment: create the objects in the thread that uses them
# (Don't pass a COM object from one thread to another)

load "threads.ring"
load "wincom.ring"

# The results of the threads
aResults = list(4)
oMutex = new_mtx_t()

func main
	mtx_init(oMutex, 0)
	oMain = com_createobject("Scripting.Dictionary")
	com_invoke(oMain, "Add", "main", 1)

	nThreads = len(aResults)
	aThreads = list(nThreads)
	for nThread = 1 to nThreads
		aThreads[nThread] = new_thrd_t()
		thrd_create(aThreads[nThread], "worker(" + nThread + ")")
	next
	for nThread = 1 to nThreads
		nResult = 0
		thrd_join(aThreads[nThread], :nResult)
	next

	for cResult in aResults
		? cResult
	next
	? "Main thread: Count = " + com_get(oMain, "Count")

func worker nThread
	# Objects created in this thread (COM is initialized for the thread automatically)
	oDict = com_createobject("Scripting.Dictionary")
	for nIndex = 1 to 1000
		com_set(oDict, "Item", "key" + (nIndex % 25), nIndex * nThread)
	next
	# Errors are raised in the thread that caused them
	nErrors = 0
	for nIndex = 1 to 10
		try
			com_invoke(oDict, "Add", "key1", 1)
		catch
			nErrors++
		done
	next
	oShell = com_createobject("WScript.Shell")
	cOS = com_invoke(oShell, "ExpandEnvironmentStrings", "%OS%")
	cResult = "Thread " + nThread + ": Count = " + com_get(oDict, "Count") +
		  ", key24 = " + com_get(oDict, 'Item("key24")') + ", errors caught = " + nErrors + ", OS = " + cOS
	mtx_lock(oMutex)
	aResults[nThread] = cResult
	mtx_unlock(oMutex)
	thrd_exit(0)

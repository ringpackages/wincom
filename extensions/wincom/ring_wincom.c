/*
**  WinCOM - Windows COM (Component Object Model) for the Ring Programming Language
**
**  The extension uses COM Automation (late binding using the IDispatch interface)
**  So any COM object that supports Automation can be used (ADO, ADOX, Scripting, WScript, Office, WMI, ...)
**
**  Ring values are converted to VARIANT values (and back) automatically
**      Number          <--> VT_I4 (integers) or VT_R8 (other numbers), all number types are returned as numbers
**      String          <--> VT_BSTR (Ring strings are UTF-8, COM strings are UTF-16)
**      List            <--> SAFEARRAY (Nested lists for arrays with many dimensions)
**      COM Object      <--> VT_DISPATCH (C pointer, type: COMObject, Release() is called by the Ring GC)
**      Null/Empty      ---> "" (Use com_lasttype() to know the type)
**      Date            ---> String (YYYY-MM-DD or YYYY-MM-DD HH:MM:SS)
**      Boolean         ---> 1 or 0
**      Byte Array      ---> String (binary data)
**  Other VARIANT types (Null, Boolean, Date, Currency, Missing, ...) are created using com_variant()
**
**  Member names can be a path: "Fields.Count", 'Fields("Name").Value', "Item(1).Name"
**
**  COM errors (including the description provided by the object) are raised as Ring errors
**  So they can be caught using try/catch
*/

#ifndef _CRT_SECURE_NO_WARNINGS
	#define _CRT_SECURE_NO_WARNINGS
#endif

#ifndef COBJMACROS
	#define COBJMACROS
#endif

#include "ring.h"

#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <ocidl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>

/* From olectl.h (Not included because it causes warnings) */
#ifndef CONNECT_E_NOCONNECTION
	#define CONNECT_E_NOCONNECTION ((HRESULT)0x80040200L)
#endif

/* C pointer types */
#define WINCOM_OBJECT "COMObject"
#define WINCOM_UNKNOWN "COMUnknown"
#define WINCOM_VARIANT "COMVariant"
#define WINCOM_BYREF "COMByRef"
#define WINCOM_NAMED "COMNamed"
#define WINCOM_EVENTS "COMEvents"

/* Modes used by com_dispatch() */
#define WINCOM_MODE_GET 1
#define WINCOM_MODE_INVOKE 2
#define WINCOM_MODE_SET 3
#define WINCOM_MODE_SETREF 4

/* Limits */
#define WINCOM_MAXNAME 256
#define WINCOM_MAXPATHARGS 16
#define WINCOM_MAXDEPTH 32
#define WINCOM_ERRORSIZE 2048

/* Error messages */
#define WINCOM_ERROR_PREFIX "Error, WinCOM: "
#define WINCOM_ERROR_OOM "Error, WinCOM: Out of memory"
#define WINCOM_ERROR_INIT "Error, WinCOM: Can't initialize COM (CoInitializeEx)"
#define WINCOM_ERROR_PARACOUNT "Error, WinCOM: Bad parameters count"
#define WINCOM_ERROR_NOTOBJECT "Error, WinCOM: The first parameter must be a COM object"
#define WINCOM_ERROR_NAME "Error, WinCOM: The member name must be a string"
#define WINCOM_ERROR_RINGOBJECT                                                                                        \
	"Error, WinCOM: Ring objects can't be passed to COM (use the pObject attribute of COMObject)"

/* Named argument (created by com_named()) */
typedef struct WinCOMNamed {
	BSTR cName;
	VARIANT vValue;
} WinCOMNamed;

/*
**  ByRef argument (created by com_byref())
**  The holder is reference counted: a call that passes it keeps it alive (even if com_release() is used
**  by an event handler during the call)
*/
typedef struct WinCOMByRef {
	LONG nRefCount;
	VARIANT vValue;
} WinCOMByRef;

/* Argument passed to a member: the value, the name (NULL for positional arguments) and the ByRef holder */
typedef struct WinCOMArg {
	VARIANT vValue;
	BSTR cName;
	WinCOMByRef *pByRef;
} WinCOMArg;

/* State of each thread (Each thread has its own COM apartment) */
#define WINCOM_TLS __declspec(thread)
static WINCOM_TLS int nWinCOMInit = 0;
static WINCOM_TLS VARTYPE nWinCOMLastType = VT_EMPTY;
static WINCOM_TLS unsigned long nWinCOMLastHResult = 0;
static WINCOM_TLS char cWinCOMError[WINCOM_ERRORSIZE];

/*
**  The first error raised by an event handler (Raised again by the com_*() function that caused the event)
**  nWinCOMCallDepth is the number of active com_*() calls that can cause events
**  Events that happen outside these calls (like in the event loop of a GUI library) print their errors
*/
static WINCOM_TLS int lWinCOMEventError = 0;
static WINCOM_TLS char cWinCOMEventError[WINCOM_ERRORSIZE];
static WINCOM_TLS int nWinCOMCallDepth = 0;

static char *wincom_strdup(const char *cStr);
static void wincom_callandreturn(void *pPointer, IDispatch *pDispatch, const char *cPath, int nMode,
				 struct WinCOMArg *aArgs, int nArgs);

/* Start a call that can cause events */
static void wincom_begincall(void) { nWinCOMCallDepth++; }

/*
**  End a call that can cause events: raise the error of an event handler (if any)
**  Returns 1 if an error was raised
*/
static int wincom_endcall(void *pPointer) {
	char cMessage[WINCOM_ERRORSIZE];
	if (nWinCOMCallDepth > 0) {
		nWinCOMCallDepth--;
	}
	if (!lWinCOMEventError) {
		return 0;
	}
	lWinCOMEventError = 0;
	memcpy(cMessage, cWinCOMEventError, sizeof(cMessage));
	RING_API_ERROR(cMessage);
	return 1;
}

/*
**  Initialize COM in the current thread (Single-Threaded Apartment, like Visual Basic)
**  If the thread is already initialized (for example by Qt) using another mode, we use it
*/
static int wincom_init(void) {
	HRESULT hr;
	if (nWinCOMInit) {
		return 1;
	}
	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	if (SUCCEEDED(hr) || (hr == RPC_E_CHANGED_MODE)) {
		nWinCOMInit = 1;
		return 1;
	}
	return 0;
}

/* Free functions (Called by the Ring Garbage Collector) */

static void wincom_freeobject(void *pState, void *pPtr) {
	IUnknown *pUnknown;
	(void)pState;
	pUnknown = (IUnknown *)pPtr;
	if (pUnknown != NULL) {
		IUnknown_Release(pUnknown);
	}
}

static void wincom_freevariant(void *pState, void *pPtr) {
	(void)pState;
	if (pPtr != NULL) {
		VariantClear((VARIANT *)pPtr);
		free(pPtr);
	}
}

static void wincom_releasebyref(WinCOMByRef *pByRef) {
	if ((pByRef != NULL) && (InterlockedDecrement(&pByRef->nRefCount) == 0)) {
		VariantClear(&pByRef->vValue);
		free(pByRef);
	}
}

static void wincom_freebyref(void *pState, void *pPtr) {
	(void)pState;
	wincom_releasebyref((WinCOMByRef *)pPtr);
}

static void wincom_freenamed(void *pState, void *pPtr) {
	WinCOMNamed *pNamed;
	(void)pState;
	pNamed = (WinCOMNamed *)pPtr;
	if (pNamed != NULL) {
		SysFreeString(pNamed->cName);
		VariantClear(&pNamed->vValue);
		free(pNamed);
	}
}

static void wincom_freeevents(void *pState, void *pPtr);

/* String conversion */

/* UTF-8 to UTF-16, strings that are not valid UTF-8 are converted using the ANSI code page */
static UINT wincom_codepage(const char *cStr, int nSize) {
	if (nSize == 0) {
		return CP_UTF8;
	}
	if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, cStr, nSize, NULL, 0) == 0) {
		return CP_ACP;
	}
	return CP_UTF8;
}

/* UTF-8 to BSTR (The string can contain the NUL character) */
static BSTR wincom_tobstr(const char *cStr, int nSize) {
	UINT nCodePage;
	int nLen;
	BSTR cOut;
	if (nSize <= 0) {
		return SysAllocStringLen(L"", 0);
	}
	nCodePage = wincom_codepage(cStr, nSize);
	nLen = MultiByteToWideChar(nCodePage, 0, cStr, nSize, NULL, 0);
	if (nLen <= 0) {
		return SysAllocStringLen(L"", 0);
	}
	cOut = SysAllocStringLen(NULL, (UINT)nLen);
	if (cOut == NULL) {
		return NULL;
	}
	MultiByteToWideChar(nCodePage, 0, cStr, nSize, cOut, nLen);
	return cOut;
}

/* UTF-16 to UTF-8 (malloc), *pSize is the size of the output */
static char *wincom_toutf8(const WCHAR *cStr, int nLen, int *pSize) {
	int nSize;
	char *cOut;
	*pSize = 0;
	if ((cStr == NULL) || (nLen <= 0)) {
		cOut = (char *)malloc(1);
		if (cOut != NULL) {
			cOut[0] = '\0';
		}
		return cOut;
	}
	nSize = WideCharToMultiByte(CP_UTF8, 0, cStr, nLen, NULL, 0, NULL, NULL);
	if (nSize <= 0) {
		nSize = 0;
	}
	cOut = (char *)malloc((size_t)nSize + 1);
	if (cOut == NULL) {
		return NULL;
	}
	if (nSize > 0) {
		WideCharToMultiByte(CP_UTF8, 0, cStr, nLen, cOut, nSize, NULL, NULL);
	}
	cOut[nSize] = '\0';
	*pSize = nSize;
	return cOut;
}

/* Append UTF-8 text to the error message (Long messages are truncated without cutting a UTF-8 character) */
static void wincom_appenderrora(const char *cStr) {
	size_t nUsed, nLen;
	nUsed = strlen(cWinCOMError);
	nLen = strlen(cStr);
	if (nUsed + nLen >= WINCOM_ERRORSIZE) {
		nLen = WINCOM_ERRORSIZE - nUsed - 1;
		/* Continuation bytes are 10xxxxxx */
		while ((nLen > 0) && ((((unsigned char)cStr[nLen]) & 0xC0) == 0x80)) {
			nLen--;
		}
	}
	memcpy(cWinCOMError + nUsed, cStr, nLen);
	cWinCOMError[nUsed + nLen] = '\0';
}

/* Append UTF-16 text to the error message (A NUL character ends the text) */
static void wincom_appenderror(const WCHAR *cStr, int nLen) {
	char *cUTF8;
	int nSize;
	if ((cStr == NULL) || (nLen <= 0)) {
		return;
	}
	cUTF8 = wincom_toutf8(cStr, nLen, &nSize);
	if (cUTF8 != NULL) {
		wincom_appenderrora(cUTF8);
		free(cUTF8);
	}
}

static char *wincom_strdup(const char *cStr) {
	size_t nLen;
	char *cOut;
	nLen = strlen(cStr);
	cOut = (char *)malloc(nLen + 1);
	if (cOut != NULL) {
		memcpy(cOut, cStr, nLen + 1);
	}
	return cOut;
}

/* Remove spaces, dots and new lines from the end of the error message */
static void wincom_trimerror(void) {
	size_t nLen;
	nLen = strlen(cWinCOMError);
	while ((nLen > 0) && ((cWinCOMError[nLen - 1] == ' ') || (cWinCOMError[nLen - 1] == '\r') ||
			      (cWinCOMError[nLen - 1] == '\n') || (cWinCOMError[nLen - 1] == '.'))) {
		cWinCOMError[--nLen] = '\0';
	}
}

/*
**  Build the error message: Error, WinCOM: <Member> : <Source> : <Description> (0x80040E14)
**  pExcepInfo is used if the object provided the error information (DISP_E_EXCEPTION)
*/
static void wincom_seterror(const char *cMember, HRESULT hr, EXCEPINFO *pExcepInfo) {
	WCHAR *cMessage;
	char cCode[32];
	DWORD nLen;
	cWinCOMError[0] = '\0';
	wincom_appenderrora(WINCOM_ERROR_PREFIX);
	if ((cMember != NULL) && (cMember[0] != '\0')) {
		wincom_appenderrora(cMember);
		wincom_appenderrora(" : ");
	}
	if ((hr == DISP_E_EXCEPTION) && (pExcepInfo != NULL)) {
		if (pExcepInfo->pfnDeferredFillIn != NULL) {
			pExcepInfo->pfnDeferredFillIn(pExcepInfo);
		}
		if (pExcepInfo->scode != 0) {
			hr = pExcepInfo->scode;
		} else if (pExcepInfo->wCode != 0) {
			hr = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, pExcepInfo->wCode);
		}
		if ((pExcepInfo->bstrSource != NULL) && (SysStringLen(pExcepInfo->bstrSource) > 0)) {
			wincom_appenderror(pExcepInfo->bstrSource, (int)SysStringLen(pExcepInfo->bstrSource));
			wincom_appenderrora(" : ");
		}
		if ((pExcepInfo->bstrDescription != NULL) && (SysStringLen(pExcepInfo->bstrDescription) > 0)) {
			wincom_appenderror(pExcepInfo->bstrDescription,
					   (int)SysStringLen(pExcepInfo->bstrDescription));
			wincom_trimerror();
			goto addcode;
		}
	}
	/* Use the system message */
	cMessage = NULL;
	nLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
				  FORMAT_MESSAGE_IGNORE_INSERTS,
			      NULL, (DWORD)hr, 0, (LPWSTR)&cMessage, 0, NULL);
	if ((nLen > 0) && (cMessage != NULL)) {
		wincom_appenderror(cMessage, (int)nLen);
		wincom_trimerror();
	} else if (HRESULT_FACILITY(hr) == FACILITY_CONTROL) {
		/* Errors raised by objects written in Visual Basic (Like Err.Raise) or the Scripting Runtime */
		sprintf(cCode, "Error number %lu", (unsigned long)HRESULT_CODE(hr));
		wincom_appenderrora(cCode);
	} else {
		wincom_appenderrora("Unknown error");
	}
	if (cMessage != NULL) {
		LocalFree(cMessage);
	}
addcode:
	sprintf(cCode, " (0x%08lX)", (unsigned long)hr);
	wincom_appenderrora(cCode);
	nWinCOMLastHResult = (unsigned long)hr;
}

static void wincom_seterrortext(const char *cText) {
	cWinCOMError[0] = '\0';
	wincom_appenderrora(cText);
}

/* Clear the EXCEPINFO strings */
static void wincom_clearexcepinfo(EXCEPINFO *pExcepInfo) {
	if (pExcepInfo->bstrSource != NULL) {
		SysFreeString(pExcepInfo->bstrSource);
	}
	if (pExcepInfo->bstrDescription != NULL) {
		SysFreeString(pExcepInfo->bstrDescription);
	}
	if (pExcepInfo->bstrHelpFile != NULL) {
		SysFreeString(pExcepInfo->bstrHelpFile);
	}
	memset(pExcepInfo, 0, sizeof(EXCEPINFO));
}

/* Ring values to VARIANT */

static void wincom_numbertovariant(double nValue, VARIANT *pVar) {
	VariantInit(pVar);
	if ((nValue == floor(nValue)) && (nValue >= -2147483648.0) && (nValue <= 2147483647.0)) {
		V_VT(pVar) = VT_I4;
		V_I4(pVar) = (LONG)nValue;
	} else {
		V_VT(pVar) = VT_R8;
		V_R8(pVar) = nValue;
	}
}

static int wincom_stringtovariant(const char *cStr, int nSize, VARIANT *pVar) {
	VariantInit(pVar);
	V_BSTR(pVar) = wincom_tobstr(cStr, nSize);
	if (V_BSTR(pVar) == NULL) {
		wincom_seterrortext(WINCOM_ERROR_OOM);
		return 0;
	}
	V_VT(pVar) = VT_BSTR;
	return 1;
}

static int wincom_listtovariant(List *pList, VARIANT *pVar, int nDepth);

/* Convert the item (nIndex) of a Ring list to VARIANT */
static int wincom_itemtovariant(List *pList, unsigned int nIndex, VARIANT *pVar, int nDepth) {
	VariantInit(pVar);
	if (ring_list_isdouble(pList, nIndex)) {
		wincom_numbertovariant(ring_list_getdouble(pList, nIndex), pVar);
		return 1;
	}
	if (ring_list_isint(pList, nIndex)) {
		wincom_numbertovariant((double)ring_list_getint(pList, nIndex), pVar);
		return 1;
	}
	if (ring_list_isstring(pList, nIndex)) {
		return wincom_stringtovariant(ring_list_getstring(pList, nIndex),
					      (int)ring_list_getstringsize(pList, nIndex), pVar);
	}
	if (ring_list_islist(pList, nIndex)) {
		return wincom_listtovariant(ring_list_getlist(pList, nIndex), pVar, nDepth + 1);
	}
	/* Other types (pointers, functions) */
	V_VT(pVar) = VT_EMPTY;
	return 1;
}

/*
**  Convert a Ring list to VARIANT
**  C Pointer (COMObject)  --> VT_DISPATCH
**  C Pointer (COMVariant) --> Copy of the VARIANT
**  Other lists            --> SAFEARRAY of VARIANT (nested lists --> VARIANT that contains SAFEARRAY)
*/
static int wincom_listtovariant(List *pList, VARIANT *pVar, int nDepth) {
	const char *cType;
	void *pPtr;
	SAFEARRAY *pArray;
	VARIANT *pData;
	unsigned int x, nSize;
	HRESULT hr;
	VariantInit(pVar);
	if (nDepth > WINCOM_MAXDEPTH) {
		wincom_seterrortext("Error, WinCOM: The list has too many levels");
		return 0;
	}
	if (ring_list_iscpointerlist(pList)) {
		cType = ring_list_getstring(pList, RING_CPOINTER_TYPE);
		pPtr = ring_list_getpointer(pList, RING_CPOINTER_POINTER);
		if (strcmp(cType, WINCOM_OBJECT) == 0) {
			/* NULL pointer is (Nothing) in Visual Basic */
			V_VT(pVar) = VT_DISPATCH;
			V_DISPATCH(pVar) = (IDispatch *)pPtr;
			if (pPtr != NULL) {
				IDispatch_AddRef((IDispatch *)pPtr);
			}
			return 1;
		}
		if (strcmp(cType, WINCOM_UNKNOWN) == 0) {
			V_VT(pVar) = VT_UNKNOWN;
			V_UNKNOWN(pVar) = (IUnknown *)pPtr;
			if (pPtr != NULL) {
				IUnknown_AddRef((IUnknown *)pPtr);
			}
			return 1;
		}
		if ((strcmp(cType, WINCOM_VARIANT) == 0) || (strcmp(cType, WINCOM_BYREF) == 0)) {
			/* COMVariant: copy of the VARIANT, COMByRef: inside lists (arrays) the value is copied */
			if (pPtr == NULL) {
				return 1;
			}
			hr = VariantCopy(pVar, (strcmp(cType, WINCOM_BYREF) == 0) ? &((WinCOMByRef *)pPtr)->vValue
										  : (VARIANT *)pPtr);
			if (FAILED(hr)) {
				wincom_seterror("", hr, NULL);
				return 0;
			}
			return 1;
		}
		if (strcmp(cType, WINCOM_NAMED) == 0) {
			wincom_seterrortext("Error, WinCOM: Named arguments (com_named()) can't be used inside lists");
			return 0;
		}
		cWinCOMError[0] = '\0';
		wincom_appenderrora("Error, WinCOM: The C pointer type is not supported: ");
		wincom_appenderrora(cType);
		return 0;
	}
	if (ring_list_isobject(pList)) {
		wincom_seterrortext(WINCOM_ERROR_RINGOBJECT);
		return 0;
	}
	/* Array */
	nSize = ring_list_getsize(pList);
	pArray = SafeArrayCreateVector(VT_VARIANT, 0, nSize);
	if (pArray == NULL) {
		wincom_seterrortext(WINCOM_ERROR_OOM);
		return 0;
	}
	if (nSize > 0) {
		hr = SafeArrayAccessData(pArray, (void **)&pData);
		if (FAILED(hr)) {
			SafeArrayDestroy(pArray);
			wincom_seterror("", hr, NULL);
			return 0;
		}
		for (x = 1; x <= nSize; x++) {
			if (!wincom_itemtovariant(pList, x, &pData[x - 1], nDepth)) {
				SafeArrayUnaccessData(pArray);
				SafeArrayDestroy(pArray);
				return 0;
			}
		}
		SafeArrayUnaccessData(pArray);
	}
	V_VT(pVar) = VT_ARRAY | VT_VARIANT;
	V_ARRAY(pVar) = pArray;
	return 1;
}

/* Convert the Ring parameter (nPara) to VARIANT */
static int wincom_paratovariant(void *pPointer, int nPara, VARIANT *pVar) {
	VariantInit(pVar);
	if (RING_API_ISNUMBER(nPara)) {
		wincom_numbertovariant(RING_API_GETNUMBER(nPara), pVar);
		return 1;
	}
	if (RING_API_ISSTRING(nPara)) {
		return wincom_stringtovariant(RING_API_GETSTRING(nPara), RING_API_GETSTRINGSIZE(nPara), pVar);
	}
	if (RING_API_ISLIST(nPara)) {
		return wincom_listtovariant(RING_API_GETLIST(nPara), pVar, 0);
	}
	wincom_seterrortext("Error, WinCOM: The parameter type is not supported");
	return 0;
}

/* Get the COM object (IDispatch) from the Ring parameter, NULL if the parameter is not a COM object */
static IDispatch *wincom_getobject(void *pPointer, int nPara) {
	List *pList;
	if (!RING_API_ISLIST(nPara)) {
		return NULL;
	}
	pList = RING_API_GETLIST(nPara);
	if (!ring_list_iscpointerlist(pList)) {
		return NULL;
	}
	if (strcmp(ring_list_getstring(pList, RING_CPOINTER_TYPE), WINCOM_OBJECT) != 0) {
		return NULL;
	}
	return (IDispatch *)ring_list_getpointer(pList, RING_CPOINTER_POINTER);
}

/* Get the C pointer of the type cType from a Ring list, NULL if the list is not this type */
static void *wincom_getlistpointer(List *pList, const char *cType) {
	if (!ring_list_iscpointerlist(pList)) {
		return NULL;
	}
	if (strcmp(ring_list_getstring(pList, RING_CPOINTER_TYPE), cType) != 0) {
		return NULL;
	}
	return ring_list_getpointer(pList, RING_CPOINTER_POINTER);
}

/*
**  Convert a Ring list to an argument
**  COMByRef (com_byref()) --> VT_BYREF | VT_VARIANT (The member can change the value)
**  COMNamed (com_named()) --> Named argument
*/
static int wincom_listtoarg(List *pList, WinCOMArg *pArg) {
	WinCOMByRef *pByRef;
	WinCOMNamed *pNamed;
	HRESULT hr;
	VariantInit(&pArg->vValue);
	pArg->cName = NULL;
	pArg->pByRef = NULL;
	pByRef = (WinCOMByRef *)wincom_getlistpointer(pList, WINCOM_BYREF);
	if (pByRef != NULL) {
		/* The argument keeps the holder alive until the call ends */
		InterlockedIncrement(&pByRef->nRefCount);
		pArg->pByRef = pByRef;
		V_VT(&pArg->vValue) = VT_BYREF | VT_VARIANT;
		V_VARIANTREF(&pArg->vValue) = &pByRef->vValue;
		return 1;
	}
	pNamed = (WinCOMNamed *)wincom_getlistpointer(pList, WINCOM_NAMED);
	if (pNamed != NULL) {
		hr = VariantCopy(&pArg->vValue, &pNamed->vValue);
		if (FAILED(hr)) {
			wincom_seterror("", hr, NULL);
			return 0;
		}
		pArg->cName = SysAllocString(pNamed->cName);
		if (pArg->cName == NULL) {
			VariantClear(&pArg->vValue);
			wincom_seterrortext(WINCOM_ERROR_OOM);
			return 0;
		}
		return 1;
	}
	if (ring_list_iscpointerlist(pList) && (ring_list_getpointer(pList, RING_CPOINTER_POINTER) == NULL) &&
	    ((strcmp(ring_list_getstring(pList, RING_CPOINTER_TYPE), WINCOM_BYREF) == 0) ||
	     (strcmp(ring_list_getstring(pList, RING_CPOINTER_TYPE), WINCOM_NAMED) == 0))) {
		wincom_seterrortext("Error, WinCOM: The argument was released (com_release())");
		return 0;
	}
	return wincom_listtovariant(pList, &pArg->vValue, 0);
}

/* Convert the Ring parameter (nPara) to an argument */
static int wincom_paratoarg(void *pPointer, int nPara, WinCOMArg *pArg) {
	pArg->cName = NULL;
	pArg->pByRef = NULL;
	if (RING_API_ISLIST(nPara)) {
		return wincom_listtoarg(RING_API_GETLIST(nPara), pArg);
	}
	return wincom_paratovariant(pPointer, nPara, &pArg->vValue);
}

/* Convert the item (nIndex) of a Ring list to an argument */
static int wincom_itemtoarg(List *pList, unsigned int nIndex, WinCOMArg *pArg) {
	pArg->cName = NULL;
	pArg->pByRef = NULL;
	if (ring_list_islist(pList, nIndex)) {
		return wincom_listtoarg(ring_list_getlist(pList, nIndex), pArg);
	}
	return wincom_itemtovariant(pList, nIndex, &pArg->vValue, 0);
}

/* VARIANT to Ring values */

/* Date to string: YYYY-MM-DD or YYYY-MM-DD HH:MM:SS */
static void wincom_datetostring(DATE dDate, char *cOut) {
	SYSTEMTIME vTime;
	if (!VariantTimeToSystemTime(dDate, &vTime)) {
		cOut[0] = '\0';
		return;
	}
	if ((vTime.wHour == 0) && (vTime.wMinute == 0) && (vTime.wSecond == 0)) {
		sprintf(cOut, "%04d-%02d-%02d", vTime.wYear, vTime.wMonth, vTime.wDay);
	} else {
		sprintf(cOut, "%04d-%02d-%02d %02d:%02d:%02d", vTime.wYear, vTime.wMonth, vTime.wDay, vTime.wHour,
			vTime.wMinute, vTime.wSecond);
	}
}

/* Read a number of nMin to nMax digits */
static int wincom_readdigits(const char **pStr, int nMin, int nMax, int *pValue) {
	int nCount;
	*pValue = 0;
	nCount = 0;
	while ((**pStr >= '0') && (**pStr <= '9') && (nCount < nMax)) {
		*pValue = (*pValue * 10) + (**pStr - '0');
		(*pStr)++;
		nCount++;
	}
	return nCount >= nMin;
}

/* String to date: YYYY-MM-DD, YYYY-MM-DD HH:MM, YYYY-MM-DD HH:MM:SS (or YYYY-MM-DDTHH:MM:SS) */
static int wincom_stringtodate(const char *cStr, DATE *pDate) {
	static const int aDays[12] = {31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
	SYSTEMTIME vTime;
	int nYear, nMonth, nDay, nHour, nMinute, nSecond, lLeap;
	nHour = 0;
	nMinute = 0;
	nSecond = 0;
	if (!wincom_readdigits(&cStr, 4, 4, &nYear) || (*cStr++ != '-') || !wincom_readdigits(&cStr, 1, 2, &nMonth) ||
	    (*cStr++ != '-') || !wincom_readdigits(&cStr, 1, 2, &nDay)) {
		return 0;
	}
	if ((*cStr == ' ') || (*cStr == 'T')) {
		cStr++;
		if (!wincom_readdigits(&cStr, 1, 2, &nHour) || (*cStr++ != ':') ||
		    !wincom_readdigits(&cStr, 1, 2, &nMinute)) {
			return 0;
		}
		if (*cStr == ':') {
			cStr++;
			if (!wincom_readdigits(&cStr, 1, 2, &nSecond)) {
				return 0;
			}
		}
	}
	if (*cStr != '\0') {
		return 0;
	}
	if ((nMonth < 1) || (nMonth > 12) || (nDay < 1) || (nDay > aDays[nMonth - 1]) || (nHour > 23) ||
	    (nMinute > 59) || (nSecond > 59) || (nYear < 100)) {
		return 0;
	}
	lLeap = ((nYear % 4) == 0) && (((nYear % 100) != 0) || ((nYear % 400) == 0));
	if ((nMonth == 2) && (nDay == 29) && !lLeap) {
		return 0;
	}
	memset(&vTime, 0, sizeof(SYSTEMTIME));
	vTime.wYear = (WORD)nYear;
	vTime.wMonth = (WORD)nMonth;
	vTime.wDay = (WORD)nDay;
	vTime.wHour = (WORD)nHour;
	vTime.wMinute = (WORD)nMinute;
	vTime.wSecond = (WORD)nSecond;
	return SystemTimeToVariantTime(&vTime, pDate) ? 1 : 0;
}

/* Number value of a VARIANT, returns 0 if the VARIANT is not a number */
static int wincom_variantnumber(VARIANT *pVar, double *pNumber) {
	VARIANT vTemp;
	switch (V_VT(pVar)) {
	case VT_I1:
		*pNumber = (double)V_I1(pVar);
		return 1;
	case VT_UI1:
		*pNumber = (double)V_UI1(pVar);
		return 1;
	case VT_I2:
		*pNumber = (double)V_I2(pVar);
		return 1;
	case VT_UI2:
		*pNumber = (double)V_UI2(pVar);
		return 1;
	case VT_I4:
		*pNumber = (double)V_I4(pVar);
		return 1;
	case VT_UI4:
		*pNumber = (double)V_UI4(pVar);
		return 1;
	case VT_INT:
		*pNumber = (double)V_INT(pVar);
		return 1;
	case VT_UINT:
		*pNumber = (double)V_UINT(pVar);
		return 1;
	case VT_I8:
		*pNumber = (double)V_I8(pVar);
		return 1;
	case VT_UI8:
		*pNumber = (double)V_UI8(pVar);
		return 1;
	case VT_R4:
		*pNumber = (double)V_R4(pVar);
		return 1;
	case VT_R8:
		*pNumber = V_R8(pVar);
		return 1;
	case VT_BOOL:
		*pNumber = V_BOOL(pVar) ? 1.0 : 0.0;
		return 1;
	case VT_ERROR:
		*pNumber = (double)V_ERROR(pVar);
		return 1;
	case VT_CY:
	case VT_DECIMAL:
		VariantInit(&vTemp);
		if (SUCCEEDED(VariantChangeType(&vTemp, pVar, 0, VT_R8))) {
			*pNumber = V_R8(&vTemp);
			return 1;
		}
		*pNumber = 0;
		return 1;
	}
	return 0;
}

static void wincom_addvariant(void *pState, List *pList, VARIANT *pVar, int nDepth);

/* Add a COM object to a Ring list (as C pointer, Release() is called by the Ring GC), takes the reference */
static void wincom_addobject(void *pState, List *pList, IUnknown *pUnknown, const char *cType) {
	List *pList2;
	Item *pItem;
	pList2 = ring_list_newlist_gc(pState, pList);
	ring_list_addpointer_gc(pState, pList2, pUnknown);
	ring_list_addstring_gc(pState, pList2, cType);
	ring_list_addint_gc(pState, pList2, RING_CPOINTERSTATUS_NOTASSIGNED);
	pItem = ring_list_getitem_gc(pState, pList2, RING_CPOINTER_POINTER);
	ring_vm_gc_setfreefunc(pItem, wincom_freeobject);
}

/* Types that can be read from a SAFEARRAY into the data part of a VARIANT */
static int wincom_issimpletype(VARTYPE nType) {
	switch (nType) {
	case VT_I1:
	case VT_UI1:
	case VT_I2:
	case VT_UI2:
	case VT_I4:
	case VT_UI4:
	case VT_INT:
	case VT_UINT:
	case VT_I8:
	case VT_UI8:
	case VT_R4:
	case VT_R8:
	case VT_BOOL:
	case VT_ERROR:
	case VT_CY:
	case VT_DATE:
	case VT_BSTR:
	case VT_DISPATCH:
	case VT_UNKNOWN:
		return 1;
	}
	return 0;
}

/*
**  Add the items of a SAFEARRAY to a Ring list
**  nDim is the dimension (1 = the first dimension (left), like Visual Basic)
**  aIndex is the index vector used by SafeArrayGetElement() (aIndex[nDim - 1] is the index of nDim)
*/
static void wincom_addarray(void *pState, List *pList, SAFEARRAY *pArray, VARTYPE nType, UINT nDim, LONG *aIndex,
			    int nDepth) {
	LONG nLower, nUpper, x;
	UINT nDims;
	VARIANT vItem;
	List *pSubList;
	nDims = SafeArrayGetDim(pArray);
	if (FAILED(SafeArrayGetLBound(pArray, nDim, &nLower)) || FAILED(SafeArrayGetUBound(pArray, nDim, &nUpper))) {
		return;
	}
	for (x = nLower; x <= nUpper; x++) {
		aIndex[nDim - 1] = x;
		if (nDim < nDims) {
			pSubList = ring_list_newlist_gc(pState, pList);
			wincom_addarray(pState, pSubList, pArray, nType, nDim + 1, aIndex, nDepth);
			continue;
		}
		VariantInit(&vItem);
		if (nType == VT_VARIANT) {
			if (FAILED(SafeArrayGetElement(pArray, aIndex, &vItem))) {
				VariantInit(&vItem);
			}
		} else if (nType == VT_DECIMAL) {
			/* DECIMAL uses all the VARIANT memory (including the type) */
			if (SUCCEEDED(SafeArrayGetElement(pArray, aIndex, &V_DECIMAL(&vItem)))) {
				V_VT(&vItem) = VT_DECIMAL;
			}
		} else if (wincom_issimpletype(nType)) {
			/* The data of the other types starts at the same address (union) */
			if (SUCCEEDED(SafeArrayGetElement(pArray, aIndex, &V_UI1(&vItem)))) {
				V_VT(&vItem) = nType;
			}
		}
		wincom_addvariant(pState, pList, &vItem, nDepth + 1);
		VariantClear(&vItem);
	}
}

/* Check if the VARIANT is a one dimension byte array */
static int wincom_isbytearray(VARIANT *pVar) {
	return (V_VT(pVar) == (VT_ARRAY | VT_UI1)) && (V_ARRAY(pVar) != NULL) && (SafeArrayGetDim(V_ARRAY(pVar)) == 1);
}

/* Access the bytes of a byte array (call SafeArrayUnaccessData() after using them), NULL if empty */
static char *wincom_bytearraydata(SAFEARRAY *pArray, LONG *pSize) {
	LONG nLower, nUpper;
	char *pData;
	*pSize = 0;
	if (FAILED(SafeArrayGetLBound(pArray, 1, &nLower)) || FAILED(SafeArrayGetUBound(pArray, 1, &nUpper))) {
		return NULL;
	}
	if (nUpper < nLower) {
		return NULL;
	}
	if (FAILED(SafeArrayAccessData(pArray, (void **)&pData))) {
		return NULL;
	}
	*pSize = nUpper - nLower + 1;
	return pData;
}

/* Add a VARIANT to a Ring list */
static void wincom_addvariant(void *pState, List *pList, VARIANT *pVar, int nDepth) {
	double nNumber;
	char cDate[64];
	char *cStr;
	int nSize;
	LONG nBytes;
	LONG aIndex[64];
	IDispatch *pDispatch;
	VARIANT vTemp;
	SAFEARRAY *pArray;
	if (V_VT(pVar) & VT_BYREF) {
		VariantInit(&vTemp);
		if (SUCCEEDED(VariantCopyInd(&vTemp, pVar))) {
			wincom_addvariant(pState, pList, &vTemp, nDepth);
		} else {
			ring_list_addstring_gc(pState, pList, "");
		}
		VariantClear(&vTemp);
		return;
	}
	if (wincom_variantnumber(pVar, &nNumber)) {
		ring_list_adddouble_gc(pState, pList, nNumber);
		return;
	}
	if (V_VT(pVar) & VT_ARRAY) {
		pArray = V_ARRAY(pVar);
		if (pArray == NULL) {
			ring_list_newlist_gc(pState, pList);
			return;
		}
		if (wincom_isbytearray(pVar)) {
			cStr = wincom_bytearraydata(pArray, &nBytes);
			if (cStr != NULL) {
				ring_list_addstring2_gc(pState, pList, cStr, (unsigned int)nBytes);
				SafeArrayUnaccessData(pArray);
			} else {
				ring_list_addstring_gc(pState, pList, "");
			}
			return;
		}
		if ((nDepth > WINCOM_MAXDEPTH) || (SafeArrayGetDim(pArray) > 64) || (SafeArrayGetDim(pArray) == 0)) {
			ring_list_newlist_gc(pState, pList);
			return;
		}
		wincom_addarray(pState, ring_list_newlist_gc(pState, pList), pArray, (VARTYPE)(V_VT(pVar) & VT_TYPEMASK),
				1, aIndex, nDepth);
		return;
	}
	switch (V_VT(pVar)) {
	case VT_BSTR:
		cStr = wincom_toutf8(V_BSTR(pVar), (int)SysStringLen(V_BSTR(pVar)), &nSize);
		if (cStr != NULL) {
			ring_list_addstring2_gc(pState, pList, cStr, (unsigned int)nSize);
			free(cStr);
		} else {
			ring_list_addstring_gc(pState, pList, "");
		}
		return;
	case VT_DATE:
		wincom_datetostring(V_DATE(pVar), cDate);
		ring_list_addstring_gc(pState, pList, cDate);
		return;
	case VT_DISPATCH:
		if (V_DISPATCH(pVar) == NULL) {
			ring_list_addstring_gc(pState, pList, "");
			return;
		}
		IDispatch_AddRef(V_DISPATCH(pVar));
		wincom_addobject(pState, pList, (IUnknown *)V_DISPATCH(pVar), WINCOM_OBJECT);
		return;
	case VT_UNKNOWN:
		if (V_UNKNOWN(pVar) == NULL) {
			ring_list_addstring_gc(pState, pList, "");
			return;
		}
		if (SUCCEEDED(IUnknown_QueryInterface(V_UNKNOWN(pVar), &IID_IDispatch, (void **)&pDispatch))) {
			wincom_addobject(pState, pList, (IUnknown *)pDispatch, WINCOM_OBJECT);
		} else {
			IUnknown_AddRef(V_UNKNOWN(pVar));
			wincom_addobject(pState, pList, V_UNKNOWN(pVar), WINCOM_UNKNOWN);
		}
		return;
	}
	/* VT_EMPTY, VT_NULL and other types */
	VariantInit(&vTemp);
	if ((V_VT(pVar) != VT_EMPTY) && (V_VT(pVar) != VT_NULL) &&
	    SUCCEEDED(VariantChangeType(&vTemp, pVar, VARIANT_ALPHABOOL, VT_BSTR))) {
		cStr = wincom_toutf8(V_BSTR(&vTemp), (int)SysStringLen(V_BSTR(&vTemp)), &nSize);
		if (cStr != NULL) {
			ring_list_addstring2_gc(pState, pList, cStr, (unsigned int)nSize);
			free(cStr);
			VariantClear(&vTemp);
			return;
		}
	}
	VariantClear(&vTemp);
	ring_list_addstring_gc(pState, pList, "");
}

/* Return a VARIANT to Ring, the VARIANT is cleared */
static void wincom_retvariant(void *pPointer, VARIANT *pVar) {
	double nNumber;
	char cDate[64];
	char *cStr;
	int nSize;
	LONG nBytes;
	LONG aIndex[64];
	List *pList;
	SAFEARRAY *pArray;
	IDispatch *pDispatch;
	IUnknown *pUnknown;
	VARIANT vTemp;
	if (V_VT(pVar) & VT_BYREF) {
		VariantInit(&vTemp);
		VariantCopyInd(&vTemp, pVar);
		VariantClear(pVar);
		*pVar = vTemp;
	}
	nWinCOMLastType = V_VT(pVar);
	if (wincom_variantnumber(pVar, &nNumber)) {
		VariantClear(pVar);
		RING_API_RETNUMBER(nNumber);
		return;
	}
	switch (V_VT(pVar)) {
	case VT_BSTR:
		cStr = wincom_toutf8(V_BSTR(pVar), (int)SysStringLen(V_BSTR(pVar)), &nSize);
		VariantClear(pVar);
		if (cStr == NULL) {
			RING_API_ERROR(WINCOM_ERROR_OOM);
			return;
		}
		RING_API_RETSTRING2(cStr, nSize);
		free(cStr);
		return;
	case VT_DATE:
		wincom_datetostring(V_DATE(pVar), cDate);
		VariantClear(pVar);
		RING_API_RETSTRING(cDate);
		return;
	case VT_DISPATCH:
		pDispatch = V_DISPATCH(pVar);
		if (pDispatch == NULL) {
			RING_API_RETSTRING("");
			return;
		}
		/* Take the reference (Don't call VariantClear) */
		V_VT(pVar) = VT_EMPTY;
		RING_API_RETMANAGEDCPOINTER(pDispatch, WINCOM_OBJECT, wincom_freeobject);
		return;
	case VT_UNKNOWN:
		pUnknown = V_UNKNOWN(pVar);
		if (pUnknown == NULL) {
			RING_API_RETSTRING("");
			return;
		}
		if (SUCCEEDED(IUnknown_QueryInterface(pUnknown, &IID_IDispatch, (void **)&pDispatch))) {
			VariantClear(pVar);
			RING_API_RETMANAGEDCPOINTER(pDispatch, WINCOM_OBJECT, wincom_freeobject);
			return;
		}
		V_VT(pVar) = VT_EMPTY;
		RING_API_RETMANAGEDCPOINTER(pUnknown, WINCOM_UNKNOWN, wincom_freeobject);
		return;
	}
	if (wincom_isbytearray(pVar)) {
		cStr = wincom_bytearraydata(V_ARRAY(pVar), &nBytes);
		if (cStr != NULL) {
			RING_API_RETSTRING2(cStr, (int)nBytes);
			SafeArrayUnaccessData(V_ARRAY(pVar));
		} else {
			RING_API_RETSTRING("");
		}
		VariantClear(pVar);
		return;
	}
	if (V_VT(pVar) & VT_ARRAY) {
		/* The items are added to the output list (Returned by reference, so it's not copied) */
		pList = RING_API_NEWLIST;
		pArray = V_ARRAY(pVar);
		if ((pArray != NULL) && (SafeArrayGetDim(pArray) > 0) && (SafeArrayGetDim(pArray) <= 64)) {
			wincom_addarray(RING_API_STATE, pList, pArray, (VARTYPE)(V_VT(pVar) & VT_TYPEMASK), 1, aIndex, 0);
		}
		VariantClear(pVar);
		RING_API_RETLISTBYREF(pList);
		return;
	}
	/* VT_EMPTY, VT_NULL and other types */
	pList = RING_API_NEWLIST;
	wincom_addvariant(RING_API_STATE, pList, pVar, 0);
	VariantClear(pVar);
	if ((ring_list_getsize(pList) == 1) && ring_list_isstring(pList, 1)) {
		RING_API_RETSTRING2(ring_list_getstring(pList, 1), (int)ring_list_getstringsize(pList, 1));
	} else {
		RING_API_RETSTRING("");
	}
}

/* Calling members */

/*
**  Call a member of the COM object
**  aArgs are in the natural order (first argument first), for property put the last argument is the value
*/
static HRESULT wincom_invokemember(IDispatch *pDispatch, const char *cName, WORD nFlags, WinCOMArg *aArgs, int nArgs,
				   VARIANT *pResult, EXCEPINFO *pExcepInfo) {
	DISPID nDispID;
	DISPID *aIDs, *aNamedIDs;
	DISPPARAMS vParams;
	VARIANT *aVars;
	OLECHAR **aNames;
	UINT nArgErr;
	HRESULT hr;
	int x, k, nNamed, nPositional, nPut;
	/* The named arguments are after the positional arguments (checked by wincom_callpath()) */
	nNamed = 0;
	for (x = 0; x < nArgs; x++) {
		if (aArgs[x].cName != NULL) {
			nNamed++;
		}
	}
	nPositional = nArgs - nNamed;
	/* For property put the value is the last positional argument (DISPID_PROPERTYPUT) */
	nPut = (nFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) ? 1 : 0;
	if (nPut && ((nPositional == 0) || (nNamed > 0))) {
		return E_INVALIDARG;
	}
	aNames = (OLECHAR **)calloc((size_t)nNamed + 1, sizeof(OLECHAR *));
	aIDs = (DISPID *)calloc((size_t)nNamed + 1, sizeof(DISPID));
	aNamedIDs = (DISPID *)calloc((size_t)nNamed + 1, sizeof(DISPID));
	aVars = (VARIANT *)calloc((size_t)nArgs + 1, sizeof(VARIANT));
	if ((aNames == NULL) || (aIDs == NULL) || (aNamedIDs == NULL) || (aVars == NULL)) {
		hr = E_OUTOFMEMORY;
		goto done;
	}
	/* Get the member ID (The empty name is the default member) and the IDs of the named arguments */
	if (cName[0] == '\0') {
		if (nNamed > 0) {
			hr = DISP_E_UNKNOWNNAME;
			goto done;
		}
		nDispID = DISPID_VALUE;
	} else if ((_stricmp(cName, "_NewEnum") == 0) && (nNamed == 0)) {
		nDispID = DISPID_NEWENUM;
	} else {
		aNames[0] = wincom_tobstr(cName, (int)strlen(cName));
		if (aNames[0] == NULL) {
			hr = E_OUTOFMEMORY;
			goto done;
		}
		for (x = 0; x < nNamed; x++) {
			aNames[x + 1] = aArgs[nPositional + x].cName;
		}
		hr = IDispatch_GetIDsOfNames(pDispatch, &IID_NULL, aNames, (UINT)nNamed + 1, LOCALE_USER_DEFAULT, aIDs);
		SysFreeString(aNames[0]);
		if (FAILED(hr)) {
			goto done;
		}
		nDispID = aIDs[0];
	}
	/* rgvarg: the named arguments, then the positional arguments in reverse order */
	k = 0;
	if (nPut) {
		aVars[k] = aArgs[nArgs - 1].vValue;
		aNamedIDs[k] = DISPID_PROPERTYPUT;
		k++;
	}
	for (x = 0; x < nNamed; x++) {
		aVars[k] = aArgs[nPositional + x].vValue;
		aNamedIDs[k] = aIDs[x + 1];
		k++;
	}
	for (x = nPositional - 1 - nPut; x >= 0; x--) {
		aVars[k++] = aArgs[x].vValue;
	}
	vParams.rgvarg = nArgs > 0 ? aVars : NULL;
	vParams.cArgs = (UINT)nArgs;
	vParams.rgdispidNamedArgs = (nNamed + nPut) > 0 ? aNamedIDs : NULL;
	vParams.cNamedArgs = (UINT)(nNamed + nPut);
	nArgErr = 0;
	if (pResult != NULL) {
		VariantInit(pResult);
	}
	/* Clear the error information of the thread (Some objects use it to fill EXCEPINFO, so an old error appears) */
	SetErrorInfo(0, NULL);
	hr = IDispatch_Invoke(pDispatch, nDispID, &IID_NULL, LOCALE_USER_DEFAULT, nFlags, &vParams, pResult, pExcepInfo,
			      &nArgErr);
done:
	free(aNames);
	free(aIDs);
	free(aNamedIDs);
	free(aVars);
	return hr;
}

/* Parse a number or a string (between quotes) in a path */
static int wincom_parseliteral(const char **pPath, VARIANT *pVar) {
	const char *cStr;
	const char *cStart;
	char cQuote;
	char *cEnd;
	char *cBuffer;
	size_t nLen, nOut;
	double nNumber;
	cStr = *pPath;
	VariantInit(pVar);
	if ((*cStr == '"') || (*cStr == '\'')) {
		cQuote = *cStr;
		cStr++;
		cStart = cStr;
		/* The quote is written two times inside the string */
		nLen = 0;
		while (*cStr != '\0') {
			if (*cStr == cQuote) {
				if (cStr[1] == cQuote) {
					cStr += 2;
					nLen++;
					continue;
				}
				break;
			}
			cStr++;
			nLen++;
		}
		if (*cStr != cQuote) {
			return 0;
		}
		cBuffer = (char *)malloc(nLen + 1);
		if (cBuffer == NULL) {
			return 0;
		}
		nOut = 0;
		while (cStart < cStr) {
			cBuffer[nOut++] = *cStart;
			if ((*cStart == cQuote) && (cStart[1] == cQuote)) {
				cStart++;
			}
			cStart++;
		}
		cBuffer[nOut] = '\0';
		V_BSTR(pVar) = wincom_tobstr(cBuffer, (int)nOut);
		free(cBuffer);
		if (V_BSTR(pVar) == NULL) {
			return 0;
		}
		V_VT(pVar) = VT_BSTR;
		*pPath = cStr + 1;
		return 1;
	}
	nNumber = strtod(cStr, &cEnd);
	if (cEnd == cStr) {
		return 0;
	}
	wincom_numbertovariant(nNumber, pVar);
	*pPath = cEnd;
	return 1;
}

static void wincom_skipspaces(const char **pPath) {
	while ((**pPath == ' ') || (**pPath == '\t')) {
		(*pPath)++;
	}
}

/*
**  Parse one part of the path: Name or Name(Arg1, Arg2, ...)
**  The arguments are numbers or strings: 'Fields("Name")', "Item(1)", "Cells(2, 3)"
**  Returns 1 on success, *pLast is 1 if this is the last part
*/
static void wincom_clearpathargs(VARIANT *aPathArgs, int nPathArgs);

static int wincom_parsepart(const char **pPath, char *cName, VARIANT *aArgs, int *pArgs, int *pLast) {
	const char *cStr;
	int nLen;
	cStr = *pPath;
	*pArgs = 0;
	*pLast = 0;
	wincom_skipspaces(&cStr);
	nLen = 0;
	while ((*cStr != '\0') && (*cStr != '.') && (*cStr != '(') && (*cStr != ' ') && (*cStr != '\t')) {
		if (nLen >= WINCOM_MAXNAME - 1) {
			return 0;
		}
		cName[nLen++] = *cStr;
		cStr++;
	}
	cName[nLen] = '\0';
	wincom_skipspaces(&cStr);
	if (*cStr == '(') {
		cStr++;
		wincom_skipspaces(&cStr);
		if (*cStr != ')') {
			while (1) {
				wincom_skipspaces(&cStr);
				if ((*pArgs >= WINCOM_MAXPATHARGS) || !wincom_parseliteral(&cStr, &aArgs[*pArgs])) {
					wincom_clearpathargs(aArgs, *pArgs);
					*pArgs = 0;
					return 0;
				}
				(*pArgs)++;
				wincom_skipspaces(&cStr);
				if (*cStr == ',') {
					cStr++;
					continue;
				}
				break;
			}
		}
		if (*cStr != ')') {
			wincom_clearpathargs(aArgs, *pArgs);
			*pArgs = 0;
			return 0;
		}
		cStr++;
		wincom_skipspaces(&cStr);
	}
	if (*cStr == '.') {
		cStr++;
		/* A path can't end with a dot ("Fields.") */
		wincom_skipspaces(&cStr);
		if (*cStr == '\0') {
			wincom_clearpathargs(aArgs, *pArgs);
			*pArgs = 0;
			return 0;
		}
	} else if (*cStr == '\0') {
		*pLast = 1;
	} else {
		wincom_clearpathargs(aArgs, *pArgs);
		*pArgs = 0;
		return 0;
	}
	/* The empty name is allowed only for the default member, i.e. "(1)" or "" */
	if ((nLen == 0) && (*pLast == 0) && (*pArgs == 0)) {
		return 0;
	}
	*pPath = cStr;
	return 1;
}

static void wincom_cleararguments(WinCOMArg *aArgs, int nArgs) {
	int x;
	if (aArgs == NULL) {
		return;
	}
	for (x = 0; x < nArgs; x++) {
		/* VariantClear() doesn't free the value of VT_BYREF arguments (It's owned by the com_byref() holder) */
		VariantClear(&aArgs[x].vValue);
		if (aArgs[x].cName != NULL) {
			SysFreeString(aArgs[x].cName);
		}
		wincom_releasebyref(aArgs[x].pByRef);
	}
	free(aArgs);
}

/* Get the object from VARIANT (VT_DISPATCH or VT_UNKNOWN), returns NULL if the VARIANT is not an object */
static IDispatch *wincom_variantobject(VARIANT *pVar) {
	IDispatch *pDispatch;
	if ((V_VT(pVar) == VT_DISPATCH) && (V_DISPATCH(pVar) != NULL)) {
		pDispatch = V_DISPATCH(pVar);
		IDispatch_AddRef(pDispatch);
		return pDispatch;
	}
	if ((V_VT(pVar) == VT_UNKNOWN) && (V_UNKNOWN(pVar) != NULL)) {
		if (SUCCEEDED(IUnknown_QueryInterface(V_UNKNOWN(pVar), &IID_IDispatch, (void **)&pDispatch))) {
			return pDispatch;
		}
	}
	return NULL;
}

static void wincom_clearpathargs(VARIANT *aPathArgs, int nPathArgs) {
	int x;
	for (x = 0; x < nPathArgs; x++) {
		VariantClear(&aPathArgs[x]);
	}
}

/*
**  Check if a property put failed because the property doesn't accept the value using DISPATCH_PROPERTYPUT
**  (Then com_set() tries DISPATCH_PROPERTYPUTREF), the error code can be in EXCEPINFO (DISP_E_EXCEPTION)
*/
static int wincom_isputerror(HRESULT hr, EXCEPINFO *pExcepInfo) {
	if (hr == DISP_E_EXCEPTION) {
		if (pExcepInfo->pfnDeferredFillIn != NULL) {
			pExcepInfo->pfnDeferredFillIn(pExcepInfo);
			pExcepInfo->pfnDeferredFillIn = NULL;
		}
		hr = pExcepInfo->scode;
	}
	return (hr == DISP_E_MEMBERNOTFOUND) || (hr == DISP_E_BADPARAMCOUNT) || (hr == DISP_E_TYPEMISMATCH);
}

/* Check the order of the arguments: the named arguments must be after the positional arguments */
static int wincom_checknamedarguments(WinCOMArg *aArgs, int nArgs, int nMode) {
	int x, lNamed;
	lNamed = 0;
	for (x = 0; x < nArgs; x++) {
		if (aArgs[x].cName != NULL) {
			lNamed = 1;
		} else if (lNamed) {
			wincom_seterrortext("Error, WinCOM: Named arguments (com_named()) must be after the other arguments");
			return 0;
		}
	}
	if (lNamed && ((nMode == WINCOM_MODE_SET) || (nMode == WINCOM_MODE_SETREF))) {
		wincom_seterrortext("Error, WinCOM: Named arguments (com_named()) can't be used to set properties");
		return 0;
	}
	return 1;
}

/*
**  The engine of com_get(), com_invoke(), com_set(), com_setref() and com_dispatch()
**  aExtra are the arguments after the path (taken by this function)
**  Returns 1 on success, 0 on error (cWinCOMError contains the error message)
*/
static int wincom_callpath(IDispatch *pDispatch, const char *cPath, int nMode, WinCOMArg *aExtra, int nExtra,
			   VARIANT *pResult) {
	char cName[WINCOM_MAXNAME];
	VARIANT aPathArgs[WINCOM_MAXPATHARGS];
	WinCOMArg aPathItems[WINCOM_MAXPATHARGS];
	WinCOMArg *aAll;
	VARIANT vResult;
	EXCEPINFO vExcepInfo, vExcepInfo2;
	IDispatch *pCurrent, *pNext;
	const char *cStr;
	int nPathArgs, lLast, nAll, x, lValueIsObject;
	WORD nFlags;
	HRESULT hr, hr2;
	cStr = cPath;
	VariantInit(pResult);
	if (!wincom_checknamedarguments(aExtra, nExtra, nMode)) {
		wincom_cleararguments(aExtra, nExtra);
		return 0;
	}
	pCurrent = pDispatch;
	IDispatch_AddRef(pCurrent);
	while (1) {
		if (!wincom_parsepart(&cStr, cName, aPathArgs, &nPathArgs, &lLast)) {
			cWinCOMError[0] = '\0';
			wincom_appenderrora("Error, WinCOM: Bad member name or path: ");
			wincom_appenderrora(cPath);
			IDispatch_Release(pCurrent);
			wincom_cleararguments(aExtra, nExtra);
			return 0;
		}
		memset(&vExcepInfo, 0, sizeof(EXCEPINFO));
		if (!lLast) {
			/* Get the object of this part */
			for (x = 0; x < nPathArgs; x++) {
				aPathItems[x].vValue = aPathArgs[x];
				aPathItems[x].cName = NULL;
				aPathItems[x].pByRef = NULL;
			}
			hr = wincom_invokemember(pCurrent, cName, DISPATCH_PROPERTYGET | DISPATCH_METHOD, aPathItems,
						 nPathArgs, &vResult, &vExcepInfo);
			wincom_clearpathargs(aPathArgs, nPathArgs);
			if (FAILED(hr)) {
				wincom_seterror(cName, hr, &vExcepInfo);
				wincom_clearexcepinfo(&vExcepInfo);
				IDispatch_Release(pCurrent);
				wincom_cleararguments(aExtra, nExtra);
				return 0;
			}
			wincom_clearexcepinfo(&vExcepInfo);
			pNext = wincom_variantobject(&vResult);
			VariantClear(&vResult);
			IDispatch_Release(pCurrent);
			if (pNext == NULL) {
				cWinCOMError[0] = '\0';
				wincom_appenderrora("Error, WinCOM: ");
				wincom_appenderrora(cName);
				wincom_appenderrora(" : The value is not a COM object (path: ");
				wincom_appenderrora(cPath);
				wincom_appenderrora(")");
				wincom_cleararguments(aExtra, nExtra);
				return 0;
			}
			pCurrent = pNext;
			continue;
		}
		/* The last part: path arguments + extra arguments (aAll takes the values) */
		nAll = nPathArgs + nExtra;
		aAll = NULL;
		if (nAll > 0) {
			aAll = (WinCOMArg *)malloc(sizeof(WinCOMArg) * (size_t)nAll);
			if (aAll == NULL) {
				wincom_clearpathargs(aPathArgs, nPathArgs);
				IDispatch_Release(pCurrent);
				wincom_cleararguments(aExtra, nExtra);
				wincom_seterrortext(WINCOM_ERROR_OOM);
				return 0;
			}
			for (x = 0; x < nPathArgs; x++) {
				aAll[x].vValue = aPathArgs[x];
				aAll[x].cName = NULL;
				aAll[x].pByRef = NULL;
			}
			for (x = 0; x < nExtra; x++) {
				aAll[nPathArgs + x] = aExtra[x];
			}
		}
		if (aExtra != NULL) {
			free(aExtra);
			aExtra = NULL;
		}
		switch (nMode) {
		case WINCOM_MODE_SET:
			nFlags = DISPATCH_PROPERTYPUT;
			break;
		case WINCOM_MODE_SETREF:
			nFlags = DISPATCH_PROPERTYPUTREF;
			break;
		default:
			nFlags = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
		}
		if ((nMode == WINCOM_MODE_SET) || (nMode == WINCOM_MODE_SETREF)) {
			/* The value is the last argument after the path (not an argument written in the path) */
			if (nAll == nPathArgs) {
				wincom_cleararguments(aAll, nAll);
				IDispatch_Release(pCurrent);
				wincom_seterrortext("Error, WinCOM: The value is missing");
				return 0;
			}
			hr = wincom_invokemember(pCurrent, cName, nFlags, aAll, nAll, NULL, &vExcepInfo);
			lValueIsObject = (V_VT(&aAll[nAll - 1].vValue) == VT_DISPATCH) ||
					 (V_VT(&aAll[nAll - 1].vValue) == VT_UNKNOWN);
			if ((nMode == WINCOM_MODE_SET) && lValueIsObject && wincom_isputerror(hr, &vExcepInfo)) {
				/*
				**  Like (Set x.Property = Object) in Visual Basic
				**  Only when the property doesn't accept the object using DISPATCH_PROPERTYPUT
				**  (Errors raised by the property itself are not retried, so the property runs one time)
				**  If the property doesn't support DISPATCH_PROPERTYPUTREF we keep the first error
				*/
				memset(&vExcepInfo2, 0, sizeof(EXCEPINFO));
				hr2 = wincom_invokemember(pCurrent, cName, DISPATCH_PROPERTYPUTREF, aAll, nAll, NULL,
							  &vExcepInfo2);
				if (SUCCEEDED(hr2) || ((hr2 != DISP_E_MEMBERNOTFOUND) && (hr2 != DISP_E_BADPARAMCOUNT))) {
					wincom_clearexcepinfo(&vExcepInfo);
					vExcepInfo = vExcepInfo2;
					hr = hr2;
				} else {
					wincom_clearexcepinfo(&vExcepInfo2);
				}
			}
		} else {
			hr = wincom_invokemember(pCurrent, cName, nFlags, aAll, nAll, pResult, &vExcepInfo);
		}
		wincom_cleararguments(aAll, nAll);
		IDispatch_Release(pCurrent);
		if (FAILED(hr)) {
			wincom_seterror(cName[0] == '\0' ? "(Default)" : cName, hr, &vExcepInfo);
			wincom_clearexcepinfo(&vExcepInfo);
			VariantClear(pResult);
			return 0;
		}
		wincom_clearexcepinfo(&vExcepInfo);
		nWinCOMLastHResult = (unsigned long)hr;
		return 1;
	}
}

/* Convert the parameters (nStart to RING_API_PARACOUNT) to an array of arguments (malloc) */
static int wincom_getparaarguments(void *pPointer, int nStart, WinCOMArg **pArgs, int *pCount) {
	WinCOMArg *aArgs;
	int nCount, x;
	*pArgs = NULL;
	*pCount = 0;
	nCount = RING_API_PARACOUNT - nStart + 1;
	if (nCount <= 0) {
		return 1;
	}
	aArgs = (WinCOMArg *)malloc(sizeof(WinCOMArg) * (size_t)nCount);
	if (aArgs == NULL) {
		wincom_seterrortext(WINCOM_ERROR_OOM);
		return 0;
	}
	for (x = 0; x < nCount; x++) {
		if (!wincom_paratoarg(pPointer, nStart + x, &aArgs[x])) {
			wincom_cleararguments(aArgs, x);
			return 0;
		}
	}
	*pArgs = aArgs;
	*pCount = nCount;
	return 1;
}

/* Convert the items of a Ring list to an array of arguments (malloc) */
static int wincom_getlistarguments(List *pList, WinCOMArg **pArgs, int *pCount) {
	WinCOMArg *aArgs;
	unsigned int nCount, x;
	*pArgs = NULL;
	*pCount = 0;
	nCount = ring_list_getsize(pList);
	if (nCount == 0) {
		return 1;
	}
	aArgs = (WinCOMArg *)malloc(sizeof(WinCOMArg) * (size_t)nCount);
	if (aArgs == NULL) {
		wincom_seterrortext(WINCOM_ERROR_OOM);
		return 0;
	}
	for (x = 1; x <= nCount; x++) {
		if (!wincom_itemtoarg(pList, x, &aArgs[x - 1])) {
			wincom_cleararguments(aArgs, (int)x - 1);
			return 0;
		}
	}
	*pArgs = aArgs;
	*pCount = (int)nCount;
	return 1;
}

/* Check the first two parameters (COM object and member name/path) */
static IDispatch *wincom_checkobjectandname(void *pPointer, int nMinParameters) {
	IDispatch *pDispatch;
	if (RING_API_PARACOUNT < nMinParameters) {
		RING_API_ERROR(WINCOM_ERROR_PARACOUNT);
		return NULL;
	}
	pDispatch = wincom_getobject(pPointer, 1);
	if (pDispatch == NULL) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return NULL;
	}
	if (!RING_API_ISSTRING(2)) {
		RING_API_ERROR(WINCOM_ERROR_NAME);
		return NULL;
	}
	return pDispatch;
}

static void wincom_callwithparameters(void *pPointer, int nMode) {
	IDispatch *pDispatch;
	WinCOMArg *aArgs;
	int nArgs;
	pDispatch = wincom_checkobjectandname(pPointer, ((nMode == WINCOM_MODE_SET) || (nMode == WINCOM_MODE_SETREF)) ? 3 : 2);
	if (pDispatch == NULL) {
		return;
	}
	if (!wincom_getparaarguments(pPointer, 3, &aArgs, &nArgs)) {
		RING_API_ERROR(cWinCOMError);
		return;
	}
	wincom_callandreturn(pPointer, pDispatch, RING_API_GETSTRING(2), nMode, aArgs, nArgs);
}

/*
**  Call the member then return the result (or raise the error)
**  The path is copied, because event handlers (Ring code) can run during the call and change the variables
**  An error in an event handler is raised here (after the call)
*/
static void wincom_callandreturn(void *pPointer, IDispatch *pDispatch, const char *cPath, int nMode, WinCOMArg *aArgs,
				 int nArgs) {
	VARIANT vResult;
	char *cPathCopy;
	int lDone;
	cPathCopy = wincom_strdup(cPath);
	if (cPathCopy == NULL) {
		wincom_cleararguments(aArgs, nArgs);
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	wincom_begincall();
	lDone = wincom_callpath(pDispatch, cPathCopy, nMode, aArgs, nArgs, &vResult);
	free(cPathCopy);
	if (wincom_endcall(pPointer)) {
		VariantClear(&vResult);
		return;
	}
	if (!lDone) {
		RING_API_ERROR(cWinCOMError);
		return;
	}
	if ((nMode == WINCOM_MODE_GET) || (nMode == WINCOM_MODE_INVOKE)) {
		wincom_retvariant(pPointer, &vResult);
	}
}

/* Ring Functions */

/* com_init() --> 1 if COM is initialized (Called automatically by the other functions) */
RING_FUNC(ring_com_init) {
	RING_API_RETNUMBER(wincom_init());
}

/* Get CLSID from ProgID ("ADODB.Connection") or from CLSID string ("{...}") */
static HRESULT wincom_getclsid(const char *cName, int nSize, CLSID *pCLSID) {
	BSTR cWide;
	HRESULT hr;
	cWide = wincom_tobstr(cName, nSize);
	if (cWide == NULL) {
		return E_OUTOFMEMORY;
	}
	if (cWide[0] == L'{') {
		hr = CLSIDFromString(cWide, pCLSID);
	} else {
		hr = CLSIDFromProgID(cWide, pCLSID);
	}
	SysFreeString(cWide);
	return hr;
}

/* Check that the function has one string parameter, then initialize COM (returns 0 after raising an error) */
static int wincom_checkstringparameter(void *pPointer) {
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return 0;
	}
	if (!RING_API_ISSTRING(1)) {
		RING_API_ERROR(RING_API_BADPARATYPE);
		return 0;
	}
	if (!wincom_init()) {
		RING_API_ERROR(WINCOM_ERROR_INIT);
		return 0;
	}
	return 1;
}

/* Return the COM object, or raise the error (cName is the ProgID or the moniker) */
static void wincom_returnobject(void *pPointer, IDispatch *pDispatch, HRESULT hr, const char *cName) {
	nWinCOMLastHResult = (unsigned long)hr;
	if (FAILED(hr)) {
		wincom_seterror(cName, hr, NULL);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	RING_API_RETMANAGEDCPOINTER(pDispatch, WINCOM_OBJECT, wincom_freeobject);
}

/* Create the object on another computer (DCOM), the class must be registered on this computer too (or use a CLSID) */
static HRESULT wincom_createremote(CLSID *pCLSID, const char *cServer, int nSize, IDispatch **ppDispatch) {
	COSERVERINFO vServer;
	MULTI_QI vQI;
	BSTR cWide;
	HRESULT hr;
	cWide = wincom_tobstr(cServer, nSize);
	if (cWide == NULL) {
		return E_OUTOFMEMORY;
	}
	memset(&vServer, 0, sizeof(COSERVERINFO));
	vServer.pwszName = cWide;
	memset(&vQI, 0, sizeof(MULTI_QI));
	vQI.pIID = &IID_IDispatch;
	hr = CoCreateInstanceEx(pCLSID, NULL, CLSCTX_REMOTE_SERVER | CLSCTX_LOCAL_SERVER, &vServer, 1, &vQI);
	SysFreeString(cWide);
	if (SUCCEEDED(hr)) {
		hr = vQI.hr;
	}
	if (SUCCEEDED(hr)) {
		*ppDispatch = (IDispatch *)vQI.pItf;
	} else if (vQI.pItf != NULL) {
		IUnknown_Release(vQI.pItf);
	}
	return hr;
}

/* com_createobject(cProgID, [cServer]) --> COM object (cServer: the computer name for DCOM) */
RING_FUNC(ring_com_createobject) {
	CLSID vCLSID;
	IDispatch *pDispatch;
	HRESULT hr;
	int lRemote;
	if ((RING_API_PARACOUNT < 1) || (RING_API_PARACOUNT > 2)) {
		RING_API_ERROR(WINCOM_ERROR_PARACOUNT);
		return;
	}
	if (!RING_API_ISSTRING(1) || ((RING_API_PARACOUNT == 2) && !RING_API_ISSTRING(2))) {
		RING_API_ERROR(RING_API_BADPARATYPE);
		return;
	}
	if (!wincom_init()) {
		RING_API_ERROR(WINCOM_ERROR_INIT);
		return;
	}
	lRemote = (RING_API_PARACOUNT == 2) && (RING_API_GETSTRINGSIZE(2) > 0);
	pDispatch = NULL;
	hr = wincom_getclsid(RING_API_GETSTRING(1), RING_API_GETSTRINGSIZE(1), &vCLSID);
	if (SUCCEEDED(hr)) {
		if (lRemote) {
			hr = wincom_createremote(&vCLSID, RING_API_GETSTRING(2), RING_API_GETSTRINGSIZE(2), &pDispatch);
		} else {
			hr = CoCreateInstance(&vCLSID, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, &IID_IDispatch,
					      (void **)&pDispatch);
		}
	}
	wincom_returnobject(pPointer, pDispatch, hr, RING_API_GETSTRING(1));
}

/* com_getobject(cName) --> COM object (Using a moniker: "winmgmts:", file name, ...) */
RING_FUNC(ring_com_getobject) {
	BSTR cWide;
	IDispatch *pDispatch;
	HRESULT hr;
	if (!wincom_checkstringparameter(pPointer)) {
		return;
	}
	cWide = wincom_tobstr(RING_API_GETSTRING(1), RING_API_GETSTRINGSIZE(1));
	if (cWide == NULL) {
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	pDispatch = NULL;
	hr = CoGetObject(cWide, NULL, &IID_IDispatch, (void **)&pDispatch);
	SysFreeString(cWide);
	wincom_returnobject(pPointer, pDispatch, hr, RING_API_GETSTRING(1));
}

/* com_getactiveobject(cProgID) --> COM object (The running instance, like Excel.Application) */
RING_FUNC(ring_com_getactiveobject) {
	CLSID vCLSID;
	IUnknown *pUnknown;
	IDispatch *pDispatch;
	HRESULT hr;
	if (!wincom_checkstringparameter(pPointer)) {
		return;
	}
	pDispatch = NULL;
	hr = wincom_getclsid(RING_API_GETSTRING(1), RING_API_GETSTRINGSIZE(1), &vCLSID);
	if (SUCCEEDED(hr)) {
		hr = GetActiveObject(&vCLSID, NULL, &pUnknown);
		if (SUCCEEDED(hr)) {
			hr = IUnknown_QueryInterface(pUnknown, &IID_IDispatch, (void **)&pDispatch);
			IUnknown_Release(pUnknown);
		}
	}
	wincom_returnobject(pPointer, pDispatch, hr, RING_API_GETSTRING(1));
}

/* com_progidexists(cProgID) --> 1 if the ProgID (or CLSID) is registered */
RING_FUNC(ring_com_progidexists) {
	CLSID vCLSID;
	if (!wincom_checkstringparameter(pPointer)) {
		return;
	}
	RING_API_RETNUMBER(SUCCEEDED(wincom_getclsid(RING_API_GETSTRING(1), RING_API_GETSTRINGSIZE(1), &vCLSID)) ? 1
														 : 0);
}

/* com_get(oObject, cPath, [Arguments...]) --> Property value */
RING_FUNC(ring_com_get) { wincom_callwithparameters(pPointer, WINCOM_MODE_GET); }

/* com_invoke(oObject, cPath, [Arguments...]) --> Result of the method */
RING_FUNC(ring_com_invoke) { wincom_callwithparameters(pPointer, WINCOM_MODE_INVOKE); }

/* com_set(oObject, cPath, [Arguments...], vValue) */
RING_FUNC(ring_com_set) { wincom_callwithparameters(pPointer, WINCOM_MODE_SET); }

/* com_setref(oObject, cPath, [Arguments...], oValue) */
RING_FUNC(ring_com_setref) { wincom_callwithparameters(pPointer, WINCOM_MODE_SETREF); }

/* com_dispatch(oObject, cPath, nMode, aArguments) --> Result (nMode: COM_GET, COM_INVOKE, COM_SET, COM_SETREF) */
RING_FUNC(ring_com_dispatch) {
	IDispatch *pDispatch;
	WinCOMArg *aArgs;
	int nArgs, nMode;
	if (RING_API_PARACOUNT != 4) {
		RING_API_ERROR(RING_API_MISS4PARA);
		return;
	}
	pDispatch = wincom_checkobjectandname(pPointer, 4);
	if (pDispatch == NULL) {
		return;
	}
	if (!RING_API_ISNUMBER(3) || !RING_API_ISLIST(4)) {
		RING_API_ERROR(RING_API_BADPARATYPE);
		return;
	}
	nMode = (int)RING_API_GETNUMBER(3);
	if ((nMode < WINCOM_MODE_GET) || (nMode > WINCOM_MODE_SETREF)) {
		RING_API_ERROR("Error, WinCOM: Bad mode (Use COM_GET, COM_INVOKE, COM_SET or COM_SETREF)");
		return;
	}
	if (!wincom_getlistarguments(RING_API_GETLIST(4), &aArgs, &nArgs)) {
		RING_API_ERROR(cWinCOMError);
		return;
	}
	wincom_callandreturn(pPointer, pDispatch, RING_API_GETSTRING(2), nMode, aArgs, nArgs);
}

/* com_items(oObject) --> List of the items of a collection (Using IEnumVARIANT, like For Each in Visual Basic) */
RING_FUNC(ring_com_items) {
	IDispatch *pDispatch;
	IEnumVARIANT *pEnum;
	VARIANT vResult, vItem;
	EXCEPINFO vExcepInfo;
	ULONG nFetched;
	List *pList;
	HRESULT hr;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pDispatch = wincom_getobject(pPointer, 1);
	if (pDispatch == NULL) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return;
	}
	memset(&vExcepInfo, 0, sizeof(EXCEPINFO));
	wincom_begincall();
	hr = wincom_invokemember(pDispatch, "_NewEnum", DISPATCH_METHOD | DISPATCH_PROPERTYGET, NULL, 0, &vResult,
				 &vExcepInfo);
	if (FAILED(hr)) {
		wincom_seterror("_NewEnum", hr, &vExcepInfo);
		wincom_clearexcepinfo(&vExcepInfo);
		if (wincom_endcall(pPointer)) {
			return;
		}
		RING_API_ERROR(cWinCOMError);
		return;
	}
	wincom_clearexcepinfo(&vExcepInfo);
	pEnum = NULL;
	hr = E_NOINTERFACE;
	if (((V_VT(&vResult) == VT_UNKNOWN) || (V_VT(&vResult) == VT_DISPATCH)) && (V_UNKNOWN(&vResult) != NULL)) {
		hr = IUnknown_QueryInterface(V_UNKNOWN(&vResult), &IID_IEnumVARIANT, (void **)&pEnum);
	}
	VariantClear(&vResult);
	if (FAILED(hr)) {
		wincom_seterror("_NewEnum", hr, NULL);
		if (wincom_endcall(pPointer)) {
			return;
		}
		RING_API_ERROR(cWinCOMError);
		return;
	}
	pList = RING_API_NEWLIST;
	while (1) {
		VariantInit(&vItem);
		nFetched = 0;
		hr = IEnumVARIANT_Next(pEnum, 1, &vItem, &nFetched);
		if ((hr != S_OK) || (nFetched == 0)) {
			VariantClear(&vItem);
			break;
		}
		wincom_addvariant(RING_API_STATE, pList, &vItem, 0);
		VariantClear(&vItem);
	}
	IEnumVARIANT_Release(pEnum);
	if (wincom_endcall(pPointer)) {
		return;
	}
	RING_API_RETLISTBYREF(pList);
}

/* com_isobject(vValue) --> 1 if the value is a COM object */
RING_FUNC(ring_com_isobject) {
	List *pList;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	if (RING_API_ISLIST(1)) {
		pList = RING_API_GETLIST(1);
		if (ring_list_iscpointerlist(pList) &&
		    (strcmp(ring_list_getstring(pList, RING_CPOINTER_TYPE), WINCOM_OBJECT) == 0) &&
		    (ring_list_getpointer(pList, RING_CPOINTER_POINTER) != NULL)) {
			RING_API_RETNUMBER(1);
			return;
		}
	}
	RING_API_RETNUMBER(0);
}

/* com_release(oObject) --> Release the COM object now (The variable can't be used after that) */
RING_FUNC(ring_com_release) {
	List *pList;
	Item *pItem;
	void *pPtr;
	const char *cType;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	if (!RING_API_ISLIST(1) || !ring_list_iscpointerlist(RING_API_GETLIST(1))) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return;
	}
	pList = RING_API_GETLIST(1);
	cType = ring_list_getstring(pList, RING_CPOINTER_TYPE);
	if ((strcmp(cType, WINCOM_OBJECT) != 0) && (strcmp(cType, WINCOM_UNKNOWN) != 0) &&
	    (strcmp(cType, WINCOM_VARIANT) != 0) && (strcmp(cType, WINCOM_BYREF) != 0) &&
	    (strcmp(cType, WINCOM_NAMED) != 0) && (strcmp(cType, WINCOM_EVENTS) != 0)) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return;
	}
	/* The item is shared by the copies of the variable, so all of them become NULL */
	pItem = ring_list_getitem(pList, RING_CPOINTER_POINTER);
	pPtr = pItem->data.pPointer;
	if (pPtr != NULL) {
		pItem->data.pPointer = NULL;
		wincom_begincall();
		if (strcmp(cType, WINCOM_VARIANT) == 0) {
			wincom_freevariant(NULL, pPtr);
		} else if (strcmp(cType, WINCOM_BYREF) == 0) {
			wincom_freebyref(NULL, pPtr);
		} else if (strcmp(cType, WINCOM_NAMED) == 0) {
			wincom_freenamed(NULL, pPtr);
		} else if (strcmp(cType, WINCOM_EVENTS) == 0) {
			wincom_freeevents(NULL, pPtr);
		} else {
			wincom_freeobject(NULL, pPtr);
		}
		wincom_endcall(pPointer);
	}
}

/*
**  Create a one dimension array of nBaseType (like VT_BSTR or VT_I4) from the items of a Ring list
**  Each item is converted to nBaseType (Strings using the invariant locale)
**  Returns 1 on success, 0 on error (cWinCOMError contains the error message)
*/
static int wincom_listtotypedarray(List *pList, VARTYPE nBaseType, VARIANT *pVar) {
	SAFEARRAY *pArray;
	VARIANT vItem, vConverted;
	DECIMAL vDecimal;
	LONG nIndex;
	unsigned int x, nSize;
	HRESULT hr;
	VariantInit(pVar);
	if (!wincom_issimpletype(nBaseType) && (nBaseType != VT_DECIMAL)) {
		wincom_seterrortext("Error, WinCOM: The array type is not supported");
		return 0;
	}
	nSize = ring_list_getsize(pList);
	pArray = SafeArrayCreateVector(nBaseType, 0, nSize);
	if (pArray == NULL) {
		wincom_seterrortext(WINCOM_ERROR_OOM);
		return 0;
	}
	for (x = 1; x <= nSize; x++) {
		if (ring_list_islist(pList, x) && !ring_list_iscpointerlist(ring_list_getlist(pList, x))) {
			SafeArrayDestroy(pArray);
			wincom_seterrortext("Error, WinCOM: Typed arrays can't contain lists (use VT_ARRAY + VT_VARIANT)");
			return 0;
		}
		if (!wincom_itemtovariant(pList, x, &vItem, 1)) {
			SafeArrayDestroy(pArray);
			return 0;
		}
		VariantInit(&vConverted);
		hr = VariantChangeTypeEx(&vConverted, &vItem, LOCALE_INVARIANT, 0, nBaseType);
		VariantClear(&vItem);
		if (FAILED(hr)) {
			SafeArrayDestroy(pArray);
			wincom_seterror("com_variant()", hr, NULL);
			return 0;
		}
		nIndex = (LONG)x - 1;
		/* SafeArrayPutElement() copies the data (BSTR, objects) */
		if ((nBaseType == VT_BSTR) || (nBaseType == VT_DISPATCH) || (nBaseType == VT_UNKNOWN)) {
			hr = SafeArrayPutElement(pArray, &nIndex, V_BYREF(&vConverted));
		} else if (nBaseType == VT_DECIMAL) {
			/* DECIMAL uses the VARIANT type field (wReserved) */
			vDecimal = V_DECIMAL(&vConverted);
			vDecimal.wReserved = 0;
			hr = SafeArrayPutElement(pArray, &nIndex, &vDecimal);
		} else {
			hr = SafeArrayPutElement(pArray, &nIndex, &V_UI1(&vConverted));
		}
		VariantClear(&vConverted);
		if (FAILED(hr)) {
			SafeArrayDestroy(pArray);
			wincom_seterror("com_variant()", hr, NULL);
			return 0;
		}
	}
	V_VT(pVar) = (VARTYPE)(VT_ARRAY | nBaseType);
	V_ARRAY(pVar) = pArray;
	return 1;
}

/* com_variant(nVarType, vValue) --> VARIANT (C pointer) that can be passed to COM methods */
RING_FUNC(ring_com_variant) {
	VARIANT *pVar;
	VARIANT vValue;
	VARTYPE nType;
	DATE dDate;
	SAFEARRAY *pArray;
	void *pData;
	const char *cStr;
	int nSize;
	HRESULT hr;
	if (RING_API_PARACOUNT != 2) {
		RING_API_ERROR(RING_API_MISS2PARA);
		return;
	}
	if (!RING_API_ISNUMBER(1)) {
		RING_API_ERROR(RING_API_BADPARATYPE);
		return;
	}
	nType = (VARTYPE)(int)RING_API_GETNUMBER(1);
	pVar = (VARIANT *)malloc(sizeof(VARIANT));
	if (pVar == NULL) {
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	VariantInit(pVar);
	hr = S_OK;
	if ((nType == VT_EMPTY) || (nType == VT_NULL)) {
		V_VT(pVar) = nType;
	} else if (nType == VT_ERROR) {
		/* Missing optional parameter (or other error code) */
		V_VT(pVar) = VT_ERROR;
		V_ERROR(pVar) = DISP_E_PARAMNOTFOUND;
		if (RING_API_ISNUMBER(2) && (RING_API_GETNUMBER(2) != 0)) {
			/* Signed (-2147352572) or unsigned (2147614724, like com_lasthresult()) 32-bit values */
			if ((RING_API_GETNUMBER(2) < -2147483648.0) || (RING_API_GETNUMBER(2) > 4294967295.0)) {
				free(pVar);
				RING_API_ERROR("Error, WinCOM: The error code must be a 32-bit number");
				return;
			}
			V_ERROR(pVar) = (SCODE)(ULONG)(LONGLONG)RING_API_GETNUMBER(2);
		}
	} else if ((nType == (VT_ARRAY | VT_UI1)) && RING_API_ISSTRING(2)) {
		/* Binary data */
		cStr = RING_API_GETSTRING(2);
		nSize = RING_API_GETSTRINGSIZE(2);
		pArray = SafeArrayCreateVector(VT_UI1, 0, (ULONG)nSize);
		if (pArray == NULL) {
			free(pVar);
			RING_API_ERROR(WINCOM_ERROR_OOM);
			return;
		}
		if (nSize > 0) {
			if (SUCCEEDED(SafeArrayAccessData(pArray, &pData))) {
				memcpy(pData, cStr, (size_t)nSize);
				SafeArrayUnaccessData(pArray);
			}
		}
		V_VT(pVar) = VT_ARRAY | VT_UI1;
		V_ARRAY(pVar) = pArray;
	} else if (((nType & ~VT_TYPEMASK) == VT_ARRAY) && ((nType & VT_TYPEMASK) != VT_VARIANT) && RING_API_ISLIST(2) &&
		   !ring_list_iscpointerlist(RING_API_GETLIST(2))) {
		/* Typed array (like VT_ARRAY + VT_BSTR) from a list of values */
		if (!wincom_listtotypedarray(RING_API_GETLIST(2), (VARTYPE)(nType & VT_TYPEMASK), pVar)) {
			free(pVar);
			RING_API_ERROR(cWinCOMError);
			return;
		}
	} else if ((nType == VT_DATE) && RING_API_ISSTRING(2) && wincom_stringtodate(RING_API_GETSTRING(2), &dDate)) {
		V_VT(pVar) = VT_DATE;
		V_DATE(pVar) = dDate;
	} else if ((nType == VT_BOOL) && RING_API_ISNUMBER(2)) {
		V_VT(pVar) = VT_BOOL;
		V_BOOL(pVar) = RING_API_GETNUMBER(2) != 0 ? VARIANT_TRUE : VARIANT_FALSE;
	} else {
		if (!wincom_paratovariant(pPointer, 2, &vValue)) {
			free(pVar);
			RING_API_ERROR(cWinCOMError);
			return;
		}
		if (V_VT(&vValue) == nType) {
			*pVar = vValue;
		} else {
			/* Strings are converted using the invariant locale (The decimal separator is .) */
			hr = VariantChangeTypeEx(pVar, &vValue, LOCALE_INVARIANT, 0, nType);
			if (FAILED(hr) && (V_VT(&vValue) == VT_BSTR)) {
				hr = VariantChangeTypeEx(pVar, &vValue, LOCALE_USER_DEFAULT, 0, nType);
			}
			VariantClear(&vValue);
		}
	}
	if (FAILED(hr)) {
		VariantClear(pVar);
		free(pVar);
		wincom_seterror("com_variant()", hr, NULL);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	RING_API_RETMANAGEDCPOINTER(pVar, WINCOM_VARIANT, wincom_freevariant);
}

/*
**  com_matrix(aRows) --> Two dimensions array (C pointer, like com_variant())
**  aRows is a list of rows, each row is a list of values: [[1, 2], [3, 4]]
**  The array is aRows[nRow][nColumn] (Like Range.Value in Microsoft Excel), the indexes start from 1
**  Short rows are completed using Empty values
*/
RING_FUNC(ring_com_matrix) {
	List *pRows, *pRow;
	VARIANT *pVar, *pCell;
	SAFEARRAY *pArray;
	SAFEARRAYBOUND aBounds[2];
	LONG aIndex[2];
	unsigned int nRows, nColumns, nRow, nColumn, nSize;
	HRESULT hr;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	if (!RING_API_ISLIST(1)) {
		RING_API_ERROR(RING_API_BADPARATYPE);
		return;
	}
	pRows = RING_API_GETLIST(1);
	if (ring_list_iscpointerlist(pRows) || ring_list_isobject(pRows)) {
		RING_API_ERROR("Error, WinCOM: com_matrix() expects a list of rows");
		return;
	}
	nRows = ring_list_getsize(pRows);
	nColumns = 0;
	for (nRow = 1; nRow <= nRows; nRow++) {
		nSize = 1;
		if (ring_list_islist(pRows, nRow)) {
			pRow = ring_list_getlist(pRows, nRow);
			if (!ring_list_iscpointerlist(pRow) && !ring_list_isobject(pRow)) {
				nSize = ring_list_getsize(pRow);
			}
		}
		if (nSize > nColumns) {
			nColumns = nSize;
		}
	}
	/* The first dimension is the rows (aBounds[0] is the left dimension) */
	aBounds[0].lLbound = 1;
	aBounds[0].cElements = nRows;
	aBounds[1].lLbound = 1;
	aBounds[1].cElements = nColumns;
	pArray = SafeArrayCreate(VT_VARIANT, 2, aBounds);
	if (pArray == NULL) {
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	for (nRow = 1; nRow <= nRows; nRow++) {
		pRow = NULL;
		if (ring_list_islist(pRows, nRow)) {
			pRow = ring_list_getlist(pRows, nRow);
			if (ring_list_iscpointerlist(pRow) || ring_list_isobject(pRow)) {
				pRow = NULL;
			}
		}
		nSize = pRow != NULL ? ring_list_getsize(pRow) : 1;
		for (nColumn = 1; nColumn <= nSize; nColumn++) {
			/* SafeArrayPtrOfIndex() indexes are in the order of the dimensions */
			aIndex[0] = (LONG)nRow;
			aIndex[1] = (LONG)nColumn;
			hr = SafeArrayPtrOfIndex(pArray, aIndex, (void **)&pCell);
			if (FAILED(hr)) {
				SafeArrayDestroy(pArray);
				wincom_seterror("com_matrix()", hr, NULL);
				RING_API_ERROR(cWinCOMError);
				return;
			}
			if (pRow != NULL) {
				if (!wincom_itemtovariant(pRow, nColumn, pCell, 1)) {
					SafeArrayDestroy(pArray);
					RING_API_ERROR(cWinCOMError);
					return;
				}
			} else if (!wincom_itemtovariant(pRows, nRow, pCell, 1)) {
				SafeArrayDestroy(pArray);
				RING_API_ERROR(cWinCOMError);
				return;
			}
		}
	}
	pVar = (VARIANT *)malloc(sizeof(VARIANT));
	if (pVar == NULL) {
		SafeArrayDestroy(pArray);
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	VariantInit(pVar);
	V_VT(pVar) = VT_ARRAY | VT_VARIANT;
	V_ARRAY(pVar) = pArray;
	RING_API_RETMANAGEDCPOINTER(pVar, WINCOM_VARIANT, wincom_freevariant);
}

/* com_lasttype() --> VARIANT type of the last value returned by com_get(), com_invoke() or com_dispatch() */
RING_FUNC(ring_com_lasttype) { RING_API_RETNUMBER((double)nWinCOMLastType); }

/* com_lasthresult() --> HRESULT of the last call (0 = S_OK), as unsigned number (0x80020006 = 2147614726) */
RING_FUNC(ring_com_lasthresult) { RING_API_RETNUMBER((double)nWinCOMLastHResult); }

/* Get the type information of the object, NULL if not available */
static ITypeInfo *wincom_gettypeinfo(IDispatch *pDispatch) {
	ITypeInfo *pTypeInfo;
	UINT nCount;
	nCount = 0;
	if (FAILED(IDispatch_GetTypeInfoCount(pDispatch, &nCount)) || (nCount == 0)) {
		return NULL;
	}
	pTypeInfo = NULL;
	if (FAILED(IDispatch_GetTypeInfo(pDispatch, 0, LOCALE_USER_DEFAULT, &pTypeInfo))) {
		return NULL;
	}
	return pTypeInfo;
}

static void wincom_addbstr(void *pState, List *pList, BSTR cStr) {
	char *cUTF8;
	int nSize;
	cUTF8 = wincom_toutf8(cStr, cStr != NULL ? (int)SysStringLen(cStr) : 0, &nSize);
	if (cUTF8 != NULL) {
		ring_list_addstring2_gc(pState, pList, cUTF8, (unsigned int)nSize);
		free(cUTF8);
	} else {
		ring_list_addstring_gc(pState, pList, "");
	}
}

/* com_typename(oObject) --> The interface name (like _Connection or _Recordset), "" if not available */
RING_FUNC(ring_com_typename) {
	IDispatch *pDispatch;
	ITypeInfo *pTypeInfo;
	BSTR cName;
	List *pList;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pDispatch = wincom_getobject(pPointer, 1);
	if (pDispatch == NULL) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return;
	}
	pTypeInfo = wincom_gettypeinfo(pDispatch);
	if (pTypeInfo == NULL) {
		RING_API_RETSTRING("");
		return;
	}
	cName = NULL;
	ITypeInfo_GetDocumentation(pTypeInfo, MEMBERID_NIL, &cName, NULL, NULL, NULL);
	ITypeInfo_Release(pTypeInfo);
	pList = RING_API_NEWLIST;
	wincom_addbstr(RING_API_STATE, pList, cName);
	if (cName != NULL) {
		SysFreeString(cName);
	}
	RING_API_RETSTRING2(ring_list_getstring(pList, 1), (int)ring_list_getstringsize(pList, 1));
}

/* Add the description and the parameter names of a member: cDescription, [cParameter1, cParameter2, ...] */
static void wincom_addmemberinfo(void *pState, List *pItem, ITypeInfo *pTypeInfo, MEMBERID nMemberID, BSTR cDoc,
				 UINT nParams) {
	BSTR *aNames;
	UINT nNames, x;
	List *pParams;
	wincom_addbstr(pState, pItem, cDoc);
	pParams = ring_list_newlist_gc(pState, pItem);
	if (nParams == 0) {
		return;
	}
	aNames = (BSTR *)calloc((size_t)nParams + 1, sizeof(BSTR));
	if (aNames == NULL) {
		return;
	}
	nNames = 0;
	if (SUCCEEDED(ITypeInfo_GetNames(pTypeInfo, nMemberID, aNames, nParams + 1, &nNames))) {
		/* aNames[0] is the member name */
		for (x = 1; x < nNames; x++) {
			wincom_addbstr(pState, pParams, aNames[x]);
		}
		for (x = 0; x < nNames; x++) {
			SysFreeString(aNames[x]);
		}
	}
	free(aNames);
}

/*
**  com_members(oObject) --> List of [cName, cKind, nParameters, cDescription, aParameterNames]
**  cKind is "method", "get" (property get), "put" (property put), "putref" or "property" (variable)
*/
RING_FUNC(ring_com_members) {
	IDispatch *pDispatch;
	ITypeInfo *pTypeInfo;
	TYPEATTR *pTypeAttr;
	FUNCDESC *pFuncDesc;
	VARDESC *pVarDesc;
	BSTR cName, cDoc;
	List *pList, *pItem;
	const char *cKind;
	UINT x;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pDispatch = wincom_getobject(pPointer, 1);
	if (pDispatch == NULL) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return;
	}
	pList = RING_API_NEWLIST;
	pTypeInfo = wincom_gettypeinfo(pDispatch);
	if (pTypeInfo == NULL) {
		RING_API_RETLISTBYREF(pList);
		return;
	}
	if (SUCCEEDED(ITypeInfo_GetTypeAttr(pTypeInfo, &pTypeAttr))) {
		for (x = 0; x < pTypeAttr->cFuncs; x++) {
			if (FAILED(ITypeInfo_GetFuncDesc(pTypeInfo, x, &pFuncDesc))) {
				continue;
			}
			/* Hide the members of IUnknown and IDispatch */
			if (!(pFuncDesc->wFuncFlags & (FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN))) {
				cName = NULL;
				cDoc = NULL;
				if (SUCCEEDED(ITypeInfo_GetDocumentation(pTypeInfo, pFuncDesc->memid, &cName, &cDoc, NULL,
									  NULL))) {
					switch (pFuncDesc->invkind) {
					case INVOKE_PROPERTYGET:
						cKind = "get";
						break;
					case INVOKE_PROPERTYPUT:
						cKind = "put";
						break;
					case INVOKE_PROPERTYPUTREF:
						cKind = "putref";
						break;
					default:
						cKind = "method";
					}
					pItem = ring_list_newlist_gc(RING_API_STATE, pList);
					wincom_addbstr(RING_API_STATE, pItem, cName);
					ring_list_addstring_gc(RING_API_STATE, pItem, cKind);
					ring_list_adddouble_gc(RING_API_STATE, pItem, (double)pFuncDesc->cParams);
					wincom_addmemberinfo(RING_API_STATE, pItem, pTypeInfo, pFuncDesc->memid, cDoc,
							     (UINT)pFuncDesc->cParams);
					SysFreeString(cName);
					SysFreeString(cDoc);
				}
			}
			ITypeInfo_ReleaseFuncDesc(pTypeInfo, pFuncDesc);
		}
		for (x = 0; x < pTypeAttr->cVars; x++) {
			if (FAILED(ITypeInfo_GetVarDesc(pTypeInfo, x, &pVarDesc))) {
				continue;
			}
			cName = NULL;
			cDoc = NULL;
			if (SUCCEEDED(ITypeInfo_GetDocumentation(pTypeInfo, pVarDesc->memid, &cName, &cDoc, NULL, NULL))) {
				pItem = ring_list_newlist_gc(RING_API_STATE, pList);
				wincom_addbstr(RING_API_STATE, pItem, cName);
				ring_list_addstring_gc(RING_API_STATE, pItem, "property");
				ring_list_adddouble_gc(RING_API_STATE, pItem, 0);
				wincom_addmemberinfo(RING_API_STATE, pItem, pTypeInfo, pVarDesc->memid, cDoc, 0);
				SysFreeString(cName);
				SysFreeString(cDoc);
			}
			ITypeInfo_ReleaseVarDesc(pTypeInfo, pVarDesc);
		}
		ITypeInfo_ReleaseTypeAttr(pTypeInfo, pTypeAttr);
	}
	ITypeInfo_Release(pTypeInfo);
	RING_API_RETLISTBYREF(pList);
}

/* com_refcount(oObject) --> The reference count of the COM object (for testing) */
RING_FUNC(ring_com_refcount) {
	IDispatch *pDispatch;
	ULONG nCount;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pDispatch = wincom_getobject(pPointer, 1);
	if (pDispatch == NULL) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return;
	}
	IDispatch_AddRef(pDispatch);
	nCount = IDispatch_Release(pDispatch);
	RING_API_RETNUMBER((double)nCount);
}

/* ByRef and named arguments */

/* com_byref(vValue) --> Argument passed by reference (The member can change the value, read it using com_byrefvalue()) */
RING_FUNC(ring_com_byref) {
	WinCOMByRef *pByRef;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pByRef = (WinCOMByRef *)malloc(sizeof(WinCOMByRef));
	if (pByRef == NULL) {
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	pByRef->nRefCount = 1;
	if (!wincom_paratovariant(pPointer, 1, &pByRef->vValue)) {
		free(pByRef);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	RING_API_RETMANAGEDCPOINTER(pByRef, WINCOM_BYREF, wincom_freebyref);
}

/* com_byrefvalue(oByRef) --> The value of the argument (after calling the member) */
RING_FUNC(ring_com_byrefvalue) {
	WinCOMByRef *pByRef;
	VARIANT vCopy;
	HRESULT hr;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pByRef = RING_API_ISLIST(1) ? (WinCOMByRef *)wincom_getlistpointer(RING_API_GETLIST(1), WINCOM_BYREF) : NULL;
	if (pByRef == NULL) {
		RING_API_ERROR("Error, WinCOM: The parameter must be created by com_byref()");
		return;
	}
	VariantInit(&vCopy);
	hr = VariantCopy(&vCopy, &pByRef->vValue);
	if (FAILED(hr)) {
		wincom_seterror("com_byrefvalue()", hr, NULL);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	wincom_retvariant(pPointer, &vCopy);
}

/* com_named(cName, vValue) --> Named argument (Like Name:=Value in Visual Basic) */
RING_FUNC(ring_com_named) {
	WinCOMNamed *pNamed;
	if (RING_API_PARACOUNT != 2) {
		RING_API_ERROR(RING_API_MISS2PARA);
		return;
	}
	if (!RING_API_ISSTRING(1) || (RING_API_GETSTRINGSIZE(1) == 0)) {
		RING_API_ERROR("Error, WinCOM: The argument name must be a string");
		return;
	}
	pNamed = (WinCOMNamed *)malloc(sizeof(WinCOMNamed));
	if (pNamed == NULL) {
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	pNamed->cName = wincom_tobstr(RING_API_GETSTRING(1), RING_API_GETSTRINGSIZE(1));
	if (pNamed->cName == NULL) {
		free(pNamed);
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	if (!wincom_paratovariant(pPointer, 2, &pNamed->vValue)) {
		SysFreeString(pNamed->cName);
		free(pNamed);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	RING_API_RETMANAGEDCPOINTER(pNamed, WINCOM_NAMED, wincom_freenamed);
}

/* com_issame(oObject1, oObject2) --> 1 if the two values are the same COM object (COM identity, IUnknown) */
RING_FUNC(ring_com_issame) {
	IUnknown *aObjects[2], *aIdentity[2];
	List *pList;
	int x, lSame;
	if (RING_API_PARACOUNT != 2) {
		RING_API_ERROR(RING_API_MISS2PARA);
		return;
	}
	for (x = 0; x < 2; x++) {
		aObjects[x] = NULL;
		if (RING_API_ISLIST(x + 1)) {
			pList = RING_API_GETLIST(x + 1);
			aObjects[x] = (IUnknown *)wincom_getlistpointer(pList, WINCOM_OBJECT);
			if (aObjects[x] == NULL) {
				aObjects[x] = (IUnknown *)wincom_getlistpointer(pList, WINCOM_UNKNOWN);
			}
		}
		if (aObjects[x] == NULL) {
			RING_API_ERROR("Error, WinCOM: The parameters must be COM objects");
			return;
		}
	}
	lSame = 0;
	if (SUCCEEDED(IUnknown_QueryInterface(aObjects[0], &IID_IUnknown, (void **)&aIdentity[0]))) {
		if (SUCCEEDED(IUnknown_QueryInterface(aObjects[1], &IID_IUnknown, (void **)&aIdentity[1]))) {
			lSame = aIdentity[0] == aIdentity[1];
			IUnknown_Release(aIdentity[1]);
		}
		IUnknown_Release(aIdentity[0]);
	}
	RING_API_RETNUMBER(lSame);
}

/* COM Events (Connection Points) */

/* Event handler: the event name ("*" for all the events) and the Ring code */
typedef struct WinCOMHandler {
	char *cName;
	char *cCode;
} WinCOMHandler;

/* The event sink: a COM object (IDispatch) that receives the events and runs the Ring code */
typedef struct WinCOMSink {
	const IDispatchVtbl *lpVtbl;
	LONG nRefCount;
	IID vIID;
	ITypeInfo *pTypeInfo;
	IConnectionPoint *pPoint;
	DWORD nCookie;
	VM *pVM;
	WinCOMHandler *aHandlers;
	int nHandlers;
} WinCOMSink;

/* The current event (Events can be nested, so each event keeps the previous one) */
typedef struct WinCOMEvent {
	const char *cName;
	DISPPARAMS *pParams;
	VARIANT *pResult;
	struct WinCOMEvent *pPrev;
} WinCOMEvent;

static WINCOM_TLS WinCOMEvent *pWinCOMEvent = NULL;

static void wincom_freehandlers(WinCOMSink *pSink) {
	int x;
	for (x = 0; x < pSink->nHandlers; x++) {
		free(pSink->aHandlers[x].cName);
		free(pSink->aHandlers[x].cCode);
	}
	free(pSink->aHandlers);
	pSink->aHandlers = NULL;
	pSink->nHandlers = 0;
}

static HRESULT STDMETHODCALLTYPE wincom_sink_queryinterface(IDispatch *pThis, REFIID riid, void **ppObject) {
	WinCOMSink *pSink;
	pSink = (WinCOMSink *)pThis;
	if (ppObject == NULL) {
		return E_POINTER;
	}
	if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch) || IsEqualIID(riid, &pSink->vIID)) {
		*ppObject = pThis;
		InterlockedIncrement(&pSink->nRefCount);
		return S_OK;
	}
	*ppObject = NULL;
	return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE wincom_sink_addref(IDispatch *pThis) {
	return (ULONG)InterlockedIncrement(&((WinCOMSink *)pThis)->nRefCount);
}

static ULONG STDMETHODCALLTYPE wincom_sink_release(IDispatch *pThis) {
	WinCOMSink *pSink;
	LONG nCount;
	pSink = (WinCOMSink *)pThis;
	nCount = InterlockedDecrement(&pSink->nRefCount);
	if (nCount == 0) {
		wincom_freehandlers(pSink);
		if (pSink->pTypeInfo != NULL) {
			ITypeInfo_Release(pSink->pTypeInfo);
		}
		if (pSink->pPoint != NULL) {
			IConnectionPoint_Release(pSink->pPoint);
		}
		free(pSink);
	}
	return (ULONG)nCount;
}

static HRESULT STDMETHODCALLTYPE wincom_sink_gettypeinfocount(IDispatch *pThis, UINT *pCount) {
	(void)pThis;
	if (pCount == NULL) {
		return E_POINTER;
	}
	*pCount = 0;
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE wincom_sink_gettypeinfo(IDispatch *pThis, UINT nInfo, LCID nLCID,
							  ITypeInfo **ppTypeInfo) {
	(void)pThis;
	(void)nInfo;
	(void)nLCID;
	if (ppTypeInfo != NULL) {
		*ppTypeInfo = NULL;
	}
	return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE wincom_sink_getidsofnames(IDispatch *pThis, REFIID riid, LPOLESTR *aNames,
							    UINT nNames, LCID nLCID, DISPID *aIDs) {
	(void)pThis;
	(void)riid;
	(void)aNames;
	(void)nNames;
	(void)nLCID;
	(void)aIDs;
	return E_NOTIMPL;
}

/* The name of the event (from the type information), "#DISPID" if not available */
static char *wincom_eventname(WinCOMSink *pSink, DISPID nDispID) {
	BSTR cName;
	UINT nNames;
	char *cOut;
	int nSize;
	cName = NULL;
	nNames = 0;
	if ((pSink->pTypeInfo != NULL) && SUCCEEDED(ITypeInfo_GetNames(pSink->pTypeInfo, nDispID, &cName, 1, &nNames)) &&
	    (nNames > 0) && (cName != NULL)) {
		cOut = wincom_toutf8(cName, (int)SysStringLen(cName), &nSize);
		SysFreeString(cName);
		return cOut;
	}
	cOut = (char *)malloc(32);
	if (cOut != NULL) {
		sprintf(cOut, "#%ld", (long)nDispID);
	}
	return cOut;
}

static HRESULT STDMETHODCALLTYPE wincom_sink_invoke(IDispatch *pThis, DISPID nDispID, REFIID riid, LCID nLCID,
						     WORD nFlags, DISPPARAMS *pParams, VARIANT *pResult,
						     EXCEPINFO *pExcepInfo, UINT *pArgErr) {
	WinCOMSink *pSink;
	WinCOMEvent vEvent;
	IErrorInfo *pErrorInfo;
	char *cName, *cCode, *cCodeCopy;
	size_t nSize;
	int x;
	(void)riid;
	(void)nLCID;
	(void)nFlags;
	(void)pExcepInfo;
	(void)pArgErr;
	pSink = (WinCOMSink *)pThis;
	if ((pSink->nHandlers == 0) || (pSink->pVM == NULL)) {
		return S_OK;
	}
	cName = wincom_eventname(pSink, nDispID);
	if (cName == NULL) {
		return E_OUTOFMEMORY;
	}
	/* Find the handler (The event name, or "*" for all the events) */
	cCode = NULL;
	for (x = 0; x < pSink->nHandlers; x++) {
		if (_stricmp(pSink->aHandlers[x].cName, cName) == 0) {
			cCode = pSink->aHandlers[x].cCode;
			break;
		}
	}
	if (cCode == NULL) {
		for (x = 0; x < pSink->nHandlers; x++) {
			if (strcmp(pSink->aHandlers[x].cName, "*") == 0) {
				cCode = pSink->aHandlers[x].cCode;
				break;
			}
		}
	}
	if (cCode == NULL) {
		free(cName);
		return S_OK;
	}
	/*
	**  The handler can change or remove the handlers, so we use a copy of the code
	**  The code runs inside try/catch: an error in the handler is raised again by the function that caused the event
	*/
	nSize = strlen(cCode) + 128;
	cCodeCopy = (char *)malloc(nSize);
	if (cCodeCopy == NULL) {
		free(cName);
		return E_OUTOFMEMORY;
	}
	snprintf(cCodeCopy, nSize, "try\n%s\ncatch\nwincom_seteventerror(cCatchError)\ndone\n", cCode);
	cCode = cCodeCopy;
	vEvent.cName = cName;
	vEvent.pParams = pParams;
	vEvent.pResult = pResult;
	vEvent.pPrev = pWinCOMEvent;
	pWinCOMEvent = &vEvent;
	/* The handler can disconnect the events (and the Ring GC can delete the C pointer) */
	IDispatch_AddRef(pThis);
	/*
	**  The object that raised the event can have error information for the current call
	**  The handler can call members (they clear it), so we keep it and set it again after the handler
	*/
	pErrorInfo = NULL;
	if (GetErrorInfo(0, &pErrorInfo) != S_OK) {
		pErrorInfo = NULL;
	}
	ring_vm_runcode(pSink->pVM, cCode);
	SetErrorInfo(0, pErrorInfo);
	if (pErrorInfo != NULL) {
		IErrorInfo_Release(pErrorInfo);
	}
	pWinCOMEvent = vEvent.pPrev;
	IDispatch_Release(pThis);
	free(cCode);
	free(cName);
	return S_OK;
}

static const IDispatchVtbl vWinCOMSinkVtbl = {
    wincom_sink_queryinterface, wincom_sink_addref,	     wincom_sink_release, wincom_sink_gettypeinfocount,
    wincom_sink_gettypeinfo,	wincom_sink_getidsofnames, wincom_sink_invoke};

/* Disconnect the sink from the object */
static void wincom_disconnect(WinCOMSink *pSink) {
	IConnectionPoint *pPoint;
	if (pSink->pPoint != NULL) {
		pPoint = pSink->pPoint;
		pSink->pPoint = NULL;
		IConnectionPoint_Unadvise(pPoint, pSink->nCookie);
		IConnectionPoint_Release(pPoint);
	}
	wincom_freehandlers(pSink);
}

static void wincom_freeevents(void *pState, void *pPtr) {
	WinCOMSink *pSink;
	(void)pState;
	pSink = (WinCOMSink *)pPtr;
	if (pSink != NULL) {
		pSink->pVM = NULL;
		wincom_disconnect(pSink);
		IDispatch_Release((IDispatch *)pSink);
	}
}

/* The type library that contains the type information of the object, NULL if not available */
static ITypeLib *wincom_gettypelib(IDispatch *pDispatch) {
	ITypeInfo *pTypeInfo;
	ITypeLib *pTypeLib;
	UINT nIndex;
	pTypeInfo = wincom_gettypeinfo(pDispatch);
	if (pTypeInfo == NULL) {
		return NULL;
	}
	pTypeLib = NULL;
	if (FAILED(ITypeInfo_GetContainingTypeLib(pTypeInfo, &pTypeLib, &nIndex))) {
		pTypeLib = NULL;
	}
	ITypeInfo_Release(pTypeInfo);
	return pTypeLib;
}

/* The IID of the default events interface of the object */
static HRESULT wincom_defaultsource(IDispatch *pDispatch, IConnectionPointContainer *pContainer, IID *pIID) {
	IProvideClassInfo2 *pInfo2;
	IProvideClassInfo *pInfo;
	ITypeInfo *pClassInfo, *pSourceInfo;
	TYPEATTR *pAttr, *pSourceAttr;
	HREFTYPE nRefType;
	IEnumConnectionPoints *pEnum;
	IConnectionPoint *pPoint;
	ULONG nFetched;
	INT nFlags;
	UINT x;
	HRESULT hr;
	/* IProvideClassInfo2 */
	if (SUCCEEDED(IDispatch_QueryInterface(pDispatch, &IID_IProvideClassInfo2, (void **)&pInfo2))) {
		hr = IProvideClassInfo2_GetGUID(pInfo2, GUIDKIND_DEFAULT_SOURCE_DISP_IID, pIID);
		IProvideClassInfo2_Release(pInfo2);
		if (SUCCEEDED(hr)) {
			return S_OK;
		}
	}
	/* The default source interface of the class (IProvideClassInfo) */
	if (SUCCEEDED(IDispatch_QueryInterface(pDispatch, &IID_IProvideClassInfo, (void **)&pInfo))) {
		hr = IProvideClassInfo_GetClassInfo(pInfo, &pClassInfo);
		IProvideClassInfo_Release(pInfo);
		if (SUCCEEDED(hr)) {
			hr = E_FAIL;
			if (SUCCEEDED(ITypeInfo_GetTypeAttr(pClassInfo, &pAttr))) {
				for (x = 0; x < pAttr->cImplTypes; x++) {
					if (FAILED(ITypeInfo_GetImplTypeFlags(pClassInfo, x, &nFlags)) ||
					    ((nFlags & (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)) !=
					     (IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE))) {
						continue;
					}
					if (SUCCEEDED(ITypeInfo_GetRefTypeOfImplType(pClassInfo, x, &nRefType)) &&
					    SUCCEEDED(ITypeInfo_GetRefTypeInfo(pClassInfo, nRefType, &pSourceInfo))) {
						if (SUCCEEDED(ITypeInfo_GetTypeAttr(pSourceInfo, &pSourceAttr))) {
							*pIID = pSourceAttr->guid;
							ITypeInfo_ReleaseTypeAttr(pSourceInfo, pSourceAttr);
							hr = S_OK;
						}
						ITypeInfo_Release(pSourceInfo);
					}
					break;
				}
				ITypeInfo_ReleaseTypeAttr(pClassInfo, pAttr);
			}
			ITypeInfo_Release(pClassInfo);
			if (SUCCEEDED(hr)) {
				return S_OK;
			}
		}
	}
	/* The first connection point */
	hr = IConnectionPointContainer_EnumConnectionPoints(pContainer, &pEnum);
	if (FAILED(hr)) {
		return hr;
	}
	pPoint = NULL;
	nFetched = 0;
	hr = IEnumConnectionPoints_Next(pEnum, 1, &pPoint, &nFetched);
	IEnumConnectionPoints_Release(pEnum);
	if ((hr != S_OK) || (nFetched == 0) || (pPoint == NULL)) {
		return CONNECT_E_NOCONNECTION;
	}
	hr = IConnectionPoint_GetConnectionInterface(pPoint, pIID);
	IConnectionPoint_Release(pPoint);
	return hr;
}

/* Find the IID of an interface using its name (like "ConnectionEvents") in the type library of the object */
static HRESULT wincom_findinterface(IDispatch *pDispatch, const char *cName, int nSize, IID *pIID) {
	ITypeLib *pTypeLib;
	ITypeInfo *pTypeInfo;
	TYPEATTR *pAttr;
	BSTR cWide, cTypeName;
	UINT x, nCount;
	HRESULT hr;
	cWide = wincom_tobstr(cName, nSize);
	if (cWide == NULL) {
		return E_OUTOFMEMORY;
	}
	if (cWide[0] == L'{') {
		hr = IIDFromString(cWide, pIID);
		SysFreeString(cWide);
		return hr;
	}
	hr = TYPE_E_ELEMENTNOTFOUND;
	pTypeLib = wincom_gettypelib(pDispatch);
	if (pTypeLib != NULL) {
		nCount = ITypeLib_GetTypeInfoCount(pTypeLib);
		for (x = 0; x < nCount; x++) {
			cTypeName = NULL;
			if (FAILED(ITypeLib_GetDocumentation(pTypeLib, (INT)x, &cTypeName, NULL, NULL, NULL)) ||
			    (cTypeName == NULL)) {
				continue;
			}
			if ((_wcsicmp(cTypeName, cWide) == 0) && SUCCEEDED(ITypeLib_GetTypeInfo(pTypeLib, x, &pTypeInfo))) {
				if (SUCCEEDED(ITypeInfo_GetTypeAttr(pTypeInfo, &pAttr))) {
					*pIID = pAttr->guid;
					ITypeInfo_ReleaseTypeAttr(pTypeInfo, pAttr);
					hr = S_OK;
				}
				ITypeInfo_Release(pTypeInfo);
			}
			SysFreeString(cTypeName);
			if (SUCCEEDED(hr)) {
				break;
			}
		}
		ITypeLib_Release(pTypeLib);
	}
	SysFreeString(cWide);
	return hr;
}

/*
**  com_connectevents(oObject, [cInterface]) --> Events object (C pointer)
**  cInterface is the name of the events interface (like "ConnectionEvents") or its IID ("{...}")
**  The default events interface of the object is used if cInterface is not passed
*/
RING_FUNC(ring_com_connectevents) {
	IDispatch *pDispatch;
	IConnectionPointContainer *pContainer;
	IConnectionPoint *pPoint;
	ITypeLib *pTypeLib;
	WinCOMSink *pSink;
	IID vIID;
	HRESULT hr;
	const char *cMember;
	if ((RING_API_PARACOUNT < 1) || (RING_API_PARACOUNT > 2)) {
		RING_API_ERROR(WINCOM_ERROR_PARACOUNT);
		return;
	}
	pDispatch = wincom_getobject(pPointer, 1);
	if (pDispatch == NULL) {
		RING_API_ERROR(WINCOM_ERROR_NOTOBJECT);
		return;
	}
	if ((RING_API_PARACOUNT == 2) && !RING_API_ISSTRING(2)) {
		RING_API_ERROR("Error, WinCOM: The events interface must be a string");
		return;
	}
	cMember = "com_connectevents()";
	hr = IDispatch_QueryInterface(pDispatch, &IID_IConnectionPointContainer, (void **)&pContainer);
	if (FAILED(hr)) {
		wincom_seterror(cMember, hr, NULL);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	if ((RING_API_PARACOUNT == 2) && (RING_API_GETSTRINGSIZE(2) > 0)) {
		hr = wincom_findinterface(pDispatch, RING_API_GETSTRING(2), RING_API_GETSTRINGSIZE(2), &vIID);
		cMember = RING_API_GETSTRING(2);
	} else {
		hr = wincom_defaultsource(pDispatch, pContainer, &vIID);
	}
	pPoint = NULL;
	if (SUCCEEDED(hr)) {
		hr = IConnectionPointContainer_FindConnectionPoint(pContainer, &vIID, &pPoint);
	}
	IConnectionPointContainer_Release(pContainer);
	if (FAILED(hr)) {
		wincom_seterror(cMember, hr, NULL);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	pSink = (WinCOMSink *)calloc(1, sizeof(WinCOMSink));
	if (pSink == NULL) {
		IConnectionPoint_Release(pPoint);
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	pSink->lpVtbl = &vWinCOMSinkVtbl;
	pSink->nRefCount = 1;
	pSink->vIID = vIID;
	pSink->pVM = (VM *)pPointer;
	/* The type information of the events interface (for the event names) */
	pTypeLib = wincom_gettypelib(pDispatch);
	if (pTypeLib != NULL) {
		if (FAILED(ITypeLib_GetTypeInfoOfGuid(pTypeLib, &vIID, &pSink->pTypeInfo))) {
			pSink->pTypeInfo = NULL;
		}
		ITypeLib_Release(pTypeLib);
	}
	hr = IConnectionPoint_Advise(pPoint, (IUnknown *)pSink, &pSink->nCookie);
	if (FAILED(hr)) {
		IConnectionPoint_Release(pPoint);
		IDispatch_Release((IDispatch *)pSink);
		wincom_seterror(cMember, hr, NULL);
		RING_API_ERROR(cWinCOMError);
		return;
	}
	pSink->pPoint = pPoint;
	RING_API_RETMANAGEDCPOINTER(pSink, WINCOM_EVENTS, wincom_freeevents);
}

static WinCOMSink *wincom_getsink(void *pPointer, int nPara) {
	WinCOMSink *pSink;
	pSink = RING_API_ISLIST(nPara) ? (WinCOMSink *)wincom_getlistpointer(RING_API_GETLIST(nPara), WINCOM_EVENTS) : NULL;
	if (pSink == NULL) {
		RING_API_ERROR("Error, WinCOM: The first parameter must be created by com_connectevents()");
	}
	return pSink;
}

/* com_onevent(oEvents, cEvent, cCode) --> Run the Ring code when the event happens ("" removes the handler) */
RING_FUNC(ring_com_onevent) {
	WinCOMSink *pSink;
	WinCOMHandler *aHandlers;
	char *cName, *cCode;
	int x;
	if (RING_API_PARACOUNT != 3) {
		RING_API_ERROR(RING_API_MISS3PARA);
		return;
	}
	pSink = wincom_getsink(pPointer, 1);
	if (pSink == NULL) {
		return;
	}
	if (!RING_API_ISSTRING(2) || !RING_API_ISSTRING(3) || (RING_API_GETSTRINGSIZE(2) == 0)) {
		RING_API_ERROR(RING_API_BADPARATYPE);
		return;
	}
	if (pSink->pPoint == NULL) {
		RING_API_ERROR("Error, WinCOM: The events are disconnected");
		return;
	}
	/* Remove the current handler */
	for (x = 0; x < pSink->nHandlers; x++) {
		if (_stricmp(pSink->aHandlers[x].cName, RING_API_GETSTRING(2)) == 0) {
			free(pSink->aHandlers[x].cName);
			free(pSink->aHandlers[x].cCode);
			pSink->aHandlers[x] = pSink->aHandlers[pSink->nHandlers - 1];
			pSink->nHandlers--;
			break;
		}
	}
	if (RING_API_GETSTRINGSIZE(3) == 0) {
		return;
	}
	cName = wincom_strdup(RING_API_GETSTRING(2));
	cCode = wincom_strdup(RING_API_GETSTRING(3));
	aHandlers = (WinCOMHandler *)realloc(pSink->aHandlers, sizeof(WinCOMHandler) * ((size_t)pSink->nHandlers + 1));
	if ((cName == NULL) || (cCode == NULL) || (aHandlers == NULL)) {
		free(cName);
		free(cCode);
		if (aHandlers != NULL) {
			pSink->aHandlers = aHandlers;
		}
		RING_API_ERROR(WINCOM_ERROR_OOM);
		return;
	}
	pSink->aHandlers = aHandlers;
	pSink->aHandlers[pSink->nHandlers].cName = cName;
	pSink->aHandlers[pSink->nHandlers].cCode = cCode;
	pSink->nHandlers++;
}

/* com_disconnectevents(oEvents) --> Stop receiving the events */
RING_FUNC(ring_com_disconnectevents) {
	WinCOMSink *pSink;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pSink = wincom_getsink(pPointer, 1);
	if (pSink == NULL) {
		return;
	}
	wincom_disconnect(pSink);
}

/* com_eventnames(oEvents) --> List of the event names (from the type information) */
RING_FUNC(ring_com_eventnames) {
	WinCOMSink *pSink;
	TYPEATTR *pAttr;
	FUNCDESC *pFuncDesc;
	BSTR cName;
	List *pList;
	UINT x;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pSink = wincom_getsink(pPointer, 1);
	if (pSink == NULL) {
		return;
	}
	pList = RING_API_NEWLIST;
	if ((pSink->pTypeInfo != NULL) && SUCCEEDED(ITypeInfo_GetTypeAttr(pSink->pTypeInfo, &pAttr))) {
		for (x = 0; x < pAttr->cFuncs; x++) {
			if (FAILED(ITypeInfo_GetFuncDesc(pSink->pTypeInfo, x, &pFuncDesc))) {
				continue;
			}
			if (!(pFuncDesc->wFuncFlags & (FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN))) {
				cName = NULL;
				if (SUCCEEDED(ITypeInfo_GetDocumentation(pSink->pTypeInfo, pFuncDesc->memid, &cName, NULL,
									  NULL, NULL))) {
					wincom_addbstr(RING_API_STATE, pList, cName);
					SysFreeString(cName);
				}
			}
			ITypeInfo_ReleaseFuncDesc(pSink->pTypeInfo, pFuncDesc);
		}
		ITypeInfo_ReleaseTypeAttr(pSink->pTypeInfo, pAttr);
	}
	RING_API_RETLISTBYREF(pList);
}

static WinCOMEvent *wincom_getevent(void *pPointer) {
	if (pWinCOMEvent == NULL) {
		RING_API_ERROR("Error, WinCOM: There is no active event (Use this function in the event handler)");
	}
	return pWinCOMEvent;
}

/* wincom_seteventerror(cError) --> Keep the error of the event handler (Used by the code that runs the handler) */
RING_FUNC(ring_wincom_seteventerror) {
	if ((RING_API_PARACOUNT != 1) || !RING_API_ISSTRING(1)) {
		return;
	}
	if (nWinCOMCallDepth == 0) {
		/* The event didn't happen during a com_*() call (Like events in the event loop of a GUI library) */
		printf("\nError, WinCOM: Event %s : %s\n", pWinCOMEvent != NULL ? pWinCOMEvent->cName : "",
		       RING_API_GETSTRING(1));
		fflush(stdout);
		return;
	}
	if (lWinCOMEventError) {
		return;
	}
	cWinCOMEventError[0] = '\0';
	snprintf(cWinCOMEventError, WINCOM_ERRORSIZE, "Error, WinCOM: Event %s : %s",
		 pWinCOMEvent != NULL ? pWinCOMEvent->cName : "", RING_API_GETSTRING(1));
	lWinCOMEventError = 1;
}

/* com_eventname() --> The name of the current event */
RING_FUNC(ring_com_eventname) {
	WinCOMEvent *pEvent;
	pEvent = wincom_getevent(pPointer);
	if (pEvent == NULL) {
		return;
	}
	RING_API_RETSTRING(pEvent->cName);
}

/*
**  The argument nIndex (0 = the first argument) of the event, NULL if not passed
**  Positional arguments are stored in reverse order after the named arguments
**  Some programs (like Microsoft Office) pass the arguments as named arguments: the name is the position
*/
static VARIANT *wincom_eventarg(DISPPARAMS *pParams, UINT nIndex) {
	UINT x, nPositional;
	if ((pParams == NULL) || (pParams->rgvarg == NULL)) {
		return NULL;
	}
	if (pParams->rgdispidNamedArgs != NULL) {
		for (x = 0; x < pParams->cNamedArgs; x++) {
			if (pParams->rgdispidNamedArgs[x] == (DISPID)nIndex) {
				return &pParams->rgvarg[x];
			}
		}
	}
	nPositional = pParams->cArgs - pParams->cNamedArgs;
	if (nIndex < nPositional) {
		return &pParams->rgvarg[pParams->cArgs - 1 - nIndex];
	}
	return NULL;
}

/* com_eventargs() --> List of the arguments of the current event */
RING_FUNC(ring_com_eventargs) {
	WinCOMEvent *pEvent;
	VARIANT *pArg;
	List *pList;
	UINT x;
	pEvent = wincom_getevent(pPointer);
	if (pEvent == NULL) {
		return;
	}
	pArg = NULL;
	pList = RING_API_NEWLIST;
	if (pEvent->pParams != NULL) {
		for (x = 0; x < pEvent->pParams->cArgs; x++) {
			pArg = wincom_eventarg(pEvent->pParams, x);
			if (pArg != NULL) {
				wincom_addvariant(RING_API_STATE, pList, pArg, 0);
			} else {
				ring_list_addstring_gc(RING_API_STATE, pList, "");
			}
		}
	}
	RING_API_RETLISTBYREF(pList);
}

/* Write a value to a VT_BYREF argument (the type of the argument is not changed) */
static HRESULT wincom_writebyref(VARIANT *pTarget, VARIANT *pValue) {
	VARIANT vConverted;
	VARTYPE nType;
	HRESULT hr;
	nType = (VARTYPE)(V_VT(pTarget) & ~VT_BYREF);
	if (!(V_VT(pTarget) & VT_BYREF) || (V_BYREF(pTarget) == NULL)) {
		return DISP_E_TYPEMISMATCH;
	}
	if (nType == VT_VARIANT) {
		hr = VariantClear(V_VARIANTREF(pTarget));
		if (SUCCEEDED(hr)) {
			hr = VariantCopy(V_VARIANTREF(pTarget), pValue);
		}
		return hr;
	}
	VariantInit(&vConverted);
	hr = VariantChangeType(&vConverted, pValue, 0, nType);
	if (FAILED(hr)) {
		return hr;
	}
	hr = S_OK;
	switch (nType) {
	case VT_I1:
		*V_I1REF(pTarget) = V_I1(&vConverted);
		break;
	case VT_UI1:
		*V_UI1REF(pTarget) = V_UI1(&vConverted);
		break;
	case VT_I2:
		*V_I2REF(pTarget) = V_I2(&vConverted);
		break;
	case VT_UI2:
		*V_UI2REF(pTarget) = V_UI2(&vConverted);
		break;
	case VT_I4:
		*V_I4REF(pTarget) = V_I4(&vConverted);
		break;
	case VT_UI4:
		*V_UI4REF(pTarget) = V_UI4(&vConverted);
		break;
	case VT_INT:
		*V_INTREF(pTarget) = V_INT(&vConverted);
		break;
	case VT_UINT:
		*V_UINTREF(pTarget) = V_UINT(&vConverted);
		break;
	case VT_I8:
		*V_I8REF(pTarget) = V_I8(&vConverted);
		break;
	case VT_UI8:
		*V_UI8REF(pTarget) = V_UI8(&vConverted);
		break;
	case VT_R4:
		*V_R4REF(pTarget) = V_R4(&vConverted);
		break;
	case VT_R8:
		*V_R8REF(pTarget) = V_R8(&vConverted);
		break;
	case VT_CY:
		*V_CYREF(pTarget) = V_CY(&vConverted);
		break;
	case VT_DATE:
		*V_DATEREF(pTarget) = V_DATE(&vConverted);
		break;
	case VT_BOOL:
		*V_BOOLREF(pTarget) = V_BOOL(&vConverted);
		break;
	case VT_ERROR:
		*V_ERRORREF(pTarget) = V_ERROR(&vConverted);
		break;
	case VT_BSTR:
		/* Take the string */
		SysFreeString(*V_BSTRREF(pTarget));
		*V_BSTRREF(pTarget) = V_BSTR(&vConverted);
		V_VT(&vConverted) = VT_EMPTY;
		break;
	case VT_DISPATCH:
		if (*V_DISPATCHREF(pTarget) != NULL) {
			IDispatch_Release(*V_DISPATCHREF(pTarget));
		}
		*V_DISPATCHREF(pTarget) = V_DISPATCH(&vConverted);
		V_VT(&vConverted) = VT_EMPTY;
		break;
	case VT_UNKNOWN:
		if (*V_UNKNOWNREF(pTarget) != NULL) {
			IUnknown_Release(*V_UNKNOWNREF(pTarget));
		}
		*V_UNKNOWNREF(pTarget) = V_UNKNOWN(&vConverted);
		V_VT(&vConverted) = VT_EMPTY;
		break;
	default:
		hr = DISP_E_TYPEMISMATCH;
	}
	VariantClear(&vConverted);
	return hr;
}

/* com_seteventarg(nIndex, vValue) --> Change a ByRef argument of the current event (like adStatus in ADO events) */
RING_FUNC(ring_com_seteventarg) {
	WinCOMEvent *pEvent;
	VARIANT vValue;
	VARIANT *pArg;
	double nIndex;
	HRESULT hr;
	if (RING_API_PARACOUNT != 2) {
		RING_API_ERROR(RING_API_MISS2PARA);
		return;
	}
	pEvent = wincom_getevent(pPointer);
	if (pEvent == NULL) {
		return;
	}
	if (!RING_API_ISNUMBER(1)) {
		RING_API_ERROR(RING_API_BADPARATYPE);
		return;
	}
	nIndex = RING_API_GETNUMBER(1);
	pArg = NULL;
	if ((pEvent->pParams != NULL) && (nIndex >= 1) && (nIndex <= (double)pEvent->pParams->cArgs)) {
		pArg = wincom_eventarg(pEvent->pParams, (UINT)nIndex - 1);
	}
	if (pArg == NULL) {
		RING_API_ERROR("Error, WinCOM: Bad event argument index");
		return;
	}
	if (!wincom_paratovariant(pPointer, 2, &vValue)) {
		RING_API_ERROR(cWinCOMError);
		return;
	}
	hr = wincom_writebyref(pArg, &vValue);
	VariantClear(&vValue);
	if (FAILED(hr)) {
		wincom_seterror("com_seteventarg()", hr, NULL);
		wincom_appenderrora(" (The argument must be passed by reference)");
		RING_API_ERROR(cWinCOMError);
	}
}

/* com_seteventresult(vValue) --> The return value of the current event (few events use it) */
RING_FUNC(ring_com_seteventresult) {
	WinCOMEvent *pEvent;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	pEvent = wincom_getevent(pPointer);
	if (pEvent == NULL) {
		return;
	}
	if (pEvent->pResult == NULL) {
		RING_API_ERROR("Error, WinCOM: The event doesn't have a return value");
		return;
	}
	VariantClear(pEvent->pResult);
	if (!wincom_paratovariant(pPointer, 1, pEvent->pResult)) {
		RING_API_ERROR(cWinCOMError);
	}
}

/*
**  com_processmessages([nMilliseconds]) --> Process the Windows messages (Events from other programs like Excel)
**  Without a parameter (or 0) the waiting messages are processed, else the function waits nMilliseconds
**  Returns the number of processed messages
*/
RING_FUNC(ring_com_processmessages) {
	MSG vMessage;
	DWORD nStart, nElapsed, nWait;
	double nCount, nMilliseconds;
	if (RING_API_PARACOUNT > 1) {
		RING_API_ERROR(WINCOM_ERROR_PARACOUNT);
		return;
	}
	nMilliseconds = 0;
	if (RING_API_PARACOUNT == 1) {
		if (!RING_API_ISNUMBER(1)) {
			RING_API_ERROR(RING_API_BADPARATYPE);
			return;
		}
		nMilliseconds = RING_API_GETNUMBER(1);
		if (nMilliseconds < 0) {
			nMilliseconds = 0;
		}
		if (nMilliseconds > 86400000.0) {
			nMilliseconds = 86400000.0;
		}
	}
	nWait = (DWORD)nMilliseconds;
	nCount = 0;
	nStart = GetTickCount();
	wincom_begincall();
	while (1) {
		while (PeekMessageW(&vMessage, NULL, 0, 0, PM_REMOVE)) {
			if (vMessage.message == WM_QUIT) {
				/* Keep WM_QUIT for the main message loop */
				PostQuitMessage((int)vMessage.wParam);
				if (wincom_endcall(pPointer)) {
					return;
				}
				RING_API_RETNUMBER(nCount);
				return;
			}
			TranslateMessage(&vMessage);
			DispatchMessageW(&vMessage);
			nCount++;
		}
		nElapsed = GetTickCount() - nStart;
		if (nElapsed >= nWait) {
			break;
		}
		MsgWaitForMultipleObjectsEx(0, NULL, nWait - nElapsed, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
	}
	if (wincom_endcall(pPointer)) {
		return;
	}
	RING_API_RETNUMBER(nCount);
}

/* Add the constants (enumerations and modules) of a type library: [cGroup, cName, vValue] */
static void wincom_addconstants(void *pState, List *pList, ITypeLib *pTypeLib) {
	ITypeInfo *pTypeInfo;
	TYPEATTR *pAttr;
	VARDESC *pVarDesc;
	TYPEKIND nKind;
	BSTR cGroup, cName;
	List *pItem;
	UINT x, y, nCount;
	nCount = ITypeLib_GetTypeInfoCount(pTypeLib);
	for (x = 0; x < nCount; x++) {
		if (FAILED(ITypeLib_GetTypeInfoType(pTypeLib, x, &nKind)) ||
		    ((nKind != TKIND_ENUM) && (nKind != TKIND_MODULE))) {
			continue;
		}
		if (FAILED(ITypeLib_GetTypeInfo(pTypeLib, x, &pTypeInfo))) {
			continue;
		}
		cGroup = NULL;
		ITypeLib_GetDocumentation(pTypeLib, (INT)x, &cGroup, NULL, NULL, NULL);
		if (SUCCEEDED(ITypeInfo_GetTypeAttr(pTypeInfo, &pAttr))) {
			for (y = 0; y < pAttr->cVars; y++) {
				if (FAILED(ITypeInfo_GetVarDesc(pTypeInfo, y, &pVarDesc))) {
					continue;
				}
				if ((pVarDesc->varkind == VAR_CONST) && (pVarDesc->lpvarValue != NULL)) {
					cName = NULL;
					if (SUCCEEDED(ITypeInfo_GetDocumentation(pTypeInfo, pVarDesc->memid, &cName, NULL, NULL,
										  NULL))) {
						pItem = ring_list_newlist_gc(pState, pList);
						wincom_addbstr(pState, pItem, cGroup);
						wincom_addbstr(pState, pItem, cName);
						wincom_addvariant(pState, pItem, pVarDesc->lpvarValue, 0);
						SysFreeString(cName);
					}
				}
				ITypeInfo_ReleaseVarDesc(pTypeInfo, pVarDesc);
			}
			ITypeInfo_ReleaseTypeAttr(pTypeInfo, pAttr);
		}
		if (cGroup != NULL) {
			SysFreeString(cGroup);
		}
		ITypeInfo_Release(pTypeInfo);
	}
}

/*
**  com_constants(vSource) --> List of [cEnumeration, cName, vValue]
**  vSource is a COM object (its type library is used) or a type library file name (like "msado15.dll")
*/
RING_FUNC(ring_com_constants) {
	IDispatch *pDispatch;
	ITypeLib *pTypeLib;
	BSTR cFile;
	List *pList;
	HRESULT hr;
	if (RING_API_PARACOUNT != 1) {
		RING_API_ERROR(RING_API_MISS1PARA);
		return;
	}
	if (!wincom_init()) {
		RING_API_ERROR(WINCOM_ERROR_INIT);
		return;
	}
	pTypeLib = NULL;
	if (RING_API_ISSTRING(1)) {
		cFile = wincom_tobstr(RING_API_GETSTRING(1), RING_API_GETSTRINGSIZE(1));
		if (cFile == NULL) {
			RING_API_ERROR(WINCOM_ERROR_OOM);
			return;
		}
		hr = LoadTypeLibEx(cFile, REGKIND_NONE, &pTypeLib);
		SysFreeString(cFile);
		if (FAILED(hr)) {
			wincom_seterror(RING_API_GETSTRING(1), hr, NULL);
			RING_API_ERROR(cWinCOMError);
			return;
		}
	} else {
		pDispatch = wincom_getobject(pPointer, 1);
		if (pDispatch == NULL) {
			RING_API_ERROR("Error, WinCOM: The parameter must be a COM object or a type library file name");
			return;
		}
		pTypeLib = wincom_gettypelib(pDispatch);
		if (pTypeLib == NULL) {
			RING_API_ERROR("Error, WinCOM: The object doesn't have a type library");
			return;
		}
	}
	pList = RING_API_NEWLIST;
	wincom_addconstants(RING_API_STATE, pList, pTypeLib);
	ITypeLib_Release(pTypeLib);
	RING_API_RETLISTBYREF(pList);
}

/* Register the functions */

RING_LIBINIT {
	RING_API_REGISTER("com_init", ring_com_init);
	RING_API_REGISTER("com_createobject", ring_com_createobject);
	RING_API_REGISTER("com_getobject", ring_com_getobject);
	RING_API_REGISTER("com_getactiveobject", ring_com_getactiveobject);
	RING_API_REGISTER("com_progidexists", ring_com_progidexists);
	RING_API_REGISTER("com_get", ring_com_get);
	RING_API_REGISTER("com_invoke", ring_com_invoke);
	RING_API_REGISTER("com_set", ring_com_set);
	RING_API_REGISTER("com_setref", ring_com_setref);
	RING_API_REGISTER("com_dispatch", ring_com_dispatch);
	RING_API_REGISTER("com_items", ring_com_items);
	RING_API_REGISTER("com_isobject", ring_com_isobject);
	RING_API_REGISTER("com_release", ring_com_release);
	RING_API_REGISTER("com_variant", ring_com_variant);
	RING_API_REGISTER("com_lasttype", ring_com_lasttype);
	RING_API_REGISTER("com_lasthresult", ring_com_lasthresult);
	RING_API_REGISTER("com_typename", ring_com_typename);
	RING_API_REGISTER("com_members", ring_com_members);
	RING_API_REGISTER("com_refcount", ring_com_refcount);
	RING_API_REGISTER("com_byref", ring_com_byref);
	RING_API_REGISTER("com_byrefvalue", ring_com_byrefvalue);
	RING_API_REGISTER("com_named", ring_com_named);
	RING_API_REGISTER("com_issame", ring_com_issame);
	RING_API_REGISTER("com_connectevents", ring_com_connectevents);
	RING_API_REGISTER("com_onevent", ring_com_onevent);
	RING_API_REGISTER("com_disconnectevents", ring_com_disconnectevents);
	RING_API_REGISTER("com_eventnames", ring_com_eventnames);
	RING_API_REGISTER("com_eventname", ring_com_eventname);
	RING_API_REGISTER("com_eventargs", ring_com_eventargs);
	RING_API_REGISTER("com_seteventarg", ring_com_seteventarg);
	RING_API_REGISTER("com_seteventresult", ring_com_seteventresult);
	RING_API_REGISTER("com_processmessages", ring_com_processmessages);
	RING_API_REGISTER("com_matrix", ring_com_matrix);
	RING_API_REGISTER("com_constants", ring_com_constants);
	RING_API_REGISTER("wincom_seteventerror", ring_wincom_seteventerror);
}

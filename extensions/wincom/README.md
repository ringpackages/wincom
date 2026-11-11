# WinCOM Extension

## Introduction

This extension provides COM (Component Object Model) Automation for the Ring programming language on MS-Windows.

The extension uses late binding (the `IDispatch` interface, like VBScript and Visual Basic): it gets the member ID using `GetIDsOfNames()`, converts the Ring values to `VARIANT` values, calls `Invoke()`, then converts the result back to a Ring value. COM objects are returned as managed C pointers (type `COMObject`): the Ring garbage collector calls `Release()` when the last copy of the variable is deleted.

The extension is self-contained: it uses the Ring extension API and the COM API (`ole32.dll`, `oleaut32.dll`) only.

Programs load the extension (and the Ring code of the package) using `wincom.ring`:

```ring
load "wincom.ring"

oDict = com_createobject("Scripting.Dictionary")
com_invoke(oDict, "Add", "key", "value")
? com_get(oDict, 'Item("key")')
```

## Files

| File | Description |
|------|-------------|
| `ring_wincom.c` | The extension source code (C + COM API), written by hand: the functions and their registration (`RING_LIBINIT`) |
| `wincom.ring` | Loads the extension library, `globals.ring`, `functions.ring`, `comobject.ring` and `comevents.ring` (used by `load "wincom.ring"` through `ring/bin/load/wincom.ring`) |
| `globals.ring` | Global constants (`WINCOM_VERSION`, `COM_GET`, ..., `VT_*` types, ADO and ADOX constants) |
| `functions.ring` | Functions written in Ring (`com_null()`, `com_bool()`, `com_date()`, `com_wasnull()`, `com_rowsfromcolumns()`, `com_version()`, ...) |
| `comobject.ring` | The `COMObject` class |
| `comevents.ring` | The `COMEvents` class |
| `buildvc.bat`, `buildvc_x64.bat` | Build using Visual C/C++ (32-bit and 64-bit), output: `ring/bin/ring_wincom.dll` |

The extension doesn't use the Ring code generator (like `ringinternet` and `ringmurmurhash`), because the functions accept any number of parameters of any type.

To add a function: write it in `ring_wincom.c` using `RING_FUNC(ring_com_name)`, register it in `RING_LIBINIT` at the end of the file (`RING_API_REGISTER("com_name", ring_com_name);`), then build.

## Functions

Written in C: `com_init()`, `com_createobject()`, `com_getobject()`, `com_getactiveobject()`, `com_progidexists()`, `com_get()`, `com_set()`, `com_setref()`, `com_invoke()`, `com_dispatch()`, `com_items()`, `com_isobject()`, `com_release()`, `com_variant()`, `com_lasttype()`, `com_lasthresult()`, `com_typename()`, `com_members()`, `com_refcount()`, `com_byref()`, `com_byrefvalue()`, `com_named()`, `com_issame()`, `com_connectevents()`, `com_onevent()`, `com_disconnectevents()`, `com_eventnames()`, `com_eventname()`, `com_eventargs()`, `com_seteventarg()`, `com_seteventresult()`, `com_processmessages()`, `com_matrix()`, `com_constants()`

Internal: `wincom_seteventerror()` (used by the code that runs the event handlers)

Written in Ring: `com_null()`, `com_empty()`, `com_missing()`, `com_bool()`, `com_date()`, `com_currency()`, `com_double()`, `com_bytes()`, `com_wasnull()`, `com_rowsfromcolumns()`, `com_version()`

See the package README for the description of each function.

## Implementation Notes

- COM is initialized (`CoInitializeEx()`, Single-Threaded Apartment) by the first function that needs it. If the thread was initialized before using another mode (`RPC_E_CHANGED_MODE`), the current mode is used.
- Errors are raised after freeing all the resources of the call (`VARIANT` values, `BSTR` strings, objects), so catching them using `try/catch` doesn't leak memory.
- For property put, the value is passed as the named argument `DISPID_PROPERTYPUT`. `com_set()` tries `DISPATCH_PROPERTYPUTREF` if the value is an object and `DISPATCH_PROPERTYPUT` fails (like `Set` in Visual Basic).
- Arrays are returned using `RING_API_RETLISTBYREF()` (the list is not copied). COM objects inside lists are managed C pointers too.
- `com_release()` sets the shared pointer item to `NULL`, so the copies of the variable see the object as released and the garbage collector doesn't release it again.
- The state (last error, last type, current event) is thread local, because each thread has its own COM apartment.
- The member path is copied before the call: event handlers (Ring code) can run during the call and change the Ring variables.
- Events: the sink is a C object that implements `IDispatch` and is registered using `IConnectionPoint::Advise()`. The event name comes from the type information of the events interface. `Invoke()` runs the handler using `ring_vm_runcode()` inside `try/catch` (`wincom_seteventerror()` keeps the error, then the `com_*()` function that caused the event raises it). The functions that can cause events (calling members, `com_items()`, `com_release()`, `com_processmessages()`) count the active calls; an error from an event that happens outside them is printed, so it's never raised later by an unrelated call. The arguments are read by position: positional arguments are stored in reverse order, and Microsoft Office passes them as named arguments (the name is the position).
- Before each `Invoke()` the error information of the thread is cleared (`SetErrorInfo(0, NULL)`): some objects (like the Scripting Runtime) fill `EXCEPINFO` from it, so an old error (from another call) appeared in the new error message.
- `com_createobject(cProgID, cServer)` uses `CoCreateInstanceEx()` with `COSERVERINFO` (DCOM); the ProgID is converted to the CLSID using the registry of this computer.
- ByRef arguments are passed as `VT_BYREF | VT_VARIANT` that points to the `VARIANT` of the `com_byref()` holder. The holder is reference counted: each call that passes it keeps it alive until the call ends, even if an event handler uses `com_release()` during the call.
- `com_set()` uses `DISPATCH_PROPERTYPUTREF` again only when `DISPATCH_PROPERTYPUT` failed with `DISP_E_MEMBERNOTFOUND`, `DISP_E_BADPARAMCOUNT` or `DISP_E_TYPEMISMATCH` (also inside `EXCEPINFO`), so the errors raised by the property itself are not retried.

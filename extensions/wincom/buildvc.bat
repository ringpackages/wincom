cls
setlocal enableextensions enabledelayedexpansion
call ../../language/build/locatevc.bat auto %1
cl /c %ringcflags% ring_wincom.c -I"..\..\language\include"
link %ringldflags% ring_wincom.obj ..\..\lib\ring.lib ole32.lib oleaut32.lib uuid.lib user32.lib /DLL /OUT:..\..\bin\ring_wincom.dll /SUBSYSTEM:CONSOLE,"5.01"
del *.obj
endlocal

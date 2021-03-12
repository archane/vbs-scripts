Dim objShell
Set objShell = WScript.CreateObject ("WScript.shell")
set fso = createobject("scripting.filesystemobject")
strPath = "D:\Scripts\Serverlist.txt"
Set serverlist = fso.opentextfile(strPath)

Do Until serverlist.AtEndOfStream
	imported_text = serverlist.readline
	temp_array = split(imported_text,"|")
	servicename = temp_array(1)
	servername = temp_array(0)
	Set wmi = GetObject("winmgmts://"& servername &"/root/cimv2")
	state = wmi.Get("Win32_Service.Name='" & servicename & "'").State
	'WScript.Echo servername & ": " & servicename & ": " & state
	WScript.Sleep 1000

loop

objShell.run "cmd /c iisreset pi-onbaseunty01"
WScript.Sleep 1000
Set objShell = Nothing
objShell.run "cmd /c iisreset pi-onbaseunty02"
WScript.Sleep 1000
Set objShell = Nothing

------pipe file like this-----
pi-onbaseimpt01|OnBase Pol
pi-onbaseimpt02|OnBase Pol
pi-onbaseimpt03|OnBase Pol
pi-onbaseimpt03|OnBase Pol1
pi-onbaseimpt03|OnBase DDS Service
pi-onbaseimpt04|OnBase Pol
pi-onbaseimpt05|OnBase Pol
pi-onbaseimpt07|OnBase Pol
pi-onbaseunty02|Hyland.Core.Distribution.NTService
pi-onbasewf01|Hyland Unity Scheduler_UnityScheduler
pi-onbasewf01|Hyland.Core.Workflow.NTService
pi-onbasewf01|Hyland.Core.Timers.NTService

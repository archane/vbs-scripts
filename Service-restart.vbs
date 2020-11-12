set fso = createobject("scripting.filesystemobject")
'strPath = "D:\Scripts\service-list.txt"
strPath = "D:\Scripts\service-list-test.txt"
Set serverlist = fso.opentextfile(strPath)

Do Until serverlist.AtEndOfStream
	imported_text = serverlist.readline
	temp_array = split(imported_text,"|")
	servicename = temp_array(1)
	servername = temp_array(0)
curr_time=now
dim log
log = curr_time
logname = "d:\Scripts\logs\Distribution-service.txt"
Set wmi = GetObject("winmgmts://"& servername &"/root/cimv2")
state = wmi.Get("Win32_Service.Name='" & serviceName & "'").State

if state = "Running" then
	CreateObject("Shell.Application").serviceStop servicename, True
	log = log & vbCrlf & "Stopping Service: " & servicename & " On Server:" & servername
else
	log = log & vbCrlf & "Service was already stopped"
end if

WScript.Sleep 3000
state = ""
state = wmi.Get("Win32_Service.Name='" & serviceName & "'").State

if state = "Running" then
	log = log & vbCrlf & "Service did NOT stop as expected, restart was not performed"
else
	CreateObject("Shell.Application").serviceStart servicename, True
	log = log & vbCrlf & "Starting Service: " & servicename
	WScript.Sleep 3000
	state = ""
	state = wmi.Get("Win32_Service.Name='" & serviceName & "'").State
	if state = "Running" then
		log = log & vbCrlf & "Service was successfully started: " & servicename & " On Server:" & servername
	else
		log = log & vbCrlf & "Service did not Start correctly or in a timely manner"
	end if
end if

set fso = createobject("scripting.filesystemobject")
set act = fso.opentextfile(logname,8,True)
act.writeline log
act.close

loop
act.close
serverlist.close

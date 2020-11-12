'--------------------------------------
' Created By Frank Jensen
' V1.1
' Used to detect files older than X minutes
'
'	Domain update 1/17/19 Fjensen
'
'--------------------------------------
'dirlist = array("\\phx.dsfcu.local\Nautilus\import$\workflow\scripts\Dev\test\") 'no trailing backslash Test location
dirlist = array("\\pi-gpsql02\GP\GPdata\Payables\DSCU\Imports\Importing\") 'no trailing backslash
nMaxFileAge = 30 'anything older than this many minutes will cause alert
Dim WshNetwork
Set WshNetwork = CreateObject("WScript.Network")
ComputerName = WshNetwork.ComputerName
for each x in dirlist     'Loop through Array'

'Set objects & error catching
	On Error Resume Next
	Dim fso 
	Dim objFolder
	Dim objFile
	Dim objSubfolder
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFolder = fso.GetFolder(x)


'DELETE all files in TempFolder Path older than x days
	For Each objFile In objFolder.files
    		If DateDiff("n", objFile.DateCreated,Now) > nMaxFileAge Then
        		Call send_mail(objfile.name,x,objFile.DateCreated)
        	End If
	Next
Next

Sub send_mail(file_name,file_path,file_date)
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "InterDyn Import File stuck for GP" 
objMessage.From = "OnBaseAlerts@desertfinancial.com" 
objMessage.To = "frank.jensen@switchthink.com,howard.hodge@switchthink.com,christina.bridwell@desertfinancial.com,Denise.sierra@desertfinancial.com,Nancy.Hong@desertfinancial.com,Erica.Eve@desertfinancial.com,Stephanie.Alejandro@desertfinancial.com" 
'objMessage.To = "frank.jensen@desertfinancial.com"
objMessage.TextBody = "Below is the path and file that has been stuck for more than 30min." & vbcr & vbcr

'export_list is anything you want to add to the msg

objMessage.TextBody = objMessage.TextBody & file_name & vbcr & file_date & vbcr & file_path & vbcr & vbcr

objMessage.TextBody = objMessage.TextBody & "To resolve this please delete the file and rerun the export process from OnBase." & vbcr
objMessage.TextBody = objMessage.TextBody & "If this problem persists, please contact InterDyn support as the GP Integration is stuck." & vbcr
objMessage.TextBody = objMessage.TextBody & vbcr & vbcr & "Email sent from: " & ComputerName

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.desertfinancial.com"
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update

objMessage.Send

End Sub

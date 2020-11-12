CONST bytesTokb = 1024
totalsize = 0
filecount = 0
dirlist = array("\\phx.dsfcu.local\Nautilus\import$\Keystone\Receipts\Index\archive","\\phx.dsfcu.local\Nautilus\import$\Keystone\Forms\Index\archive","\\phx.dsfcu.local\Nautilus\import$\Keystone\Forms\Loan\csv\archive","\\phx.dsfcu.local\Nautilus\import$\Keystone\Forms\Upload\archive","\\phx.dsfcu.local\Nautilus\import$\Keystone\Reports\BACKUP","\\phx.dsfcu.local\Nautilus\import$\Keystone\Reports\Index","\\phx.dsfcu.local\Nautilus\import$\Keystone\Reports\proc\backup") 'no trailing backslash
'dirlist = array("D:\scripts\temp")
nMaxFileAge = 15 'anything older than this many days will be removed
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
    		If DateDiff("d", objFile.DateLastModified,Now) > nMaxFileAge Then
			totalsize = totalsize + objFile.Size
			filecount = filecount + 1
        		objFile.Delete True
        	End If
	Next
Next
msgbox "File Size: " & CINT(totalsize / bytesTokb) & "kb - Total files deleted = " & filecount
wscript.echo "File Size: " & CINT(totalsize / bytesTokb) & "kb - Total files deleted = " & filecount

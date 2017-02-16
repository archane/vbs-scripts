'--------------------------------------
' Created By Frank Jensen
' V1.0
' Used to delete files and subfolder older than X days
'
'
'--------------------------------------
dirlist = array("first_folder","second_folder") 'no trailing backslash
nMaxFileAge = 30 'anything older than this many days will be removed
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
    		If DateDiff("d", objFile.DateCreated,Now) > nMaxFileAge Then
        		objFile.Delete True
        	End If
	Next

'DELETE all subfolders in TempFolder Path older than x days
	For Each objSubfolder In objFolder.Subfolders
    		If DateDiff("d", objSubfolder.DateCreated,Now) > nMaxFileAge Then
            		objSubfolder.Delete True      
        	End If
	Next
Next

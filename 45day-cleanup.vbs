dirlist = array("\\phx.dsfcu.local\nautilus\import$\Prod Split Backup","\\phx.dsfcu.local\nautilus\import$\Dev Split Backup","\\ppr-nas01\archive$\FSI_Reports\PROCESSED","\\phx-fs-01\dfs\Mortgage_Servicing\For Imaging Final Title & Recorded DOT\Processed","\\phx-fs-01\dfs\Mortgage_Servicing\HUD-1 Settlement Statements\Processed") 'no trailing backslash
nMaxFileAge = 45 'anything older than this many days will be removed
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

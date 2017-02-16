'---------------------------------------'
' Created by Frank Jensen
' V1.0
'Removed header from a document, can remove as many lines fro the beginning as needed
'
'
'---------------------------------------'
Const FOR_READING = 1 
Const FOR_WRITING = 2 
inputfile = WScript.Arguments(0) 'for use in OnBase'
outputfile = WScript.Arguments(1) 'for use in OnBase'
iNumberOfLinesToDelete = 1 'change to effect number of lines removed from final file'

Set objFS = CreateObject("Scripting.FileSystemObject")
msgbox inputfile 
Set objTS = objFS.OpenTextFile(inputfile, FOR_READING) 
strContents = objTS.ReadAll 
objTS.Close 
 
arrLines = Split(strContents, vbNewLine) 
Set objTS = objFS.OpenTextFile(outputfile, FOR_WRITING, True) 
 
For i=0 To UBound(arrLines) 
   If i > (iNumberOfLinesToDelete - 1) Then 
      objTS.WriteLine arrLines(i) 
   End If 
Next

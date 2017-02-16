'--------------------------------------------'
' Created by Frank Jensen
' V1.0
' Used for comparing 2 files line by line and creating a list of the differences
'
'
'-------------------------------------------'
Const FOR_READING = 1 
Const FOR_WRITING = 2 
first_compare_file = "C:\Users\frjensen\Documents\QCD\dates.csv"
second_compare_file = "C:\Users\frjensen\Documents\QCD\remaining.csv"
outputfile = "C:\Users\frjensen\Documents\QCD\with-dates.csv"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(first_compare_file, FOR_READING) 
 
     	Set output = objFS.opentextfile(outputfile,FOR_WRITING,True)
     		Do Until objTS.AtEndOfStream 
     		strContents = objTS.readline
     		account = split(strContents,",")
     		Set file = objFS.OpenTextFile(second_compare_file, FOR_READING, True)
     		Do Until file.AtEndOfStream
     			imported_text = file.readline
     			if imported_text = account(1) then
     			output.writeline strContents
     			end if
     		loop
     		file.close
     		loop
     	file.close
output.close
msgbox "done"
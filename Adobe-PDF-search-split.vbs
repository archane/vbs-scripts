'************************************************************************
'Created by Frank Jensen
' Ver1.1 deployed on 1/7/14
'
' change list
' v1.1 added email function for alerts when issues arise
' This script opens a PDF, searches for a specific string, and moved the page it finds the text on to the front of the PDF using PDF toolkit
'
'************************************************************************
on error resume next 'I hate having to use this'
'Option Explicit
 
Dim acroApp, acroAVDoc, acroPDDoc, acroRect, acroPDPage, acroPageView,gPDFPath, nElem, newAApdf,newpdf
Dim acroHiList, acroPDTextSel,objShell, pagelocation,parts,adobe_file,teststring,currpage
Dim fso, strPath, filelog, holdingfolder,processingfolder,procfolder,holdfiles,holdfolder,folderIdx 
Dim curr_time, start_time,testtime, objMessage,body_text,subject_text,body_text2,subject_text2
Set objShell = WScript.CreateObject ("WScript.shell")      'for running the CLI'
set fso = createobject("scripting.filesystemobject")
start_time=now
'holdingfolder = "c:\Scripts\test\"  'source of pdf files'
'processingfolder = "c:\Scripts\process\"  'destination of new pdf files'

holdingfolder = "UNC path here"  'source of pdf files'
processingfolder = "different UNC path here"  'destination of new pdf files'
strPath = "c:\Scripts\OAO_log.txt" 'log file'
Set filelog = fso.opentextfile(strPath,8,True)
filelog.writeline "being processing at - " & start_time  
Set holdfolder = fso.GetFolder(holdingfolder)
Set holdfiles = holdfolder.Files 'file list to process'
For each folderIdx In holdfiles
          Set procfolder = fso.GetFolder(processingfolder)  
          'do while procfolder.files.Count <>0 'waiting for the files to move''
               'WScript.Sleep 10000 '10 second pause'
          'loop           
gPDFPath =holdingfolder & folderIdx.Name ' ** Initialize Acrobat by creating App object
Set acroApp = CreateObject("AcroExch.App")' ** show acrobat
Set acroAVDoc = CreateObject("AcroExch.AVDoc")' ** open the PDF this is required for the SDK to actually look at the file
     If acroAVDoc.Open( gPDFPath,"Accessing PDF's") Then 
          If acroAVDoc.IsValid = False Then ExitTest()
               acroAVDoc.BringToFront()
               Call acroAVDoc.Maximize(True)
               Set acroPDDoc = acroAVDoc.GetPDDoc()
               Set acroHiList = CreateObject("AcroExch.HiliteList")' ** Create a hilite that includes all possible text on the page
               Call acroHiList.Add( 0, 32000 ) ' selects the entire page of text'
               Set acroPageView = acroAVDoc.GetAVPageView()
          for currpage = 0 to acroPDDoc.GetNumPages() ' loop through pages to find text'
          curr_time=now
          testtime=datediff("s",start_time,curr_time)
               if testtime >=600 then
                    filelog.writeline "Programming running more than 10mins, exiting now"
                    subject_text = "OAO script has run beyond 10min"
                    body_text = "Verify that Adobe is not held open and that cscript is not running on the server"
                    AcroApp.CloseAllDocs()   'begin closing and variable handling'
                    AcroApp.Exit()
                    filelog.close
                    call email(body_text,subject_text) 'emails when adobe appears to be hung'                    
                    exit for
               end if
               Call acroPageView.Goto( currpage )
               Set acroPDPage = acroPageView.GetPage()
               Set acroPDTextSel = acroPDPage.CreatePageHilite( acroHiList )
                    If acroPDTextSel Is Nothing Then 'detects if the page has no text'
                         filelog.writeline "No text to highlight for  -->"& acroPDDoc.GetFileName()
                    subject_text = "No text to highlight for  -->"& acroPDDoc.GetFileName()
                    body_text = "Verify that the file is not blank. If the file is not blank, import manually. "
                    AcroApp.CloseAllDocs()   'begin closing and variable handling'
                    AcroApp.Exit()
                    call email(body_text,subject_text) 'emails if the file is blank or missing the required data'
                    exit for 
                    else
               ' ** Set that as the current text selection and show it
               Call acroAVDoc.SetTextSelection( acroPDTextSel )
               Call acroAVDoc.ShowTextSelect()' ** Get the number of words in the text selection and the first word in selection
                    If acroPDTextSel.GetNumText > 0 Then
                         teststring = Replace(acroPDTextSel.GetText( 0 ) & acroPDTextSel.GetText( 1 ), vbCrLf, "")
                              If (teststring) = "ACCOUNT APPLICATION" then                               'this is the search string check here'
                                   filelog.writeline "File Name ---> "& acroPDDoc.GetFileName()
                                   pagelocation = acroPDTextSel.GetPage() +1
                                   filelog.writeline "Current Selection Page ---> "& pagelocation
                                   parts = split(acroPDDoc.GetFileName(),".")
                                   adobe_file = parts(0)
                                   Call acroPDTextSel.Destroy() 'removes current selected text for the next pass on a new doc'
                                   AcroApp.CloseAllDocs()   'begin closing and variable handling'
                                   AcroApp.Exit()
                                   Set acroPDTextSel = Nothing : Set acroRect = Nothing : Set AcroApp =  Nothing : Set AcroAVDoc =  Nothing
                                   Set acroHiList = Nothing : Set acroPageView = Nothing : Set acroPDPage = Nothing   
                                   objShell.run "pdftk " & gPDFPath & " cat " & pagelocation & "-" & pagelocation+1 &" output "& holdingfolder &"AA_" & adobe_file & ".pdf"
                                   newAApdf = "AA_" & adobe_file & ".pdf"
                                   WScript.Sleep 5000 'pause so adobe and pdftk can close
                                   objShell.run "pdftk " & gPDFPath & " cat 1-" & pagelocation-1 & " " & pagelocation+2 & "-end output "& holdingfolder &"org_" & adobe_file & ".pdf"                            
                                   newpdf = "org_" & adobe_file & ".pdf"
                                   WScript.Sleep 5000
                                   objShell.run "pdftk " & holdingfolder & newAApdf & " " & holdingfolder & newpdf & " cat output "& processingfolder & adobe_file & ".pdf"
                                   WScript.Sleep 5000
                                   filelog.writeline "move - " & holdingfolder & folderIdx.Name & " ->> " & holdingfolder & folderIdx.Name
                                   fso.deletefile holdingfolder & "AA_" & adobe_file & ".pdf"
                                   fso.deletefile holdingfolder & "org_" & adobe_file & ".pdf" 
                                   fso.deletefile holdingfolder & folderIdx.Name 
                                   exit for
                              end if
                    Else
                         filelog.writeline "2No text to highlight for  -->"& acroPDDoc.GetFileName()
                         subject_text = "Frank something weird happened"
                         body_text = "Scripted exited, check the log"
                         AcroApp.CloseAllDocs()   'begin closing and variable handling'
                         AcroApp.Exit()
                         call email(body_text,subject_text) 
                    End If
                    end if 
          next
     End If
next
filelog.writeline "ending process at - " & now
Set acroPDTextSel = Nothing : Set acroRect = Nothing : Set AcroApp =  Nothing:Set AcroAVDoc =  Nothing
Set acroHiList = Nothing : Set acroPageView = Nothing : Set acroPDPage = Nothing : Set objShell = Nothing
filelog.close

function email(body_text2,subject_text2)
                    Set objMessage = CreateObject("CDO.Message") 
                    objMessage.Subject = subject_text2 
                    objMessage.From = "email address here" 
                    objMessage.To = "comma delimited email list here" 

                    objMessage.TextBody = objMessage.TextBody

                    objMessage.Configuration.Fields.Item _
                    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
                    objMessage.Configuration.Fields.Item _
                    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail server goes here"
                    objMessage.Configuration.Fields.Item _
                    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
                    objMessage.Configuration.Fields.Item _
                    ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
                    objMessage.Configuration.Fields.Update

                    objMessage.Send
end function

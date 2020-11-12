'Written by Frank Jensen 7/23/2015'
'conversion file for C2 to add fullpath to file name and to change form name for Doc Type'

set fso = createobject("scripting.filesystemobject")
Set form_name = CreateObject ("System.Collections.ArrayList")
strPath = "\\phx.dsfcu.local\Nautilus\import$\Keystone\Forms\Index\" 'path for index files to read'
storePath = "\\phx.dsfcu.local\Nautilus\import$\Keystone\Forms\Index\" 'path to write new index files'
pdfpath = "\\phx.dsfcu.local\Nautilus\import$\Keystone\Forms\" 'path to add to file names for Onbase'

with form_name  'Array has to be listed like this for Indexof function to work'
.Add "test"
.Add "Stop Payment Personal Check"
.Add "Roth IRA Withdrawal Authorization"
.Add "Roth IRA Rollover Deposit Certification"
.Add "IRA Beneficiary Designation Form"
.Add "IRA Withdrawal Authorization"
.Add "IRA Waiver of RMD for Current Year"
.Add "IRA Waiver of Minimum Distribution"
.Add "IRA Rollover Request"
.Add "IRA Required Minimum Distribution Form"
.Add "Required Minimum Distribution Checklist"
.Add "HSA Withdrawal Authorization"  
.Add "HSA Transfer Request Form"          
.Add "HSA Deposit Rollover Certification"  
.Add "HSA Beneficiary Designation"
.Add "HSA Authorized Signer Form"         
.Add "HE Interest Rate Change Notice - Annual"                      
.Add "HE Interest Rate Change Notice - Monthly"
.Add "ESA Withdrawal Authorization Form"
.Add "Certificate Signature Card"
.Add "Health Savings Account Application"
.Add "AUTMA Application"
.Add "IRA Simplifier Account Application"
.Add "Roth IRA Simplifier Account Application"
.Add "Business Account Master Application"
.Add "Secure Savings Account Application"
.Add "Organization Account Agreement"
.Add "Business Certification of Identity"
.Add "Upload Document"
.Add "Account Application"
.Add "Card Overdraft Coverage"
.Add "ATM Deposit Adjustment Letter"
.Add "Traditional IRA Transfer Request"
.Add "Stop Payment Personal Check - Joint"
.Add "Stop Payment Confirmation Personal Check"
.Add "Stop Payment Cancellation Personal Check"
.Add "Replacement Agreement"
.Add "Restricted Account Agreement"
.Add "Card Overdraft Coverage"
.Add "IRA Deposit Rollover Certification"
.Add "Business Acct Public Funds Agreement"
.Add "Business Account Auth Agreement"
.Add "Organization Account Agreement Savings"
.Add "membership savings"
.Add "Roth IRA Transfer Request"
.Add "DBA Business Master Application"
.Add "W9"
.Add "Notice of Adverse Action"
.Add "Name Change Request"
.Add "Health Savings Account Application"
.Add "Traditional IRA Transfer Request"
.Add "Acknowledge Fraud Claim"
.Add "Acknowledge Visa Dispute 1"
.Add "Acknowledge Visa Dispute 2"
.Add "Acknowledge Visa Dispute Final"
.Add "Add'l Info Needed - Debit Card Dispute"
.Add "ATM Cash Deposit Dispute Acknowlegement"
.Add "ATM Cash Deposit Dispute Resolution"
.Add "ATM Deposit Return"
.Add "Bank By Mail Returned Deposit"
.Add "Chargeback Item Loan Notice"
.Add "Chargeback Notice"
.Add "Confirm of Electronic Funds Transfer"
.Add "Deposit Adjustment Letter"
.Add "DES Levy 2"
.Add "DES Levy with $250.00 exemption"
.Add "Domestic Levy Letter"
.Add "Ebranch Close Letter"
.Add "Electronic Sorry Letter"
.Add "Error Resolution Letter"
.Add "External Electronic Funds Trans Auth 1"
.Add "External Electronic Funds Trans Auth 2"
.Add "Final Account Resolution"
.Add "Final Demand Letter"
.Add "Final VISA Fraud Resolution"
.Add "Foreign Item Collection Fee Notification"
.Add "GAP 2"
.Add "HSA contribution Change Form"
.Add "IRA Notice of Withholding"
.Add "IRS Notice of Levy"
.Add "Large Dollar Collection Item FinalCredit"
.Add "Memorial or Benefit Closed Acct Notice"
.Add "Miscommunication Letter"
.Add "New Account Error Final"
.Add "Notice of Hold Letter"
.Add "Overdrawn Account: 10 Day Notice"
.Add "Pay Off Quote"
.Add "Returned Electronic Transfer: Letter 1"
.Add "Returned Electronic Transfer: Letter 2"
.Add "Sorry Letter Desert Schools Error"
.Add "Stop Unauth Electronic Activity Request"
.Add "Unauth Remotely Created Ck Final Credit"
.Add "US Treasury Notification of Death"
.Add "AZ Dept of Revenue Levy Notice"
.Add "Z_Denial of Account"
.Add "AZ Dept of Economic Security Levy Notice"
.Add "AZ Dept of Transportation Levy Notice"
.Add "HELOC Fixed Rate Option Cancellation"
.Add "HELOC Fixed Rate Option Change In Terms"
.Add "HELOC Fixed Rate Option Voucher"
.Add "Beneficial Ownership Certification"
.Add "Returned Electronic Transfer: OLB"
.Add "Business And Organization Account App"
.Add "External Electronic Funds Transfer Auth"
.Add "External Transfer Form"
.Add "Membership Closed Letter"
.Add "LOC Suspension Adverse Action Notice"
.Add "Business Signature Addendum"


end with

doctype_name = array("test2","3580","458","454","461","455","461","463","452","462","463","279","282","281","280","3480","3853","3853","460","270","272","287","450","456","437","283","285","438","123","283","3401","3821","451","3580","3580","3580","284","3967","3401","453","3291","439","285","283","457","437","294","3728","3465","272","451","4086","4086","4086","4090","4086","4086","4090","2414","2546","2839","2427","2083","2551","4104","4104","4104","Ebranch Close Letter","506","2428","2429","2936","2774","511","4090","2776","GAP2","4071","464","4104","2434","Memorial or Benefit Closed Acct Notice","505","New Account Error Final","494","499","423","495","495","506","4090","2624","4105","4104","3728","4104","4104","4239","4240","4241","4245","4305","437","290","4536","4540","3294") 'EXACT number of the doc type in Onbase'

Set CurFolder = fso.GetFolder(strPath)
Set CurFiles = CurFolder.Files

For Each CurFolderIdx in CurFiles
	TARGET = CurFolderIdx.Name
     Set file = fso.opentextfile(strPath & TARGET)
     file_split = split(CurFolderIdx.name,".")
	 	If LCase(Mid(TARGET, InStrRev(TARGET, "."))) = ".txt" Then
	     Set outputfile = fso.opentextfile(strPath & LEFT(TARGET, (LEN(TARGET)-4)) & ".csv",2,True)

     Do Until file.AtEndOfStream
          imported_text = file.readline
          split_text = split(imported_text,"|")
          'msgbox split_text(0)
          intIndex = form_name.IndexOf (split_text(0),0)
          If doctype_name(intIndex) = "3728" or doctype_name(intIndex) = "3465" then     
               outputfile.writeline doctype_name(intIndex) & "|" & split_text(1) & "|" & split_text(2) & "|" & pdfpath & split_text(3)
          else
               outputfile.writeline doctype_name(intIndex) & "|" & split_text(1) & "|" & split_text(2) & "|" & split_text(3) & "|" & split_text(4) & "|" & pdfpath & split_text(5)
          end if
     loop
          
     file.close
     outputfile.close
     If doctype_name(intIndex) = "123" then
          fso.Copyfile strPath & file_split(0) & ".csv", pdfpath & "Uploaded\" & file_split(0) & ".csv", True
          fso.DeleteFile(strPath & file_split(0) & ".csv")
     end if
     fso.Copyfile strPath & CurFolderIdx.name, strPath & "archive\" & CurFolderIdx.name, True 'physical move of the file'
     fso.DeleteFile(strPath & CurFolderIdx.name)
     end if
next
file.close
outputfile.close

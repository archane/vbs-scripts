
'Written by Frank Jensen 7/23/2015'
'conversion file for C2 to add fullpath to file name and to change form name for Doc Type'

set fso = createobject("scripting.filesystemobject")
Set form_name = CreateObject ("System.Collections.ArrayList")
strPath = "\\phx.dsfcu.local\Nautilus\import$\Keystone\Reports\" 'path for index files to read'
storePath = "\\phx.dsfcu.local\Nautilus\import$\Keystone\reports\proc\" 'path to write new index files'

with form_name  'Array has to be listed like this for Indexof function to work'
.Add "Akcelerant_Framework_Collateral_File"
.Add "Akcelerant_Framework_Person_Address_Link_File"
.Add "Akcelerant_Framework_Payment_File"
.Add "Akcelerant_Framework_Address_File"
.Add "Akcelerant_Framework_Loan_File"
.Add "Akcelerant_Framework_Share_File"
.Add "Akcelerant_Framework_Relationship_File"
.Add "Akcelerant_Framework_Contact_File"
.Add "Akcelerant_Framework_Person_File"
.Add "Patriot_Officer_Savings_Account_File"
.Add "Patriot_Officer_Debit_Card_File"
.Add "Patriot_Officer_Member_Information_File"
.Add "Patriot_Officer_Member_Account_Cross_Reference_File"
.Add "Patriot_Officer_Loan_Account_File"
.Add "Patriot_Officer_Demand_Account_File"
.Add "Patriot_Officer_Certificate_Account_File"
.Add "Patriot_Officer_Branch_File"
.Add "Patriot_Officer_ATM_Card_File"
.Add "Patriot_Officer_Account_Mailing_Address_File"
.Add "Consumer_Deposit_Fee_Posting"
.Add "Consumer_Deposit_Fee_Posting_Exceptions"
.Add "Fusion_ATM_Refund_Post"
.Add "Fusion_ATM_Refund_Post_Exceptions"
.Add "Business_Service_Fees"
.Add "Business_Service_Fees_Exceptions"
.Add "RR_ATM_Refund_Post"
.Add "RR_ATM_Refund_Post_Exceptions"
.Add "Daily_Notice_Production"
.Add "Harland_MCIF_Card_File"
.Add "Harland_MCIF_Main_File"
.Add "Harland_MCIF_Extra_File"
.Add "Household_New_Person_Evaluation"
.Add "Database_Stats_Exceptions"
.Add "FDI_Dealer_Track_Exceptions"
.Add "External_Loan_Delete_Exceptions"
.Add "Escrow_Analysis_Job_Analyze_Exceptions"
.Add "Health_Care_Payment_Notice_Letters_Exceptions"
.Add "HELOC_Annual_Fee"
.Add "HELOC_Annual_Fee_Exceptions"
.Add "Indirect_Dealer_Invoice_Letter"
.Add "Dividend_Accrual_Exceptions"
.Add "Reg_D_Update"
.Add "Reg_D_Reset_From_File"
.Add "Reg_D_Reset_From_File_Exceptions"
.Add "BRD_Monthly_Service_Fee"
.Add "Card_Plastic_Design_Update"
.Add "Card_Plastic_Design_Update_Exceptions"
.Add "Override_Report_Exceptions"
.Add "Health_Care_Payment_Notice_Letters"
.Add "Dade_Daily_Load_Extract"
.Add "Alogent_Blue_Point_MRDC_Post"
.Add "Alogent_Blue_Point_MRDC_Post_Exceptions"
.Add "FICS_MS_Import_Exceptions"
.Add "Late_Fee_Assessment_Transaction_Exceptions"
.Add "Business_Sweep_Exceptions"
.Add "Business_Sweep"
.Add "FICS_MS_Joint_Dat5_Exceptions"
.Add "FICS_Commercial_Loan_Import_Exceptions"
.Add "Delay_A_Pay_Post_Job_Exceptions"
.Add "Delay_A_Pay_Post_Job"
.Add "Zero_Balance_Closure_Posting_Report_Exceptions"
.Add "Business_Analysis_Exceptions"
.Add "Business_Analysis"
.Add "Reg_D_Update_Exceptions"
.Add "IRS_File_Maintenance_5498_SA"
.Add "IRS_File_Maintenance"
.Add "Heloc_Promo"
.Add "Heloc_Promo_Job_Exceptions"
.Add "IRS_5498_Mass_File_Cleanup"
.Add "IRS_5498_Mass_File_Cleanup_Exceptions"
.Add "Reg_D_Reset"
.Add "Reg_D_Reset_Exceptions"
.Add "Tax_Plan_RMD_Update_Exceptions"
.Add "FICS_MS_Import_Exceptions"

end with


Set CurFolder = fso.GetFolder(strPath)
Set CurFiles = CurFolder.Files

For Each CurFolderIdx in CurFiles
     TARGET = CurFolderIdx.Name
	file_split = split(TARGET,".")
     intIndex = form_name.IndexOf (file_split(1),0)
     if intIndex  <> "-1" then
     'fso.copyfile strPath & CurFolderIdx.name, strPath & "to_delete\" & CurFolderIdx.name, True 'physical move of the file'
     fso.DeleteFile(strPath & CurFolderIdx.name) 
     else
     Set file = fso.opentextfile(strPath & TARGET)
     file_split = split(CurFolderIdx.name,".")
	 	If LCase(Mid(TARGET, InStrRev(TARGET, "."))) = ".txt" Then
	     Set outputfile = fso.opentextfile(storepath & LEFT(TARGET, (LEN(TARGET)-4)) & ".txt",2,True)
          
     file_date = mid(TARGET,1,8)
     imported_text = file.readall
     outputfile.writeline file_date
     outputfile.write imported_text
          
     file.close
     outputfile.close
     'fso.Copyfile strPath & CurFolderIdx.name, strPath & "archive\" & CurFolderIdx.name, True 'physical move of the file'
     fso.DeleteFile(strPath & CurFolderIdx.name)
     end if
     end if
next
file.close
'msgbox "done"

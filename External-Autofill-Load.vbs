Dim objShell
Set objShell = WScript.CreateObject ("WScript.shell")
set fso = createobject("scripting.filesystemobject")

If fso.FileExists("\\phx-fs-01\dfs\Document_Solutions\Autofill\KeystoneAutofill.txt") Then
	fso.Copyfile "\\phx-fs-01\dfs\Document_Solutions\Autofill\KeystoneAutofill.txt","\\phx.dsfcu.local\Nautilus\import$\workflow\Prod\Autofill\KeystoneAutofill.txt"
	WScript.Sleep 50000 'just to give the system time to copy

	objShell.run "cmd /c sqlcmd -S pi-onbasesql01 -U hsi -P wstinol -d DSFCU -Q ""Exec dbo.Keystone_acct_af"" -o  d:\Scripts\Logs\AF_Import.txt -s"","" -C -h-1 -I"
	WScript.Sleep 50000 'since the sql doesn't cause the script to wait

	Set objShell = Nothing
	curr_time=now
	curr_date=Replace(FormatDateTime(curr_time,2),"/","")

	fso.Copyfile "\\phx-fs-01\dfs\Document_Solutions\Autofill\KeystoneAutofill.txt", "\\phx-fs-01\dfs\Document_Solutions\Autofill\backup\KeystoneAutofill_"& curr_date & ".txt", True
	fso.DeleteFile "\\phx-fs-01\dfs\Document_Solutions\Autofill\KeystoneAutofill.txt"
	fso.DeleteFile "\\phx.dsfcu.local\Nautilus\import$\workflow\Prod\Autofill\KeystoneAutofill.txt"

	objShell.run "cmd /c iisreset pi-onbaseunty01"
	WScript.Sleep 1000
	objShell.run "cmd /c iisreset pi-onbaseunty02"
	WScript.Sleep 1000

End If 

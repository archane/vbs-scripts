'Simple VBScript to run stored procedures in an MS DB
'Created by Frank Jensen
'V1.0
'
Dim objShell
Set objShell = WScript.CreateObject ("WScript.shell")
objShell.run "cmd /c sqlcmd -S <servername> -U <username> -P <password> -d <databasename> -Q ""Exec hsi.APupdate"" -C"
Set objShell = Nothing

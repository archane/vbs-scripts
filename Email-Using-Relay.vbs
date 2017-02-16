                    
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "Email Test for Relay server" 
objMessage.From = "archane@gmail.com" 
objMessage.To = "archane@gmail.com" 
objMessage.TextBody = "Below is the list of failed reports to verify." & vbcr & vbcr

'export_list is anything you want to add to the msg

objMessage.TextBody = objMessage.TextBody & export_list

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.domainhere.org"
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update

objMessage.Send

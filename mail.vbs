Option Explicit
 
Dim oPrm
Dim oMsg
Dim sUrl

Dim fromAddress
Dim fromPassword
Dim toAddress

Dim subj
Dim body
Dim attach 'Full Path

fromAddress = "from@gmail.com"
fromPassword = "password"

toAddress = "to@gmail.com"
'toAddress = "to@yahoo.co.jp"


subj = "åèñº"
body = "ñ{ï∂"
attach = "C:\vbs_mail\test.txt"


Set oPrm = WScript.Arguments
Set oMsg = CreateObject("CDO.Message")
 
oMsg.From = fromAddress
'oMsg.To = oPrm(0)
oMsg.To = toAddress
oMsg.Subject = subj
oMsg.TextBody = body
oMsg.AddAttachment(attach)
 
sUrl = "http://schemas.microsoft.com/cdo/configuration/"
 
With oMsg.Configuration.Fields
    .Item(sUrl & "sendusing") = 2
    .Item(sUrl & "smtpserver") = "smtp.gmail.com"
    .Item(sUrl & "smtpserverport") = 465
    .Item(sUrl & "smtpauthenticate") = 1
    .Item(sUrl & "smtpusessl") = true
    .Item(sUrl & "sendusername") = fromAddress
    .Item(sUrl & "sendpassword") = fromPassword
    .Update
end With
 
oMsg.Send
 
Set oMsg = Nothing
Set oPrm = Nothing
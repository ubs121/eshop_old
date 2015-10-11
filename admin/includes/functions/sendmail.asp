<%
Private sub SendMail()
	select case strMailMethod
		case "cdonts":
			SendCdonts()
		case "cdosys":
			SendCdosys()
		case "dundas":
			SendDundas()
		case "jmail":
			SendJmail()
		case "persits":
			SendPersits()
		case "aspmail":
			SendAspmail()
	end select
end sub

private sub sendAspmail()
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	
	Mailer.FromAddress = MailFrom
	if len(strMailOut) > 0 then
		Mailer.RemoteHost = strMailOut
	else
		Mailer.RemoteHost = "localhost"
	end if
	
	Mailer.AddRecipient "", MailTo
	Mailer.Subject = MailSubject
	Mailer.BodyText = MailBody
	Mailer.ContentType = "text/html"
	
	Mailer.sendmail
end sub

private sub SendPersits()
  Const CdoBodyFormatText = 0
  Const CdoBodyFormatHTML = 0
  Const CdoMailFormatMime = 0
  Dim Message 'As New cdonts.NewMail
 
  'Create CDO message object
  Set Message = Server.CreateObject("Persits.MailSender")
  With Message
  
  if len(strMailOut) > 0 then
 	Message.Host = strMailOut
  end if

 Message.FromName = MailFrom ' Specify sender's name
 Message.AddAddress MailTo

 Message.Subject = MailSubject
 Message.IsHTML = True
 Message.Body = MailBody & Chr(13) & Chr(10)
 On Error Resume Next
 Message.Send 
  
    'Send the message
    .Send
  End With
End Sub

Private sub SendCdosys()
	' Create the e-mail server object
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
	' Outgoing SMTP server
	if len(strMailOut) = 0 then
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	else
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailOut
	end if
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	objCDOSYSCon.Fields.Update
	
	' Update the CDOSYS Configuration
	Set objCDOSYSMail.Configuration = objCDOSYSCon
	objCDOSYSMail.From = MailFrom
	objCDOSYSMail.To = MailTo
	objCDOSYSMail.Subject = MailSubject
	objCDOSYSMail.HTMLBody = MailBody
	objCDOSYSMail.Send
	'Close the server mail object
	Set objCDOSYSMail = Nothing
	Set objCDOSYSCon = Nothing 
end sub

private sub SendCdonts()
  Const CdoBodyFormatText = 0
  Const CdoBodyFormatHTML = 0
  Const CdoMailFormatMime = 0
  Dim Message 'As New cdonts.NewMail
  
  'Create CDO message object
  Set Message = CreateObject("cdonts.NewMail")
  With Message
    
    'Set email adress, subject And body
    .To = MailTo
    .Subject = MailSubject
    .Body = MailBody
    
    'set mail And body format
    .MailFormat = CdoMailFormatHTML
    .BodyFormat = CdoBodyFormatHTML
    
    'Set sender address If specified.
    .From = MailFrom
    
    'Send the message
    .Send
  End With
End Sub

private sub SendDundas()
	dim objDundasMail
	Set objDundasMail = Server.CreateObject("Dundas.Mailer")
	
	objDundasMail.TOs.Add MailTo
	objDundasMail.Subject = MailSubject
	objDundasMail.FromAddress = MailFrom
	if len(strMailOut) > 0 then
		objDundasMail.SMTPRelayServers.Add strMailOut
	end if
	objDundasMail.HTMLBody = MailBody
	
	objDundasMail.SendMail
	
	set objDundasMail = nothing
end sub

private sub SendJMail()
	set objJmail = Server.CreateOBject( "JMail.Message" )
	
	objJmail.logging = false
	objJmail.silent  = true
	
	objJmail.from = MailFrom
	objJmail.AddRecipient MailTo
	objJmail.subject = MailSubject
	objJmail.HTMLbody = MailBody
	
	if len(strMailOut) > 0 then
		objJmail.send(strMailOut)
	end if
	
	set objJmail = nothing
end sub
%>
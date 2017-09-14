<%@LANGUAGE="VBSCRIPT"%>

<%
	Dim strFromName, strMemberName, strToEmail, strFromEmail, strSubject, strFormat, strEmailBody

	strToEmail 		= "Ed Agnes <eagnes@cwaustral.com.au>"
	strFromName 	= Request.Form("Name")
	strFromEmail 	= Request.Form("Email")
	strSubject 		= Request.Form("Subject")
	strMessage		= Request.Form("Message")
	strFormat 		= "HTML"
	strEmailBody 	= bodytext


	bodytextFooter	= getFileContents("VESAEmailFooter.asp")
	bodytext 		= getFileContents("VESAEmailHeader.asp")

	bodytext = bodytext & "<p>Hi <strong>Ed</strong>,</p>" & vbCrLf _
						& "<p>You have an enquiry from <strong>" & strFromName & "</strong> regarding" & strSubject & ".</p>" & vbCrLf _
						& "<p>" & vbCrLf _
						& "Enquiry/Feedback:<br />" & vbCrLf _
						& "<strong>" & strMessage & "</strong>" & vbCrLf _
						& "</p>" & vbCrLf _
						& "<p>" & vbCrLf _
						& "Contact details:<br />" & vbCrLf _
						& "Email: <strong><a href='mailto:" & strFromEmail & "'>" & strFromEmail & "</a></strong><br />" & vbCrLf _
						& "</p>" & vbCrLf _
						& "<p>" & vbCrLf _
						& "Thank You.<br />" & vbCrLf _
						& "<strong>VESA Database Administrator</strong>" & vbCrLf _
						& "</p>" & vbCrLf _
						& "<p>" & vbCrLf _
						& "--<br />" & vbCrLf _
						& "This e-mail was sent from the VESA Members Database Conatct Us Form (http://cwmedia.com.au/database/vesa/Members/VESAContact.asp)" & vbCrLf _
						& "</p>"

	bodytext = bodytext & bodytextFooter

	Call SendMail(strFromName, strMemberName, strToEmail, strFromEmail, strSubject, strFormat, strEmailBody)


Function SendMail(strFromName, strMemberName, strToEmail, strFromEmail, strSubject, strFormat, strEmailBody)
	
	'Dimension variables
	Dim cdoSendUsingPickup
	Dim cdoSendUsingPort
	Dim cdoAnonymous
	Dim cdoBasic
	Dim cdoNTLM
	Dim schemas

	schemas = "http://schemas.microsoft.com/cdo/configuration/"

	cdoSendUsingPickup 	= 1 	'Send message using the local SMTP service pickup directory.
	cdoSendUsingPort 	= 2 	'Send the message using the network (SMTP over the network).

	cdoAnonymous 		= 0		'Do not authenticate
	cdoBasic 			= 1 	'basic (clear-text) authentication
	cdoNTLM 			= 2 	'NTLM

	
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = strSubject
	objMessage.From = strFromName & "<" & strFromEmail & ">"
			
	objMessage.To = "<" & strToEmail & ">"
			
	If strFormat = "HTML" Then
		objMessage.HTMLBody = strEmailBody
	Else 
		objMessage.TextBody = strEmailBody
	End If 

	'==This section provides the configuration information for the remote SMTP server.

	objMessage.Configuration.Fields.Item _
	(schemas & "sendusing") = cdoSendUsingPort

	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	(schemas & "smtpserver") = "mail.carmacloud.com"

	'Type of authentication, NONE, Basic (Base64 encoded), NTLM
	objMessage.Configuration.Fields.Item _
	(schemas & "smtpauthenticate") = cdoBasic

	'Your UserID on the SMTP server
	objMessage.Configuration.Fields.Item _
	(schemas & "sendusername") = "newsletter@cwa.carmacloud.com"

	'Your password on the SMTP server
	objMessage.Configuration.Fields.Item _
	(schemas & "sendpassword") = "!cwa673!"

	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	(schemas & "smtpserverport") = 587

	'Use SSL for the connection (False or True)
	objMessage.Configuration.Fields.Item _
	(schemas & "smtpusessl") = true

	'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
	objMessage.Configuration.Fields.Item _
	(schemas & "smtpconnectiontimeout") = 60

	objMessage.Configuration.Fields.Update

	'==End remote SMTP server configuration section==

	objMessage.Send

	Set objMessage = Nothing	
End Function

'***** Read and return the contents of the file as a string *****
'--------------------------------------------------------------------------------------------------
'Pass the name of the file to the function.
Function getFileContents(strIncludeFile)
	Dim objFSO
	Dim objText
	Dim strPage

	'Instantiate the FileSystemObject Object.
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	'Open the file and pass it to a TextStream Object (objText). The "MapPath" function of the Server Object is used to get the physical path for the file.
	Set objText = objFSO.OpenTextFile(Server.MapPath(strIncludeFile))

	'Read and return the contents of the file as a string.
	getFileContents = objText.ReadAll

	objText.Close
	Set objText = Nothing
	Set objFSO = Nothing
End Function
%>
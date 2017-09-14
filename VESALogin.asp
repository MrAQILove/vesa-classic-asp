<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE FILE="include/include.asp"-->
<!--#INCLUDE FILE="include/adovbs.inc"-->
<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = True 

Session("UserLoggedIn") = ""

login		= Request.Form("login")
LoginType	= Request.Form("LoginType")

If login = "true" And LoginType = "AdminLogin" Then 
    Call CheckAdminLogin
ElseIf login = "true" And LoginType = "VESAUnitLogin" Then 
    Call CheckVESAUnitLogin
Else
    Response.Redirect "AdminLogin.asp"
End If

'**************************************************************************************************
' Admin Login 
'**************************************************************************************************

Sub CheckAdminLogin
	Dim objRS
	Dim strSQL
	Dim strUserName, strPassword, strDatabaseName

	strUserName		= Replace(Request.Form("txtUserName"), "'", "''")
	strPassword		= Replace(Request.Form("txtPassword"), "'", "''")
	strDatabaseName	= Replace(Request.Form("DatabaseName"), "'", "''")

	EstablishConnection()

	'- Look up the user name/password.
	strSQL = "SELECT * FROM MembersDB_tblUsers U" 
	strSQL = strSQL & " INNER JOIN MembersDB_tblUserRights UR on U.UserID = UR.UserID "
	strSQL = strSQL & " INNER JOIN MembersDB_tblUserAccess UA on UR.AccessID = UA.AccessID"
	strSQL = strSQL & " WHERE U.UserName='" & strUserName & "'" 
	strSQL = strSQL & " AND U.Password='" & strPassword & "'"
	strSQL = strSQL & " AND U.DatabaseName='" & strDatabaseName & "'"

	Set objRS = Server.CreateObject("ADODB.Recordset")
	Set objRS.ActiveConnection = Conn   
	objRS.Open strSQL, Conn, 1, 1

	On Error Resume Next

	If Err <> 0 Then
		DisplayMessage("<p>There is an error connecting into the database please click back button below and try again. If this is not the frst time you've seen this message, please contact the site administrator.</p>")

	Else 
		'- See if we got anything.
		If CLng(objRS.Fields(0)) < 1 Then
			DisplayMessage("<p>Please check that your user name password is correct.</p>")

		Else
			Select Case objRS.Fields("AccessRights")
				Case "Level 1"
					Session("VESAID") = 1
					Session("UserLoggedIn") = "true"
					Session("AccessRights") = "Level 1"
					Session("User") = strUserName
					Response.Redirect "VESAMain.asp"

				Case "Level 4"
					Session("VESAID") = 1
					Session("UserLoggedIn") = "true" 
					Session("AccessRights") = "Level 4"
					Session("User") = strUserName
					Response.Redirect "VESAViewAllMembers.asp"
			End Select 
		End If
	End If

	objRS.Close
	set objRS = nothing
	CloseConnection()
End Sub

'**************************************************************************************************
' VESA Unit Login 
'**************************************************************************************************
Sub CheckVESAUnitLogin
	Dim objRS
	Dim strSQL
	Dim strVESAUnitID, strPassword

	strVESAUnitID	= Request.Form("VESAUnitID")
	strPassword		= Replace(Request.Form("Password"), "'", "''")

	If strVESAUnitID = "0" Then
		DisplayMessage("<p>Please make sure that you have selected a valid Branch.</p>")
  
	Else
		EstablishConnection()

		'- Look up the user name/password.
		strSQL = "SELECT * FROM VESA_tblUnit" 
		strSQL = strSQL & " WHERE VESAUnitID='" & strVESAUnitID & "'" 
		strSQL = strSQL & " AND Password='" & strPassword & "'"

		Set objRS = Server.CreateObject("ADODB.Recordset")
		Set objRS.ActiveConnection = Conn   
		objRS.Open strSQL, Conn, 1, 1

		On Error Resume Next

		If Err <> 0 Then
			DisplayMessage("<P>There is an error connecting into the database please click back button below and try again. If this is not the frst time you've seen this message, please contact the site administrator.</P>")
		Else 
			'- See if we got anything.
			If CLng(objRS.Fields(0)) < 1 Then
				DisplayMessage("<p>Please check that your password is correct.</p>")
			Else
				Session("VESAID") = 1
				Session("UserLoggedIn") = "true" 
				Session("AccessRights") = "Level 5"
				Session("User") = "Unit"
				Session("VESAUnitID") = strVESAUnitID
				Response.Redirect "VESAMain.asp"
			End If
		End If

		objRS.Close
		Set objRS = nothing
		CloseConnection()
	End If 
End Sub


Sub DisplayMessage(strMessage) %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title>VESA Members Database : Incorrect Password</title>
	<meta name="keywords" content="" />
	<meta name="VESA Members Database" content="" />
	<link rel="stylesheet" href="css/default.css" type="text/css" media="screen" />
	<link rel="stylesheet" href="css/buttons.css">
	<link rel="stylesheet" href="css/forms.css">
	<link rel="stylesheet" href="css/base.css">
	<script language="JavaScript">
	<!--
	<% If LoginType = "AdminLogin" Then %>
	   function goBack() {
		  document.location.href = "AdminLogin.asp";
	   }
	<% Else %>
	   function goBack() {
		  document.location.href = "VESAUnitLogin.asp";
	   }
	<% End if %>
	//-->
	</script>
</head>
<body>
<div id="wrapper">
	<% If Session("UserLoggedIn") <> "true" Then %>
		<div id="menu">
			<ul id="main">
				<li><a href="AdminLogin.asp">Home</a></li>
				<li><a href="VESAUnitLogin.asp">Unit Login</a></li>
				<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
				<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
				<li class="current_page_item"><a href="contactUs.html">Contact Us</a></li>
			</ul>
		</div>

	<% Else %>
		<div id="menu">
			<ul id="main">
				<li><a href="VESAMain.asp">Home</a></li>
				<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
				<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
				<li class="current_page_item"><a href="contactUs.html">Contact Us</a></li>
			</ul>
		</div>
	<% End If %>
	<!-- start header -->
	
	<div id="header">
		<div id="logo">
			<h1><a href="#"><span></span></a></h1>
			<p></p>
		</div>
	</div>
	<!-- end header -->
	
	<!-- start page -->
	<div id="page">
		<div id="sidebar1" class="sidebar">
			<ul>
				<li> <br />
				<img src="images/Phoenix-Logo.jpg" width="220" height="218" alt="" />
				</li>
			</ul>
		</div>
    
		<!-- start content -->
		<div id="content">
			<div class="post">
				<h1 class="title"><a href="#">Error!</a></h1>
				<p class="byline">We've encountered an error when you tried logging into the database.</p>
				<div class="entry">
					<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>
						<p><%=strMessage%></p>
						<p>
						If you're having difficulties logging in or you may have forgotten your password, please email the VESA Database Administrator on <b><a href="mailto:eagnes@cwaustral.com.au">eagnes@cwaustral.com.au</a>. 
						</p>
						</td>
					</tr>

					<tr><td><img src="images/spacer.gif" width="1" height="10" border="0"></td></tr>
					
					<tr>
						<td><button type="button" class="pure-button" onClick="goBack()">Back</button></td>
					</tr>

					<tr><td><img src="images/spacer.gif" width="1" height="20" border="0"></td></tr>

					<tr>
						<td>
						<b>HELP:</b> The password you entered must be case-sensitive. This means, for example, that mypwd (all lowercase) is different from MyPwD (mixed case) or MYPWD (all uppercase). If you do not remember the exact way in which you specified your password, we can email you your password. Please email the VESA Database Administrator on <b><a href="mailto:eagnes@cwaustral.com.au">eagnes@cwaustral.com.au</a>.
						</td>
					</tr>
					</table>
				</div>
			</div>
		</div>
		<!-- end content -->
    
		<div style="clear: both;">&nbsp;</div>
	</div>
	<!-- end page -->
</div>
<div id="footer">
  <p class="copyright">&copy;&nbsp;&nbsp;2008 - <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
</div>
</body>
</html>
<% End Sub %>
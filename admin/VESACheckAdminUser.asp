<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/database/vesa/Members/include/include.asp"-->
<!--#include virtual="/database/vesa/Members/include/functions.asp"-->

<%
Response.Expires = -1000 ' Makes the browser not cache this page
Response.Buffer = True ' Buffers the content so our Response.Redirect will work

If Session("UserLoggedIn") <> "true" Then
   Response.Redirect "../AdminLogin.asp"

Else   
	Call checkAdmininistrationUser()
End If 

Sub checkAdmininistrationUser() %>
<!DOCtype html PUBLIC "-//W3C//Dtd XHTML 1.0 Strict//EN" "http://www.w3.org/tr/xhtml1/Dtd/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>VESA Members Database : Edit Administration User</title>
<meta name="keywords" content="" />
<meta name="VESA Members Database" content="" />
<link href="../css/default.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript">
<!--
function stopSubmit() {
   return false;
}
//-->
</script>
</head>
<body onload="doSubmit()">
<div id="wrapper">
	<div id="menu">
		<ul id="main">
			<li><a href="../VESAMain.asp">Home</a></li>
			<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
			<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
			<li><a href="contactUs.html">Contact Us</a></li>
		</ul>
	</div>
	
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
				<li> 
					<h2>VESA Members</h2>
					<ul>
						<li><a href="../VESAMain.asp">Home</a></li>
						<li><a href="../VESASearch.asp">Search for Member</a></li>
						<li><a href="../VESAViewAllMembers.asp">View All Members</a></li>
						<li><a href="../VESAAddNewMember.asp">Add a New Member</a></li>
						<li><a href="../VESADeleteMember.asp">Delete a Member</a></li>
						<li><a href="../VESAViewHistory.asp">View History</a></li>
					</ul>
				</li>

				<li> 
					<h2>Admin Members</h2>
					<ul>
						<li><a href="VESAViewInactiveMembers.asp">View Inactive Members</a></li>
						<li><a href="ViewAllVESAUnits.asp">View All VESA Units</a></li>
						<li><a href="AddNewVESAUnit.asp">Add a New VESA Unit</a></li>
						<li><a href="DeleteVESAUnit.asp">Delete a VESA Unit</a></li>
						<li><a href="VESAViewAllAdminUsers.asp">View All Admin Users</a></li>
						<li><a href="VESAAddAdminUser.asp">Add an Admin User</a></li>
						<li><a href="VESADeleteAdminUser.asp">Delete an Admin User</a></li>
						<li><a href="../AdminLogin.asp">Log Out</a></li>
					</ul>
				</li>
			</ul>
		</div>
		
		<!-- start content -->
		<div id="content">
			<div class="post">
				<div class="entry">
					<%
						'***** Start user input validation *************************************************
						Session("Surname")			= Request("Surname")
						Session("Firstname")		= Request("Firstname")
						Session("EmailAddress")		= Request("EmailAddress")
						Session("Username")			= Request("Username")
						Session("AccessID")			= Request("AccessID")
						   
						'- Here is one way of checking for an empty text box
						'- Surname -----------------------------------------------------------------------
						If Not Len(Request("Surname")) > 0 Then 
							Session("badSurname") = "T" 
							Session("Errors") = Session("Errors") + 1 
						End If

						If ereg(Request("Surname"), "[^a-zA-Z\s]", true) = true Then
							Session("badSurname") = "T1" 
							Session("Errors") = Session("Errors") + 1 
						End If
						'-----------------------------------------------------------------------------------

						'- Firstname -----------------------------------------------------------------------
						If Not Len(Request("Firstname")) > 0 Then 
							Session("badFirstname") = "T" 
							Session("Errors") = Session("Errors") + 1 
						End If

						If ereg(Request("Firstname"), "[^a-zA-Z\s]", True) = True Then
							Session("badFirstname") = "T1" 
							Session("Errors") = Session("Errors") + 1 
						End If
						'-----------------------------------------------------------------------------------

						'- Email ---------------------------------------------------------------------------
						If Not Len(Request("EmailAddress")) > 0 Then 
							Session("badEmailAddress") = "T" 
							Session("Errors") = Session("Errors") + 1 
						End If

						If CheckIsEmail(Request("EmailAddress")) = False Then
							Session("badEmailAddress") = "T1" 
							Session("Errors") = Session("Errors") + 1 
						End If
						'-----------------------------------------------------------------------------------

						'- Username ------------------------------------------------------------------------
						If Not Len(Request("Username")) > 0 Then 
							Session("badUsername") = "T" 
							Session("Errors") = Session("Errors") + 1 
						End If

						If ereg(Request("Username"), "[^a-zA-Z\s]", True) = True Then
							Session("badUsername") = "T1" 
							Session("Errors") = Session("Errors") + 1 
						End If
						'-----------------------------------------------------------------------------------

						'- Access Rights ------------------------------------------------------------------
						If Request("AccessID")= "Please Choose" Then 
							Session("badAccessID") = "T" 
							Session("Errors") = Session("Errors") + 1 
						End If
						'-----------------------------------------------------------------------------------

						'-----------------------------------------------------------------------------------
						If Session("Errors") > 0 Then 
							'- there were errors, so send back to form 
							Response.Redirect "VESAEditAdminUser.asp" 
						Else
							Call CheckUpdatedAdministrationUser() 
						End If 
						'----------------------------------------------------------------------------------
						'***** End of user input validation for Administration User ***********************
						%>     
					</div>
				</div>
			</div>
			<!-- end content -->
			
			<div style="clear: both;">&nbsp;</div>
		</div>
		<!-- end page -->
	</div>
	<div id="footer">
		<p class="copyright">&copy;&nbsp;&nbsp;2008 -  <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
	</div>
</body>
</html>
<% End Sub

'--------------------------------------------------------------------------------------------------
'***** Display the Header according to the action *****		
'--------------------------------------------------------------------------------------------------
Sub DisplayHeader(error_msg, error_flag, username)
	Select Case error_flag 
		Case False 
			Response.Write "<h1 class=""title""><a href=""#"">Success!</a></h1>" & vbCrLf
			Response.Write "<p class=""byline""><b>You have successfully updated the " & username & "'s details.</b></p>"
		Case True 	
			Response.Write "<h1 class=""title""><a href=""#"">Error!</a></h1>" & vbCrLf
			Response.Write "<p class=""byline""><b>" & error_msg & "</b></p>"
	End Select  
End Sub
'--------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine DISPLAY SUCCESS MESSAGE *****		
'----------------------------------------------------------------------------------
Sub DisplaySuccessMessage(strMessage)
	Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""0"">"
	Response.Write "<tr>"
	Response.Write "<td>"
	Response.Write "<p><b>" & strMessage & "</b></p>"
	Response.Write "<tr><td><img src=""../../images/pixel.gif"" width=""10"" height=""30"" border=""0""></td></tr>"
	Response.Write "</table>"
	End Sub
'----------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
'***** Subroutine DISPLAY ERROR MESSAGE *****		
'--------------------------------------------------------------------------------------------------
Sub DisplayErrorMessage(strMessage)
	Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""0"">"
    Response.Write "<tr>"
    Response.Write "<td>"
	Response.Write "<p><b>" & strMessage & "</b></p>"
	Response.Write "<tr><td><img src=""../../images/pixel.gif"" width=""10"" height=""30"" border=""0""></td></tr>"
	Response.Write "</table>"
End Sub
'--------------------------------------------------------------------------------------------------

Sub CheckUpdatedAdministrationUser()
	Dim rs
	Dim strSQL
	Dim blnPassThrough

	Dim objValueActionType, objValueOldPassword, objValueNewPassword1, objValueNewPassword2, objValueAdminUserUserID, objValueAdminUserSurname, objValueAdminUserFirstName, objValueAdminUserEmailAddress, objValueAdminUserUsername, objValueAdminUserAccessID
	Dim error_flag, error_msg

	objValueOldPassword				= Request.Form("OldPassword")
	objValueNewPassword1			= Request.Form("NewPassword1")
	objValueNewPassword2			= Request.Form("NewPassword2")

	objValueActionType				= Request.Form("ActionType")
	objValueAdminUserUserID			= Request.Form("UserID")
	objValueAdminUserSurname		= Request.Form("Surname")
	objValueAdminUserFirstName		= Request.Form("FirstName")
	objValueAdminUserEmailAddress	= Request.Form("EmailAddress")
	objValueAdminUserUsername		= Request.Form("Username")
	objValueAdminUserAccessID		= Request.Form("AccessID")
%>
	<form name="EditAdminUser" id="EditAdminUser" action="../VESASave.asp" method="post" onSubmit="return stopSubmit()">
	<input type="hidden" name="ActionType" value="<%=objValueActionType%>">
	<input type="hidden" name="AdminUserUserID" value="<%=objValueAdminUserUserID%>">
	<input type="hidden" name="AdminUserSurname" value="<%=objValueAdminUserSurname%>">
	<input type="hidden" name="AdminUserFirstName" value="<%=objValueAdminUserFirstName%>">
	<input type="hidden" name="AdminUserEmailAddress" value="<%=objValueAdminUserEmailAddress%>">
	<input type="hidden" name="AdminUserUsername" value="<%=objValueAdminUserUsername%>">
	<input type="hidden" name="AdminUserPassword" value="<%=objValueNewPassword2%>">
	<input type="hidden" name="AdminUserAccessID" value="<%=objValueAdminUserAccessID%>">
	
	<table cellspacing="0" cellpadding="0">
	<tr>
	<td>
	<%

	error_flag = "False" '' Setting error flag
	error_msg = "" '' Error message

	Dim RExp : Set RExp = new RegExp
		with RExp
		.Pattern = "^[a-zA-Z0-9]{3,8}$"
		.IgnoreCase = True
		.Global = True
	End with

	EstablishConnection()

	strSQL = "SELECT * FROM MembersDB_tblUsers U" 
	strSQL = strSQL & " INNER JOIN MembersDB_tblUserRights UR on U.UserID = UR.UserID "
	strSQL = strSQL & " INNER JOIN MembersDB_tblUserAccess UA on UR.AccessID = UA.AccessID"
	strSQL = strSQL & " WHERE U.UserID='" & objValueAdminUserUserID & "'"
	strSQL = strSQL & " AND Password='" & objValueOldPassword & "'"
	strSQL = strSQL & " ORDER BY U.UserID ASC"

	Set rsResult = Server.CreateObject("ADODB.Recordset")
	rsResult.Open strSQL, Conn

	If rsResult.EOF Then
		blnPassThrough = False
		Call DisplayHeader("Sorry, that password does not exist. Please click back on and enter a different password.", true, Request.Form("Username"))
		DisplayErrorMessage("<p><b>Sorry, that password does not exist. Please click back on and enter a different password.</b></p>")

	Else
		If (Not (RExp.test(objValueNewPassword1) And RExp.test(objValueNewPassword2))) Then
			error_flag = True
			error_msg = "&nbsp;&nbsp;<font color=""#ff0000"">** (Please enter valid data only.) **</font>"
		End If

		If (objValueNewPassword1 <> objValueNewPassword2) Then
			error_flag = True
			error_msg = error_msg + "&nbsp;&nbsp;<font color=""#ff0000"">** (Passwords are not matching) **</font>"
		End If

		If error_flag = False Then
			Call DisplaySuccessMessage("Are you sure you want to change the password?")
			blnPassThrough = True
		Else
			Call DisplayHeader("Are you sure you want to change the password?", False, Request.Form("Username"))
			Call DisplayErrorMessage(error_msg)
			blnPassThrough = True
		End If

		'------------------------------------------------------------------------------------------
		rsResult.MoveFirst
		Do Until rsResult.EOF
			rsResult.MoveNext
		Loop
%>
         <span>Below is the information that you wish to add:</span><br /><br />
	     <table cellpadding="0" cellspacing="0">
		 <tr>
			 <td><div align="left"><strong><label for="Name">Name:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			 <td><%=objValueAdminUserFirstName%>&nbsp;<%=objValueAdminUserSurname%></td>
		 </tr>

		 <tr>
			 <td><div align="left"><strong><label for="Email">Email:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			 <td><%=objValueAdminUserEmailAddress%></td>
		 </tr>

		 <tr>
			 <td valign="top" class="blacktext"><div align="left"><strong><label for="Username">Username:</label></strong></div></td>
			 <td><img src="../../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
			 <td><%=objValueAdminUserUsername%></font></td>
		 </tr>

		 <tr>
			 <td valign="top" class="blacktext"><div align="left"><strong><label for="Old Password">Old Password:</label></strong></div></td>
			 <td><img src="../../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
			 <td><%=objValueOldPassword%></font></td>
		 </tr>

		 <tr>
			 <td><div align="left"><strong><label for="Password">New Password:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			 <td>
			 <%
			 If (Not (RExp.test(objValueNewPassword1) And RExp.test(objValueNewPassword2))) Then
				Response.Write "&nbsp;&nbsp;<font color=""#ff0000"">** (Please enter valid data only) **</font>"
			 End If

			 If (objValueNewPassword1 <> objValueNewPassword2) Then
				Response.Write "&nbsp;&nbsp;<font color=""#ff0000"">** (Passwords are not matching) **</font>"
			 End If 
			 %>
			 </td>
		 </tr>

		 <tr>
			 <td><div align="left"><strong><label for="Password">Confirm Password:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			 <td>
			 <%
			 If (Not (RExp.test(objValueNewPassword1) And RExp.test(objValueNewPassword2))) Then
				Response.Write "&nbsp;&nbsp;<font color=""#ff0000"">** (Please enter valid data only) **</font>"
			 End If

			 If (objValueNewPassword1 <> objValueNewPassword2) Then
				Response.Write "&nbsp;&nbsp;<font color=""#ff0000"">** (Passwords are not matching) **</font>"
			 End If 
			 %>
			 </td>
		 </tr>

		 <tr>
			 <td valign="top" class="blacktext"><div align="left"><strong><label for="Access Rights">Access Rights:</label></strong></div></td>
			 <td><img src="../../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
			 <td>Level <%=objValueAdminUserAccessID%></td>
		 </tr>
		 </table>
	  </td>
	  </tr>

	  <tr height="5"><td><img src="../images/spacer.gif" width="1" height="5" alt="" /></td></tr>
		
<%

	  rsResult.Close
	  Set rsResult = Nothing
	  CloseConnection()
		'------------------------------------------------------------------------------------------
	End If
	
	If blnPassThrough = True Then
		Call DisplaySuccessMessage("Please wait...Adding New Administration User - " & Request.Form("Firstname") & "&nbsp;" & Request.Form("Surname"))
	Else
%>
		<tr>
		<td><input type="image" name="back" class="back-btn" src="http://www.roscripts.com/images/btn.gif" alt="back" title="back" onClick="goback()" /></td>
		</tr>
	<% End If %>
  </table>
  </form>

  <script type="text/javascript">
  <!--
  function doSubmit() 
  {
     <% If blnPassThrough = true Then %>
           document.EditAdminUser.submit();
     <% End If %>
  }

  function goback() {
     history.go(-1); 
  }
  //-->
  </script>
<% 
End Sub
	
'----- Regular Expression -------------------------------------------------------------------------   
Function ereg(strOriginalString, strPattern, varIgnoreCase)
	'Function matches pattern, returns true or false
	'varIgnoreCase must be trUE (match is case insensitive) or FALSE (match is case sensitive)
	Dim objRegExp : Set objRegExp = new RegExp
	with objRegExp
		.Pattern = strPattern
		.IgnoreCase = varIgnoreCase
		.Global = true
	End with
	ereg = objRegExp.test(strOriginalString)
	Set objRegExp = Nothing 
 End Function 
'--------------------------------------------------------------------------------------------------

Function kCheckRegExp(vPattern, vStr)
	Dim oRegExp
	Set oRegExp = New RegExp
	oRegExp.Pattern = vPattern
	oRegExp.IgnoreCase = False
	kCheckRegExp = oRegExp.Test(vStr)
	Set oRegExp = Nothing
End Function

Function kLeachRegExp(vStr1, vPattern, vStr2)
	Dim oRegExp
	Set oRegExp = New RegExp
	oRegExp.Pattern = vPattern
	oRegExp.IgnoreCase = True
	oRegExp.Global = True
	kLeachRegExp = oRegExp.Replace(vStr1, vStr2)
	Set oRegExp = Nothing
End Function

'--------------------------------------------------------------------------------------------------
'***** CHECK IS VALID EMAIL ADDRESS FORMAT *****		
'--------------------------------------------------------------------------------------------------
Function CheckIsEmail(strEmailAddress)
   Dim regEx, retVal
   Set regEx = New RegExp
   '# note the regex below must be on one single line
   '# it is shown wrapped on this web page
   regEx.Pattern = "^(([^<>()[\]\\.,;:\s@""]+(\.[^<>()[\]\\.,;:\s@""]+)*)|("".+""))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$"
   regEx.IgnoreCase = true
   blnReturnValue = regEx.Test(strEmailAddress)
   CheckIsEmail = blnReturnValue
End Function
%>
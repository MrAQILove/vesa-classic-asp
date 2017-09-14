<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/database/vesa/Members/include/include.asp"-->
<!--#include virtual="/database/vesa/Members/include/functions.asp"--> 

<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- buffers the content so our Response.Redirect will work
Response.buffer = true 

Response.AddHeader "Pragma", "no-store"
Response.CacheControl = "no-store"

If Session("UserLoggedIn") <> "true" Then
   Response.Redirect "../AdminLogin.asp" 

Else   
   Call checkAdmin()
End If

Sub checkAdmin()
%>
	<!DOCtype html PUbLIC "-//W3C//Dtd XHTML 1.0 Strict//EN" "http://www.w3.org/tr/xhtml1/Dtd/xhtml1-strict.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<% If objValueActiontype = "AddAdministrationUser" Then %>
			<title>VESA Members Database : Checking Added New Admin Member</title>
		<% Else %>
			<title>VESA Members Database : Checking Added New VESA Unit</title>
		<% End If %>
		<meta name="keywords" content="" />
		<meta name="VESA Members Database" content="" />
		<link href="../css/default.css" rel="stylesheet" type="text/css" media="screen" />
		<link rel="stylesheet" href="../css/buttons.css">
		<link rel="stylesheet" href="../css/forms.css">
		<link rel="stylesheet" href="../css/base.css">
		<link rel="stylesheet" href="../css/grids.css">
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
					<li><a href="http://www.cwmedia.com.au/">Countrywide Media</a></li>
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
						<%
							If Request.Form("Actiontype") = "AddAdministrationUser" Then
								Response.Write "<h1 class=""title""><a href=""#"">Checking Added New Admin User</a></h1>"
								Response.Write "<p class=""byline""><b>The following Member is already in the database:</b></p>"
							Else
								Response.Write "<h1 class=""title""><a href=""#"">Checking a newly added VESA Unit</a></h1>"
								Response.Write "<p class=""byline""><b>The following VESA Unit is already in the database:</b></p>"
							End If 
							
							Response.Write "<div class=""entry"">"
			
							If Request.Form("Actiontype") = "AddAdministrationUser" Then
								'***** Start user input validation ********************************
								Session("Surname")		= Request("Surname")
								Session("Firstname")	= Request("Firstname")
								Session("EmailAddress")	= Request("EmailAddress")
								Session("Username")		= Request("Username")
								Session("Password")		= Request("Password")
								Session("Password")		= Request("Password")
								Session("UserAccessRights")	= Request("UserAccessRights")
				   
								'- Here is one way of checking for an empty text box
								'- Surname --------------------------------------------------------
								If Not Len(Request("Surname")) > 0 Then 
									Session("badSurname") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If

								If ereg(Request("Surname"), "[^a-zA-Z\s]", true) = true Then
									Session("badSurname") = "T1" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'- Firstname ------------------------------------------------------
								If Not Len(Request("Firstname")) > 0 Then 
									Session("badFirstname") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If

								If ereg(Request("Firstname"), "[^a-zA-Z\s]", True) = True Then
									Session("badFirstname") = "T1" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'- Email ----------------------------------------------------------
								If Not Len(Request("EmailAddress")) > 0 Then 
									Session("badEmailAddress") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If

								If CheckIsEmail(Request("EmailAddress")) = False Then
									Session("badEmailAddress") = "T1" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'- Username -------------------------------------------------------
								If Not Len(Request("Username")) > 0 Then 
									Session("badUsername") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If

								If ereg(Request("Username"), "[^a-zA-Z\s]", True) = True Then
									Session("badUsername") = "T1" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'- Password -------------------------------------------------------
								If Not Len(Request("Password")) > 0 Then 
									Session("badPassword") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If

								If ereg(Request("Password"), "[^0-9a-zA-Z\s]", True) = True Then
									Session("badPassword") = "T1" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'- Access Rights --------------------------------------------------
								If Request("UserAccessRights")= "Please Choose" Then 
									Session("badUserAccessRights") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'------------------------------------------------------------------
								If Session("Errors") > 0 Then 
									'- there were errors, so send back to form 
									Response.Redirect "VESAAddAdminUser.asp" 
								Else
									Call CheckAddedNewAdmin() 
								End If 
								'------------------------------------------------------------------
								'***** End of user input validation for Administration User *******
				
							Else
								'***** Start user input validation for VESA Unit/Distribution *****
								Session("VESAUnitName")			= Request("VESAUnitName")
								Session("VESAUnitPassword")		= Request("VESAUnitPassword")
								Session("UnitEmailAddress")		= Request("UnitEmailAddress")
								Session("SESRegionID")			= Request("SESRegionID")
							   
								'- Here is one way of checking for an empty text box
								'- VESA Unit/Distribution Name ------------------------------------
								If ereg(Request("VESAUnitName"), "[^a-zA-Z\s]", true) = true Then
									Session("badVESAUnitName") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'- VESA Unit/Distribution Password --------------------------------
								If Not Len(Request("VESAUnitPassword")) > 0 Then 
									Session("badVESAUnitPassword") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'- SES Region -----------------------------------------------------
								If Request("SESRegionID")= "0" Then 
									Session("badSESRegionID") = "T" 
									Session("Errors") = Session("Errors") + 1 
								End If
								'------------------------------------------------------------------

								'------------------------------------------------------------------
								If Session("Errors") > 0 Then 
									'- there were errors, so send back to form 
									Response.Redirect "AddNewVESAUnit.asp" 
								Else
									Call CheckAddedUnit() 
								End If 
								'------------------------------------------------------------------
								'***** End of user input validation *******************************
							End If 
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

Sub DisplayMessage()
   Response.Write "<tr>"
   Response.Write "<td>Please wait..."
   
   If Request.Form("ActionType") = "AddAdministrationUser" Then
      Response.Write "Adding New Administration User - " & Request.Form("Firstname") & "&nbsp;" & Request.Form("Surname")
   Else
      Response.Write "Adding New VESA Unit - " & Request.Form("VESAUnitName")
   End If 

   Response.Write "</td>"
   Response.Write "</tr>"
End Sub

Sub CheckAddedNewAdmin()
	Dim objValueActionType, objValueAdminUserDatabaseName, objValueAdminUserSurname, objValueAdminUserFirstName, objValueAdminUserEmailAddress, objValueAdminUserUsername, objValueAdminUserPassword, objValueAdminUserAccessID
	
	Dim strPassword

	objValueActionType				= Request.Form("ActionType")
	objValueAdminUserDatabaseName	= Request.Form("DatabaseName")
	objValueAdminUserSurname		= Request.Form("Surname")
	objValueAdminUserFirstName		= Request.Form("FirstName")
	objValueAdminUserEmailAddress	= Request.Form("EmailAddress")
	objValueAdminUserUsername		= Request.Form("Username")
	objValueAdminUserPassword		= Request.Form("Password")
	objValueAdminUserAccessID		= Request.Form("UserAccessRights")
%>
	<form name="CheckAdminUserForm" id="CheckAdminUserForm" action="../VESASave.asp" method="post" onSubmit="return stopSubmit()">
	<input type="hidden" name="ActionType" value="<%=objValueActionType%>">
	<input type="hidden" name="AdminUserDatabaseName" value="<%=objValueDatabasename%>">
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
	Dim strSQL
	Dim rsResult
	Dim blnPassThrough

	EstablishConnection()

	strSQL = "SELECT * FROM CWM_tblUsers U" 
	strSQL = strSQL & " INNER JOIN CWM_tblUserRights UR on U.UserID = UR.UserID "
	strSQL = strSQL & " INNER JOIN CWM_tblUserAccess UA on UR.AccessID = UA.AccessID"
	strSQL = strSQL & " WHERE U.UserName='" & objValueUsername & "'" 
	strSQL = strSQL & " AND U.Password='" & objValuePassword & "'"
	strSQL = strSQL & " AND U.DatabaseName='VESA_tblMembers'"
   
	Set rsResult = Server.CreateObject("ADODb.Recordset")
	rsResult.Open strSQL, Conn

	If not rsResult.EOF Then
		blnPassThrough = False

		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			   
		rsResult.MoveFirst
		Do Until rsResult.EOF 
%>
			<tr>
			<td><div align="left"><strong><label for="User ID">User ID:</label></strong></div></td>
			<td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			<td colspan="3"><div align="left"><font color="#0000a0"><strong><%=rsResult.Fields("UserID")%></strong></div></td>
			</tr>

			<tr>
			<td><div align="left"><strong><label for="Name">Name:</label></strong></div></td>
			<td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			<td><%=rsResult.Fields("FirstName")%>&nbsp;<%=rsResult.Fields("Surname")%></td>
			</tr>

			<tr>
			<td><div align="left"><strong><label for="Email">Email:</label></strong></div></td>
			<td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			<td><%=rsResult.Fields("EmailAddress")%></td>
			</tr>

			<tr>
			<td valign="top" class="blacktext"><div align="left"><strong><label for="Username">Username:</label></strong></div></td>
			<td><img src="../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
			<td><%=rsResult.Fields("Username")%> &nbsp;&nbsp;<font color="#ff0000">** (Same USERNAME) **</font></td>
			</tr>

			<tr>
			<td><div align="left"><strong><label for="Password">Password:</label></strong></div></td>
			<td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			<td>
			<%
		    sPassword = rsResult.Fields("Password")
			strPassword = kLeachRegExp("" & rsResult.Fields("Password")& "", "[^()?<>.*?]", "*") 

			Response.Write strPassword
			%> &nbsp;&nbsp;<font color="#ff0000">** (Same PASSWORD) **</font>
			</td>
			</tr>

			<tr>
			<td valign="top" class="blacktext"><div align="left"><strong><label for="Access Rights">Access Rights:</label></strong></div></td>
			<td><img src="../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
			<td><%=rsResult.Fields("AccessRights")%></td>
			</tr>
<%
			rsResult.MoveNext
		Loop
%>
        </table>
      </td>
      </tr>

      <tr height="5"><td><img src="../images/spacer.gif" width="1" height="20" alt="" /></td></tr>

	  <tr>
	  <td>
	     <span>Below is the information that you wish to add:</span><br /><br />
	     <table cellpadding="0" cellspacing="0">
		 <tr>
			 <td><div align="left"><strong><label for="Name">Name:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			 <td><%=objValueFirstname%>&nbsp;<%=objValueSurname%></td>
		 </tr>

		 <tr>
			 <td><div align="left"><strong><label for="Email">Email:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			 <td><%=objValueEmail%></td>
		 </tr>

		 <tr>
			 <td valign="top" class="blacktext"><div align="left"><strong><label for="Username">Username:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
			 <td><%=objValueUsername%> &nbsp;&nbsp;<font color="#ff0000">** (Same USERNAME) **</font></td>
		 </tr>

		 <tr>
			 <td><div align="left"><strong><label for="Password">Password:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
			 <td>
			 <%
				strPassword = kLeachRegExp("" & objValuePassword & "", "[^()?<>.*?]", "*")
				Response.Write strPassword
			 %> &nbsp;&nbsp;<font color="#ff0000">** (Same PASSWORD) **</font>
			 </td>
		 </tr>

		 <tr>
			 <td valign="top" class="blacktext"><div align="left"><strong><label for="Access Rights">Access Rights:</label></strong></div></td>
			 <td><img src="../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
			 <td>Level <%=objValueUserAccessRights%></td>
		 </tr>
		 </table>
	  </td>
	  </tr>

	  <tr height="5"><td colspan="3"><img src="../images/spacer.gif" width="1" height="5" alt="" /></td></tr>

<%
     rsResult.Close
     Set rsResult = nothing
     CloseConnection()

  Else
     blnPassThrough = true
  End If

  If blnPassThrough = true Then
     Call DisplayMessage()
  Else
%>
	<tr>
		<td><button type="button" class="pure-button" onClick="goBack()">Back to Main</button></td> 
	</tr>
  <% End If %>
  </table>
  </form>

  <script type="text/javascript">
  <!--
  function doSubmit() 
  {
     <% If blnPassThrough = true Then %>
           document.CheckAdminUserForm.submit();
     <% End If %>
  }

  function goBack() {
     history.go(-1); 
  }
  //-->
  </script>
<% 
End Sub


Sub CheckAddedUnit()
	Dim objValueActionType, objValueVESAUnitName, objValueVESAUnitPassword, objValueUnitEmailAddress, objValueUnitSESRegionID, objValueIsUnitSES 

	objValueActionType			= Request.Form("ActionType")
	objValueVESAUnitName		= Request.Form("VESAUnitName")
	objValueVESAUnitPassword	= Request.Form("VESAUnitPassword")
	objValueUnitEmailAddress	= Request.Form("UnitEmailAddress")
	objValueUnitSESRegionID		= Request.Form("SESRegionID")
	objValueIsUnitSES			= Request.Form("IsUnitSES")

%>
	<form name="CheckVESAUnitForm" id="CheckVESAUnitForm" action="../VESASave.asp" method="post" onSubmit="return stopSubmit()">
	<input type="hidden" name="ActionType" value="<%=objValueActionType%>">
	<input type="hidden" name="VESAUnit" value="<%=objValueVESAUnitName%>">
	<input type="hidden" name="VESAUnitPassword" value="<%=objValueVESAUnitPassword%>">
	<input type="hidden" name="UnitEmailAddress" value="<%=objValueUnitEmailAddress%>">
	<input type="hidden" name="UnitSESRegionID" value="<%=objValueUnitSESRegionID%>">
	<input type="hidden" name="IsUnitSES" value="<%=objValueIsUnitSES%>">
   
	<table cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<%
		Dim strSQL
		Dim rsResult
		Dim blnPassThrough

		EstablishConnection()

		strSQL = "SELECT * FROM VESA_tblUnit U"
		strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
		strSQL = strSQL & " WHERE VESAUnit='" & objValueVESAUnitName & "'" 
		'strSQL = strSQL & " AND Password='" & objValueVESAUnitPassword & "'"
		'strSQL = strSQL & " AND IsActive = '1'"
	   
		Set rsResult = Server.CreateObject("ADODb.Recordset")
		rsResult.Open strSQL, Conn

		If not rsResult.EOF Then
			blnPassThrough = False
		%>
		<span class="formTitle">The following VESA Unit is already in the database. Please enter a different name.</span><br /><br />
			<table border="0" cellpadding="0" cellspacing="0" width="350">

			<%		   
			rsResult.MoveFirst
			Do Until rsResult.EOF 
			%>
				<tr>
					<td><div align="left"><strong><label for="VESA Unit ID">VESA Unit ID:</label></strong></div></td>
					<td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
					<td><div align="left"><font color="#0000a0"><strong><%=rsResult.Fields("VESAUnitID")%></strong></div></td>
				</tr>

				<tr>
					<td><div align="left"><strong><label for="VESA Unit">VESA Unit:</label></strong></div></td>
					<td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
					<td><%=rsResult.Fields("VESAUnit")%></td>
				</tr>

				<!--<tr>
					<td><div align="left"><strong><label for="Password">Password:</label></strong></div></td>
					<td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
					<td><%=rsResult.Fields("Password")%>&nbsp;&nbsp;<font color="#ff0000">** (Same PASSWORD) **</font></td>
				</tr>-->

				<tr>
					<td><div align="left"><strong><label for="Email Address">Email Address:</label></strong></div></td>
					<td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
					<td><div align="left"><font color="#0000a0"><strong><%=rsResult.Fields("EmailAddress")%></strong></div></td>
				</tr>

				<tr>
					<td valign="TOP"><div align="left"><strong><label for="SES Region">SES Region:</label></strong></div></td>
					<td><img src="../images/spacer.gif" width="10" height="1" alt=""></td>
					<td><%Call showdbSelectedValue(Conn, rsVESAUnitID, "VESA_tblSESRegion", "SESRegion", "SESRegionID", "" & rsResult.Fields("SESRegionID") & "")%></td>
				</tr>

				<tr height="10"><td colspan="3"><img src="../images/spacer.gif" width="1" height="10" border="0"></td></tr>
			<%
				rsResult.MoveNext
			Loop
			%>
			</table>
		</td>
	</tr>
	<%
    rsResult.Close
    Set rsResult = nothing
    CloseConnection()
	   
	Else
		blnPassThrough = True
	End If
 
	If blnPassThrough = True Then
		Response.Write	"<tr>" &_
						"<td>Please wait - Adding Member</td>" &_
						"</tr>"
	Else
	%>
		<tr height="5"><td colspan="3"><img src="../images/spacer.gif" width="1" height="5" alt="" /></td></tr>

		<tr>
			<td><button type="button" class="pure-button" onClick="goBack()">Go back</button></td>  
		</tr>
   <% End If %>
   </table>
   </form>
	<script type="text/javascript">
	<!--
	function doSubmit() 
	{
		<% If blnPassThrough = true Then %>
			document.CheckVESAUnitForm.submit();
		<% End If %>
	}

	function goBack() {
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


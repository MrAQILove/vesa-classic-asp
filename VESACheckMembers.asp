<%@LANGUAGE="VbSCRIPT"%>
<!--#INCLUDE FILE="include/include.asp"-->
<!--#INCLUDE FILE="include/functions.asp"-->

<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- buffers the content so our Response.Redirect will work
Response.buffer = true 

Response.AddHeader "Pragma", "no-store"
Response.CacheControl = "no-store"

If Session("UserLoggedIn") <> "true" Then
   If Session("AccessRights") = "Level 1" Then
      Response.Redirect "AdminLogin.asp"
   Else
      Response.Redirect "VESAUnitLogin.asp"
   End If 

Else
   Call CheckMembers()
End If

Sub CheckMembers()
%>
<!DOCtype html PUBLIC "-//W3C//Dtd XHTML 1.0 Strict//EN" "http://www.w3.org/tr/xhtml1/Dtd/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>VESA Members Database : Checking Added New Member</title>
<meta name="keywords" content="" />
<meta name="VESA Members Database" content="" />
<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript">
<!--
function stopSubmit() {
   return false;
}

// view members 
function viewMembers() {
   if (<%=Session("VESAID")%> == 1) {
      document.viewMembersForm.submit();
   }
}

// add member 
function addMember() {
   if (<%=Session("VESAID")%> == 1) {
      document.addMemberForm.submit();
   }
}

// delete member 
function deleteMember() {
   if (<%=Session("VESAID")%> == 1) {
      document.deleteMemberForm.submit();
   }
}
//-->
</script>
</head>
<body onload="doSubmit()">
<!--<body>-->
<div id="wrapper">
	<div id="menu">
		<ul id="main">
			<li><a href="VESAMain.asp">Home</a></li>
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
  <% Select Case Session("AccessRights")
	      Case "Level 1" 
	%>
    <div id="sidebar1" class="sidebar">
      <ul>
        <li> 
		  <h2>VESA Members</h2>
          <ul>
		    <li><a href="VESAMain.asp">Home</a></li>
			<li><a href="VESASearch.asp">Search for Member</a></li>
			<li><a href="VESAViewAllMembers.asp">View All Members</a></li>
            <li><a href="VESAAddNewMember.asp">Add a New Member</a></li>		
			<% If Request.Form("ActionType") = "Add" Then %>
				<li class="hover"><a href="#">Checking Added New Member</a></li>
			<% Else %>
				<li class="hover"><a href="#">Checking Edited Member</a></li>
			<% End If %>
			<li><a href="VESADeleteMember.asp">Delete a Member</a></li>
            <li><a href="VESAViewHistory.asp">View History</a></li>
          </ul>
        </li>

		<li> 
		  <h2>Admin Members</h2>
          <ul>
		    <li><a href="admin/VESAViewInactiveMembers.asp">View Inactive Members</a></li>
		    <li><a href="admin/ViewAllVESAUnits.asp">View All VESA Units</a></li>
			<li><a href="admin/AddNewVESAUnit.asp">Add a New VESA Unit</a></li>
			<li><a href="admin/DeleteVESAUnit.asp">Delete a VESA Unit</a></li>
		    <li><a href="admin/VESAViewAllAdminUsers.asp">View All Admin Users</a></li>
            <li><a href="admin/VESAAddAdminUser.asp">Add an Admin User</a></li>
            <li><a href="admin/VESADeleteAdminUser.asp">Delete an Admin User</a></li>
            <li><a href="AdminLogin.asp">Log Out</a></li>
          </ul>
        </li>
      </ul>
    </div>
	<% 
	      Case Else
	%>
	         <div id="sidebar1" class="sidebar">
      <ul>
        <li> 
		  <h2>
          <% 
		  EstablishConnection()

		  Dim strVESAUnit
		  Call showVESAUnit(Conn, rs, "VESA_tblUnit", "" & Session("VESAUnitID") & "", strVESAUnit)
		  
		  Response.Write strVESAUnit & " Members"
		  
		  CloseConnection()
		  %> 
		  </h2>
          <ul>
			<ul>
				<% Call displaySelectedMenu(Request.ServerVariables("SCRIPT_NAME")) %>
				<li><a href="VESAUnitLogin.asp">Log Out</a></li>
          </ul>
        </li>
      </ul>
    </div>
	<%
	   End Select
	%>
    <!-- start content -->
    <div id="content">
      <div class="post">
        <% Call DisplayHeader() %>
        <div class="entry">
        <%
			   '***** Start user input validation *************************************************
			   If Request.Form("ActionType") = "Add" Then
					Session("Surname_Organization")		= Request("Surname_Organization")
					Session("FirstName")				= Request("FirstName")
					Session("Address")					= Request("Address") 
					Session("Suburb")					= Request("Suburb")
					Session("StateID")					= Request("StateID")
					Session("Postcode")					= Request("Postcode")
					Session("MembershipNumber")			= Request("MembershipNumber")
					Session("MemberEmailAddress")		= Request("MemberEmailAddress")
					Session("VESAUnitID")				= Request("VESAUnitID")

				Else
					Session("RecipientID")				= Request("RecipientID")
					Session("Surname_Organization")		= Request("Surname_Organization")
					Session("FirstName")				= Request("FirstName")
					Session("Address")					= Request("Address") 
					Session("Suburb")					= Request("Suburb")
					Session("StateID")					= Request("StateID")
					Session("Postcode")					= Request("Postcode")
					Session("MembershipNumber")			= Request("MembershipNumber")
					Session("MemberEmailAddress")		= Request("MemberEmailAddress")
					Session("VESAUnitID")				= Request("VESAUnitID")
					Session("PhoenixCopies")			= Request("PhoenixCopies")
					Session("VESAPocketDiary")			= Request("VESAPocketDiary")
					Session("VESAWallCalendar")			= Request("VESAWallCalendar")
					Session("VESAWallCalendar")			= Request("VESAWallCalendar")
					Session("SESRegionID")				= Request("SESRegionID")
					Session("SESRegion")				= Request("SESRegion")
				End If 
			   
			   '- Here is one way of checking for an empty text box
			   '- Surname/Organization ------------------------------------------------------------
			   If ereg(Request("Surname_Organization"), "[^a-zA-Z\/\-\s]", true) = true Then
			      Session("badSurname_Organization") = "T" 
				  Session("Errors") = Session("Errors") + 1 
			   End If
			   '-----------------------------------------------------------------------------------

			   '- First Name ----------------------------------------------------------------------
			   If ereg(Request("FirstName"), "[^a-zA-Z\/\-\s]", true) = true Then
			      Session("badFirstName") = "T" 
				  Session("Errors") = Session("Errors") + 1 
			   End If
			   '-----------------------------------------------------------------------------------
			   
			   '- Address -------------------------------------------------------------------------
			   If ereg(Request("Address"), "[^0-9a-zA-Z\/\-\,\s]", true) = true Then
			      Session("badAddress") = "T" 
				  Session("Errors") = Session("Errors") + 1 
			   End If
			   '-----------------------------------------------------------------------------------

			   '- Suburb --------------------------------------------------------------------------
			   If ereg(Request("Suburb"), "[^a-zA-Z\-\s]", true) = true Then
			      Session("badSuburb") = "T" 
				  Session("Errors") = Session("Errors") + 1 
			   End If
			   '-----------------------------------------------------------------------------------

			   '- State ------------------------------------------------------------------------
			   If Request("StateID")= "" Then 
			      Session("badStateID") = "T" 
				  Session("Errors") = Session("Errors") + 1 
			   End If
			   '--------------------------------------------------------------------------------

			   '- Postcode ---------------------------------------------------------------------
			   If Len(Request("Postcode")) <= 3 Then 
			      Session("badPostcode") = "T" 
				  Session("Errors") = Session("Errors") + 1 
			   End If

			   If ereg(Request("Postcode"), "[^0-9\s]", true) = true Then
			      Session("badPostcode") = "T1" 
				  Session("Errors") = Session("Errors") + 1 
			   End If
			   '-----------------------------------------------------------------------------------

			   '- Membership Number ---------------------------------------------------------------
			   If Session("AccessRights") = "Level 1" Then
			      If ereg(Request("MembershipNumber"), "[^0-9\s]", true) = true Then
			         Session("badMembershipNumber") = "T1" 
				     Session("Errors") = Session("Errors") + 1 
			      End If
			   End If 
			   '-----------------------------------------------------------------------------------

			   '- Email Address -------------------------------------------------------------------
				If Request("MemberEmailAddress") <> "" Then
					If CheckIsEmail(Request("MemberEmailAddress")) = False Then
						Session("badMemberEmailAddress") = "T" 
						Session("Errors") = Session("Errors") + 1 
					End If
				End If 
				'-----------------------------------------------------------------------------------

			   '- VESA Unit -----------------------------------------------------------------------
			   If Request("VESAUnitID")= "Please Choose" Then 
			      Session("badVESAUnitID") = "T" 
				  Session("Errors") = Session("Errors") + 1 
			   End If
			   '-----------------------------------------------------------------------------------
			   
			   '-----------------------------------------------------------------------------------
			   If Session("Errors") > 0 Then 
			      '- there were errors, so send back to form 
				  If Request.Form("ActionType") = "Add" Then
				     Response.Redirect "VESAAddNewMember.asp" 
				  Else
				     Response.Redirect "VESAEdit.asp"
				  End If 
			   Else
			      Call CheckMember() 
			   End If 
			   '-----------------------------------------------------------------------------------
			   '***** End of user input validation ************************************************
			   
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

Sub DisplayHeader()
   If Request.Form("ActionType") = "Update" Then
      Response.Write "<h1 class=""title""><a href=""#"">Checking Edited Member</a></h1>"
      Response.Write "<p class=""byline""><b>The information below is currently on the database:</b></p>"
   
   Else
      Response.Write "<h1 class=""title""><a href=""#"">Checking Added New Member</a></h1>"
      Response.Write "<p class=""byline""><b>The following member is already in the database:</b></p>"
   End If 
End Sub

Sub DisplayMessage()
	Response.Write "<tr>"
	Response.Write "<td><b>Please wait... <br /><br />"

	Select Case Request.Form("ActionType")
		Case "Add"
			If Request.Form("Surname_Organization") <> "" And Request.Form("FirstName") <> "" Then
				Response.Write "Adding Member - " & Request.Form("FirstName") & "&nbsp;" & Request.Form("Surname_Organization")
			Else
				Response.Write "Adding Member - " & Request.Form("Surname_Organization")
			End If
			
		Case "Edit"
			Response.Write "Editing Member - "	
	End Select 
   

   'If Not IsNull(rsResult.Fields(2)) And Not IsNull(rsResult.Fields(3)) Then
	'	Response.Write "Edit Member - " & FirstName & "&nbsp;" & Surname_Organization

	'Else
'		Response.Write "Edit Member - " & Surname_Organization     
 '  End If
   
   Response.Write "</b></td>"
   Response.Write "</tr>"
End Sub

Sub CheckMember()
	Dim objValueRecipientID, objValueActionType, objValueSurname_Organization, objValueFirstName, objValueAddress, objValueSuburb, objValuePostcode, objValueStateID
	Dim objValueMembershipNumber, objValueMemberEmailAddress, objValuePhoenixCopies, objValueVESAPocketDiary, objValueVESAWallCalendar, objValueSESRegionID   

	objValueRecipientID				= Request.Form("RecipientID")
	objValueActionType				= Request.Form("ActionType")
	objValueSurname_Organization	= Request.Form("Surname_Organization")
	objValueFirstName				= Request.Form("FirstName")
	objValueAddress					= Request.Form("Address")
	objValueSuburb					= Request.Form("Suburb")
	objValuePostcode				= Request.Form("Postcode")
	objValueStateID					= Request.Form("StateID")
	objValueMembershipNumber		= Request.Form("MembershipNumber")
	objValueMemberEmailAddress		= Request.Form("MemberEmailAddress")
	objValuePhoenixCopies			= Request.Form("PhoenixCopies")
	objValueVESAPocketDiary			= Request.Form("VESAPocketDiary")
	objValueVESAWallCalendar		= Request.Form("VESAWallCalendar")
	objValueVESAUnitID				= Request.Form("VESAUnitID")
%>
	<form name="CheckForm" id="CheckForm" action="VESASave.asp" method="post" onSubmit="return stopSubmit()">
	<input type="hidden" name="ActionType" value="<%=objValueActionType%>">
	<% If objValueActionType = "Update" Then %>    
		<input type="hidden" name="RecipientID" value="<%=objValueRecipientID%>">
	<% End If %>
	<input type="hidden" name="Surname_Organization" value="<%=objValueSurname_Organization%>">
	<input type="hidden" name="FirstName" value="<%=objValueFirstName%>">
	<input type="hidden" name="Address" value="<%=objValueAddress%>">
	<input type="hidden" name="Suburb" value="<%=objValueSuburb%>">
	<input type="hidden" name="Postcode" value="<%=objValuePostcode%>">
	<input type="hidden" name="StateID" value="<%=objValueStateID%>">
	<input type="hidden" name="MembershipNumber" value="<%=objValueMembershipNumber%>">
	<input type="hidden" name="MemberEmailAddress" value="<%=objValueMemberEmailAddress%>">
	<input type="hidden" name="PhoenixCopies" value="<%=objValuePhoenixCopies%>">
	<input type="hidden" name="VESAPocketDiary" value="<%=objValueVESAPocketDiary%>">
	<input type="hidden" name="VESAWallCalendar" value="<%=objValueVESAWallCalendar%>">
	<input type="hidden" name="VESAUnitID" value="<%=objValueVESAUnitID%>">
   
	<% If objValueActionType = "Update" Then 
		blnPassThrough = True
	%>
		 </form>
	<%

		 If Session("AccessRights") = "Level 5" Then
			Call displayFORMLinks()
		 End If
	%>	 
		 <script type="text/javascript">
		 <!--
		 function doSubmit() {
			<% If blnPassThrough = true Then %>
				document.CheckForm.submit();
			<% End If %>
		 }
		 //-->
		 </script>
	
	<% 
		Else
			If Session("AccessRights") = "Level 1" Then
				blnPassThrough = True
	%>
				</form>
		 
				<script type="text/javascript">
				<!--
				function doSubmit() {
					<% If blnPassThrough = true Then %>
						document.CheckForm.submit();
					<% End If %>
				}
				//-->
				</script>

	<%		Else
				Call checkMembersDatabase()
			End If 
	End If 
End Sub

Sub checkMembersDatabase()
%>
	<table cellspacing="0" cellpadding="0">
	<tr>
	<td>
	<%
	Dim strSQL
	Dim intblueLightID
	Dim rsResult
	Dim blnPassThrough

	EstablishConnection()

	strSQL = "SELECT * FROM VESA_tblMembers M" 
	strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
	strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
	strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
	strSQL = strSQL & " WHERE MembershipNumber = '" & Request.Form("MembershipNumber") & "'"
   
	Set rsResult = Server.CreateObject("ADODB.Recordset")
	rsResult.Open strSQL, Conn

	If not rsResult.EOF Then
		blnPassThrough = False

		Response.Write "You are using a <font color=""#ff0000"">** (MEMBERSHIP NUMBER) **</font> that a current member already have.<br /><br />" & vbCrLf
		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			   
		rsResult.MoveFirst
		Do Until rsResult.EOF 
	%>
			<tr>
				<td><div align="left"><strong><label for="Surname/Organization">Surname/Organization:</label></strong></div></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt=""></td>
				<td><%=Request.Form("Surname_Organization")%></td>
			</tr>

			<tr>
				<td><div align="left"><strong><label for="First Name">First Name:</label></strong></div></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt=""></td>
				<td><%=Request.Form("FirstName")%></td>
			</tr>

			<tr>
				<td valign="top"><div align="left"><strong><label for="Address">Address:</label></strong></div></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt=""></td>
				<td>
				<%=Request.Form("Address")%> <br />
				<%=Request.Form("Suburb")%>&nbsp;<%=Request.Form("Postcode")%> <br />
				<%=Request.Form("State_Name")%>
				</td>
			</tr>

			<tr>
				<td><div align="left"><strong><label for="Membership Number">Membership Number:</label></strong></div></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt=""></td>
				<td><%=Request.Form("MembershipNumber")%>&nbsp;&nbsp;<font color="#ff0000">** (Same MEMBERSHIP NUMBER) **</font></td>
			</tr>

			<tr>
				<td valign="top"><div align="left"><strong><label for="Email Address">Email Address:</label></strong></div></td>	   
				<td><img src="images/spacer.gif" width="5" height="1" alt="" /></td>
				<td colspan="3">
				<%
				If Request.Form("EmailAddress") = "" Then
					Response.Write "<font color=""#ff0000""><i>No Email Address given</i></font>"
				Else
					Response.Write rstSearch.Fields("EmailAddress")
				End If			
				%>
				</td>
			</tr>

			<!--
  			'----------------------------------------------------------------------------------
			'***** Show these editable fields if the login User is the Administrator *****		
			'----------------------------------------------------------------------------------
			-->
			<% If Session("AccessRights") = "Level 1" Then %>
			<tr>
				<td valign="top"><div align="left"><strong><label for="PhoenixCopies">Phoenix Copies:</label></strong></div></td>	   
				<td><img src="images/spacer.gif" width="5" height="1" alt="" /></td>
				<td colspan="3"><%=Request.Form("PhoenixCopies")%></td>
			</tr>

			<tr>
				<td valign="top"><div align="left"><strong><label for="VESAPocketDiary">VESA Pocket Diary:</label></strong></div></td>	   
				<td><img src="images/spacer.gif" width="5" height="1" alt="" /></td>
				<td colspan="3"><%=Request.Form("VESAPocketDiary")%></td>
			</tr>

			<tr>
				<td valign="top"><div align="left"><strong><label for="VESAWallCalendar">VESA Wall Calendar:</label></strong></div></td>	   
				<td><img src="images/spacer.gif" width="5" height="1" alt="" /></td>
				<td colspan="3"><%=Request.Form("VESAWallCalendar")%></td>
			</tr>

			<tr>
				<td Valign="TOP"><div align="left"><strong><label for="VESA Unit">VESA Unit:</label></strong></div></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt=""></td>
				<td><% Call showdbSelectedValue(Conn, rsVESAUnitID, "VESA_tblUnit", "VESAUnit", "VESAUnitID", "" & Request.Form("VESAUnitID") & "")%></td>
			</tr>

			<tr>
				<td Valign="TOP"><div align="left"><strong><label for="SES Region">SES Region:</label></strong></div></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt=""></td>
				<td><% Call showdbSelectedValue(Conn, rsVESAUnitID, "VESA_tblSESRegion", "SESRegion", "SESRegionID", "" & Request.Form("SESRegionID") & "")%></td>
			</tr>

			<!--
			'----------------------------------------------------------------------------------
			'***** Otherwise only show this field if the User is a VESA Unit user/member *****		
			'----------------------------------------------------------------------------------
			-->
			<% Else %>
			<tr>
				<td Valign="TOP"><div align="left"><strong><label for="VESA Unit">VESA Unit:</label></strong></div></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt=""></td>
				<td><%Call showdbSelectedValue(Conn, rsVESAUnitID, "VESA_tblUnit", "VESAUnit", "VESAUnitID", "" & Request.Form("VESAUnitID") & "")%></td>
			</tr>
			<% End If %>
<%
         rsResult.MoveNext
      Loop
%>
			</table>
		</td>
		</tr>

		<tr height="5"><td colspan="3"><img src="images/spacer.gif" width="1" height="5" alt="" /></td></tr>

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
			<td><input type="image" name="back" class="back-btn" src="http://www.roscripts.com/images/btn.gif" alt="back" title="back" onClick="goback()" /></td>
		</tr>
  <% End If %>
  </table>
  </form>

  <script type="text/javascript">
  <!--
  function doSubmit() {
     <% If blnPassThrough = true Then %>
           document.CheckForm.submit();
     <% End If %>
  }

  function goback() {
     history.go(-1); 
  }
  //-->
  </script>

<% End Sub

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


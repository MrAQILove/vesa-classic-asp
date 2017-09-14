<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="include/include.asp"-->
<!--#include file="include/functions.asp"-->

<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = True 

If Session("UserLoggedIn") <> "true" Then
   If Session("AccessRights") = "Level 1" Then
      Response.Redirect "AdminLogin.asp"
   Else
      Response.Redirect "VESAUnitLogin.asp"
   End If 

Else
	EstablishConnection()  
	Call ControlPanel()   
	CloseConnection()
End If 

Sub ControlPanel() 
%>
	<!DOCTYPE html>
	<html>
	<head>
		<meta charset="UTF-8">
		<!--[if lt IE 9]>
		  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
		<![endif]-->
		<title>VESA Members Database</title>
		<meta name="VESA Members Database" content="" />
		<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
		<link rel="stylesheet" href="css/buttons.css">
		<link rel="stylesheet" href="css/forms.css">
		<link rel="stylesheet" href="css/base.css">
		<link rel="stylesheet" href="css/grids.css">
		<script type="text/javascript">
		<!--
		// view members 
		function viewMembers() 
		{
		   if (<%=Session("VESAID")%> == 1) {
			  document.viewMembersForm.submit();
		   }
		}

		// add member 
		function addMember() 
		{
		   if (<%=Session("VESAID")%> == 1) {
			  document.addMemberForm.submit();
		   }
		}

		// delete member 
		function deleteMember() 
		{
		   if (<%=Session("VESAID")%> == 1) {
			  document.deleteMemberForm.submit();
		   }
		}
		
		//Member Selected
		function memberSelected(strUnit)
		{
			if(<%=Session("VESAID")%> == 1) 
			{
				document.EditUnit.VESAUnitID.value = strUnit;
				document.EditUnit.submit();
			}
		}
		//-->
	</script>
	</head>
	<body>
		<div id="wrapper">
			<!-- start menu -->
			<nav id="menu">
				<ul id="main">
					<li class="current_page_item"><a href="VESAMain.asp">Home</a></li>
					<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
					<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
					<li><a href="contactUs.html">Contact Us</a></li>
				</ul>
			</nav>
			<!-- end menu -->
			
			<!-- start header -->
			<header id="header">
				<div id="logo">
					<h1><a href="#"><span></span></a></h1>
					<p></p>
				</div>
			</header>
			<!-- end header -->
			
			<!-- start section -->
			<section id="page">
				<%
				If Session("AccessRights") = "Level 1" Then
					Call adminMainMenu()
				Else 
					Call VESAUnitMainMenu()
				End If 
				%>
				
				<!-- start article -->
				<article id="content">
					<div class="post">
						<% Call displayWelcome(Session("User")) %>
						<div class="entry">
						<%
							If Session("AccessRights") = "Level 5" Then
								Call displayFORMLinks()
							Else
								Call viewAllUnits()
							End If 
						%>
						</div>
					</div>
				</article>
				<!-- end article -->
				
				<div style="clear: both;">&nbsp;</div>
			</section>
		</div>
		
		<!-- start footer -->
		<footer id="footer">
			<p class="copyright">&copy;&nbsp;&nbsp;2008 - <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
		</footer>
		<!-- end footer -->
	</body>
	</html>
<% End Sub 

Sub adminMainMenu()
%>
	<aside id="sidebar1" class="sidebar">
		<ul>
			<li> 
				<h2>VESA Members</h2>
				<ul>
					<li class="hover"><a href="#">Home</a></li>
					<li><a href="VESASearch.asp">Search for Member</a></li>
					<li><a href="VESAViewAllMembers.asp">View All Members</a></li>
					<li><a href="VESAAddNewMember.asp">Add a New Member</a></li>
					<li><a href="VESADeleteMember.asp">Delete a Member</a></li>
					<li><a href="VESAViewHistory.asp">View History</a></li>
					<li><a href="admin/VESAViewInactiveMembers.asp">View Inactive Members</a></li>
				</ul>
			</li>

			<li> 
				<h2>VESA Units</h2>
				<ul>
					<li><a href="admin/ViewAllVESAUnits.asp">View All VESA Units</a></li>
					<li><a href="admin/AddNewVESAUnit.asp">Add a New VESA Unit</a></li>
					<li><a href="admin/DeleteVESAUnit.asp">Delete a VESA Unit</a></li>
					<li><a href="admin/VESAViewInactiveUnits.asp">View All Inactive Units</a></li>
				</ul>
			</li>

			<li> 
				<h2>Admin Members</h2>
				<ul>
					<li><a href="admin/VESAViewAllAdminUsers.asp">View All Admin Users</a></li>
					<li><a href="admin/VESAAddAdminUser.asp">Add an Admin User</a></li>
					<li><a href="admin/VESADeleteAdminUser.asp">Delete an Admin User</a></li>
				</ul>
			</li>

			<li><a href="AdminLogin.asp"><h2>Log Out</h2></a></li>
		</ul>
	</aside>
<% End Sub

Sub viewAllUnits()
%>
	<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr>
		<td align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<!--/* Active Units */-->
					<form class="pure-form pure-form-aligned" name="VESAMain" method="post" action="VESAOutput.asp">
					<input type="hidden" name="search" value="12">

						<fieldset>
							<legend style="color:#ff0000; font-weight:bold">ACTIVE UNITS</legend>
							<div class="pure-control-group">
								<label>No. of Active Units:</label>
								<span style="color:#0000a0; font-weight:bold"><% Call unitCount("1")%></span>
							</div>

							<div class="pure-control-group">
								<label>VESA Units:</label>
								<% Call ActiveVESAUnit(Conn, rsActiveVESAUnit, "VESA_tblUnit", "VESAUnitID", "VESAUnit", "searchForVESAUnit")%>
							</div>

							<div class="pure-controls">
								<button type="submit" class="pure-button">View Members</button>
							</div>
						</fieldset>
					</form>
					<!--/* End of Active Units */-->
				</td>
			</tr>

			<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>
			
			<tr>
				<td>
					<!--/* Inactive Units */-->
					<form class="pure-form pure-form-aligned" name="VESAMain" method="post" action="VESAOutput.asp">
					<input type="hidden" name="search" value="13">

						<fieldset>
							<legend style="color:#ff0000; font-weight:bold">INACTIVE UNITS</legend>
							<div class="pure-control-group">
								<label>No. of Inactive Units:</label>
								<span style="color:#0000a0; font-weight:bold"><% Call unitCount("0")%></span>
							</div>

							<div class="pure-control-group">
								<label>VESA Units:</label>
								<% Call InactiveVESAUnit(Conn, rsActiveVESAUnit, "VESA_tblUnit", "VESAUnitID", "VESAUnit", "searchForVESAUnit")%>
							</div>

							<div class="pure-controls">
								<button type="submit" class="pure-button">View Members</button>
							</div>
						</fieldset>
					</form>
					<!--/* End of Inactive Units */-->
				</td>
			</tr>		 
			</table>   
		</td>
	</tr>
   </table>
   <!--/* End Here */-->
<% End Sub 

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'		admin/DeleteVESAUnit.asp
'		(Show dropdown list for VESA Unit field)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub ActiveVESAUnit(c, r, table, dbField1, dbField2, formField)

   '-- SQL Statement
   strSQL = "SELECT * FROM " & table & " WHERE IsActive = '1' ORDER BY " & dbField2 & " ASC"
      
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   Response.Write "<input type=""hidden"" name=""active"" value=""1"">"
   Response.Write "<select id=""" & formField & """ name=""" & formField & """ class=""pure-input-medium"">"
   Response.Write "<option value="""">Please Chooose</option>"

   '-- loop and build each database entry as a selectable option
   While r.EOF = false
      Response.Write "<option value=""" & r.Fields("" & dbField1 & "").Value & """>" _ 
	  & r.Fields("" & dbField2 & "").Value & "</option>" & vbCrLf

      '-- Move recordset to the next value
      r.movenext
   Wend
   '--END OF MAIN CODE BLOCK

   '-- close select/form tags
   Response.Write "</select>" & vbCrlf
   Response.Write "<noscript><input type=""submit"" value=""submit""></noscript>" & vbCrlf
End Sub
'-----------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'		admin/DeleteVESAUnit.asp
'		(Show dropdown list for VESA Unit field)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub InactiveVESAUnit(c, r, table, dbField1, dbField2, formField)

   '-- SQL Statement
   strSQL = "SELECT * FROM " & table & " WHERE IsActive = '0' ORDER BY " & dbField1 & " ASC"
      
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   Response.Write "<input type=""hidden"" name=""inactive"" value=""1"">"
   Response.Write "<select id=""" & formField & """ name=""" & formField & """ class=""pure-input-medium"">"
   Response.Write "<option value="""">Please Chooose</option>"

   '-- loop and build each database entry as a selectable option
   While r.EOF = false
      Response.Write "<option value=""" & r.Fields("" & dbField1 & "").Value & """>" _ 
	  & r.Fields("" & dbField2 & "").Value & "</option>" & vbCrLf

      '-- Move recordset to the next value
      r.movenext
   Wend
   '--END OF MAIN CODE BLOCK

   '-- close select/form tags
   Response.Write "</select>" & vbCrlf
End Sub
'-----------------------------------------------------------------------------------------------

Sub unitCount(strIsActive)
	
	'-----Count the number of VESA Units
	Set ObjRs = Server.CreateObject("ADODB.Recordset")
	ObjRs.Open "SELECT * FROM VESA_tblUnit WHERE IsActive = '" & strIsActive & "'", Conn, 1, 1
	ObjRsCount = ObjRs.RecordCount
	ObjRs.Close

	Response.Write ObjRsCount

End Sub 
%>
<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/database/vesa/Members/include/include.asp"-->
<!--#include virtual="/database/vesa/Members/include/functions.asp"--> 
<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = True 

If Session("UserLoggedIn") <> "true" Then
   Response.Redirect "../AdminLogin.asp" 

Else
   '- Constants ripped from adovbs.inc:
   Const adOpenStatic = 3
   Const adLockReadOnly = 1
   Const adCmdText = &H0001

   '- Our own constants:
   Const PAGE_SIZE = 300  ' The size of our pages.

   '- Declare our variables... always good practice!
   Dim rstSearch		' ADO recordset
   Dim strSQL			' The SQL Query we build on the fly
   Dim iPageCurrent		' The page we're currently on
   Dim iPageCount		' Number of pages of records
   Dim iRecordCount		' Count of the records returned
   Dim I				' Standard looping variable

   '- Retrieve page to show or default to the first
   If Request.QueryString("page") = "" Then
      iPageCurrent = 1
   Else
      iPageCurrent = CInt(Request.QueryString("page"))
   End If

   EstablishConnection()
   
   '- Build our query based on the input.
   strSQL = "SELECT * FROM VESA_tblUnit U"
   strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
   'strSQL = strSQL & " WHERE IsUnitSES = '1'"
   strSQL = strSQL & " WHERE IsActive = '1'" 
   strSQL = strSQL & " ORDER BY VESAUnitID ASC"

   '- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
   Set rstSearch = Server.CreateObject("ADODB.Recordset")
   rstSearch.PageSize  = PAGE_SIZE
   rstSearch.CacheSize = PAGE_SIZE

   '- Open our recordset
   rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

   '- Get a count of the number of records and pages for use in building the header and footer text.
   iRecordCount = rstSearch.RecordCount
   iPageCount   = rstSearch.PageCount

   DisplayHeader("VESA Members Database : Displaying All VESA Units")
      
   Call OutputPage()

   '- Close our recordset and connection and dispose of the objects
   rstSearch.Close
   Set rstSearch = Nothing
   
   CloseConnection()

End If

Sub DisplayHeader(strMessage) %>
	<!DOCTYPE html>
		<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
			  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
			<title><%=strMessage%></title>
			<meta name="keywords" content="" />
			<meta name="eMag Members Database" content="" />
			<link href="../css/default.css" rel="stylesheet" type="text/css" media="screen" />
			<link rel="stylesheet" href="../css/buttons.css">
			<link rel="stylesheet" href="../css/forms.css">
			<link rel="stylesheet" href="../css/base.css">
			<script language="javascript">
			<!--
			// Export Member List into an Excel File 
			function exportFile(){
			   document.ExportForm.submit();
			}
			//Member Selected
			function unitSelected(strUnit)
			{
			   if(<%=Session("VESAID")%> == 1) 
			   {
				  document.EditForm.VESAUnitID.value = strUnit;
				  document.EditForm.submit();
			   }
			}

			// Delete Member
			function deleteSelected()
			{
			   var ctr;
			   
			   ctr = 0;
			   
			   // check for single checkbox by seeing if an array has been created
			   var cblength = document.forms['MultiDeleteUnit'].elements['DoDeleteVESAUnit'].length;
			   if(typeof cblength == "undefined")
			   {
				  if(document.forms['MultiDeleteUnit'].elements['DoDeleteVESAUnit'].checked == true) ctr++;
			   }
			   else
			   {
				  for(i = 0; i < document.forms['MultiDeleteUnit'].elements['DoDeleteVESAUnit'].length; i++)
				  {
					 if(document.forms['MultiDeleteUnit'].elements['DoDeleteVESAUnit'][i].checked) ctr++;
				  }
				}
							  
			   if (ctr == 1) 
			   {
				   var answer;
				   answer = confirm('Are you sure you want to delete this unit?');
				   if (answer)
				   {
					  document.MultiDeleteUnit.submit();
					  return false;   
				   }

				   //else {;}
				}
				
				else if (ctr > 1) 
				{
				   var answer;
				   answer = confirm('Are you sure you want to delete ' + ctr + ' units?');
				   if (answer)
				   {
					  document.MultiDeleteUnit.submit();
					  return false;
				   }

				   //else {;}
				}
				
				else 
				{
				   confirm("No unit selected for deletion");
				   return true;
				}
			}
			//-->
			</script>
		</head>
<% End Sub

Sub OutputPage() %>
	<body>
		<div id="wrapper">
			<nav id="menu">
				<ul id="main">
					<li><a href="../VESAMain.asp">Home</a></li>
					<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
					<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
					<li><a href="contactus.asp">Contact Us</a></li>
				</ul>
			</nav>
			
			<!-- start header -->
			<header>
				<div id="logo">
					<h1><a href="#"><span></span></a></h1>
					<p></p>
				</div>
			</header>
			<!-- end header -->
			
			<!-- start page -->
			<section>
				<%
					If Session("AccessRights") = "Level 1" Then
						Call adminMainMenu()
					End If 
				%>
				
				<!-- start content -->
				<article>
					<% Call viewAllUnits() %>
				</article>
				<!-- end content -->
				<div style="clear: both;">&nbsp;</div>
			</section>
			<!-- end page -->

		</div>
		<footer id="footer">
			<p class="copyright">&copy;&nbsp;&nbsp;2008 - <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
		</footer>
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
					<li><a href="../VESAMain.asp">Home</a></li>
					<li><a href="../VESASearch.asp">Search for Member</a></li>
					<li><a href="../VESAViewAllMembers.asp">View All Members</a></li>
					<li><a href="../VESAAddNewMember.asp">Add a New Member</a></li>
					<li><a href="../VESADeleteMember.asp">Delete a Member</a></li>
					<li><a href="../VESAViewHistory.asp">View History</a></li>
					<li><a href="VESAViewInactiveMembers.asp">View Inactive Members</a></li>
				</ul>
			</li>

			<li> 
				<h2>VESA Units</h2>
				<ul>
					<li class="hover"><a href="#">View All VESA Units</a></li>
					<li><a href="AddNewVESAUnit.asp">Add a New VESA Unit</a></li>
					<li><a href="DeleteVESAUnit.asp">Delete a VESA Unit</a></li>
					<li><a href="VESAViewInactiveUnits.asp">View All Inactive Units</a></li>
				</ul>
			</li>

			<li> 
				<h2>Admin Members</h2>
				<ul>
					<li><a href="VESAViewAllAdminUsers.asp">View All Admin Users</a></li>
					<li><a href="VESAAddAdminUser.asp">Add an Admin User</a></li>
					<li><a href="VESADeleteAdminUser.asp">Delete an Admin User</a></li>
				</ul>
			</li>

			<li><a href="../AdminLogin.asp"><h2>Log Out</h2></a></li>
		</ul>
	</aside>
<% End Sub

Sub viewAllUnits()
%>
	<div class="entry">
		<!--/* Start Here */-->
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
		<tr>
			<td align="center">
				<table border="0" align="center" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td>
					<!--/* header */-->
					<h1 class="title"><a href="#">Displaying All Active VESA Units</a></h1>
					<p class="byline">
					<b>
					<% 
						Response.Write "The VESA Member's database has <b><font color=""#ff0000"">"
						If iRecordCount > 1 Then
							Response.Write iRecordCount & "</font></b> active units!" & vbCrLf
						Else
							Response.Write iRecordCount & "</font></b> active unit!" & vbCrLf
						End If 	   
					%>	
					</b>
					</p>
					
					<%
					'- Check page count to prevent bombing when zero results are returned!-----------------
					If iRecordCount = 0 Then
						Response.Write "<p class=""byline""><b>No records found!</b></p>"
						Response.Write "</td></tr>"
						Response.Write "</table>"

					Else
						rstSearch.AbsolutePage = iPageCurrent
					%>
					<td>
				</tr>

				<tr height="10"><td><img src="../images/spacer.gif" width="1" height="10" border="0"></td></tr>

				<tr>
					<td>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
						<tr>
							<td><strong><font color="#ff0000">There are <%= iRecordCount %> active units.</font></strong></td>
							<td align="right">
							<span style="color: #ff0000; font-weight:bold">
							<%
							If iRecordCount > PAGE_SIZE Then
								Response.Write "Displaying " & PAGE_SIZE & " of " & iRecordCount & " Records &nbsp;&bull;&nbsp; Page " & iPageCurrent & " of " & iPageCount & ":"
							Else
								If iRecordCount <> 1 Then
									Response.Write "Displaying " & iRecordCount & " Records &nbsp;&bull;&nbsp; Page " & iPageCurrent & " of " & iPageCount & ":"
								Else
									Response.Write "Displaying " & iRecordCount & " Record &nbsp;&bull;&nbsp; Page " & iPageCurrent & " of " & iPageCount & ":"
								End If 
							End If 
							%>
							</span>
							</td>
						</tr>
						</table>
					</td>
				</tr>

				<tr><td valign="top"><img src="../images/spacer.gif" width="1" height="10" alt="" /></td></tr>

				<tr>
					<td bgcolor="#eeeeee">
						<!--/* Output Search */-->
						<form name="MultiDeleteUnit" id="MultiDeleteUnit" action="../VESASave.asp" method="post" onSubmit="return stopSubmit()">
						<input type="hidden" name="VESAID" value="<%=Session("VESAID")%>">
						<input type="hidden" name="ActionType" value="DeleteVESAUnit">
						<table id="main_table" border="0" align="center" cellspacing="2" cellpadding="1" width="100%">
						<tr align="center" height="30">
							<% If Session("User") <> "Webmaster" Then %>
								<td class="tab_header_cell"><b>VESA Unit ID</b></td>
							<% End If %>
							<td class="tab_header_cell"><font color="#0000a0"><b>Edit</b></font></td>
							<td class="tab_header_cell"><font color="#0000a0"><b>Delete</b></font></td>
							<td class="tab_header_cell"><b>VESA Units</b></td>

							<% If (Session("User") = "Webmaster") Or (Session("User") = "Administrator" ) Then %>
								<td class="tab_header_cell"><b>Password</b></td>
							<% End If %>
							
							<td class="tab_header_cell"><b>No. of Members</b></td>
							<td class="tab_header_cell"><b>Email Address</b></td>
							<td class="tab_header_cell"><b>SES Region</b></td>
							<td class="tab_header_cell"><b>Active</b></td>
						</tr>
			   
						<%
						Do While Not rstSearch.EOF And rstSearch.AbsolutePage = iPageCurrent
							VESAUnitID		= rstSearch("VESAUnitID") & ""
							IDArray			= CInt(rstSearch("VESAUnitID")) & ""
							VESAUnit		= rstSearch("VESAUnit") & ""
							Password		= rstSearch("Password") & ""
							EmailAddress	= rstSearch("EmailAddress") & ""
							SESRegion		= rstSearch("SESRegion") & ""
							IsActive		= rstSearch("IsActive") & ""
							
							j = j + 1
							Response.write "<tr height=""20"" class=""listTableText" & (j And 1) & """>"
							
							If Session("User") <> "Webmaster" Then %>
								<td align="center"><%=VESAUnitID%></td>
							<% End If %>
							<td align="center"><a href="javascript:unitSelected(<%=VESAUnitID%>)"><img src="../images/edit.gif" width="16" height="16" border="0" alt="<%=VESAUnitID%>" /></a></td>
							<td align="center"><input type="checkbox" id="DoDeleteVESAUnit" name="DoDeleteVESAUnit" value="<%=IDArray%>"></td>
							<td align="center"><%=VESAUnit%></td>
							<% If (Session("User") = "Webmaster") Or (Session("User") = "Administrator" ) Then %>
								<td align="center"><%=Password%></td>
							<% End If %>
							<td align="center">
							<%
							'Password = kLeachRegExp("" & rstSearch("Password") & "", "[^()?<>.*?]", "*") 
							'Response.Write Password

							Call memberCount(VESAUnitID)
							%>
							</td>
							<%
							Response.Write "<td align=""center"">"
							If Not IsNull(rstSearch.Fields("EmailAddress")) Then
								Response.Write rstSearch("EmailAddress")
							Else
								Response.Write "<font color=""#ff0000"">No Email Address given</font>"
							End If 
							Response.Write "</td>"

							Response.Write "<td align=""center"">"
							If Not IsNull(rstSearch.Fields("SESRegion")) Then
								Response.Write rstSearch("SESRegion")
							Else
								Response.Write "<font color=""#ff0000"">No SES Region given</font>"
							End If 
							Response.Write "</td>"

							Response.Write "<td align=""center"">"
							If rstSearch.Fields("IsActive") = "1" Then
								Response.Write "YES"
							Else
								Response.Write "<font color=""#ff0000"">NO</font>"
							End If 
							Response.Write "</td>"
							%>
						</tr>
						<%
							rstSearch.MoveNext
						Loop
						%>
						</table>
						</form>
					</td>
				</tr>
			
				<tr><td valign="top"><img src="../images/spacer.gif" width="1" height="10" alt="" /></td></tr>
		 
				<tr>
					<td>
						<table border="0" width="100%">
						<tr>
							<td>Page <strong><%=iPageCurrent%></strong> of <strong><%=iPageCount%></strong></td>
							<td>
							<div align="right">					
								<div class="pages">
									<% Call databasePaging() %>
								</div>
							</div>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<% End If %>

				<% If Session("AccessRights") = "Level 1" Then %>
				<tr><td valign="top"><img src="../images/spacer.gif" width="1" height="20" alt="" /></td></tr>

				<tr>
					<td width="100%">
						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td><button type="button" class="pure-button" onClick="addNewUser()">Add</button></td>
							<td><img src="../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
							<td><button type="button" class="pure-button" onClick="deleteUser()">Delete</button></td>
							
							<% If (Session("User") = "Webmaster") Or (Session("User") = "Administrator") Then %>
							<td><img src="../images/spacer.gif" width="10" height="1" border="0" alt=""></td>
							<td><button type="button" class="pure-button" onClick="exportFile()">Export as Excel</button></td>
							<% End If %>
						</tr>
						</table>
					</td>
				</tr>

				<!--/* Add User Form */-->
				<form name="AddUserForm" id="AddUserForm" action="VESAAddAdminUser.asp" method="post"></form>
				<!--/* End Here */-->

				<form name="EditForm" id="EditForm" action="EditVESAUnit.asp" method="post">
				<input type="hidden" name="VESAUnitID" value="<%=VESAUnitID%>">
				</form>

				<form name="ExportForm" id="ExportForm" action="../VESAExport.asp" method="post">
				<input type="hidden" name="search" value="All VESA Unit Password">
				</form>
				<% End If %>

				</table>
			</td>
		</tr>
		</table>
		<!--/* End Here */-->
	</div>
<% End Sub 

Sub databasePaging()
	If iPageCurrent > 1 Then 
	%>
		<a href="ViewAllVESAUnits.asp?page=<%=iPageCurrent - 1%>">&lt;&nbsp;Prev</a>
	<%
	Else
		Response.Write "<span class=""disabled"">&lt;&nbsp;Prev</span>"
	End If
							
	'--------------------------------------------------------------------------------------
	'- You can also show page numbers:
	For I = 1 To iPageCount
		'- Don't hyperlink the current page number
		If I = iPageCurrent Then
			Response.Write "<span class=""current"">" & I & "</span>" & vbCrLf
															
		Else
			Response.Write "<a href=""ViewAllVESAUnits.asp?page=" & I & """>" & I & "</a>" & vbCrLf
		End If
	'- I
	Next 
		If iPageCurrent < iPageCount Then
			Response.Write "<a href=""ViewAllVESAUnits.asp?page=" & iPageCurrent + 1 & """>Next&nbsp;&gt;</a>"
														
		Else
			Response.Write "<span class=""disabled"">Next&nbsp;&gt;</span>"
		End If
	'--------------------------------------------------------------------------------------
End Sub


Sub memberCount(strVESAUnitID)
	'strSQL = "SELECT (SELECT COUNT(*) FROM VESA_tblUnit WHERE IsActive = 1) AS Count_1, (SELECT COUNT(*) FROM VESA_tblUnit WHERE IsActive = 0) AS Count_2

	'-----Count All Members with corresponding VESAUnit
	Set ObjRs = Server.CreateObject("ADODB.Recordset")
	ObjRs.Open "SELECT * FROM VESA_tblMembers M INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID WHERE M.VESAUnitID = '" & strVESAUnitID & "'", Conn, 1, 1
	ObjRsCount = ObjRs.RecordCount
	ObjRs.Close

	Response.Write ObjRsCount

End Sub 
%>
 
   

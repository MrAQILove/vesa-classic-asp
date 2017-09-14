<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE FILE="include/include.asp"-->
<!--#INCLUDE FILE="include/functions.asp"--> 
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
   '- Constants ripped from adovbs.inc:
   Const adOpenStatic = 3
   Const adLockReadOnly = 1
   Const adCmdText = &H0001

   '- Our own constants:
   Const PAGE_SIZE = 500  ' The size of our pages.

   '- Declare our variables... always good practice!
   Dim rstSearch		' ADO recordset
   Dim strSQL			' The SQL Query we build on the fly
   Dim strSearch		' The text being looked for 1
   Dim strSearchFor		' The text being looked for 2

   Dim iPageCurrent		' The page we're currently on
   Dim iPageCount		' Number of pages of records
   Dim iRecordCount		' Count of the records returned
   Dim I				' Standard looping variable

   '- Retrieve page to show or default to the first
   strSearch		= Request.Form("search")
   strSearchFor		= Request.Form("searchFor")
   strSearchFor		= Replace(strSearchFor, "'", "''")
   
   If Request.QueryString("page") = "" Then
      iPageCurrent = 1
   Else
      iPageCurrent = CInt(Request.QueryString("page"))
   End If

   EstablishConnection()
   
   '- Build our query based on the input.
   strSQL = "SELECT * FROM VESA_tblAudit A"
   strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON A.StateID = S.StateID"
   strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON A.VESAUnitID = U.VESAUnitID"
   strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
	
	Select Case strSearch   
		Case "RecipientID"
			strSQL = strSQL & " WHERE A.RecipientID='" & strSearchFor & "'"

		Case "Surname/Organisation"
			strSQL = strSQL & " WHERE A.Surname_Organization='" & strSearchFor & "'"
		 
		Case "First Name"
			strSQL = strSQL & " WHERE A.FirstName='" & strSearchFor & "'"
		 
		Case "Address"
			strSQL = strSQL & " WHERE A.Address='" & strSearchFor & "'"
		 
		Case "Suburb"
			strSQL = strSQL & " WHERE A.Suburb='" & strSearchFor & "'"
		 
		Case "Postcode"
			strSQL = strSQL & " WHERE A.Postcode='" & strSearchFor & "'"
		 
		Case "State"
			strSQL = strSQL & " WHERE S.StateID='" & strSearchFor & "'"

		Case "Membership Number"
			strSQL = strSQL & " WHERE A.MembershipNumber='" & strSearchFor & "'"

		Case "VESA Unit"
			strSQL = strSQL & " WHERE U.VESAUnitID='" & strSearchFor & "'"

		Case "SES Region"
			strSQL = strSQL & " WHERE R.SESRegionID='" & strSearchFor & "'"
   End Select

   strSQL = strSQL & " ORDER BY AuditID DESC"

   '- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
   Set rstSearch = Server.CreateObject("ADODB.Recordset")
   rstSearch.PageSize  = PAGE_SIZE
   rstSearch.CacheSize = PAGE_SIZE

   '- Open our recordset
   rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

   '- Get a count of the number of records and pages for use in building the header and footer text.
   iRecordCount = rstSearch.RecordCount
   iPageCount   = rstSearch.PageCount

   DisplayHeader("VESA Members Database : Displaying All Audited Members")
      
   Call OutputPage()

   '- Close our recordset and connection and dispose of the objects
   rstSearch.Close
   Set rstSearch = Nothing
   
   CloseConnection()

End If

Sub DisplayHeader(strMessage) %>
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<title><%=strMessage%></title>
	<meta name="keywords" content="" />
	<meta name="VESA Members Database" content="" />
	<link href="css/databaseView.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="stylesheet" href="css/buttons.css">
	<link rel="stylesheet" href="css/forms.css">
	<link rel="stylesheet" href="css/base.css">
	<link rel="stylesheet" href="css/grids.css">
	<script language="javascript">
	<!--
	// Back to Main
	function goBack() {
	   document.location.href = "VESAMain.asp";
	}

	// Export Member List into an Excel File 
	function exportFile(){
	   document.ExportForm.submit();
	}
	//-->
	</script>
	</head>
<% End Sub

Sub OutputPage() %>
	<body>
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
				<%Call viewMembers()%>
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

Sub DisplayTitle(strSearch)
	Select Case strSearch
		Case "RecipientID"
			If strSearchFor <> "" Then 
				Response.Write "Displaying Audited Member with a Recipient ID of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the Recipient ID that you are searching for."
			End If

		Case "Surname/Organisation"
			If strSearchFor <> "" Then 
				Response.Write "Displaying All Audited Members with a SURNAME/ORGANISATION of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. <br /> Please specify the SURNAME/ORGANISATION that you are searching for."
			End If

		Case "First Name"
			If strSearchFor <> "" Then 
				Response.Write "Displaying All Audited Members with a FIRST NAME of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. <br /> Please specify the FIRST NAME that you are searching for."
			End If
					 
		Case "Address"
			If strSearchFor <> "" Then 
				Response.Write "Displaying Member with an ADDRESS of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the ADDRESS that you are searching for."
			End If
					 
		Case "Suburb"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Audited Members in the SUBURB of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the SUBURB that you are searching for."
			End If

		Case "Postcode"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Audited Members in the POSTCODE of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the POSTCODE that you are searching for."
			End If

		Case "State"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Audited Members in the STATE of <font color=""#ff0000"">"
				Select Case strSearchFor
					Case 1
						Response.Write "Australian Capital Territory"
					Case 2
						Response.Write "New South Wales"
					Case 3
						Response.Write "Northern Territory"
					Case 4
						Response.Write "Queensland"
					Case 5
						Response.Write "South Australia"
					Case 6
						Response.Write "Tasmania"
					Case 7
						Response.Write "Victoria"
					Case 8
						Response.Write "Western Australia"
				 End Select
				 Response.Write "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the STATE that you are searching for."
			End If
					
		Case "Membership Number"
			If strSearchFor <> "" Then 
				Response.Write "Displaying Member with a MEMBERSHIP NUMBER of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the MEMBERSHIP NUMBER that you are searching for."
			End If

			 
		Case "VESA Unit"
			Call showVESAUnit(Conn, rs, "VESA_tblUnit", strSearchFor, strVESAUnit)
						
			If strSearchFor <> "" Then
			%>
				<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>Displaying All Audited Members in</td>
				</tr>				 
				<tr>
					<td>
						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td>UNIT/DESIGNATION:</td>
							<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
							<%							 
							If strSearchFor = "-1" Or strSearchFor = "0" Or strSearchFor = "" Then
								Response.Write "<td><font color=""#ff0000"">No VESA Unit</font></td>" 
							Else
								Response.Write "<td><font color=""#ff0000"">" & UCase(strVESAUnit) & "</font></td>"
							End If 
							%>				 
						</tr>
						</table>
					</td>
				</tr>
				</table>
			<%  
			 
				Else 
					Response.Write "Error in your Search. Please specify the UNIT/DESIGNATION that you are searching for."
				End If

		Case "SES Region"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Audited Members in the SES REGION of <font color=""#ff0000"">"
				Select Case strSearchFor
					Case 1
						Response.Write "Central"
					Case 2
						Response.Write "South West"
					Case 3
						Response.Write "East"
					Case 4
						Response.Write "North East"
					Case 5
						Response.Write "Mid West"
					Case 6
						Response.Write "North West"
				End Select
						 Response.Write "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the SES Region that you are searching for."
			End If
     			
		Case Else
			Response.Write "Displaying All Audited Members"
	End Select
End Sub 

Sub viewMembers()
%>
	<!--/* Start Here */-->
	<table border="0" cellspacing="0" cellpadding="0" width="940">
	<tr>
		<td align="center" width="940">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
				<td>
				<!--/* header */-->
				<h1 class="title"><a href="#">
				<% 
				Call DisplayTitle(strSearch)
			   
				Response.Write "</a></h1>"

				'- Check page count to prevent bombing when zero results are returned!-----------------
				If iRecordCount = 0 Then
					Response.Write "<p class=""byline""><b>No records found!</b></p>"
					Response.Write "</td></tr>"
					Response.Write "</table>"

				Else
					rstSearch.AbsolutePage = iPageCurrent
					Response.Write "<p class=""byline""><strong><font color=""#c40000"">" & iRecordCount & " Records Found.</font></strong></p>" & vbCrLf
					Response.Write "</td></tr>"
				%>

			<tr height="10"><td><img src="images/spacer.gif" width="1" height="10" border="0"></td></tr>

			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="1140">
					<tr>
						<td><strong><font color="#ff0000">There are <%= iRecordCount %> active members.</font></strong></td>
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

			<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>
 
			<tr>
				<td>
					<table border="0" style="width:1140px !important;">
					<tr>
						<td>Page <strong><%=iPageCurrent%></strong> of <strong><%=iPageCount%></strong></td>
						<td>
						<div align="right">					
							<div class="pages" style="width:600px !important;">
								<% Call databasePaging() %>
							</div>
						</div>
						</td>
					</tr>
					</table>
				</td>
			</tr>

			<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>

			<tr>
				<td bgcolor="#eeeeee">
					<!--/* Output Search */-->
					<table id="main_table" border="0" align="center" cellspacing="2" cellpadding="1" width="1140">
					<tr align="center" height="30">
						<td class="tab_header_cell"><b>Audit Action</b></td>
						<td class="tab_header_cell"><b>Membership <br /> Number</b></td>
						<td class="tab_header_cell"><b>Name / Organisation</b></td>
						<td class="tab_header_cell"><b>Email Address</b></td>
						<td class="tab_header_cell"><b>Publication Assigned</b></td>
						<td class="tab_header_cell"><b>Unit / Designation</b></td>
						<td class="tab_header_cell"><b>SES Region</b></td>
					</tr>
			   
					<%
					Do While Not rstSearch.EOF And rstSearch.AbsolutePage = iPageCurrent
						AuditID					= rstSearch("AuditID") & ""
						ActionType				= rstSearch("ActionType") & ""
						ChangedBy				= rstSearch("ChangedBy") & ""
						ActionDateTime			= rstSearch("ActionDateTime") & ""
						RecipientID				= rstSearch("RecipientID") & ""
						Surname_Organization	= rstSearch("Surname_Organization") & ""
						FirstName				= rstSearch("FirstName") & ""
						Address					= rstSearch("Address") & ""
						Suburb					= rstSearch("Suburb") & ""
						Postcode				= rstSearch("Postcode") & ""
						State					= rstSearch("State_Name") & ""
						MembershipNumber		= rstSearch("MembershipNumber") & ""
						EmailAddress			= rstSearch("MemberEmailAddress") & ""
						PhoenixCopies			= rstSearch("PhoenixCopies") & ""
						VESAPocketDiary			= rstSearch("VESAPocketDiary") & ""
						VESAWallCalendar		= rstSearch("VESAWallCalendar") & ""
						VESAUnit				= rstSearch("VESAUnit") & ""
						SESRegion				= rstSearch("SESRegion") & ""

						Select Case ActionType
						   Case "Add"
							  objValueAction = "A"
						   Case "Update"
							  objValueAction = "U"
						   Case "Delete"
							  objValueAction = "D"
						   Case "Activate"
							  objValueAction = "AC"
						End Select
							
						j = j + 1
						Response.write "<tr height=""20"" class=""listTableText" & (j And 1) & """>"
					%>
						<td>
						<div style="padding:5px !important; width:200px !important;">
							 Audit ID: <span style="color: #0000a0; font-weight:bold"><%=AuditID%></span><br />
							 Action: <font color="#ff0000"><%=ActionType%></font><br /> 
							 <%
								 If IsNull(rstSearch.Fields("ChangedBy")) Or rstSearch.Fields("ChangedBy") = ""  Then
									Response.Write "Changed By: <span style=""color: #0000a0"">" & VESAUnit & " Administrator</span><br />"
								 Else
									Response.Write "Changed By: <span style=""color: #0000a0"">" & rstSearch.Fields("ChangedBy") & "</span><br />"
								 End If			
								 %>		 
								 Date Actioned: <span style="color: #0000a0"><%=FormatAuditDate(ActionDateTime)%></span>
							 </div>
							 </td>

							<td align="center">
							<%
							If IsNull(rstSearch.Fields("MembershipNumber")) Then
								Response.Write "<font color=""#ff0000"">No Membership <br /> Number provided</font>"
							Else
								Response.Write rstSearch.Fields("MembershipNumber")
							End If			
							%>
							</td>

							<td>
							<div style="padding:5px !important;">
								<p><span style="color: #0000a0; font-weight:bold"><%=RecipientID%></span></p>
								<p>
								<%
								If IsNull(rstSearch.Fields("FirstName")) Or rstSearch.Fields("FirstName") = ""  Then
									Response.Write UCase(Surname_Organization)
								Else
									Response.Write "<strong>" & FirstName & "&nbsp;" & UCase(Surname_Organization) & "</strong>"
								End If 
								%>
								</p>
								<p>
								<%=Address%> <br />
								<%=UCase(Suburb)%>&nbsp;<%=State%>&nbsp;<%=Postcode%>
								</p>
							</div>
							</td>
									 
							<td align="center">
							<%
							If IsNull(rstSearch.Fields("MemberEmailAddress")) Then
								Response.Write "<font color=""#ff0000"">No Email Address provided</font>"
							Else
								Response.Write LCase(rstSearch.Fields("MemberEmailAddress"))
							End If			
							%>
							</td>
							 
							<td>
							<div style="padding:5px !important; width:150px !important;">
								<div style="padding-bottom:15px !important;">
									<div style="width: 80%; float:left">Phoenix Copies:</div>
									<div style="width: 20%; float:right"><%=PhoenixCopies%></div>
								</div>

								<div style="padding-bottom:15px !important;">
									<div style="width: 80%; float:left">Pocket Diary:</div>
									<div style="width: 20%; float:right"><%=VESAPocketDiary%></div>
								</div>

								<div style="padding-bottom:15px !important;">
									<div style="width: 80%; float:left">Wall Calendar:</div>
									<div style="width: 20%; float:right"><%=VESAWallCalendar%></div>
								</div>
							 </div>
							 </td>

							 <td align="center"><font color="#ff0000"><%=UCase(VESAUnit)%></font></td>
							 <td align="center"><%=SESRegion%></td>
						</tr>
		 
						<%
							rstSearch.MoveNext
						Loop
						%>
						</table>
					</td>
				</tr>
			</table>  
		</td>
	</tr>
 
	<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>
 
	<tr>
		<td>
			<table border="0" style="width:1140px !important;">
			<tr>
				<td>Page <strong><%=iPageCurrent%></strong> of <strong><%=iPageCount%></strong></td>
				<td>
				<div align="right">					
					<div class="pages" style="width:600px !important;">
						<% Call databasePaging() %>
					</div>
				</div>
				</td>
			</tr>
			</table>
			<% End If %>
		</td>
	</tr>

	<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>

	<tr>
		<td width="100%">
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><button type="button" class="pure-button" onClick="goBack()">Back to Main</button></td>
	  
				<% If iRecordCount <> 0 Then %>
					<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
					<td><button type="button" class="pure-button" onClick="exportFile()">Export as Excel</button></td>
				<% End If %>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	<!--/* End Here */-->
   
	<% If iRecordCount <> 0 Then %>
		<form name="ExportForm" id="ExportForm" action="VESAExport.asp" method="post">
		<input type="hidden" name="search" value="Audit">
		<input type="hidden" name="searchFor" value="1">
		</form>
	<% End If 
End Sub

Sub databasePaging()
	If iPageCurrent > 1 Then 
	%>
		<a href="VESAViewHistory.asp?page=<%=iPageCurrent - 1%>">&lt;&nbsp;Prev</a>
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
			Response.Write "<a href=""VESAViewHistory.asp?page=" & I & """>" & I & "</a>" & vbCrLf
		End If
	'- I
	Next 
		If iPageCurrent < iPageCount Then
			Response.Write "<a href=""VESAViewHistory.asp?page=" & iPageCurrent + 1 & """>Next&nbsp;&gt;</a>"
														
		Else
			Response.Write "<span class=""disabled"">Next&nbsp;&gt;</span>"
		End If
	'--------------------------------------------------------------------------------------
End Sub
%>
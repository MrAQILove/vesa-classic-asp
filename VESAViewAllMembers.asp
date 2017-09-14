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
   Const PAGE_SIZE = 300	' The size of our pages.

   '- Declare our variables... always good practice!
   Dim rstSearch			' ADO recordset
   Dim strSQL				' The SQL Query we build on the fly
   Dim strSearchFor			' The text being looked for

   Dim iPageCurrent			' The page we're currently on
   Dim iPageCount			' Number of pages of records
   Dim iRecordCount			' Count of the records returned
   Dim I					' Standard looping variable

   '- Retreive the term being searched for.  I'm doing it on
   '- the QS since that allows people to bookmark results.
   '- You could just as easily have used the form collection.
   strSearch = Request.Form("search")
   
   '- Retrieve page to show or default to the first
   If Request.QueryString("page") = "" Then
      iPageCurrent = 1
   Else
      iPageCurrent = CInt(Request.QueryString("page"))
   End If

   EstablishConnection()
   
   '- Build our query based on the input.
   strSQL = "SELECT * FROM VESA_tblMembers M"
   strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
   strSQL = strSQL & " LEFT JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
   strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
   strSQL = strSQL & " WHERE (U.IsActive = '1')"
   'strSQL = strSQL & " ORDER BY Surname_Organization ASC"
   strSQL = strSQL & " ORDER BY RecipientID DESC"

   '- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
   Set rstSearch = Server.CreateObject("ADODB.Recordset")
   rstSearch.PageSize  = PAGE_SIZE
   rstSearch.CacheSize = PAGE_SIZE

   '- Open our recordset
   rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

   '- Get a count of the number of records and pages for use in building the header and footer text.
   iRecordCount = rstSearch.RecordCount
   iPageCount   = rstSearch.PageCount

   DisplayHeader("VESA Members Database : Displaying All Members")
      
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
			<link rel="stylesheet" href="css/default.css" type="text/css" media="screen" />
			<link rel="stylesheet" href="css/buttons.css">
			<link rel="stylesheet" href="css/forms.css">
			<link rel="stylesheet" href="css/base.css">
			<script language="javascript">
			<!--
			// Log out
			function logOut() {
			   document.location.href = "AdminLogin.asp";
			}

			// Back to Main
			function goBack() {
			   document.location.href = "VESAMain.asp";
			}

			// Show Member History
			function showHistory() 
			{
			   if (<%=Session("VESAID")%> == 1) {
				  document.ShowHistoryForm.submit();
			   }
			}

			// Export Member List into an Excel File 
			function exportXLSFile(){
			   document.ExportXLSForm.submit();
			}

			// Export Member List into an Excel File 
			function exportCSVFile(){
			   document.ExportCSVForm.submit();
			}
			//-->
			</script>
		</head>
<% End Sub

Sub OutputPage() %>
	<body>
		<nav id="menu">
			<ul id="main">
				<li><a href="VESAMain.asp">Home</a></li>
				<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
				<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
				<li><a href="contactUs.html">Contact Us</a></li>
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
			
		<!-- start section -->
		<section>
			<% Call viewMembers() %>
			<div style="clear: both;">&nbsp;</div>
		</section>
		<!-- end section -->
		
		<footer id="footer">
			<p class="copyright">&copy;&nbsp;&nbsp;2008 -  <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
		</footer>
		</body>
	</html>
<% End Sub 

Sub viewMembers()
%>
	<!--/* Start Here */-->
	<table border="0" cellspacing="0" cellpadding="0" border="0" width="100%">
	<tr>
		<td align="center">
			<table border="0" align="center" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td>
				<!--/* header */-->
				<h1 class="title"><a href="#">Displaying All Members</a></h1>
				<%
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
					<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td><strong><font color="#ff0000">There are <%= iRecordCount %> active members.</font></strong></td>
						<td align="right">
						<span style="color: #ff0000; font-weight:bold">
						<% Call recordsDisplay() %>
						</span>
						</td>
					</tr>
					</table>
				</td>
			</tr>

			<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>
 
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="100%">
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

			<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>

			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td><button type="button" class="pure-button" onClick="goBack()">Back</button></td>
						<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>

						<td><button type="button" class="pure-button" onClick="showHistory()">History</button></td>
						<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>

						<td><button type="button" class="pure-button" onClick="exportXLSFile()">Export as Excel</button></td>
						<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
						
						<td><button type="button" class="pure-button" onClick="exportCSVFile()">Export as CSV</button></td>
					</tr>
					</table>
				</td>
			</tr>
		 			
			<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>

			<tr>
				<td bgcolor="#eeeeee">
					<table id="main_table" border="0" align="center" cellspacing="2" cellpadding="1" width="100%">
					<tr align="center" height="30">
						<td class="tab_header_cell"><b>Membership <br /> Number</b></td>
						<td class="tab_header_cell"><b>Name / Organisation</b></td>
						<td class="tab_header_cell"><b>Email Address</b></td>
						<td class="tab_header_cell"><b>Publication Assigned</b></td>
						<td class="tab_header_cell"><b>Unit / Designation</b></td>
						<td class="tab_header_cell"><b>SES Region</b></td>
					</tr>		
			   
					 <%
					 Do While Not rstSearch.EOF And rstSearch.AbsolutePage = iPageCurrent
						RecipientID				= rstSearch("RecipientID") & ""
						IDArray					= CInt(rstSearch("RecipientID")) & ""
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
						
						j = j + 1
						Response.write "<tr height=""20"" class=""listTableText" & (j And 1) & """>"
					%>
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
								If Len(Surname_Organization) > 15 Then 
									Response.Write FirstName & "<br />" 
									Response.Write UCase(Surname_Organization)
									
								ElseIf InStr(1, FirstName, "C/-") > 0 Then
										Response.Write FirstName & "<br />" 
										Response.Write UCase(Surname_Organization)
								Else
									Response.Write "<strong>" & FirstName & "&nbsp;" & UCase(Surname_Organization) & "</strong>"
								End If 
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
							Response.Write rstSearch.Fields("MemberEmailAddress")
						End If			
						%>
						</td>
					 
						<td>
						<div style="padding:5px !important; width:150px !important;">
							<div style="padding-bottom:15px !important;">
								<div style="width: 80%; float:left">Phoenix Copies:</div>
								<div style="width: 20%; float:right">
									<% 
									If PhoenixCopies > 1 Then 
										Response.Write "<strong><font color=""#ff0000"">" & PhoenixCopies & "</font></strong>"
									Else
										Response.Write PhoenixCopies
									End If 
									%>
								</div>
							</div>

							<div style="padding-bottom:15px !important;">
								<div style="width: 80%; float:left">Pocket Diary:</div>
								<div style="width: 20%; float:right">
								<% 
									If VESAPocketDiary > 1 Then 
										Response.Write "<strong><font color=""#ff0000"">" & VESAPocketDiary & "</font></strong>"
									Else
										Response.Write VESAPocketDiary
									End If 
								%>
								</div>
							</div>

							<div style="padding-bottom:15px !important;">
								<div style="width: 80%; float:left">Wall Calendar:</div>
								<div style="width: 20%; float:right">
									<% 
									If VESAWallCalendar > 1 Then 
										Response.Write "<strong><font color=""#ff0000"">" & VESAWallCalendar & "</font></strong>"
									Else
										Response.Write VESAWallCalendar
									End If 
									%>
								</div>
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
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><button type="button" class="pure-button" onClick="goBack()">Back</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>

				<td><button type="button" class="pure-button" onClick="showHistory()">History</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>

				<td><button type="button" class="pure-button" onClick="exportXLSFile()">Export as Excel</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
				
				<td><button type="button" class="pure-button" onClick="exportCSVFile()">Export as CSV</button></td>
			</tr>
			</table>
		</td>
	</tr>
 
	
	<tr><td valign="top"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>
 
	<tr>
		<td>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
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
			<% End If %>
		</td>
	</tr>
	</table>
	<!--/* End Here */-->

	<form name="ShowHistoryForm" id="ShowHistoryForm" action="VESAAudit.asp" method="post">
	<input type="hidden" name="search" value="All">
	</form>

	<form name="ExportXLSForm" id="ExportXLSForm" action="VESAExport.asp" method="post">
	<input type="hidden" name="search" value="All">
	<input type="hidden" name="type" value="XLS">
	</form>

	<form name="ExportCSVForm" id="ExportCSVForm" action="VESAExport.asp" method="post">
	<input type="hidden" name="search" value="All">
	<input type="hidden" name="type" value="CSV">
	</form>
<% End Sub 

Sub databasePaging()
	If iPageCurrent < 10 Then
      StartPage = 1
      EndPage = 10
	Else
		  StartPage = iPageCurrent - 5
		  EndPage = iPageCurrent + 4
		  If EndPage > iPageCount Then
				EndPage = iPageCount
				StartPage = EndPage - 9
		  End If                              
	End if
							
	If iPageCurrent > 1 Then 
	%>
		<a href="VESAViewAllMembers.asp?page=<%=iPageCurrent - 1%>">&lt;&nbsp;Prev</a>
	<%
	Else
		Response.Write "<span class=""disabled"">&lt;&nbsp;Prev</span>"
	End If
	
	For I = StartPage To EndPage
		  If I <> iPageCurrent Then
				Response.Write "<a href=""VESAViewAllMembers.asp?page=" & I & """>" & I & "</a>" & vbCrLf
		  Else
				'The active page
				Response.Write "<span class=""current"">" & I & "</span>" & vbCrLf
		  End If
		  'Writes | as a separator if we're not at the last link
		  'If I <> iPageCount Then Response.Write(" | ")                          
	Next
		If iPageCurrent < iPageCount Then
			Response.Write "<a href=""VESAViewAllMembers.asp?page=" & iPageCurrent + 1 & """>Next&nbsp;&gt;</a>"
														
		Else
			Response.Write "<span class=""disabled"">Next&nbsp;&gt;</span>"
		End If
End Sub

Sub recordsDisplay()
	If iRecordCount > PAGE_SIZE Then
		
		For k = 1 to iPageCount 
			If k = iPageCurrent Then 
				
				m = 1
				n = k - 1
				
				If k = 1 Then 
					Response.Write "Displaying Records " & m & " - " & (PAGE_SIZE * k) & " of " & iRecordCount & "&nbsp;&bull;&nbsp;"
				
				ElseIf k = iPageCount Then 
					Response.Write "Displaying Records " & (n * PAGE_SIZE + m) & " - " & iRecordCount & " of " & iRecordCount & "&nbsp;&bull;&nbsp;"
				
				Else
					Response.Write "Displaying Records " & (n * PAGE_SIZE + m) & " - " & (PAGE_SIZE * k) & " of " & iRecordCount & "&nbsp;&bull;&nbsp;" 
				End If 

				Response.Write "Page " & k & " of " & iPageCount & ":"
			End If 
		Next
	Else
		If iRecordCount <> 1 Then
			Response.Write "Displaying " & iRecordCount & " Records &nbsp;&bull;&nbsp; Page " & iPageCurrent & " of " & iPageCount & ":"
		Else
			Response.Write "Displaying " & iRecordCount & " Record &nbsp;&bull;&nbsp; Page " & iPageCurrent & " of " & iPageCount & ":"
		End If 
	End If
End Sub
%>
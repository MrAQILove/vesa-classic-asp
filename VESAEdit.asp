<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE FILE="include/include.asp"-->
<!--#INCLUDE FILE="include/adovbs.inc"-->
<!--#INCLUDE FILE="include/functions.asp"-->

<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = true 

Response.AddHeader "Pragma", "no-store"
Response.CacheControl = "no-store"

If Session("UserLoggedIn") <> "true" Then
   If Session("AccessRights") = "Level 1" Then
      Response.Redirect "AdminLogin.asp"
   Else
      Response.Redirect "VESAUnitLogin.asp"
   End If  

Else
	Dim rsResult
	Dim strID
	Dim RecipientID, Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress
	Dim PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID, SESRegion, SESRegionID
	
	Dim objValueVESAUnitID
	objValueVESAUnitID = Session("VESAUnitID")
   
	strID = CLng(Request.Form("RecipientID"))

	EstablishConnection()

	strSQL = "SELECT * FROM VESA_tblMembers M"
	strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
	strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
	strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
	strSQL = strSQL & " WHERE M.RecipientID ='" & strID & "'"
				  
	Set rsResult = Server.CreateObject("ADODB.Recordset")
				  
	rsResult.Open strSQL, Conn, adOpenStatic, adLockReadOnly

	RecipientID				= rsResult("RecipientID") & ""
	Surname_Organization	= rsResult("Surname_Organization") & ""
	FirstName				= rsResult("FirstName") & ""
	Address					= rsResult("Address") & ""
	Suburb					= rsResult("Suburb") & ""
	Postcode				= rsResult("Postcode") & ""
	StateID					= rsResult("StateID") & ""
	MembershipNumber		= rsResult("MembershipNumber") & ""
	MemberEmailAddress		= rsResult("MemberEmailAddress") & ""
	PhoenixCopies			= rsResult("PhoenixCopies") & ""
	VESAPocketDiary			= rsResult("VESAPocketDiary") & ""
	VESAWallCalendar		= rsResult("VESAWallCalendar") & ""
	VESAUnitID				= rsResult("VESAUnitID") & ""
	SESRegionID				= rsResult("SESRegionID") & ""
	SESRegion				= rsResult("SESRegion") & ""

	'- Display the Edit page with no Errors
	Call EditMember()

	rsResult.Close
	Set rsResult = Nothing
	CloseConnection()
End If

Sub Displayheader()
	Select Case Session("Errors")
		Case 0
			If Not IsNull(rsResult.Fields(2)) And Not IsNull(rsResult.Fields(3)) Then
				Response.Write "Edit Member - " & FirstName & "&nbsp;" & UCase(Surname_Organization)
			Else
				Response.Write "Edit Member - " & UCase(Surname_Organization)     
			End If
	
		Case Else
			If Not IsNull(Session("Surname_Organization")) And Not IsNull(Session("FirstName")) Then
				Response.Write "Edit Member - " & Session("FirstName") & "&nbsp;" & UCase(Session("Surname_Organization"))
			Else
				Response.Write "Edit Member - " & UCase(Session("Surname_Organization"))     
			End If
	End Select 
End Sub

Sub EditMember()
%>
	<!DOCTYPE html>
	<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
				<script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
			<title>VESA Members Database :<% Call Displayheader() %> - <% Response.Write VESAWallCalendar %></title>
			<meta name="keywords" content="" />
			<meta name="VESA Members Database" content="" />
			<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
			<link rel="stylesheet" href="css/buttons.css">
			<link rel="stylesheet" href="css/forms.css">
			<link rel="stylesheet" href="css/base.css">
			<script language="JavaScript">
			<!--
			// view all members 
			function viewMembers() 
			{
			   if (<%=Session("VESAID")%> == 1) {
				  document.viewMembersForm.submit();
			   }
			}

			// delete a current member 
			function deleteMember() 
			{
			   if (<%=Session("VESAID")%> == 1) {
				  document.deleteMemberForm.submit();
			   }
			}
			//-->
			</script>
		</head>
		
		<body>
			<div id="wrapper">
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
				
				<!-- start page -->
				<section id="page">
					<%
					If Session("AccessRights") = "Level 1" Then
						Call adminMainMenu()
					Else 
						Call VESAUnitMainMenu()
					End If 
					%>
					
					<!-- start content -->
					<article>
						<h1 class="title"><a href="#"><% Call Displayheader() %></a></h1>
						<p class="byline"><b>Please fill out the form.</b></p>
						<div class="entry">
							<%
							'- On the first time that this page loads, session("errors") has no value and so equals 0. 
							'- On subsequent visits, session("errors will have a value of 1 or more if errors were made") 
							If Session("Errors") = 0 Then 
								Response.Write "<p><b>Fields that have a <font color=""#990000""><b>*</b></font> next to them are Mandatory.</b></p>" & vbCrLf
							Else
								'- Errors were made so list them in the rest of the table reset our error counter
								Session("Errors") = "0"
							%>
								<div id="sidebar3" class="sidebar4">
									<ul>
										<li> 
											<h2>There are ERRORS in your form. Please make the following changes before clicking the submit button</h2>
											<ul>
												<%
												'- These session variables are set in the second page
												'- If there were errors value is "F"  
												'- Error Ouput for LastName ---------------------------------------------------
												If Session("badSurname_Organization") = "T" Then
													Response.write "<li><a href=""#"">Enter only letters in your <b>SURNAME/ORGANIZATION</b> field. <br />" & vbCrLf
													Response.write "Avoid using (~ ! @ # % ^ & * ( ) _ + = ` { } [ ] | : ; \ "" "" ' < > , . ?) characters and numbers.</a></li>" & vbCrLf
													Session("badSurname_Organization") = "F"
												End If
												'--------------------------------------------------------------------------------

												'- Error Ouput for FirstName ----------------------------------------------------
												If Session("badFirstName") = "T" Then
													Response.write "<li><a href=""#"">Enter only letters in your <b>FIRSTNAME</b> field. <br />" & vbCrLf
													Response.write "Avoid using (~ ! @ # % ^ & * ( ) _ + = ` { } [ ] | : ; \ "" "" ' < > , . ?) characters and numbers.</a></li>" & vbCrLf
													Session("badFirstName") = "F"
												End If
												'--------------------------------------------------------------------------------

												'- Error Ouput for Address ------------------------------------------------------
												If Session("badAddress") = "T" Then
													Response.write "<li><a href=""#"">Enter only letters in your <b>ADDRESS</b> field. <br />" & vbCrLf
													Response.write "Avoid using (~ ! @ # % ^ & * ( ) _ + = ` { } [ ] | : ; \ "" "" ' < > . ?) characters.</a></li>" & vbCrLf
													Session("badAddress") = "F"
												End If
												'--------------------------------------------------------------------------------

												'- Error Ouput for Suburb -------------------------------------------------------
												If Session("badSuburb") = "T" Then
													Response.write "<li><a href=""#"">Enter only letters in your <b>SUBURB</b> field. <br />" & vbCrLf
													Response.write "Avoid using (~ ! @ # % ^ & * ( ) _ + = ` { } [ ] | : ; / \ "" "" ' < > , . ?) characters and numbers.</a></li>" & vbCrLf
													Session("badSuburb") = "F"
												End If
												'--------------------------------------------------------------------------------
									
												'- Error Ouput for Postcode -----------------------------------------------------
												If Session("badPostcode") = "T" Then
													Response.write "<li><a href=""#"">Please enter a valid <b>POSTCODE</b>. Make sure it's a 4 digit numbers.</a></li>" & vbCrLf
													Session("badPostcode") = "F"
												End If

												If Session("badPostcode") = "T1" Then
													Response.write "<li><a href=""#"">Enter only numbers in your <b>POSTCODE</b> field. <br />" & vbCrLf
													Response.write "Avoid using (~ ! @ # % ^ & * ( ) - _ + = ` { } [ ] | : ; / \ "" "" ' < > , . ?) characters.</a></li>" & vbCrLf
													Session("badPostcode") = "F1"
												End If
												'--------------------------------------------------------------------------------

												'- Error Ouput for State --------------------------------------------------------
												If Session("badStateID") = "T" Then
													Response.write "<li><a href=""#"">Please choose your <b>STATE</b>.</a></li>" & vbCrLf
													Session("badStateID") = "F"
												End If
												'--------------------------------------------------------------------------------

												'- Error Ouput for Membership Number --------------------------------------------
												If Session("AccessRights") = "Level 1" Then
													If Session("badMembershipNumber") = "T" Then
														Response.write "<li><a href=""#"">Enter only numbers in your <b>MEMBERSHIP NUMBER</b> field. <br />" & vbCrLf
														Response.write "Avoid using (~ ! @ # % ^ & * ( ) - _ + = ` { } [ ] | : ; / \ "" "" ' < > , . ?) characters and numbers.</a></li>" & vbCrLf
														Session("badMembershipNumber") = "F"
													End If
												End If
												'--------------------------------------------------------------------------------

												'- Error Ouput for Email --------------------------------------------------------
												If Session("badMemberEmailAddress") = "T" Then
													Response.write "<li><a href=""#"">Invalid <b>EMAIL</b> Address.</a></li>" & vbCrLf
													Session("badMemberEmailAddress") = "F"
												End If
												'--------------------------------------------------------------------------------

												If Session("AccessRights") = "Level 1" Then
													'- Error Ouput for VESA Unit ----------------------------------------------------
													If Session("badVESAUnitID") = "T" Then
														Response.write "<li><a href=""#"">Please choose your <b>VESA UNIT</b>.</a></li>" & vbCrLf
														Session("badVESAUnitID") = "F"
													End If
													'--------------------------------------------------------------------------------
												End If
												'- End the errors table
												%>
											</ul>
										</li>
									</ul>
								</div>
								<div style="clear: both;">&nbsp;</div>
							<% End If %>
							
							<form class="pure-form pure-form-aligned" name="AddForm" method="post" action="VESACheckMembers.asp">
								<input type="hidden" name="ActionType" value="Update">
								<input type="hidden" name="RecipientID" value="<%=RecipientID%>">
								<% If Session("AccessRights") = "Level 5" Then %>
									<input type="hidden" name="VESAUnitID" value="<%=VESAUnitID%>">
								<% End If %>
								<input type="hidden" name="SESRegionID" value="<%=SESRegionID%>">
								<input type="hidden" name="SESRegion" value="<%=SESRegion%>">

								<fieldset>
									<div class="pure-control-group">
										<label>Surname/Organization:</label>
										<input type="text" name="Surname_Organization" placeholder="Surname/Organization" class="pure-input-1-2" value="<%=Surname_Organization%>" required>
										<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
									</div>

									<div class="pure-control-group">
										<label>First Name:</label>
										<input type="text" name="FirstName" placeholder="First Name" value="<%=FirstName%>" required>
										<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
									</div>

									<div class="pure-control-group">
										<label>Address:</label>
										<input type="text" name="Address" placeholder="Address" class="pure-input-1-2" value="<%=Address%>" required>
										<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside><br />
										<aside class="pure-form-message-inline" style="margin-left:145px;">
										<code>
										Only <font color="#990000"><b>(/ - ,)</b></font> characters are valid. Avoid using (~ ! @ # % ^ & * ( ) _ + = ` { } [ ] | : ; \ "" "" ' < > . ?) characters.
										</code>
									</div>

									<div class="pure-control-group">
										<label>Suburb:</label>
										<input type="text" name="Suburb" placeholder="Suburb" value="<%=Suburb%>" required>
										<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
									</div>

									<div class="pure-control-group">
										<label>Postcode:</label>
										<input type="text" name="Postcode" placeholder="Postcode" value="<%=Postcode%>" required>
										<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
									</div>

									<div class="pure-control-group">
										<label for="state">State:</label>
										<%
										If StateID = "" Then
												Call showSelectedValue(Conn, rsState, "MembersDB_tblState", "StateID", "State_Name", "StateID") 
										Else %>
											<select id="StateID" name="StateID" class="pure-input-medium" required>
											<option value="-1" <%If StateID = "-1" Then Response.Write "class=""selectedItem"" selected"%>>Please Choose</option>
											<option value="1" <%If StateID = "1" Then Response.Write "class=""selectedItem"" selected"%>>ACT</option>
											<option value="2" <%If StateID = "2" Then Response.Write "class=""selectedItem"" selected"%>>NSW</option>
											<option value="3" <%If StateID = "3" Then Response.Write "class=""selectedItem"" selected"%>>NT</option>
											<option value="4" <%If StateID = "4" Then Response.Write "class=""selectedItem"" selected"%>>QLD</option>
											<option value="5" <%If StateID = "5" Then Response.Write "class=""selectedItem"" selected"%>>SA</option>
											<option value="6" <%If StateID = "6" Then Response.Write "class=""selectedItem"" selected"%>>TAS</option>
											<option value="7" <%If StateID = "7" Then Response.Write "class=""selectedItem"" selected"%>>VIC</option>
											<option value="8" <%If StateID = "8" Then Response.Write "class=""selectedItem"" selected"%>>WA</option>
											<option value="9" <%If StateID = "9" Then Response.Write "class=""selectedItem"" selected"%>>NZ</option>
											</select> 
										<% End If %>
										<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
									</div>

									<div class="pure-control-group">
										<label>Membership Number:</label>
										<input type="text" name="MembershipNumber" placeholder="Membership Number" value="<%=MembershipNumber%>" required>
									</div>

									<div class="pure-control-group">
										<label>Email Address:</label>
										<input type="email" name="MemberEmailAddress" placeholder="Email Address" class="pure-input-1-2" value="<%=MemberEmailAddress%>">
									</div>

									<div class="pure-control-group">
										<label>Phoenix Copies:</label>
										<input type="text" name="PhoenixCopies" placeholder="Phoenix Copies" value="<%=PhoenixCopies%>">
									</div>

									<div class="pure-control-group">
										<label>VESA Pocket Diary:</label>
										<input type="text" name="VESAPocketDiary" placeholder="VESA Pocket Diary" value="<%=VESAPocketDiary%>">
									</div>

									<div class="pure-control-group">
										<label>VESA Wall Calendar:</label>
										<input type="text" name="VESAWallCalendar" placeholder="VESA Wall Calendar" value="<%=VESAWallCalendar%>">
									</div>

									<% If Session("AccessRights") = "Level 1" Then %>
										<!--
										'----------------------------------------------------------------------------------
										'***** Show these fields if the login User is the Administrator *****		
										'----------------------------------------------------------------------------------
										-->

										<div class="pure-control-group">
											<label for="state">VESA Unit:</label>
											<%
											If VESAUnitID = "" Then
												Call showSelectedValue(Conn, rsVESAUnit, "VESA_tblUnit", "VESAUnitID", "VESAUnit", "VESAUnitID") 
											Else 
												Call selectedVESAUnitList(Conn, rsVESAUnit, "VESA_tblUnit", "VESAUnitID", "VESAUnitID", "" & VESAUnitID & "", "pure-input-medium")
											End If 
											%>
											<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
										</div>
									<% Else %>
										<!--
										'----------------------------------------------------------------------------------
										'***** Otherwise only show this field if the User is a VESA Unit user/member *****		
										'----------------------------------------------------------------------------------
										-->
										<div class="pure-control-group">
											<label>VESA Unit:</label>
											<% 
											Call showVESAUnit(Conn, rsVESAUnit, "VESA_tblUnit", objValueVESAUnitID, strVESAUnit) 
											Response.Write "<font color=""#990000""><b>" & strVESAUnit & "</b></font>"
											%>
										</div>
									<% End If %>

									<div class="pure-controls">
										<button type="submit" class="pure-button">Submit</button>
									</div>
								<fieldset>
								</form>

								<%
								If Session("AccessRights") = "Level 5" Then
									Call displayFORMLinks()
								End If 
								%>
							</div>
					</article>
					<!-- end content -->
					<div style="clear: both;">&nbsp;</div>
				</section>
				<!-- end page -->
			</div>
			<footer id="footer">
			  <p class="copyright">&copy;&nbsp;&nbsp;2008 -  <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
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
					<li><a href="VESAMain.asp">Home</a></li>
					<li><a href="VESASearch.asp">Search for Member</a></li>
					<li><a href="VESAViewAllMembers.asp">View All Members</a></li>
					<li><a href="VESAAddNewMember.asp">Add a New Member</a></li>
					<li><a href="VESADeleteMember.asp">Delete a Member</a></li>
					<li class="hover"><a href="#">Edit Member - <br />
					<%
					Select Case Session("Errors")
						Case 0
							If Not IsNull(rsResult.Fields(2)) And Not IsNull(rsResult.Fields(3)) Then
								Response.Write FirstName & "&nbsp;" & UCase(Surname_Organization)
							Else
								Response.Write UCase(Surname_Organization)     
							End If
			
						Case Else
							If Not IsNull(Session("Surname_Organization")) And Not IsNull(Session("FirstName")) Then
								Response.Write Session("FirstName") & "&nbsp;" & UCase(Session("Surname_Organization"))
							Else
								Response.Write UCase(Session("Surname_Organization"))     
							End If
					End Select 
					%>
					</a></li>
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
<% End Sub %>
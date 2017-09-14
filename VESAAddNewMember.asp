<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE FILE="include/include.asp"-->
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
	Dim objValueVESAUnitID
	objValueVESAUnitID = Session("VESAUnitID")
	
	EstablishConnection()
	Call AddNewMember()
	CloseConnection()
	
End If

Sub AddNewMember()
%>
	<!DOCTYPE html>
	<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
				<script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
			<title>VESA Members Database : Add a New Member</title>
			<meta name="keywords" content="" />
			<meta name="VESA Members Database" content="" />
			<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
			<link rel="stylesheet" href="css/buttons.css">
			<link rel="stylesheet" href="css/forms.css">
			<link rel="stylesheet" href="css/base.css">
			<link rel="stylesheet" href="css/grids.css">
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
				<h1 class="title"><a href="#">Add a New Member</a></h1>
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
						<input type="hidden" name="Actiontype" value="Add">
						<% If Not Session("AccessRights") = "Level 1" Then %>
							<input type="hidden" name="PhoenixCopies" value="1">
							<input type="hidden" name="VESAPocketDiary" value="1">
							<input type="hidden" name="VESAWallCalendar" value="1">
							<input type="hidden" name="VESAUnitID" value="<%=objValueVESAUnitID%>">
						<% End If %>

						<fieldset>
							<div class="pure-control-group">
								<label>Surname/Organization:</label>
								<input type="text" name="Surname_Organization" placeholder="Surname/Organization" class="pure-input-1-2" value="<%=Session("Surname_Organization")%>" required>
								<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
							</div>

							<div class="pure-control-group">
								<label>First Name:</label>
								<input type="text" name="FirstName" placeholder="First Name" value="<%=Session("FirstName")%>" required>
								<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
							</div>

							<div class="pure-control-group">
								<label>Address:</label>
								<input type="text" name="Address" placeholder="Address" class="pure-input-1-2" value="<%=Session("Address")%>" required>
								<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside><br />
								<aside class="pure-form-message-inline" style="margin-left:145px;">
								<code>
								Only <font color="#990000"><b>(/ - ,)</b></font> characters are valid. Avoid using (~ ! @ # % ^ & * ( ) _ + = ` { } [ ] | : ; \ "" "" ' < > . ?) characters.
								</code>
							</div>

							<div class="pure-control-group">
								<label>Suburb:</label>
								<input type="text" name="Suburb" placeholder="Suburb" value="<%=Session("Suburb")%>" required>
								<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
							</div>

							<div class="pure-control-group">
								<label>Postcode:</label>
								<input type="text" name="Postcode" placeholder="Postcode" value="<%=Session("Postcode")%>" required>
								<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
							</div>

							<div class="pure-control-group">
								<label for="state">State:</label>
								<%
								If Session("StateID") = "" Then
										Call showSelectedValue(Conn, rsState, "MembersDB_tblState", "StateID", "State_Name", "StateID") 
								Else %>
									<select id="StateID" name="StateID" class="pure-input-medium" required>
									<option value="-1" <%If Session("StateID") = "-1" Then Response.Write "class=""selectedItem"" selected"%>>Please Choose</option>
									<option value="1" <%If Session("StateID") = "1" Then Response.Write "class=""selectedItem"" selected"%>>ACT</option>
									<option value="2" <%If Session("StateID") = "2" Then Response.Write "class=""selectedItem"" selected"%>>NSW</option>
									<option value="3" <%If Session("StateID") = "3" Then Response.Write "class=""selectedItem"" selected"%>>NT</option>
									<option value="4" <%If Session("StateID") = "4" Then Response.Write "class=""selectedItem"" selected"%>>QLD</option>
									<option value="5" <%If Session("StateID") = "5" Then Response.Write "class=""selectedItem"" selected"%>>SA</option>
									<option value="6" <%If Session("StateID") = "6" Then Response.Write "class=""selectedItem"" selected"%>>TAS</option>
									<option value="7" <%If Session("StateID") = "7" Then Response.Write "class=""selectedItem"" selected"%>>VIC</option>
									<option value="8" <%If Session("StateID") = "8" Then Response.Write "class=""selectedItem"" selected"%>>WA</option>
									<option value="9" <%If Session("StateID") = "9" Then Response.Write "class=""selectedItem"" selected"%>>NZ</option>
									</select> 
								<% End If %>
								<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
							</div>

							<div class="pure-control-group">
								<label>Membership Number:</label>
								
								<% If Session("AccessRights") = "Level 5" Then %>								
									<input type="text" name="MembershipNumber" placeholder="Membership Number" value="<%=Session("MembershipNumber")%>" required>
									<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
								<% Else %>
									<input type="text" name="MembershipNumber" placeholder="Membership Number" value="<%=Session("MembershipNumber")%>">
								<% End If %>
							</div>

							<div class="pure-control-group">
								<label>Email Address:</label>
								<input type="email" name="MemberEmailAddress" placeholder="Email Address" class="pure-input-1-2" value="<%=Session("MemberEmailAddress")%>">
							</div>

							<% If Session("AccessRights") = "Level 1" Then %>
								<!--
								'----------------------------------------------------------------------------------
								'***** Show these fields if the login User is the Administrator *****		
								'----------------------------------------------------------------------------------
								-->

								<div class="pure-control-group">
									<label>Phoenix Copies:</label>
									<input type="text" name="PhoenixCopies" value="1">
								</div>

								<div class="pure-control-group">
									<label>VESA Pocket Diary:</label>
									<input type="text" name="VESAPocketDiary" value="1">
								</div>

								<div class="pure-control-group">
									<label>VESA Wall Calendar:</label>
									<input type="text" name="VESAWallCalendar" value="1">
								</div>

								<div class="pure-control-group">
									<label for="state">VESA Unit:</label>
									<%
									If Session("VESAUnitID") = "" Then
										Call showSelectedValue(Conn, rsVESAUnit, "VESA_tblUnit", "VESAUnitID", "VESAUnit", "VESAUnitID") 
									Else 
										Call selectedVESAUnitList(Conn, rsVESAUnit, "VESA_tblUnit", "VESAUnitID", "VESAUnitID", "" & Session("VESAUnitID") & "", "pure-input-medium")
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
					<li class="hover"><a href="#">Add a New Member</a></li>
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
<% End Sub %>


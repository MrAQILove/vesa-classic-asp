<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/database/vesa/Members/include/include.asp"-->
<!--#include virtual="/database/vesa/Members/include/functions.asp"--> 

<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = true 

Response.AddHeader "Pragma", "no-store"
Response.CacheControl = "no-store"

If Session("UserLoggedIn") <> "true" Then
	Response.Redirect "../AdminLogin.asp"

Else
	EstablishConnection()
	Call addNewVESAUnit()
	CloseConnection()
End If

Sub addNewVESAUnit()
%>
	<!DOCTYPE html>
	<html>
	<head>
		<meta charset="UTF-8">
		<!--[if lt IE 9]>
		  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
		<![endif]-->
		<title>VESA Members Database : Add a New VESA Unit</title>
		<meta name="keywords" content="" />
		<meta name="VESA Members Database" content="" />
		<link href="../css/default.css" rel="stylesheet" type="text/css" media="screen" />
		<link rel="stylesheet" href="../css/buttons.css">
		<link rel="stylesheet" href="../css/forms.css">
		<link rel="stylesheet" href="../css/base.css">
	</head>
	<body>
		<div id="wrapper">
			<!-- start nav -->
			<nav id="menu">
				<ul id="main">
					<li><a href="../VESAMain.asp">Home</a></li>
					<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
					<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
					<li><a href="../VESAContact.asp">Contact Us</a></li>
				</ul>
			</nav>
			<!-- end nav -->
			
			<!-- start header -->
			<header>
				<div id="logo">
					<h1><a href="#"><span></span></a></h1>
					<p></p>
				</div>
			</header>
			<!-- end header -->
			
			<!-- start section -->
			<section id="page">
				
				<!-- start aside -->
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
								<li><a href="ViewAllVESAUnits.asp">View All VESA Units</a></li>
								<li class="hover"><a href="#">Add a New VESA Unit</a></li>
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
				<!-- end aside -->
				
				<!-- start article -->
				<article id="content">
					<h1 class="title"><a href="#">Add a New VESA Unit</a></h1>
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
												'- Error Ouput for Unit ---------------------------------------------------------
												If Session("badVESAUnitName") = "T" Then
													Response.write "<li><a href=""#"">Enter only letters in your <b>VESA UNIT NAME</b> field. <br /> Avoid using (/ ' """" ! @ # $ % ^ & * -) characters and numbers.</a></li>" & vbCrLf
													Session("badVESAUnitName") = "F"
												End If
												'--------------------------------------------------------------------------------
												
												'- Error Ouput for Unit Password ------------------------------------------------
												If Session("badVESAUnitPassword") = "T" Then
													Response.write "<li><a href=""#"">The <b>PASSWORD</b> field must not be empty.</a></li>" & vbCrLf
													Session("badVESAUnitPassword") = "F"
												End If
												'--------------------------------------------------------------------------------

												'- Error Ouput for VESA Unit ----------------------------------------------------
												If Session("badSESRegionID") = "T" Then
													Response.write "<li><a href=""#"">Please choose your <b>SES Region</b>.</a></li>" & vbCrLf
													 Session("badSESRegionID") = "F"
												End If
												'--------------------------------------------------------------------------------
												'- End the errors table
												%>
										</ul>
									</li>
								</ul>
							</div>
								<div style="clear: both;">&nbsp;</div>
						<% End If %>

						<form class="pure-form pure-form-aligned" name="AddAdminUnitForm" method="post" action="VESACheckAdmin.asp">
						<input type="hidden" name="Actiontype" value="AddVESAUnit">
						
							<fieldset>
								<div class="pure-control-group">
									<label>VESA Unit:</label>
									<input type="text" name="VESAUnitName" placeholder="VESA Unit" class="pure-input-1-2" value="<%=Session("VESAUnitName")%>" required>
									<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
								</div>

								<div class="pure-control-group">
									<label for="password">Password</label>
									<input type="password" name="VESAUnitPassword" placeholder="Password" required>
									<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
								</div>

									<div class="pure-control-group">
											<label>Email Address:</label>
											<input type="email" name="UnitEmailAddress" placeholder="Email Address" class="pure-input-1-2" value="<%=Session("UnitEmailAddress")%>">
										</div>

										<div class="pure-control-group">
											<label>SES Region:</label>
											<%
											If Session("SESRegionID") = "" Then
												Call showSelectedValue(Conn, rsSESRegion, "VESA_tblSESRegion", "SESRegionID", "SESRegion", "SESRegionID") 
											Else %>
												<select id="SESRegionID" name="SESRegionID" class="pure-input-medium" required>
												<option value="0" <%If Session("SESRegionID") = "0" Then Response.Write "class=""selectedItem"" selected"%>>Please Choose</option>
												<option value="1" <%If Session("SESRegionID") = "1" Then Response.Write "class=""selectedItem"" selected"%>>Central</option>
												<option value="2" <%If Session("SESRegionID") = "2" Then Response.Write "class=""selectedItem"" selected"%>>East</option>
												<option value="3" <%If Session("SESRegionID") = "3" Then Response.Write "class=""selectedItem"" selected"%>>Mid West</option>
												<option value="4" <%If Session("SESRegionID") = "4" Then Response.Write "class=""selectedItem"" selected"%>>North East</option>
												<option value="5" <%If Session("SESRegionID") = "5" Then Response.Write "class=""selectedItem"" selected"%>>North West</option>
												<option value="6" <%If Session("SESRegionID") = "6" Then Response.Write "class=""selectedItem"" selected"%>>South West</option>
												</select> 
											<% End If %>
											<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
										</div>

										<div class="pure-control-group">
											<label>Is this a SES Region?</label>
											<input name="IsUnitSES" type="radio" value="1"> Yes
											<input name="IsUnitSES" type="radio" value="0"> No
										</div>

										<div class="pure-controls">
											<button type="submit" class="pure-button">Submit</button>
										</div>
									<fieldset>
							</form>	
						</div>
				</article>
				<!-- end article -->
				<div style="clear: both;">&nbsp;</div>
			</section>
		<!-- end section -->
		</div>
		<footer id="footer">
			<p class="copyright">&copy;&nbsp;&nbsp;2008 -  <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
		</footer>
	</body>
	</html>
<% End Sub %>

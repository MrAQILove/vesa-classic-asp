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

'If Session("UserLoggedIn") <> "true" Then
 Call DisplayHTML()

'Else
	'Connect to the database
 ''  	EstablishConnection()
	
''	Call UserLoggedIn()

   'Close database connection
 ''  CloseConnection()
'End If

Sub DisplayHTML() %>
	<!DOCTYPE html>
		<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
			  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
			<title>VESA Members Database : Contact Us</title>
			<meta name="keywords" content="" />
			<meta name="VESA Members Database" content="" />
			<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
			<link rel="stylesheet" href="css/buttons.css">
			<link rel="stylesheet" href="css/forms.css">
			<link rel="stylesheet" href="css/base.css">
		</head>
	<body>
		<div id="wrapper">
			<!-- start nav -->
			<nav id="menu">
				<ul id="main">
					<li><a href="VESAMain.asp">Home</a></li>
					<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
					<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
					<li class="hover"><a href="#">Contact Us</a></li>
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
				<%
				If Session("UserLoggedIn") <> "true" Then
					Call LogoSection()
				Else 
					Call MenuSection()
				End If
				%>
				
				<div style="clear: both;">&nbsp;</div>
			</section>
			<!-- end section -->
		</div>
		
		<footer id="footer">
			<p class="copyright">&copy;&nbsp;&nbsp;2008 - <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
		</footer>
	</body>
	</html>
<% End Sub

Sub ContctUSForm()
%>
	<h1 class="title"><a href="#">Send us a message</a></h1>
	<p class="byline">If you need any help or assistance in regards to your VESA Unit Passwords or you have any general enquiry, please fill in the form below and we shall get straight back to you.</p>
	<form class="pure-form pure-form-aligned" name="AddAdminUnitForm" method="post" action="VESAMail.asp">
		<fieldset>
			<div class="pure-control-group">
				<label>Name:</label>
				<input type="text" name="Name" placeholder="Name" class="pure-input-1-2" value="" required>
				<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
			</div>

			<div class="pure-control-group">
				<label>Email Address:</label>
				<input type="email" name="Email" placeholder="Email Address" class="pure-input-1-2" value="" required>
				<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
			</div>

			<div class="pure-control-group">
				<label>Subject:</label>
				<input type="text" name="Subject" placeholder="Subject" class="pure-input-1-2" value="" required>
				<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
			</div>

			<div class="pure-control-group">
				<label>Message/Comment:</label>
				<textarea class="pure-input-1-2" name="Message" placeholder="Message/Comment"></textarea>
				<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
			</div>
			
			<div class="pure-controls">
				<button type="submit" class="pure-button">Submit</button>
			</div>
		<fieldset>
	</form>	
<%
End Sub

Sub LogoSection() %>
	<!-- start aside -->
	<aside id="sidebar1" class="sidebar">
		<ul>
			<li> <br /><img src="images/Phoenix-Logo.jpg" width="220" height="218" alt="" /></li>
		</ul>
	</aside>
	<!-- end aside -->

	<!-- start article -->
	<article id="content">
		<% Call ContctUSForm() %>
	</article>
	<!-- end article -->
	
<% End Sub 

Sub MenuSection() %>
	<!-- start aside -->
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
	<!-- end aside -->

	<!-- start article -->
	<article id="content">
		<% Call ContctUSForm() %>
	</article>
	<!-- end article -->
<% End Sub 

Sub ErrorMessage () %>
	<p align="center"><strong>ERROR</strong></p>

	<p align="center">One or more of the fields are not completed.<br>
	Please go back and complete all fields.<br>
	Enter &quot;na&quot; if any field is not applicable.</p>
<% End Sub 

Sub ThankYouMessage()
%>
	<p align="center">&nbsp;</p>
	<p align="center"><strong>Your message has been sent.</strong></p>
	<p align="center"><strong>Thank you.</strong></p>

<% End Sub %>




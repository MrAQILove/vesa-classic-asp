<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="include/include.asp"-->
<!--#include file="include/functions.asp"--> 

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
   Call DisplaySearchForm()
End If 

Sub DisplaySearchForm() %>
	<!DOCTYPE html>
	<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
				<script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
	
	<title>VESA Members Database : Member Search</title>
	<meta name="keywords" content="" />
	<meta name="VESA Members Database" content="" />
	<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="stylesheet" href="css/buttons.css">
	<link rel="stylesheet" href="css/forms.css">
	<link rel="stylesheet" href="css/base.css">
	<script src="//ajax.googleapis.com/ajax/libs/jquery/2.0.0/jquery.min.js"></script>
	<script type="text/javascript">
		$(document).ready( function ()
		{
			/* we are assigning change event handler for select box */
			/* it will run when selectbox options are changed */
			$('#search').change(function()
			{
				/* setting currently changed option value to option variable */
				var option = $(this).find('option:selected').val();
				
				if (option === '2') { $('div#Div_1').show(250); }
				else { $('div#Div_1').hide(250); }

				if (option === '3') { $('div#Div_2').show(250); }
				else { $('div#Div_2').hide(250); }

				if (option === '4') { $('div#Div_3').show(250); }
				else { $('div#Div_3').hide(250); }

				if (option === '5') { $('div#Div_4').show(250); }
				else { $('div#Div_4').hide(250); }

				if (option === '6') { $('div#Div_5').show(250); }
				else { $('div#Div_5').hide(250); }

				if (option === '7') { $('div#Div_6').show(250); }
				else { $('div#Div_6').hide(250); }

				if (option === '8') { $('div#Div_7').show(250); }
				else { $('div#Div_7').hide(250); }

				if (option === '9') { $('div#Div_8').show(250); }
				else { $('div#Div_8').hide(250); }

				if (option === '10') { $('div#Div_9').show(250); }
				else { $('div#Div_9').hide(250); }

				if (option === '11') { $('div#Div_10').show(250); }
				else { $('div#Div_10').hide(250); }

				//if (option === '12') { $('div#Div_11').show(250); }
				//else { $('div#Div_10').hide(250); }

				//if (option === '13') { $('div#Div_12').show(250); }
				//else { $('div#Div_10').hide(250); }
			});
		});
	</script>
	<script type="text/javascript" src="javascript/resetButton.js"></script>
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
		<header id="header">
			<div id="logo">
				<h1><a href="#"><span></span></a></h1>
				<p></p>
			</div>
		</header>
		<!-- end header -->
		
		<!-- start side menu bar -->
		<section id="page">
			<%
				If Session("AccessRights") = "Level 1" Then
					Call adminMainMenu()
				Else 
					Call VESAUnitMainMenu()
				End If 
				%>
		
				<!-- start content -->
				<article id="content">
					<div class="post">
						<h1 class="title"><a href="#">Search for Members</a></h1>
						<p class="byline"><b>Please fill out the form.</b></p>
						<div class="entry">
							<form class="pure-form pure-form-aligned" name="VESASearch" method="post" action="VESAOutput.asp" onSubmit="return stopSubmit()">
							<fieldset>
								<div class="pure-control-group">
									<label for="delete">Search:</label>
									<select id="search" name="search" class="pure-input-medium" onChange="ShowReg(this.selectedIndex)">
									<option value="0">Please Choose</option>
									<option value="1">All</option>
									<option value="2">Recipient ID</option>
									<option value="3">Surname/Organization</option>
									<option value="4">First Name</option>
									<option value="5">Address</option>
									<option value="6">Suburb</option>
									<option value="7">Postcode</option>
									<option value="8">State</option>
									<option value="9">Membership Number</option>
									<option value="10">VESA Unit</option>
									<option value="11">SES Region</option>
									</select>
								</div>

								<div class="pure-control-group" id="Div_1" style="display:none;">
									<label for="search for">Search For:</label>
									<input name="searchForRecipientID" id="searchForRecipientID" type="text">
								</div>

								<div class="pure-control-group" id="Div_2" style="display:none;">
									<label for="search for">Search For:</label>
									<input name="searchForSurname_Organisation" id="searchForSurname_Organisation" type="text">
								</div>

								<div class="pure-control-group" id="Div_3" style="display:none;">
									<label for="search for">Search For:</label>
									<input name="searchForFirstName" id="searchForFirstName" type="text">
								</div>

								<div class="pure-control-group" id="Div_4" style="display:none;">
									<label for="search for">Search For:</label>
									<input name="searchForAddress" id="searchForAddress" type="text">
								</div>

								<div class="pure-control-group" id="Div_5" style="display:none;">
									<label for="search for">Search For:</label>
									<input name="searchForSuburb" id="searchForSuburb" type="text">
								</div>

								<div class="pure-control-group" id="Div_6" style="display:none;">
									<label for="search for">Search For:</label>
									<input name="searchForPostcode" id="searchForPostcode" type="text">
								</div>

								<div class="pure-control-group" id="Div_7" style="display:none;">
									<label for="search for">Search For:</label>
									<select id="searchForState" name="searchForState" class="pure-input-medium">
										<option value="0">Please Choose</option>
										<option value="1">ACT</option>
										<option value="2">NSW</option>
										<option value="3">NT</option>
										<option value="4">QLD</option>
										<option value="5">SA</option>
										<option value="6">TAS</option>
										<option value="7">VIC</option>
										<option value="8">WA</option>
										<option value="9">OUTSIDE AUSTRALIA</option>
									</select>
								</div>

								<div class="pure-control-group" id="Div_8" style="display:none;">
									<label for="search for">Search For:</label>
									<input name="searchForMembershipNumber" id="searchForMembershipNumber" type="text">
								</div>

								<div class="pure-control-group" id="Div_9" style="display:none;">
									<label for="search for">Search For:</label>
									<select id="searchForVESAUnit" name="searchForVESAUnit" class="pure-input-medium">
										<option value="0">Please Choose</option>
										<%
										Dim rs
										Dim strSQL
													  
										strSQL = "SELECT * FROM VESA_tblUnit"
										strSQL = strSQL & " WHERE IsActive = '1'"
										strSQL = strSQL & " ORDER BY VESAUnit ASC"
									   
										Set rs = Server.CreateObject("ADODB.Recordset")
									   
										EstablishConnection()
									   
										rs.Open strSQL,Conn
										rs.MoveFirst
													  
										Do Until rs.EOF
											Response.Write "<option value=""" & rs.Fields("VESAUnitID") & """>" & rs.Fields("VESAUnit") & "</option>" & vbCrLf
											rs.MoveNext
										Loop
									   
										rs.Close
										Set rs = Nothing
										CloseConnection()
										%>
									</select>
								</div>

								<div class="pure-control-group" id="Div_10" style="display:none;">
									<label for="search for">Search For:</label>
									<select id="searchForSESRegion" name="searchForSESRegion" class="pure-input-medium">
									<option value="0">Please Choose</option>
									<option value="1">Central</option>
									<option value="2">South West</option>
									<option value="3">East</option>
									<option value="4">North East</option>
									<option value="5">Mid West</option>
									<option value="6">North West</option>
									</select>
								</div>

								<div class="pure-controls">
									<button type="submit" class="pure-button">Search</button>
								</div>
							</fieldset>
							</form>
						</div>
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
					<li class="hover"><a href="#">Search for Member</a></li>
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
<% End Sub %>

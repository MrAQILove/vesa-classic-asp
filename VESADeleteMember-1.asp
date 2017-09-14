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
   'Connect to the database
   EstablishConnection()

   Call DeleteMember()

   'Close database connection
   CloseConnection()
End If

Sub DeleteMember()
%>
	<!DOCTYPE html>
	<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
				<script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
			<title>VESA Members Database : Delete a Member</title>
			<meta name="keywords" content="" />
			<meta name="VESA Members Database" content="" />
			<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
			<link rel="stylesheet" href="css/buttons.css">
			<link rel="stylesheet" href="css/forms.css">
			<link rel="stylesheet" href="css/base.css">
			<link rel="stylesheet" href="css/grids.css">
			<script type="text/javascript" src="javascript/resetButton.js"></script>
			<script type="text/javascript">
			<!--
			function stopSubmit() {
				return false;
			}

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
			function deleteMember() {
				document.DeleteForm.submit();
			}
			//-->
			</script>
			<script src="//ajax.googleapis.com/ajax/libs/jquery/2.0.0/jquery.min.js"></script>
			<script type="text/javascript">
				$(document).ready( function ()
				{
					/* we are assigning change event handler for select box */
					/* it will run when selectbox options are changed */
					$('#WhyDelete').change(function()
					{
						/* setting currently changed option value to option variable */
						var option = $(this).find('option:selected').val();
					
						if (option === '4') {
							$('div#Div_1').show(250);
						}

						else {
							$('div#Div_1').hide(250);
						}
					});
				});
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
				</div>
				
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
					<article id="content">
						<h1 class="title"><a href="#">Delete a Member</a></h1>
						<p class="byline"><b>Please fill out the form.</b></p>
						<div class="entry">
							<form class="pure-form pure-form-aligned" name="DeleteForm" method="post" action="VESASave.asp">
							<input type="hidden" name="ActionType" value="Delete">
							<fieldset>
								<div class="pure-control-group">
									<label for="delete">Delete Member:</label>
									<%
									If Session("AccessRights") = "Level 1" Then 
										Call showDeleteMember(Conn, rsDelete, "VESA_tblMembers")
									Else
										Call showDeleteMemberByVESAUnit(Conn, rsDelete, "VESA_tblMembers", "" & Session("VESAUnitID") & "")
									End If 
									%>
								</div>

								<div class="pure-control-group">
									<label for="resaon">Reason:</label>
									<select id="WhyDelete" name="WhyDelete" class="pure-input-medium">
									<option>Please Choose</option>
									<%
									'ASP Usage: Options from array to simulate a recordset
									aOptions = Array("Area Closed", "Area Amalgamated", "No Longer Interested", "Other Reason")
									For each sOption in aOptions
										Select Case sOption
											Case "Area Closed"
												sOptionS = "1"

											Case "Area Amalgamated"
												sOptionS = "2"
											  
											Case "No Longer Interested"
												sOptionS = "3"

											Case "Other Reason"
												sOptionS = "4"
										End Select
										Response.Write "<option value=""" & sOptionS & """>" & sOption & "</option>" & vbCrLf
									Next
									%>
									</select>
								</div>

								<div class="pure-control-group" id="Div_1" style="display:none;">
									<label for="resaon">Specify Reason:</label>
									<textarea id="SpecifyReason" name="SpecifyReason" class="pure-input-1-2" placeholder="Specify your reason for deletion"></textarea>
								</div>

								<div class="pure-controls">
									<button type="submit" class="pure-button">Delete Member</button>
								</div>
							</fieldset>
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
					<li class="hover"><a href="#">Delete a Member</a></li>
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




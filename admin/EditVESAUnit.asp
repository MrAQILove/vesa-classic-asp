<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/database/vesa/Members/include/include.asp"-->
<!--#include virtual="/database/vesa/Members/include/adovbs.inc"-->
<!--#include virtual="/database/vesa/Members/include/functions.asp"--> 

<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = True 

If Session("UserLoggedIn") <> "true" Then
   Response.Redirect "../AdminLogin.asp"

Else  
   Dim rsResult
   Dim strVESAUnitID
   Dim VESAUnitID, VESAUnit, Password, UnitEmailAddress, UnitSESRegionID
   
   strVESAUnitID = CLng(Request.Form("VESAUnitID"))

   EstablishConnection()

   strSQL = "SELECT * FROM VESA_tblUnit U"
   strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
   strSQL = strSQL & " WHERE VESAUnitID ='" & strVESAUnitID & "'"
				  
   Set rsResult = Server.CreateObject("ADODB.Recordset")
				  
   rsResult.Open strSQL, Conn, adOpenKeyset, AdLockOptimistic 

   VESAUnitID			= rsResult("VESAUnitID") & ""
   VESAUnit				= rsResult("VESAUnit") & ""
   Password				= rsResult("Password") & ""
   UnitEmailAddress		= rsResult("EmailAddress") & ""
   UnitSESRegionID		= rsResult("SESRegionID") & ""

   '- Print the Edit page
   Call editUnit()

   rsResult.Close
   Set rsResult = Nothing
   CloseConnection()
End If 


Sub editUnit() %>
	<!DOCTYPE html>
	<html>
	<head>
		<meta charset="UTF-8">
		<!--[if lt IE 9]>
		  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
		<![endif]-->
		<title>VESA Members Database : Edit VESA Unit/Distribution</title>
		<meta name="keywords" content="" />
		<meta name="VESA Members Database" content="" />
		<link rel="stylesheet" href="../css/default.css" type="text/css" media="screen" />
		<link rel="stylesheet" href="../css/buttons.css">
		<link rel="stylesheet" href="../css/forms.css">
		<link rel="stylesheet" href="../css/base.css">
	</head>
	
	<body>
		<div id="wrapper">
			<nav id="menu">
				<ul id="main">
					<li><a href="VESAMain.asp">Home</a></li>
					<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
					<li><a href="http://www.cwmedia.com.au/">Countrywide Media</a></li>
					<li><a href="VESAContact.asp">Contact Us</a></li>
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
			<section id="page">
				<aside id="sidebar1" class="sidebar">
					<ul>
						<li> 
							<h2>VESA Members</h2>
							<ul>
								<li><a href="../VESAMain.asp">Home</a></li>
								<li><a href="../VESASearch.asp">Search for Member</a></li>
								<li><a href=../"VESAViewAllMembers.asp">View All Members</a></li>
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
								<li><a href="AddNewVESAUnit.asp">Add a New VESA Unit</a></li>
								<li class="hover"><a href="#">Edit Unit - <br /><%=VESAUnit%></a></li>
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

				<!-- start article -->
				<article id="content">
					<h1 class="title"><a href="#">Edit <%=VESAUnit%> Unit</a></h1>
					<p class="byline"><b>Please fill out the form.</b></p>
						
					<div class="entry">
						<form class="pure-form pure-form-aligned" name="EditForm" method="post" action="../VESASave.asp">
						<input type="hidden" name="ActionType" value="UpdateVESAUnit">
						<input type="hidden" name="VESAUnitID" value="<%=VESAUnitID%>">

							<fieldset>
								<div class="pure-control-group">
									<label>VESA Unit ID:</label>
									<span style="color:#0000a0; font-weight:bold"><%=VESAUnitID%></span>
								</div>

								<div class="pure-control-group">
									<label>VESA Unit:</label>
									<input type="text" name="VESAUnit" value="<%=VESAUnit%>">
									<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
								</div>

								<div class="pure-control-group">
									<label>Password:</label>
									<input type="text" name="VESAUnitPassword" value="<%=Password%>">
									<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
								</div>

								<div class="pure-control-group">
									<label>Email Address:</label>
									<input type="email" name="UnitEmailAddress" placeholder="Email Address" class="pure-input-1-2" value="<%=UnitEmailAddress%>">
								</div>

								<div class="pure-control-group">
									<label>SES Region:</label>
									<%
									If UnitSESRegionID = "" Then
										Call showSelectedValue(Conn, rsSESRegion, "VESA_tblSESRegion", "SESRegionID", "SESRegion", "UnitSESRegionID")
									Else %>
										<select id="UnitSESRegionID" name="UnitSESRegionID" class="pure-input-medium">
										<option <%If UnitSESRegionID = "" Then Response.Write "class=""selectedItem"" selected"%>>Please Choose</option>
										<option value="1" <%If UnitSESRegionID = "1" Then Response.Write "class=""selectedItem"" selected"%>>Central</option>
										<option value="2" <%If UnitSESRegionID = "2" Then Response.Write "class=""selectedItem"" selected"%>>East</option>
										<option value="3" <%If UnitSESRegionID = "3" Then Response.Write "class=""selectedItem"" selected"%>>Mid West</option>
										<option value="4" <%If UnitSESRegionID = "4" Then Response.Write "class=""selectedItem"" selected"%>>North East</option>
										<option value="5" <%If UnitSESRegionID = "5" Then Response.Write "class=""selectedItem"" selected"%>>North West</option>
										<option value="6" <%If UnitSESRegionID = "6" Then Response.Write "class=""selectedItem"" selected"%>>South West</option>
										</select> 
									<% End If %>
								</div>

								<div class="pure-controls">
									<button type="submit" class="pure-button">Update</button>
								</div>
							</fieldset>
						</form>  
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
<% End Sub %>

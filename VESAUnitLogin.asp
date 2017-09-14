<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="include/include.asp"-->

<!DOCTYPE html>
	<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
				<script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
			<title>VESA Members Database : Admin Login</title>
			<meta name="keywords" content="" />
			<meta name="VESA Members Database" content="" />
			<link rel="stylesheet" href="css/default.css" type="text/css" media="screen" />
			<link rel="stylesheet" href="css/buttons.css">
			<link rel="stylesheet" href="css/forms.css">
			<link rel="stylesheet" href="css/base.css">
			<script type="text/javascript" src="javascript/resetButton.js"></script>
		</head>
		
		<body>
		<div id="wrapper">
			<nav id="menu">
				<ul id="main">
					<li><a href="AdminLogin.asp">Administration Login</a></li>
					<li class="current_page_item"><a href="#">Unit Login</a></li>
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
				<aside id="sidebar1" class="sidebar">
					<ul>
						<li> <br />
						<img src="images/Phoenix-Logo.jpg" width="220" height="218" alt="" />
						<!--<ul>
						<li>&nbsp;</li>
						</ul>-->
						</li>
					</ul>
				</aside>
			
				<!-- start content -->
				<article>
					<h1 class="title"><a href="#">Welcome to the VESA Administrator's area</a></h1>
					<p class="byline">Please select your username and enter your password below:</p>
					<div class="entry">
						<form class="pure-form pure-form-aligned" name="AdminForm" method="post" action="VESALogin.asp">
							<input type="hidden" name="LoginType" value="VESAUnitLogin">
							<input type="hidden" name="login" value="true">
							
							<fieldset>
								<div class="pure-control-group">
									<label for="name">VESA Unit:</label>
									<select name="VESAUnitID" id="VESAUnitID" class="pure-input-medium" required>
									<%
									Dim objRS
									Dim strSQL
																  
									strSQL = "SELECT * FROM VESA_tblUnit"
									strSQL = strSQL & " WHERE IsUnitSES = '1'"
									strSQL = strSQL & " AND IsActive = '1'"
									strSQL = strSQL & " ORDER BY VESAUnit ASC"
												   
									Set objRS = Server.CreateObject("ADODB.Recordset")
												   
									EstablishConnection()
												   
									objRS.Open strSQL,Conn
									objRS.MoveFirst
																  
									Do Until objRS.EOF
										Response.Write "<option value=""" & objRS.Fields("VESAUnitID") & """>" & objRS.Fields("VESAUnit") & "</option>" & vbCrLf
										objRS.MoveNext
									Loop
												   
									objRS.Close
									Set objRS = Nothing
									CloseConnection()
									%>
									</select>
								</div>

								<div class="pure-control-group">
									<label for="password">Password</label>
									<input id="Password" name="Password" type="password" placeholder="Password" required>
								</div>
								
								<div class="pure-controls">
									<button type="submit" class="pure-button">Log In</button>
								</div>
							</fieldset>
						</form>
					</div>
				</article>
				<!-- end content -->
				<div style="clear: both;">&nbsp;</div>
			</div>
			<!-- end page -->
		</div>
		<footer id="footer">
		  <p class="copyright">&copy;&nbsp;&nbsp;2008 - <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
		</footer>
		</body>
</html>


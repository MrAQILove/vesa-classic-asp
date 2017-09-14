<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/database/vesa/Members/include/include.asp"-->
<!--#include virtual="/database/vesa/Members/include/functions.asp"--> 
<!--#include virtual="/database/vesa/Members/include/adovbs.inc"--> 

<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = True 

If Session("UserLoggedIn") <> "true" Then
	Response.Redirect "../AdminLogin.asp" 

Else  
	Dim rsResult
	Dim strID
	Dim Surname, FirstName, EmailAddress, Username, AdminPassword, AccessID
   
	strID = CLng(Request.Form("UserID"))

	EstablishConnection()

	strSQL = "SELECT * FROM MembersDB_tblUsers U" 
	strSQL = strSQL & " INNER JOIN MembersDB_tblUserRights UR on U.UserID = UR.UserID "
	strSQL = strSQL & " INNER JOIN MembersDB_tblUserAccess UA on UR.AccessID = UA.AccessID"
	strSQL = strSQL & " WHERE U.UserID ='" & strID & "'"
				  
	Set rsResult = Server.CreateObject("ADODB.Recordset")
				  
	rsResult.Open strSQL, Conn, adOpenStatic, adLockReadOnly

	UserID				= rsResult("UserID") & ""
	Surname				= rsResult("Surname") & ""
	FirstName			= rsResult("FirstName") & ""
	EmailAddress		= rsResult("EmailAddress") & ""
	Username			= rsResult("Username") & ""
	AdminPassword		= rsResult("Password") & ""
	AccessID			= rsResult("AccessID") & ""

	'- Print the Edit page
	Call EditAdministrationUser()

	rsResult.Close
	Set rsResult = Nothing
	CloseConnection()
End If 

Sub DisplayHTMLTitleTag()
	If Not IsNull(rsResult.Fields(2)) And Not IsNull(rsResult.Fields(3)) Then
		Response.Write "Edit Administration User - " & FirstName & "&nbsp;" & Surname

	Else
		Response.Write "Edit Administration User - " & Surname     
   End If 
End Sub

Sub EditAdministrationUser() %>
	<!DOCTYPE html>
		<html>
		<head>
			<meta charset="UTF-8">
			<!--[if lt IE 9]>
			  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
			<![endif]-->
			<title>VESA Members Database : <% Call DisplayHTMLTitleTag() %></title>
			<meta name="keywords" content="" />
			<meta name="VESA Members Database" content="" />
			<link href="../css/default.css" rel="stylesheet" type="text/css" media="screen" />
			<link rel="stylesheet" href="../css/buttons.css">
			<link rel="stylesheet" href="../css/forms.css">
			<link rel="stylesheet" href="../css/base.css">
			<script type="text/javascript" src="../javascript/resetButton.js"></script>
			<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
			<script type="text/javascript">
				$(document).ready(
				    function() {
				        $("#toggleElement").click(function() {
				            $("#showPassword").fadeToggle();
				        });
				    });
			</script>
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
								<li><a href="AddNewVESAUnit">Add a New VESA Unit</a></li>
								<li><a href="DeleteVESAUnit.asp">Delete a VESA Unit</a></li>
								<li><a href="VESAViewInactiveUnits.asp">View All Inactive Units</a></li>
							</ul>
						</li>

						<li> 
							<h2>Admin Members</h2>
							<ul>
								<li><a href="VESAViewAllAdminUsers.asp">View All Admin Users</a></li>
								<li><a href="VESAAddAdminUser.asp">Add an Admin User</a></li>
								<li class="hover"><a href="#">Edit an Admin User</a></li>
								<li><a href="VESADeleteAdminUser.asp">Delete an Admin User</a></li>
							</ul>
						</li>

						<li><a href="../AdminLogin.asp"><h2>Log Out</h2></a></li>
					</ul>
				</aside>
				<!-- end aside -->

				<!-- start article -->
				<article id="content">
					<% Call OutputEditForm() %>
				</article>
				<!-- end article -->
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

Sub OutputEditForm()
%> 
	<h1 class="title"><a href="#"><% Call DisplayHTMLTitleTag() %></a></h1>
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
							'- Error Ouput for Surname ------------------------------------------------------
							If Session("badSurname") = "T" Then
					        	Response.write "<li><a href=""#"">The <b>SURNAME</b> field must not be empty.</a></li>" & vbCrLf
								Session("badSurname") = "F"
					        End If

							If Session("badSurname") = "T1" Then
					         	Response.write "<li><a href=""#"">Enter only letters in your <b>SURNAME</b> field. <br /> Avoid using (/ ' "" "" ! @ # $ % ^ & * -) characters and numbers.</a></li>" & vbCrLf
							  	Session("badSurname") = "F1"
					        End If
							'--------------------------------------------------------------------------------

						  	'- Error Ouput for First Name ---------------------------------------------------
					      	If Session("badFirstname") = "T" Then
		                    	Response.write "<li><a href=""#"">The <b>FIRSTNAME</b> field must not be empty.</a></li>" & vbCrLf
						     	Session("badFirstname") = "F"
		                  	End If

					      	If Session("badFirstname") = "T1" Then
		                    	Response.write "<li><a href=""#"">Enter only letters in your <b>FIRSTNAME</b> field. <br /> Avoid using (/ ' "" "" ! @ # $ % ^ & * -) characters and numbers.</a></li>" & vbCrLf
						     	Session("badFirstname") = "F1"
		                  	End If
					      	'--------------------------------------------------------------------------------

					      	'- Error Ouput for Email --------------------------------------------------------
					      	If Session("badEmailAddress") = "T" Then
		                    	Response.write "<li><a href=""#"">The <b>EMAIL</b> field must not be empty.</a></li>" & vbCrLf
						     	Session("badEmailAddress") = "F"
		                  	End If

						  	If Session("badEmailAddress") = "T1" Then
		                    	Response.write "<li><a href=""#"">Invalid <b>EMAIL</b> Address.</a></li>" & vbCrLf
						     	Session("badEmailAddress") = "F1"
		                  	End If
					      	'--------------------------------------------------------------------------------

						  	'- Error Ouput for Username -----------------------------------------------------
					      	If Session("badUsername") = "T" Then
		                    	Response.write "<li><a href=""#"">The <b>USERNAME</b> field must not be empty.</a></li>" & vbCrLf
						     	Session("badUsername") = "F"
		                  	End If

					      	If Session("badUsername") = "T1" Then
		                    	Response.write "<li><a href=""#"">Enter only letters in your <b>USERNAME</b> field. <br /> Avoid using (/ ' "" "" ! @ # $ % ^ & * -) characters and numbers.</a></li>" & vbCrLf
						     	Session("badUsername") = "F1"
		                  	End If
					      	'--------------------------------------------------------------------------------

						  	'- Error Ouput for Access Rights ------------------------------------------------
					      	If Session("badAccessID") = "T" Then
		                    	Response.write "<li><a href=""#"">Please choose your <b>ACCESS RIGHTS</b>.</a></li>" & vbCrLf
						     	Session("badAccessID") = "F"
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
							
		<form class="pure-form pure-form-aligned" name="EditAdminUserForm" method="post" action="VESACheckAdminUser.asp" onSubmit="return stopSubmit()">
			<input type="hidden" name="ActionType" value="EditAdministrationUser">
			<input type="hidden" name="UserID" value="<%=UserID%>">
			
			<fieldset>
				<div class="pure-control-group">
					<label>User ID:</label>
					<font color="#0000a0"><strong><%=UserID%></strong></font>
				</div>

				<div class="pure-control-group">
					<label>Surname:</label>
					<input type="text" name="Surname" placeholder="Surname" value="<%=Surname%>" required>
					<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
				</div>

				<div class="pure-control-group">
					<label>First Name:</label>
					<input type="text" name="FirstName" placeholder="First Name" value="<%=FirstName%>" required>
					<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
				</div>

				<div class="pure-control-group">
					<label>Email Address:</label>
					<input type="email" name="EmailAddress" placeholder="EmailAddress" class="pure-input-1-2" value="<%=EmailAddress%>" required>
					<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
				</div>

				<div class="pure-control-group">
					<label>Username:</label>
					<input type="text" name="Username" placeholder="Username" value="<%=Username%>" required>
					<aside class="pure-form-message-inline"><font color="#990000"><b><code>*</code></b></font></aside>
				</div>

				<div class="pure-control-group">
					<label>Change your Password?</label>
					Yes <input id="toggleElement" type="checkbox" name="toggle" />
				</div>

				<div id="showPassword" style="display: none;">
					<div class="pure-control-group">
						<label>Old Password:</label>
						<%
					  	If IsNull(rsResult.Fields("Password")) Then
					    	Response.Write "<input type=""text"" name=""OldPassword"" value="""" />"
					  	Else
					    	'Password = kLeachRegExp("" & AdminPassword & "", "[^()?<>.*?]", "*") 
	  					 	Response.Write "<input type=""text"" name=""OldPassword"" placeholder=""OldPassword"" value=""" & AdminPassword & """>"
					  	End If 
					 	 %>
					</div>

					<div class="pure-control-group">
						<label>New Password:</label>
						<input type="text" name="NewPassword1" placeholder="New Password" value="">
					</div>

					<div class="pure-control-group">
						<label>Confirm Password:</label>
						<input type="text" name="NewPassword2" placeholder="Confirm Password" value="">
					</div>
				</div>

				<div class="pure-control-group">
					<label for="state">Access Rights:</label>
					<%
					If AccessID = "" Then
						Call showSelectedValue(Conn, rsAccessRights, "MembersDB_tblUserAccess", "AccessID", "AccessRights", "AccessID") 
					Else %>
						<select id="AccessID" name="AccessID" class="pure-input-medium">
				  		<option <%If UnitSESRegionID = "" Then Response.Write "class=""selectedItem"" selected"%>>Please Choose</option>
				  		<option value="1" <%If AccessID = "1" Then Response.Write "class=""selectedItem"" selected"%>>Level 1</option>
				  		<option value="2" <%If AccessID = "2" Then Response.Write "class=""selectedItem"" selected"%>>Level 2</option>
				  		<option value="3" <%If AccessID = "3" Then Response.Write "class=""selectedItem"" selected"%>>Level 3</option>
				  		<option value="4" <%If AccessID = "4" Then Response.Write "class=""selectedItem"" selected"%>>Level 4</option>
			   	  		<option value="5" <%If AccessID = "5" Then Response.Write "class=""selectedItem"" selected"%>>Level 5</option>
			      		</select>  
					<% End If %>
				</div>

				<div class="pure-controls">
					<button type="submit" class="pure-button">Submit</button>
				</div>
			<fieldset>
		</form>		
    </div>
<% End Sub %>


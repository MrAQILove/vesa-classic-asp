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
   Call AddNewAdministrationUser()
End If

Sub AddNewAdministrationUser()%>
<!DOCtype html PUBLIC "-//W3C//Dtd XHTML 1.0 Strict//EN" "http://www.w3.org/tr/xhtml1/Dtd/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>VESA Members Database : Add a New Administration User</title>
<meta name="keywords" content="" />
<meta name="VESA Members Database" content="" />
<link href="../css/default.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript" src="../javascript/resetButton.js"></script>
</head>
<body>
<div id="wrapper">
	 <div id="menu">
		<ul id="main">
			<li><a href="../VESAMain.asp">Home</a></li>
			<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
			<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
			<li><a href="../contactUs.html">Contact Us</a></li>
		</ul>
	</div>
	
	<!-- start header -->
	<div id="header">
		<div id="logo">
			<h1><a href="#"><span></span></a></h1>
			<p></p>
		</div>
	</div>
	<!-- end header -->
	
	<!-- start page -->
	<div id="page">
		<div id="sidebar1" class="sidebar">
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
					</ul>
				</li>

				<li> 
					<h2>Admin Members</h2>
					<ul>
						<li><a href="VESAViewInactiveMembers.asp">View Inactive Members</a></li>
						<li><a href="ViewAllVESAUnits.asp">View All VESA Units</a></li>
						<li><a href="AddNewVESAUnit.asp">Add a New VESA Unit</a></li>
						<li><a href="DeleteVESAUnit.asp">Delete a VESA Unit</a></li>
						<li><a href="VESAViewAllAdminUsers.asp">View All Admin Users</a></li>
						<li class="hover"><a href="#">Add an Admin User</a></li>
						<li><a href="VESADeleteAdminUser.asp">Delete an Admin User</a></li>
						<li><a href="../AdminLogin.asp">Log Out</a></li>
					</ul>
				</li>
			</ul>
		</div>
		
		<!-- start content -->
		<div id="content">
			<div class="post">
				<h1 class="title"><a href="#">Add a New Administration User</a></h1>
				<p class="byline"><b>Please fill out the form.</b></p>
				<div class="entry">
					<form name="AddAdminUserForm" method="post" action="VESACheckAdmin.asp">
					<input type="hidden" name="Actiontype" value="AddAdministrationUser">
					<input type="hidden" name="DatabaseName" value="VESA_tblMembers">
					<table border="0">
					<%
					'- On the first time that this page loads, session("errors") has no value and so equals 0. 
					'- On subsequent visits, session("errors will have a value of 1 or more if errors were made") 
					If Session("Errors") = 0 Then 
						Response.Write "<tr><td><b>Fields that have a <font color=""#990000""><b>*</b></font> next to them are Mandatory.</b></td></tr>" & vbCrLf
						Response.Write "<tr><td><img src=""../../../images/spacer.gif"" width=""1"" height=""10"" alt=""""></td></tr>" & vbCrLf
					
					Else
					   '- Errors were made so list them in the rest of the table reset our error counter
					   Session("Errors") = "0"
					%>
			<tr>
			<td>
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

				  '- Error Ouput for Password -----------------------------------------------------
			      If Session("badPassword") = "T" Then
                     Response.write "<li><a href=""#"">The <b>PASSWORD</b> field must not be empty.</a></li>" & vbCrLf
				     Session("badPassword") = "F"
                  End If
			      '--------------------------------------------------------------------------------

				  '- Error Ouput for Access Rights ------------------------------------------------
			      If Session("badUserAccessRights") = "T" Then
                     Response.write "<li><a href=""#"">Please choose your <b>ACCESS RIGHTS</b>.</a></li>" & vbCrLf
				     Session("badUserAccessRights") = "F"
                  End If
			      '--------------------------------------------------------------------------------
				  '- End the errors table
				  %>
	              </ul>
               </li>
               </ul>
            </div>
			</td>
			</tr>    
			<% End If %>

			<tr>
			<td>
			   <table cellpadding="0" cellspacing="0" border="0">
		       <tr>
			   <td valign="top"><div align="left"><strong><label for="Surname">Surname:</label></strong></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><div align="left"><input name="Surname" id="Surname" type="text" class="inputTextField" value="<%=Session("Surname")%>" /></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><font color="#990000"><b>*</b></font></td>
			   </tr>

			   <tr><td colspan="5"><img src="../images/spacer.gif" width="1" height="3" /></td></tr>

			   <tr>
			   <td valign="top"><div align="left"><strong><label for="First Name">First Name:</label></strong></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><div align="left"><input name="FirstName" id="FirstName" type="text" class="inputTextField" value="<%=Session("FirstName")%>" /></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><font color="#990000"><b>*</b></font></td>
			   </tr>

			   <tr><td colspan="5"><img src="../images/spacer.gif" width="1" height="3" /></td></tr>

			   <tr>
		       <td valign="top"><div align="left"><strong><label for="Email">Email:</label></strong></div></td>	   
		       <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
		       <td><div align="left"><input name="EmailAddress" id="EmailAddress" type="text" class="inputTextField" value="<%=Session("EmailAddress")%>" /></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><font color="#990000"><b>*</b></font></td>
		       </tr>

			   <tr><td colspan="5"><img src="../images/spacer.gif" width="1" height="3" /></td></tr>

			   <tr>
			   <td><div align="left"><strong><label for="Username">Username:</label></strong></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" border="0"></td>
			   <td><div align="left"><input name="Username" id="Username" type="text" class="inputTextField2" size="32" value="<%=Session("Username")%>"></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><font color="#990000"><b>*</b></font></td>
			   </tr>

			   <tr><td colspan="5"><img src="../images/spacer.gif" width="1" height="3" /></td></tr>

			   <tr>
			   <td><div align="left"><strong><label for="Password">Password:</label></strong></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" border="0"></td>
			   <td><div align="left"><input name="Password" id="Password" type="password" class="inputTextField2" size="32" value="<%=Session("Password")%>"></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><font color="#990000"><b>*</b></font></td>
			   </tr>

			   <tr><td colspan="5"><img src="../images/spacer.gif" width="1" height="3" /></td></tr>

			   <tr>
			   <td><div align="left"><strong><label for="Access Rights">Access Rights:</label></strong></div></td>
			   <td><img src="../images/spacer.gif" width="5" height="1" border="0"></td>
			   <td>
			   <div align="left">
			   <% If Session("UserAccessRights") = "" Then %>
			      <select id="UserAccessRights" name="UserAccessRights" size="1" class="inputSelection">
			      <option value="Please Choose">Please Choose</option>
			      <option value="1">Level 1</option>
			      <option value="2">Level 2</option>
				  <option value="3">Level 3</option>
				  <option value="4">Level 4</option>
				  <option value="5">Level 5</option>
			      </select>
			
			   <% Else %>
			      <select id="UserAccessRights" name="UserAccessRights" size="1" class="inputSelection">
			      <option <%If Session("UserAccessRights") = "Please Choose" Then Response.Write "class=""selectedItem"" selected"%>>Please Choose</option>
			      <option value="1" <%If Session("UserAccessRights") = "1" Then Response.Write "class=""selectedItem"" selected"%>>Level 1</option>
			      <option value="2" <%If Session("UserAccessRights") = "2" Then Response.Write "class=""selectedItem"" selected"%>>Level 2</option>
			      <option value="3" <%If Session("UserAccessRights") = "3" Then Response.Write "class=""selectedItem"" selected"%>>Level 3</option>
			      <option value="4" <%If Session("UserAccessRights") = "4" Then Response.Write "class=""selectedItem"" selected"%>>Level 4</option>
			      <option value="5" <%If Session("UserAccessRights") = "5" Then Response.Write "class=""selectedItem"" selected"%>>Level 5</option>
			      </select>
			   <% End If %>
			   </div>
			   </td>
			   <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><font color="#990000"><b>*</b></font></td>
			   </tr>

			   <tr><td colspan="5"><img src="../images/spacer.gif" width="1" height="3" /></td></tr>

		       <tr>
               <td>&nbsp;</td>
			   <td><img src="../../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td colspan="3">
			   <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0">
			      <tr>
			      <td><input type="image" name="submit" class="submit-btn" src="http://www.roscripts.com/images/btn.gif" alt="submit" title="submit" /></td>
			      <td><img src="../images/spacer.gif" width="5" height="1" alt="" /></td>
			      <td valign="top">
			      <script type="text/javascript">
			      <!--
			      var ri = new resetimage("../images/button/reset_off.gif");
			      ri.name = "resetter";
			      ri.rollover = "../images/button/reset_on.gif";
			      ri.write();
			      //-->
			      </script>
			      <noscript><input type="reset"></noscript>		
			      </td>
			      </tr>
			      </table>
               </div>
			   </td>
               </tr>
			   </table>
			</td>
			</tr>
			</table>
		    </form>  
		</div>
      </div>
    </div>
    <!-- end content -->
    <div style="clear: both;">&nbsp;</div>
  </div>
  <!-- end page -->
</div>
<div id="footer">
	<p class="copyright">&copy;&nbsp;&nbsp;2008 - <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
</div>
</body>
</html>
<% End Sub %>

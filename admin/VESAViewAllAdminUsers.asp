<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/database/vesa/Members/include/include.asp"-->
<!--#include virtual="/database/vesa/Members/include/functions.asp"-->  
<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = True 

If Session("UserLoggedIn") <> "true" Then
   Response.Redirect "../AdminLogin.asp" 

Else
   '- Constants ripped from adovbs.inc:
   Const adOpenStatic = 3
   Const adLockReadOnly = 1
   Const adCmdText = &H0001

   '- Our own constants:
   Const PAGE_SIZE = 20  ' The size of our pages.

   '- Declare our variables... always good practice!
   Dim rstSearch		' ADO recordset
   Dim strSQL			' The SQL Query we build on the fly
   Dim iPageCurrent		' The page we're currently on
   Dim iPageCount		' Number of pages of records
   Dim iRecordCount		' Count of the records returned
   Dim I				' Standard looping variable

   '- Retrieve page to show or default to the first
   If Request.QueryString("page") = "" Then
      iPageCurrent = 1
   Else
      iPageCurrent = CInt(Request.QueryString("page"))
   End If

   EstablishConnection()
   
   '- Build our query based on the input.
   strSQL = "SELECT * FROM MembersDB_tblUsers U"
   strSQL = strSQL & " INNER JOIN MembersDB_tblUserRights UR ON U.UserID = UR.UserID"
   strSQL = strSQL & " WHERE U.UserActive = '1'"
   strSQL = strSQL & " AND DatabaseName='VESA_tblMembers'"
   strSQL = strSQL & " ORDER BY U.UserID ASC"

   '- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
   Set rstSearch = Server.CreateObject("ADODB.Recordset")
   rstSearch.PageSize  = PAGE_SIZE
   rstSearch.CacheSize = PAGE_SIZE

   '- Open our recordset
   rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

   '- Get a count of the number of records and pages for use in building the header and footer text.
   iRecordCount = rstSearch.RecordCount
   iPageCount   = rstSearch.PageCount

   Call DisplayHeader("VESA Members Database : Displaying All Administration Users")
      
   Call OutputPage()

   '- Close our recordset and connection and dispose of the objects
   rstSearch.Close
   Set rstSearch = Nothing
   
   CloseConnection()

End If

Sub DisplayHeader(strMessage) %>
<!DOCTYPE html>
	<html>
	<head>
		<meta charset="UTF-8">
		<!--[if lt IE 9]>
		  <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
		<![endif]-->
		<title><%=strMessage%></title>
		<meta name="keywords" content="" />
		<meta name="VESA Members Database" content="" />
		<link href="../css/default.css" rel="stylesheet" type="text/css" media="screen" />
		<link rel="stylesheet" href="../css/buttons.css">
		<link rel="stylesheet" href="../css/forms.css">
		<link rel="stylesheet" href="../css/base.css">
		<script language="JavaScript">
		<!--
		// Stop Submit button
		function stopSubmit() {
		   return false;
		}

		//User Selected
		function userSelected(strUser)
		{
		   if(<%=Session("VESAID")%> == 1) 
		   {
		      document.EditForm.UserID.value = strUser;
			  document.EditForm.submit();
		   }
		}

		// Change Admin User Password
		function changePassword(strUserID)
		{
		   if(<%=Session("VESAID")%> == 1) 
		   {
		      document.EditForm.UserID.value = strUserID;
			  document.EditForm.submit();
		   }
		}

		// Add New Member 
		function addNewUser()
		{
		   if (<%=Session("VESAID")%> == 1) {
		      document.AddUserForm.submit();
		   }
		}

		function goBack() {
		   document.location.href = "../VESAMain.asp";
		}

		// Delete User
		function deleteUser()
		{
		   var ctr;
		   
		   ctr = 0;
		   
		   // check for single checkbox by seeing if an array has been created
		   var cblength = document.forms['DeleteUserForm'].elements['DoDeleteUser'].length;
		   if(typeof cblength == "undefined")
		   {
		      if(document.forms['DeleteUserForm'].elements['DoDeleteUser'].checked == true) ctr++;
		   }
		   else
		   {
		      for(i = 0; i < document.forms['DeleteUserForm'].elements['DoDeleteUser'].length; i++)
		      {
		         if(document.forms['DeleteUserForm'].elements['DoDeleteUser'][i].checked) ctr++;
		      }
		    }
		          		  
		   if (ctr == 1) 
		   {
		       var answer;
			   answer = confirm('Are you sure you want to delete this Administration User?');
			   if (answer)
			   {
			      document.DeleteUserForm.submit();
		          return false;   
			   }

			   //else {;}
		    }
		    
			else if (ctr > 1) 
		    {
			   var answer;
			   answer = confirm('Are you sure you want to delete ' + ctr + ' Administration User?');
			   if (answer)
			   {
			      document.DeleteUserForm.submit();
		          return false;
			   }

			   //else {;}
		    }
		    
			else 
			{
		       confirm("No Administration User selected for deletion");
		       return true;
		    }
		}

		function newImage(arg) {
			if (document.images) {
				rslt = new Image();
				rslt.src = arg;
				return rslt;
			}
		}

		function changeImages() {
			if (document.images && (preloadFlag == true)) {
				for (var i=0; i<changeImages.arguments.length; i+=2) {
					document[changeImages.arguments[i]].src = changeImages.arguments[i+1];
				}
			}
		}

		var preloadFlag = false;
		function preloadImages() {
			if (document.images) {
			   	menu0_on = newImage("../images/paging_previous_on.jpg");
				menu1_on = newImage("../images/paging_next_on.jpg");		
				preloadFlag = true;
			}
		}

		// -->
		</script>
	</head>
<% End Sub

Sub OutputPage() %>
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
									<li><a href="AddNewVESAUnit.asp">Add a New VESA Unit</a></li>
									<li><a href="DeleteVESAUnit.asp">Delete a VESA Unit</a></li>
									<li><a href="VESAViewInactiveUnits.asp">View All Inactive Units</a></li>
								</ul>
							</li>

							<li> 
								<h2>Admin Members</h2>
								<ul>
									<li class="hover"><a href="#">View All Admin Users</a></li>
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
							<%Call viewAllAdminUsers()%>
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

Sub viewAllAdminUsers()
%>
   <div class="entry">
	   <!--/* Start Here */-->
	   <table border="0" cellspacing="0" cellpadding="0" width="100%">
		   <tr>
			   <td align="center">
			      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
				  	<tr>
				  		<td>
				  			<!--/* header */-->
				  			<h1 class="title"><a href="#">Displaying All Administration Users</a></h1>
				  			<%
							'- Check page count to prevent bombing when zero results are returned!-----------------
							If iRecordCount = 0 Then
								Response.Write "<p class=""byline""><b>No records found!</b></p>"
								Response.Write "</td></tr>"
								Response.Write "</table>"

							Else
								rstSearch.AbsolutePage = iPageCurrent
								Response.Write "<p class=""byline"">&nbsp;</p>"
								Response.Write "</td></tr>"
							%>

				   			<tr height="10"><td><img src="../images/spacer.gif" width="1" height="10" border="0"></td></tr>

							<tr>
						   		<td>
						    		<table border="0" cellpadding="0" cellspacing="0" width="100%">
							  			<tr>
							  				<td>
											<% 
												Response.Write "The VESA Member's database has <b><font color=""#ff0000"">"
												If iRecordCount > 1 Then
													Response.Write iRecordCount & "</font></b> Administration Users!" & vbCrLf
												Else
													Response.Write iRecordCount & "</font></b> Administration User!" & vbCrLf
												End If 	   
											%>		  
							  				</td>
							  				<td align="right"><strong><font color="#ff0000">Displaying page <%= iPageCurrent %> of <%= iPageCount %>:</font></strong></td>
							  			</tr>
							  		</table>
				  				</td>
				  			</tr>

				  			<tr><td><img src="../images/spacer.gif" width="1" height="5" border="0"></td></tr>

						  	<tr>
						  		<td bgcolor="#eeeeee">
							    	<!--/* Output Search */-->
							     	<form name="DeleteUserForm" id="DeleteUserForm" action="../VESASave.asp" method="post" onSubmit="return stopSubmit()">
								 	<input type="hidden" name="ActionType" value="DeleteAdminUser">
								 	<table id="main_table" border="0" align="center" cellspacing="2" cellpadding="1" width="100%">
								 		<tr align="center" height="30">
								 			<td class="tab_header_cell"><b>User ID</b></td>
								 			<td class="tab_header_cell"><font color="#0000a0"><b>Edit <br /> Password</b></font></td>
								 			<td class="tab_header_cell"><font color="#0000a0"><b>Delete <br /> User</b></font></td>
								 			<td class="tab_header_cell"><b>Name</b></td>
								 			<td class="tab_header_cell"><b>Email <br /> Address</b></td>
								 			<td class="tab_header_cell"><b>Username</b></td>
								 			<td class="tab_header_cell"><b>Access <br /> Rights</b></td>
						     			</tr>
						   
									    <%
									    Do While Not rstSearch.EOF And rstSearch.AbsolutePage = iPageCurrent
											UserID			= rstSearch("UserID") & ""
											UserIDArray		= CInt(rstSearch("UserID")) & ""
											Surname			= rstSearch("Surname") & ""
											FirstName		= rstSearch("FirstName") & "" 
											EmailAddress	= rstSearch("EmailAddress") & "" 
											Username		= rstSearch("Username") & ""
											AccessRights	= rstSearch("AccessID") & ""
		
											j = j + 1
											Response.Write "<tr height=""20"" class=""listTableText" & (j And 1) & """>"
					 						Response.Write "<td align=""center"">" & UserID & "</td>"
					 						
					 						Select Case Session("AccessRights")
											    Case "Level 1"
												   Response.Write "<td align=""center""><a href=""javascript:userSelected(" & UserID & ")""><img src=""../images/edit.gif"" width=""16"" height=""16"" border=""0"" alt=""Change Password""></a></td>" & vbCrLf
												   Response.Write "<td align=""center""><input type=""checkbox"" id=""DoDeleteAdmin"" name=""DoDeleteAdmin"" value=" & UserIDArray & "></td>" & vbCrLf
												      
											    Case "Level 4"
												   Response.Write "<td align=""center"">Not Available</td>" & vbCrLf
												   Response.Write "<td align=""center""><input type=""checkbox"" id=""DoDeleteAdmin"" name=""DoDeleteAdmin"" value="""" DISABLED></td>" & vbCrLf
										    End Select 
				     					
					 						Response.Write "<td align=""center"">"
					 						
					 						If IsNull(rstSearch.Fields("FirstName")) And IsNull(rstSearch.Fields("Surname")) Then
												Response.Write "<font color=""#ff0000""><i>No Name given</i></font>" & vbCrLf
					 						Else
												Response.Write rstSearch.Fields("FirstName") & "&nbsp;" & rstSearch.Fields("Surname") & vbCrLf
					 						End If			
					 					%>
					 						</td>
										 	<td align="center">
										<%
										 	If IsNull(rstSearch.Fields("EmailAddress")) Then
												Response.Write "<font color=""#ff0000""><i>No Email Address given</i></font>"
										 	Else
												Response.Write rstSearch.Fields("EmailAddress")
											End If			
										%>
										 	</td>
											<td align="center"><%=Username%></td>
											<td align="center">
										<%
											Select Case AccessRights
										    	Case "1"
											    	Response.Write "Level 1"

												Case "2"
												   Response.Write "Level 2"

											    Case "3"
											       Response.Write "Level 3"

												Case "4"
											       Response.Write "Level 4"

											End Select
										%>  
											</td>
										</tr>
					 
					 					<%
					    					rstSearch.MoveNext
										Loop
				     					%>
					 					</table>
					 					</form>

					 					<!--/* Add User Form */-->
									   	<form name="AddUserForm" id="AddUserForm" action="VESAAddAdminUser.asp" method="post"></form>
									   	<!--/* End Here */-->

									   	<form name="EditForm" id="EditForm" action="VESAEditAdminUser.asp" method="post">
									   		<input type="hidden" name="UserID" value="<%=UserID%>">
									   	</form>
				  					</td>
				  				</tr>
				  			</table>  
			   			</td>
			   		</tr>
			 
			   		<tr><td valign="top"><img src="../images/spacer.gif" width="1" height="10" alt="" /></td></tr>
		 
				<tr>
					<td>
						<table border="0" width="100%">
						<tr>
							<td>Page <strong><%=iPageCurrent%></strong> of <strong><%=iPageCount%></strong></td>
							<td>
							<div align="right">					
								<div class="pages">
									<% Call databasePaging() %>
								</div>
							</div>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				<% End If %>

				<% If Session("AccessRights") = "Level 1" Then %>
				<tr><td valign="top"><img src="../images/spacer.gif" width="1" height="20" alt="" /></td></tr>

				<tr>
					<td width="100%">
						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td><button type="button" class="pure-button" onClick="addNewUser()">Add Admin User</button></td>
							<td><img src="../images/spacer.gif" width="10" height="1" alt="" /></td>
							<td><button type="button" class="pure-button" onClick="deleteUser()">Delete Admin User</button></td>
						</tr>
						</table>
					</td>
				</tr>
				<% End If %>

				</table>
			</td>
		</tr>
		</table>
		<!--/* End Here */-->
	</div>
<% End Sub 

Sub databasePaging()
	If iPageCurrent > 1 Then 
	%>
		<a href="VESAViewAllAdminUsers.asp?page=<%=iPageCurrent - 1%>">&lt;&nbsp;Prev</a>
	<%
	Else
		Response.Write "<span class=""disabled"">&lt;&nbsp;Prev</span>"
	End If
							
	'--------------------------------------------------------------------------------------
	'- You can also show page numbers:
	For I = 1 To iPageCount
		'- Don't hyperlink the current page number
		If I = iPageCurrent Then
			Response.Write "<span class=""current"">" & I & "</span>" & vbCrLf
															
		Else
			Response.Write "<a href=""VESAViewAllAdminUsers.asp?page=" & I & """>" & I & "</a>" & vbCrLf
		End If
	'- I
	Next 
		If iPageCurrent < iPageCount Then
			Response.Write "<a href=""VESAViewAllAdminUsers.asp?page=" & iPageCurrent + 1 & """>Next&nbsp;&gt;</a>"
														
		Else
			Response.Write "<span class=""disabled"">Next&nbsp;&gt;</span>"
		End If
	'--------------------------------------------------------------------------------------
End Sub
%>


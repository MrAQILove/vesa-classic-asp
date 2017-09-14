<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE FILE="include/include.asp"-->
<!--#INCLUDE FILE="include/adovbs.inc"-->
<!--#INCLUDE FILE="include/functions.asp"-->
<%
'- Makes the browser not cache this page
Response.Expires = -1000 

'- Buffers the content so our Response.Redirect will work
Response.Buffer = True 

If Session("UserLoggedIn") <> "true" Then
   If Session("AccessRights") = "Level 1" Then
      Response.Redirect "AdminLogin.asp"
   Else
      Response.Redirect "VESAUnitLogin.asp"
   End If 

Else
	Call VESASave()
End If

Sub VESASave() 
%>
	<!DOCTYPE html>
    <html lang="en">
    	<head>
    		<meta charset="utf-8" />
    			<!--[if lt IE 9]>
					<script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
				<![endif]-->
			<title>VESA Members Database</title>
			<meta name="keywords" content="" />
			<meta name="VESA Members Database" content="" />
			<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
			<link rel="stylesheet" href="css/buttons.css">
			<link rel="stylesheet" href="css/forms.css">
			<link rel="stylesheet" href="css/base.css">
			<script type="text/javascript">
			<!--
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
		
		<!-- start section -->
		<section>
			<% 
			Select Case Session("AccessRights")
				Case "Level 1"
					Call adminMainMenu() 
			
			  	Case "Level 5"
			%>
					<aside id="sidebar1" class="sidebar">
						<ul>
							<li> 
								<h2>
								<% 
								EstablishConnection()

								Dim strVESAUnit
								Call showVESAUnit(Conn, rs, "VESA_tblUnit", Request.Form("VESAUnitID"), strVESAUnit)
							  
								Response.Write strVESAUnit & " Members"
							  
								CloseConnection()
								%> 
								</h2>
								
								<ul>
									<% Call displaySelectedMenu(Request.ServerVariables("SCRIPT_NAME")) %>
									<li><a href="VESAUnitLogin.asp">Log Out</a></li>
								</ul>
							</li>
						</ul>
					</aside>
			<% End Select %>
		
			<!-- start article -->
			<article id="content">
				<% Call DisplayHeader() %>
				<div class="entry">
				<%
				'--- Start ---'
					If Not Session("AccessRights") = "Level 4" Then
						Call displayFORMLinks()
					End If 

					'-----Declare Global Variable-------------------------------------------------------
					Dim strSQL 
					Dim objValueActionType, objValueRecipientID, objValueSurname_Organization, objValueFirstName, objValueAddress, objValueSuburb, objValuePostcode, objValueStateID
					Dim objValueMembershipNumber, objValueMemberEmailAddress, objValuePhoenixCopies, objValueVESAPocketDiary, objValueVESAWallCalendar, objValueVESAUnitID, objValue1

					Dim objValueDoDeleteID, objValueWhyDelete, objValueSpecifyReason, objValueDoActivateID, objValueDoDeleteAdmin, objValueChangedBy
					Dim objValueVESAUnit, objValueVESAUnitPassword, objValueUnitEmailAddress, objValueUnitSESRegionID, objValueDoDeleteVESAUnit, objValueDoActivateVESAUnit
					Dim objValueAdminUserUserID, objValueAdminUserSurname, objValueAdminUserFirstName, objValueAdminUserEmailAddress, objValueAdminUserUsername, objValueAdminUserPassword, objValueAdminUserDatabaseName, objValueAdminUserAccessRights, objValueAdminUserAccessID, objValueDoDeleteAdminUser
						
					objValueActionType				= Request.Form("ActionType")

					objValueRecipientID				= Request.Form("RecipientID")
					objValueSurname_Organization	= Request.Form("Surname_Organization")
					objValueFirstName				= Request.Form("FirstName")
					objValueAddress					= Request.Form("Address")
					objValueSuburb					= Request.Form("Suburb")
					objValuePostcode				= Request.Form("Postcode")
					objValueStateID					= Request.Form("StateID")
					objValueMembershipNumber		= Request.Form("MembershipNumber")
					objValueMemberEmailAddress		= Request.Form("MemberEmailAddress")
					objValuePhoenixCopies			= Request.Form("PhoenixCopies")
					objValueVESAPocketDiary			= Request.Form("VESAPocketDiary")
					objValueVESAWallCalendar		= Request.Form("VESAWallCalendar")
					objValueVESAUnitID				= Request.Form("VESAUnitID")
					objValue1						= "1"
					
					objValueDoDeleteID				= Request.Form("DoDelete")
					objValueWhyDelete				= Request.Form("WhyDelete") 
					objValueSpecifyReason			= Request.Form("SpecifyReason")
					objValueDoActivateID			= Request.Form("DoActivate")

					Select Case objValueWhyDelete
						Case 1
							objValueWhyDelete = "Area Closed"
						Case 2
							objValueWhyDelete = "Area Amalgamated"
						Case 3
							objValueWhyDelete = "No Longer Interested"
						Case 4
							objValueWhyDelete = "Other Reason"
					End Select 

					'-----VESA Unit/Distribution Section
					objValueVESAUnit				= Request.Form("VESAUnit")
					objValueVESAUnitPassword		= Request.Form("VESAUnitPassword")
					objValueUnitEmailAddress		= Request.Form("UnitEmailAddress")
					objValueUnitSESRegionID			= Request.Form("UnitSESRegionID")
					objValueIsUnitSES				= Request.Form("IsUnitSES")
						
					objValueDoDeleteVESAUnit		= Request.Form("DoDeleteVESAUnit")
					objValueDoActivateVESAUnit		= Request.Form("DoActivateVESAUnit")
						 
					Select Case Session("User")
						Case "Administrator" 
							objValueChangedBy = "Administrator"
					End Select

					'-----Administration Section
					objValueAdminUserUserID			= Request.Form("AdminUserUserID")
					objValueAdminUserSurname		= Request.Form("AdminUserSurname")
					objValueAdminUserFirstName		= Request.Form("AdminUserFirstName")
					objValueAdminUserEmailAddress	= Request.Form("AdminUserEmailAddress")
					objValueAdminUserUsername		= Request.Form("AdminUserUsername")
					objValueAdminUserPassword		= Request.Form("AdminUserPassword")
					objValueAdminUserAccessID		= Request.Form("AdminUserAccessID")
					objValueAdminUserDatabaseName	= Request.Form("AdminUserDatabaseName")
					objValueDoDeleteAdminUser		= Request.Form("DoDeleteAdminUser")
					
						
					'----- Establish a connection to the database
					EstablishConnection()

					Select Case Request.Form("ActionType")
						'----------------------------------------------------------------------------------
						'***** Add Member *****
						'----------------------------------------------------------------------------------
						Case "Add"
							strSQL = "SET NOCOUNT ON; INSERT INTO VESA_tblMembers (Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID)"    
							strSQL = strSQL & " VALUES "
							strSQL = strSQL & "('" & objValueSurname_Organization & "',"

							'----- Check if FirstName field is empty
							If objValueFirstName <> "" Then 
								strSQL = strSQL & "'" & objValueFirstName & "',"
							Else
								strSQL = strSQL & "NULL,"
							End If

							strSQL = strSQL & "'" & objValueAddress & "',"
							strSQL = strSQL & "'" & UCase(objValueSuburb) & "',"
							strSQL = strSQL & "'" & objValuePostcode & "',"
							strSQL = strSQL & "'" & objValueStateID & "',"
							strSQL = strSQL & "'" & objValueMembershipNumber & "',"

							'----- Check if MemberEmailAddress field is empty
							If objValueMemberEmailAddress <> "" Then 
								strSQL = strSQL & "'" & objValueMemberEmailAddress & "',"
							Else
								strSQL = strSQL & "NULL,"
							End If

							strSQL = strSQL & "'" & objValuePhoenixCopies & "',"
							strSQL = strSQL & "'" & objValueVESAPocketDiary & "',"
							strSQL = strSQL & "'" & objValueVESAWallCalendar & "',"
							strSQL = strSQL & "'" & objValueVESAUnitID & "');"
							strSQL = strSQL & " SELECT SCOPE_IDENTITY()"
							
							on error resume Next

							'----- Execute the SQL statement
							newID = Conn.Execute(strSQL, intRecordsAffected, adCmdText)(0)
														   
							If Err <> 0 Then
								DisplayErrorMessage()
							Else 
								If objValueFirstName <> "" And objValueSurname_Organization <> "" Then
									DisplaySuccessMessage("<h3>" & objValueFirstName & "&nbsp;" & UCase(objValueSurname_Organization) & " have been ADDED into the database.</h3>")
									Call ClearSessionVars()

								ElseIf objValueFirstName = "" And objValueSurname_Organization <> "" Then
									DisplaySuccessMessage("<h3>" & UCase(objValueSurname_Organization) & " have been ADDED into the database.</h3>")
									Call ClearSessionVars()
								End If 
								
								'-----Insert into VESA_tblAudit TABLE
								'Call UpdateAudit(newID)
								Call UpdateAudit(Conn, newID, objValueActionType, objValueChangedBy)
							End If				
							
						'----------------------------------------------------------------------------------
						'***** Update Member *****
						'----------------------------------------------------------------------------------
						Case "Update"
							strSQL = "UPDATE VESA_tblMembers SET"
							strSQL = strSQL & " Surname_Organization='" & UCase(objValueSurname_Organization) & "',"
								
							'----- Check if FirstName field is empty
							If objValueFirstName <> "" Then 
								strSQL = strSQL & " FirstName='" & objValueFirstName & "',"
							Else
								strSQL = strSQL & " FirstName=NULL,"
							End If

							strSQL = strSQL & " Address='" & objValueAddress & "',"
							strSQL = strSQL & " Suburb='" & UCase(objValueSuburb) & "',"
							strSQL = strSQL & " Postcode='" & objValuePostcode & "',"
							strSQL = strSQL & " StateID='" & objValueStateID & "',"
							strSQL = strSQL & " MembershipNumber='" & objValueMembershipNumber & "',"

							'----- Check if MemberEmailAddress field is empty
							If objValueMemberEmailAddress <> "" Then 
								strSQL = strSQL & "MemberEmailAddress='" & objValueMemberEmailAddress & "',"
							Else
								strSQL = strSQL & "MemberEmailAddress=NULL,"
							End If

							strSQL = strSQL & " PhoenixCopies='" & objValuePhoenixCopies & "',"
							strSQL = strSQL & " VESAPocketDiary='" & objValueVESAPocketDiary & "',"
							strSQL = strSQL & " VESAWallCalendar='" & objValueVESAWallCalendar & "',"
							strSQL = strSQL & " VESAUnitID='" & objValueVESAUnitID & "'"
							strSQL = strSQL & " WHERE RecipientID='" & objValueRecipientID & "'"

							on error resume Next

							'----- Execute the SQL statement
							Conn.Execute strSQL, adCmdText

							If Err <> 0 Then
								DisplayErrorMessage()
							Else 
								If objValueFirstName <> "" And objValueSurname_Organization <> "" Then
									DisplaySuccessMessage("<h3>" & objValueFirstName & "&nbsp;" & UCase(objValueSurname_Organization) & " have been UPDATED into the database.</h3>")
									ClearSessionVars()

								ElseIf objValueFirstName = "" And objValueSurname_Organization <> "" Then
									DisplaySuccessMessage("<h3>" & UCase(objValueSurname_Organization) & " have been UPDATED into the database.</h3>")
									ClearSessionVars()
								End If 
									
								'-----Insert into VESA_tblAudit TABLE
								'Call UpdateAudit(objValueRecipientID)
								Call UpdateAudit(Conn, objValueRecipientID, objValueActionType, objValueChangedBy)

							End If
						'----------------------------------------------------------------------------------
						
						'----------------------------------------------------------------------------------
						'***** Delete Member *****
						'----------------------------------------------------------------------------------
						Case "Delete"
							If objValueDoDeleteID = "" Then 'No items to delete 
								DisplayErrorMessage()

							Else 
								'***** deleting one member *****
								If Request("DoDelete").Count = 1 Then
									strSQL = "INSERT INTO VESA_tblDeletedMembers"
									strSQL = strSQL & " (RecipientID, Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID, WhyDelete, SpecifyReason, DateDeleted, IsDeleteMember)"  
									strSQL = strSQL & " SELECT"
									strSQL = strSQL & " RecipientID, Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID,"
										
									'----- Check if WhyDelete field is empty
									If objValueSpecifyReason <> "" Then 
										strSQL = strSQL & " '" & objValueWhyDelete & "',"
									Else
										strSQL = strSQL & " NULL,"
									End If

									'----- Check if SpecifyReason field is empty
									If objValueSpecifyReason <> "" Then 
										strSQL = strSQL & " '" & objValueSpecifyReason & "',"
									Else
										strSQL = strSQL & " NULL,"
									End If

									strSQL = strSQL & " '" & SQLDateTime(Now()) & "',"
									strSQL = strSQL & " '" & objValue1 & "'"
									strSQL = strSQL & " FROM VESA_tblMembers WHERE RecipientID='" & objValueDoDeleteID & "'"
											
									on error resume Next

									'-----Execute the SQL statement
									Conn.Execute strSQL, adCmdText + adExecuteNoRecords

									'***** SQL Statement Delete from TABLE (UFUA_tblMembers) *****
									strSQL = "DELETE FROM VESA_tblMembers"
									strSQL = strSQL & " WHERE RecipientID='" & objValueDoDeleteID & "'"
									
									'----- Execute the query using the connection object.
									Conn.Execute strSQL,,adCmdText + adExecuteNoRecords
										
								Else
									'***** deleting multiple Member *****
									Dim arrToDelete
									Dim intIndex

									If len(objValueDoDeleteID) > 0 Then
										arrToDelete = Split(objValueDoDeleteID,", ")
										For intIndex = 0 to Ubound(arrToDelete)
											Call DoDelete (Conn, arrToDelete(intIndex))
										Next
									End If
								End If 
										
								   
								If Err <> 0 Then
									DisplayErrorMessage()
								Else
									If Request("DoDelete").Count = 1 Then
										DisplaySuccessMessage("<h3>" & Request("DoDelete").Count & " member was deleted. <br /> The member with a Member ID of " & objValueDoDeleteID & " has been DELETED from the database</h3>")
									Else
										DisplaySuccessMessage("<h3>" & Request("DoDelete").Count & " members were deleted. <br /> The members with a Member ID of " & objValueDoDeleteID & " has been DELETED from the database</h3>")
									End If
								End If
							End If
						'----------------------------------------------------------------------------------
						
						'----------------------------------------------------------------------------------
						'***** Activate Member *****
						'----------------------------------------------------------------------------------
						Case "Activate"
							If objValueDoActivateID = "" Then 'No items to delete 
								DisplayErrorMessage() 
									 
							Else 
								strSQL = "INSERT INTO VESA_tblMembers"
								strSQL = strSQL & " (Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID)"  
								strSQL = strSQL & " SELECT "
								strSQL = strSQL & " Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID"
								strSQL = strSQL & " FROM VESA_tblDeletedMembers WHERE RecipientID='" & objValueDoActivateID & "'"	

								on error resume Next

								'-----Execute the SQL statement
								Conn.Execute strSQL, adCmdText + adExecuteNoRecords
								   
								If Err <> 0 Then
									DisplayErrorMessage()
								Else 
									If Request("DoActivate").Count = 1 Then
										DisplaySuccessMessage("<h3>" & Request("DoActivate").Count & " member was activated. <br /> The member with a Recipient ID of " & objValueDoActivateID & " has been ACTIVATED from the database</h3>")
									Else
										DisplaySuccessMessage("<h3>" & Request("DoActivate").Count & " members were activated. <br /> The members with a Recipient ID of " & objValueDoActivateID & " has been ACTIVATED from the database</h3>")
									End If
											
									'-----Insert into VESA_tblDeletedMembers TABLE
									'Call UpdateVESA_tblDeletedMembers(objValueDoActivateID)
									Call UpdateVESA_tblDeletedMembers(Conn, objValueDoActivateID)							
								End If
							End If
						'----------------------------------------------------------------------------------
						
						'----------------------------------------------------------------------------------
						'***** Add VESA Unit/Distribution *****
						'----------------------------------------------------------------------------------
						Case "AddVESAUnit"
							strSQL = "SET NOCOUNT ON; INSERT INTO VESA_tblUnit (VESAUnit, Password, EmailAddress, SESRegionID, IsUnitSES, IsActive)"
							strSQL = strSQL & " VALUES "
							strSQL = strSQL & "('" & objValueVESAUnit & "',"
							strSQL = strSQL & "'" & objValueVESAUnitPassword & "',"
							strSQL = strSQL & "'" & objValueUnitEmailAddress & "',"
							strSQL = strSQL & "'" & objValueUnitSESRegionID & "',"
							strSQL = strSQL & "'" & objValueIsUnitSES & "',"
							strSQL = strSQL & "'" & objValue1 & "');"
							strSQL = strSQL & " SELECT SCOPE_IDENTITY()"

							on error resume Next

							'-----Execute the SQL statement
							Conn.Execute strSQL, adCmdText + adExecuteNoRecords
														   
							If Err <> 0 Then
								DisplayErrorMessage()
							Else 
								DisplaySuccessMessage("<h3>" & objValueVESAUnit & " have been ADDED into the database.</h3>")
								ClearSessionVESAUnitVars() 
							End If
						'----------------------------------------------------------------------------------

						'----------------------------------------------------------------------------------
						'***** Update VESA Unit/Distribution *****
						'----------------------------------------------------------------------------------
						Case "UpdateVESAUnit"
							strSQL = "UPDATE VESA_tblUnit SET"
							strSQL = strSQL & " VESAUnit='" & objValueVESAUnit & "', "
							strSQL = strSQL & " Password='" & objValueVESAUnitPassword & "', "
							strSQL = strSQL & " EmailAddress='" & objValueUnitEmailAddress & "', "
							strSQL = strSQL & " SESRegionID='" & objValueUnitSESRegionID & "', "
							strSQL = strSQL & " IsUnitSES='" & objValue1 & "'"
							strSQL = strSQL & " WHERE VESAUnitID='" & objValueVESAUnitID & "'"

							On Error Resume Next

							'-----Execute the SQL statement
							Conn.Execute strSQL, adCmdText

							If Err <> 0 Then
								DisplayErrorMessage()
							Else 
								DisplaySuccessMessage("<h3>" & objValueVESAUnit & " details have been UPDATED into the database.</h3>")
								ClearSessionVESAUnitVars()
							End If
						'----------------------------------------------------------------------------------

						'----------------------------------------------------------------------------------
						'***** Delete VESA Unit/Distribution *****
						'----------------------------------------------------------------------------------
						Case "DeleteVESAUnit"
							If objValueDoDeleteVESAUnit = "" Then 'No items to delete 
								DisplayErrorMessage()
									 
							Else 
								strSQL = " UPDATE VESA_tblUnit"
								strSQL = strSQL & " SET VESA_tblUnit.IsActive = '0'"
								strSQL = strSQL & " From VESA_tblUnit A, VESA_tblMembers B"
								strSQL = strSQL & " WHERE A.VESAUnitID = B.VESAUnitID AND A.VESAUnitID='" & objValueDoDeleteVESAUnit & "'"
								strSQL = strSQL & " UPDATE VESA_tblMembers"
								strSQL = strSQL & " SET VESA_tblMembers.PhoenixCopies = '0', VESA_tblMembers.VESAPocketDiary = '0', VESA_tblMembers.VESAWallCalendar = '0'"
								strSQL = strSQL & " FROM VESA_tblUnit A, VESA_tblMembers B"
								strSQL = strSQL & " WHERE A.VESAUnitID = B.VESAUnitID" 
								strSQL = strSQL & " AND A.VESAUnitID='" & objValueDoDeleteVESAUnit & "'"

								On Error Resume Next

								'-----Execute the SQL statement
								Conn.Execute strSQL, adCmdText + adExecuteNoRecords

								If Err <> 0 Then
									DisplayErrorMessage()
								Else 
									'-----Display to the user that the product have been deleted.
									If Request("DoDeleteVESAUnit").Count = 1 Then
										DisplaySuccessMessage("<h3>" & Request("DoDeleteVESAUnit").Count & " VESA Unit/Distribution was deleted. <br /> The VESA Unit/Distribution with a VESA Unit/Distribution ID of " & objValueDoDeleteVESAUnit & " has been DELETED from the database</h3>")
									Else 
										DisplaySuccessMessage("<h3>" & Request("DoDeleteVESAUnit").Count & " VESA Unit/Distribution were deleted. <br /> The VESA Unit/Distribution with a Branch ID of " & objValueDoDeleteVESAUnit & " has been DELETED from the database</h3>")
									End If 
								End If
							End If 
						'----------------------------------------------------------------------------------

						'----------------------------------------------------------------------------------
						'***** Activate VESA Unit/Distribution *****
						'----------------------------------------------------------------------------------
						Case "ActivateVESAUnit"
							If objValueDoActivateVESAUnit = "" Then 'No items to activate 
								DisplayErrorMessage()
									 
							Else 
								strSQL = " UPDATE VESA_tblUnit"
								strSQL = strSQL & " SET VESA_tblUnit.IsActive = '1'"
								strSQL = strSQL & " From VESA_tblUnit A, VESA_tblMembers B"
								strSQL = strSQL & " WHERE A.VESAUnitID = B.VESAUnitID AND A.VESAUnitID='" & objValueDoActivateVESAUnit & "'"
								strSQL = strSQL & " UPDATE VESA_tblMembers"
								strSQL = strSQL & " SET VESA_tblMembers.PhoenixCopies = '1', VESA_tblMembers.VESAPocketDiary = '0', VESA_tblMembers.VESAWallCalendar = '1'"
								strSQL = strSQL & " FROM VESA_tblUnit A, VESA_tblMembers B"
								strSQL = strSQL & " WHERE A.VESAUnitID = B.VESAUnitID" 
								strSQL = strSQL & " AND A.VESAUnitID='" & objValueDoActivateVESAUnit & "'"

								On Error Resume Next

								'-----Execute the SQL statement
								Conn.Execute strSQL, adCmdText + adExecuteNoRecords

								If Err <> 0 Then
									DisplayErrorMessage()
								Else 
									'-----Display to the user that the product have been deleted.
									If Request("DoActivateVESAUnit").Count = 1 Then
										DisplaySuccessMessage("<h3>" & Request("DoActivateVESAUnit").Count & " VESA Unit/Distribution was activated. <br /> The VESA Unit/Distribution with a VESA Unit/Distribution ID of " & objValueDoActivateVESAUnit & " has been ACTIVATED into the database</h3>")
									Else 
										DisplaySuccessMessage("<h3>" & Request("DoActivateVESAUnit").Count & " VESA Unit/Distribution were activated. <br /> The VESA Unit/Distribution with a Branch ID of " & objValueDoActivateVESAUnit & " has been ACTIVATED INTO the database</h3>")
									End If 
								End If
							End If 
						'----------------------------------------------------------------------------------

						'----------------------------------------------------------------------------------
						'***** Add Administration User *****
						'----------------------------------------------------------------------------------
						Case "AddAdministrationUser"
							strSQL = "SET NOCOUNT ON; INSERT INTO MembersDB_tblUsers (Surname, FirstName, EmailAddress, UserName, Password, RegistrationDate, DatabaseName, UserActive)"
							strSQL = strSQL & " VALUES "
							strSQL = strSQL & "('" & objValueAdminUserSurname & "',"
							strSQL = strSQL & "'" & objValueAdminUserFirstName & "',"
							strSQL = strSQL & "'" & objValueAdminUserEmailAddress & "',"
							strSQL = strSQL & "'" & objValueAdminUserUsername & "',"
							strSQL = strSQL & "'" & objValueAdminUserPassword & "',"
							strSQL = strSQL & "'" & SQLDateTime(Now()) & "',"
							strSQL = strSQL & "'" & objValueAdminUserDatabaseName & "',"
							strSQL = strSQL & "'" & objValue1 & "');"
							strSQL = strSQL & " SELECT SCOPE_IDENTITY()"

							On Error Resume Next

							'-----Execute the SQL statement
							newUserRightID = Conn.Execute(strSQL, intRecordsAffected, adCmdText)(0)

							If Err <> 0 Then
								DisplayErrorMessage()
							Else 
								DisplaySuccessMessage("<h3>Admin User - <i>" & objValueUsername & "</i> have been ADDED into the admin user database.</h3>")
								Call ClearSessionVESAAdminstrationUserVars()

								'-----Insert User Access into MembersDB_tblUserRights
								'Call UpdateUserRights(newUserRightID)
								Call UpdateUserRights(Conn, newUserRightID)
							End If
						'----------------------------------------------------------------------------------

						'----------------------------------------------------------------------------------
						'***** EDIT Administration User *****
						'----------------------------------------------------------------------------------
						Case "EditAdministrationUser"
							strSQL = "UPDATE MembersDB_tblUsers SET"
							strSQL = strSQL & " Surname='" & objValueAdminUserSurname & "', "
							strSQL = strSQL & " FirstName='" & objValueAdminUserFirstName & "', "
							strSQL = strSQL & " EmailAddress='" & objValueAdminUserEmailAddress & "', "
							strSQL = strSQL & " Username='" & objValueAdminUserUsername & "', "
							strSQL = strSQL & " Password='" & objValueAdminUserPassword & "', "
							strSQL = strSQL & " UserActive='" & objValue1 & "'"
							strSQL = strSQL & " WHERE UserID='" & objValueAdminUserUserID & "'"

							On Error Resume Next

							'-----Execute the SQL statement
							Conn.Execute strSQL, adCmdText
							
							If Err <> 0 Then
								DisplayErrorMessage()
							Else 
								DisplaySuccessMessage("<h3>" & objValueAdminUserFirstName & "&nbsp;" & objValueAdminUserSurname & " (" & objValueAdminUserUsername & ")" &   " details have been UPDATED into the database.</h3>")
								Call ClearSessionVESAAdminstrationUserVars()
								'Call UpdateUserRights(objValueAdminUserUserID)
									Call UpdateUserRights(Conn, objValueAdminUserUserID)
								End If
						'----------------------------------------------------------------------------------
						
						'----------------------------------------------------------------------------------
						'***** Delete Administration User *****
						'----------------------------------------------------------------------------------
						Case "DeleteAdministrationUser"
							If objValueDoDeleteAdminUser = "" Then 'No Administration User to delete 
								DisplayErrorMessage()
									 
							Else 
								strSQL = "UPDATE MembersDB_tblUsers SET UserActive='0'" 
								strSQL = strSQL & " WHERE UserID IN (" & objValueDoDeleteAdminUser & ")" 			

								On Error Resume Next

								'-----Execute the SQL statement
								Conn.Execute strSQL, adCmdText + adExecuteNoRecords

								If Err <> 0 Then
									DisplayErrorMessage()
								Else 
									'-----Display to the user that the product have been deleted.
									If Request("DoDeleteAdminUser").Count = 1 Then
										DisplaySuccessMessage("<h3>" & Request("DoDeleteAdminUser").Count & " member was deleted...<br />The member with a User ID of " & objValueDoDeleteAdminUser & " has been DELETED from the database</h3>")
									Else 
										DisplaySuccessMessage("<h3>" & Request("DoDeleteAdminUser").Count & " members were deleted...<br />The member with a User ID of " & objValueDoDeleteAdminUser & " has been DELETED from the database</h3>")
									End If  
								End If
							End If 
						'----------------------------------------------------------------------------------
						
						Case Else
							Response.Redirect "AdminLogin.asp"
					End Select

					CloseConnection()
					'-----------------------------------------------------------------------------------
					%>
				</div>
			</article>
			<!-- end article -->
			<div style="clear: both;">&nbsp;</div>
		</section>
		
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
<% End Sub

'--------------------------------------------------------------------------------------
'***** Display the Header according to the action *****		
'--------------------------------------------------------------------------------------
Sub DisplayHeader()
	Response.Write "<h1 class=""title""><a href=""#"">Success!</a></h1>" & vbCrLf

	Select Case Request.Form("ActionType")
		Case "Add"
			Response.Write "<p class=""byline""><b>You have successfully added a new member.</b></p>"
						   
		Case "Update"
			Response.Write "<p class=""byline""><b>You have successfully updated a member's details.</b></p>"

		Case "Delete"
			Response.Write "<p class=""byline""><b>You have successfully deleted a member.</b></p>"

		Case "Activate"
			Response.Write "<p class=""byline""><b>You have successfully activated a member.</b></p>"

		Case "AddVESAUnit"
			Response.Write "<p class=""byline""><b>You have successfully added a new VESA Unit.</b></p>"

		Case "UpdateVESAUnit"
			Response.Write "<p class=""byline""><b>You have successfully updated a VESA Unit's details.</b></p>"

		Case "DeleteVESAUnit"
			Response.Write "<p class=""byline""><b>You have successfully deleted a VESA Unit.</b></p>"

		Case "AddAdministrationUser"
			Response.Write "<p class=""byline""><b>You have successfully added a new Administration user.</b></p>"

		Case "EditAdministrationUser"
			Response.Write "<p class=""byline""><b>You have successfully updated the Administration user.</b></p>"

		Case "DeleteAdministrationUser"
			Response.Write "<p class=""byline""><b>You have successfully deleted an Administration user.</b></p>"
	End Select 
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine INSERT INTO AUDIT TABLE (VESA_tblAudit) *****
'***** Usage: UpdateAudit(Conn, auditID, objValueActionType, objValueChangedBy)		
'----------------------------------------------------------------------------------
Sub UpdateAudit(c, auditID, ActionType, ValueChangeBy)
	Dim strSQL

	strSQL = "INSERT INTO VESA_tblAudit"
	strSQL = strSQL & " (RecipientID, Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID, IsAuditMember, ActionType, ActionDateTime, ChangedBy)"  
	strSQL = strSQL & " SELECT "
	strSQL = strSQL & " '" & auditID & "', Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID,"
	strSQL = strSQL & " '" & objValue1 & "',"
	strSQL = strSQL & " '" & ActionType & "',"
	strSQL = strSQL & " '" & SQLDateTime(Now()) & "',"
	strSQL = strSQL & " '" & ValueChangeBy & "' "
									 
	If ActionType = "Delete" Then
		strSQL = strSQL & "FROM VESA_tblDeletedMembers WHERE RecipientID='" & auditID & "'"
									 
	ElseIf ActionType = "Activate" Then
		strSQL = strSQL & "FROM VESA_tblMembers WHERE RecipientID='" & auditID & "'"
								 
	Else
		strSQL = strSQL & "FROM VESA_tblMembers WHERE RecipientID='" & auditID & "'"
	End If 

	c.Execute strSQL, adCmdText + adExecuteNoRecords
												
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine Insert into VESA_tblDeletedMembers and  
'***** Delete from TABLE (VESA_tblMembers)
'***** Usage: DoDelete(Conn, intRecipientID)	 	
'----------------------------------------------------------------------------------
Sub DoDelete(c, intRecipientID)
	Dim objValueWhyDelete, objValueSpecifyReason , objValue1
	Dim strSQL

	objValueWhyDelete		= Request.Form("WhyDelete")
	objValueSpecifyReason	= Request.Form("SpecifyReason")
	objValue1				= "1"

	'***** SQL Statement Insert into TABLE (VESA_tblDeletedMembers) *****
	strSQL = "INSERT INTO VESA_tblDeletedMembers"
	strSQL = strSQL & " (RecipientID, Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress, PhoenixCopies,"
	strSQL = strSQL & " VESAPocketDiary, VESAWallCalendar, VESAUnitID, WhyDelete, SpecifyReason, DateDeleted, IsDeleteMember)"
	strSQL = strSQL & " SELECT"
	strSQL = strSQL & " RecipientID, Surname_Organization, FirstName, Address, Suburb, Postcode, StateID, MembershipNumber, MemberEmailAddress," 
	strSQL = strSQL & " PhoenixCopies, VESAPocketDiary, VESAWallCalendar, VESAUnitID,"
											
	'----- Check if WhyDelete field is empty
	If objValueWhyDelete <> "" Then 
		strSQL = strSQL & " '" & objValueWhyDelete & "',"
	Else
		strSQL = strSQL & " NULL,"
	End If

	'----- Check if SpecifyReason field is empty
	If objValueSpecifyReason <> "" Then 
		strSQL = strSQL & " '" & objValueSpecifyReason & "',"
	Else
		strSQL = strSQL & " NULL,"
	End If

	strSQL = strSQL & " '" & SQLDateTime(Now()) & "',"
	strSQL = strSQL & " '" & objValue1 & "'"
	strSQL = strSQL & " FROM VESA_tblMembers WHERE RecipientID='" & intRecipientID & "'"

	c.Execute strSQL,,adCmdText + adExecuteNoRecords
									
	'***** SQL Statement Delete from TABLE (VESA_tblMembers) *****
	strSQL = "DELETE FROM VESA_tblMembers"
	strSQL = strSQL & " WHERE RecipientID='" & intRecipientID & "'"
										
	c.Execute strSQL,,adCmdText + adExecuteNoRecords
									
	'Call UpdateAudit(intRecipientID)
	Call UpdateAudit(Conn, intRecipientID, objValueActionType, objValueChangedBy)
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine Delete from TABLE (VESA_tblDeletedMembers) *****
'***** Usage: UpdateVESA_tblDeletedMembers(Conn, activatedID)			
'----------------------------------------------------------------------------------
Sub UpdateVESA_tblDeletedMembers(c, activatedID)
	Dim strSQL
						 
	strSQL = "DELETE FROM VESA_tblDeletedMembers"
	strSQL = strSQL & " WHERE RecipientID='" & activatedID & "'"
									 
	c.Execute strSQL, adCmdText + adExecuteNoRecords
									 
	'Call UpdateAudit(activatedID)
	Call UpdateAudit(Conn, activatedID, objValueActionType, objValueChangedBy)
						
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine Insert into TABLE (MembersDB_tblUserRights) *****	
'***** Usage: UpdateUserRights(Conn, newUserRightID)	
'----------------------------------------------------------------------------------
Sub UpdateUserRights(c, newUserRightID)
	Dim strSQL

	Select Case Request.Form("ActionType")
		Case "AddAdministrationUser"
			strSQL = "INSERT INTO MembersDB_tblUserRights (UserID, AccessID)"
			strSQL = strSQL & " SELECT "
			strSQL = strSQL & " " & newUserRightID & ","
			strSQL = strSQL & "'" & objValueAdminUserAccessID & "' "
			strSQL = strSQL & "FROM MembersDB_tblUsers WHERE UserID = " & newUserRightID
										
		Case "EditAdministrationUser"
			strSQL = "UPDATE MembersDB_tblUserRights SET"
			strSQL = strSQL & " AccessID='" & objValueAdminUserAccessID & "'"
			strSQL = strSQL & " WHERE UserID='" & newUserRightID & "'"
	End Select

	on error resume Next
									
	'-----Execute the SQL statement
	c.Execute strSQL, adCmdText + adExecuteNoRecords
									
	If Err <> 0 Then
		Call DisplayErrorMessage() 
	End If 
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine CLEARING SESSION VARIABLES *****		
'----------------------------------------------------------------------------------
'- This function clears the session variables in the situaton where we don't need the values any more.
Sub ClearSessionVars()
	Session("Surname_Organization") = ""
	Session("FirstName") = ""
	Session("Address") = ""
	Session("Suburb") = ""
	Session("Postcode") = ""
	Session("StateID") = ""
	Session("MembershipNumber") = ""
	Session("MemberEmailAddress") = ""
	Session("PhoenixCopies") = ""
	Session("VESAPocketDiary") = ""
	Session("VESAWallCalendar") = ""
	If Session("AccessRights") = "Level 1" Then
		Session("VESAUnitID") = ""
	End If 
	Session("SESRegionID") = ""
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine CLEARING SESSION VARIABLES for VESA Unit/Distribution form *****		
'----------------------------------------------------------------------------------
'- This function clears the session variables in the situation where we don't need the values any more.
Sub ClearSessionVESAUnitVars()
	Session("VESAUnit") = ""
	Session("VESAUnitPassword") = ""
	Session("UnitEmailAddress") = ""
	Session("UnitSESRegionID") = ""
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine CLEARING SESSION VARIABLES for VESA Administration User form *****		
'----------------------------------------------------------------------------------
'- This function clears the session variables in the situation where we don't need the values any more.
Sub ClearSessionVESAAdminstrationUserVars()
	Session("Surname") = ""
	Session("FirstName") = ""
	Session("EmailAddress") = ""
	Session("Username") = ""
	Session("Password") = ""
	Session("DatabaseName") = ""
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine DISPLAY SUCCESS MESSAGE *****		
'----------------------------------------------------------------------------------
Sub DisplaySuccessMessage(strMessage)
	Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write "<td>" & vbCrLf
	Response.Write "<p><b>" & strMessage & "</B></p>" & vbCrLf
									 
	Select Case Session("User")
		Case "Administrator"
			Response.Write "<p><button type=""button"" class=""pure-button"" onClick=""javascript:document.location.href='VESAMain.asp'"">Back to Main</button></p>" & vbCrLf
										
		Case Else
			Select Case Request.Form("ActionType")
				Case "Add"
					Response.Write "<p><button type=""button"" class=""pure-button"" onClick=""addMember()"">Back to Add Member Form</button></p>" & vbCrLf
										
				Case "Delete"
					Response.Write "<p><button type=""button"" class=""pure-button"" onClick=""deleteMember()"">Back to Delete Member Form</button></p>" & vbCrLf

				Case Else
					Response.Write "<p><button type=""button"" class=""pure-button"" onClick=""javascript:document.location.href='VESAMain.asp'"">Back to Main</button></p>" & vbCrLf
			End Select 
	End Select
									
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr><td><img src=""images/spacer.gif"" width=""10"" height=""30"" border=""0""></td></tr>" & vbCrLf
	Response.Write "</table>"
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Subroutine DISPLAY ERROR MESSAGE *****		
'----------------------------------------------------------------------------------
Sub DisplayErrorMessage()
	Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""100%"">" & vbCrLf
	Response.Write "<tr>" & vbCrLf
	Response.Write "<td>" & vbCrLf
											  
	Select Case Request.Form("ActionType")
		Case "Delete"
			Response.Write "<p><b>You did not select any members to delete!</b></p>" & vbCrLf
								
		Case "Activate"
			Response.Write "<p><b>You did not select any members to activate!</b></p>" & vbCrLf
								
		Case "DeleteVESAUnit"
			Response.Write "<p><b>You did not select any VESA Unit/Distribution to delete!</b></p>" & vbCrLf

		Case "DeleteAdministrationUser"
			Response.Write "<p><b>You did not select any Administration User to delete!</b></p>" & vbCrLf
										
		Case Else
			Response.Write "<p><b>An error occurred in the execution of this ASP page <br />Please report the following information to the support desk.</b></p>" & vbCrLf
			Response.write "<p><b>Page Error Object</b><br />" & vbCrLf
			Response.Write "Error Number: " & Err.Number & "<br />" & vbCrLf
			Response.Write "Error Description: "  & Err.Description & "<br />" & vbCrLf 
			Response.Write "Source: " & Err.Source & "<br />" & vbCrLf 
			Response.Write "</p>" & vbCrLf
	End Select 
											  
	Response.Write "<p><button type=""button"" class=""pure-button"" onClick=""javascript:document.location.href='VESAMain.asp'"">Back to Main</button></p>" & vbCrLf
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "<tr><td><img src=""images/spacer.gif"" width=""10"" height=""30"" border=""0""></td></tr>" & vbCrLf
	Response.Write "</table>" & vbCrLf
End Sub
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'***** Function SQL Time and Date *****		
'----------------------------------------------------------------------------------
Function SQLDate(currDate)	
	Dim tempDate
							
	tempDate = fillDigit(Day(currDate))
	tempDate = tempDate & " "
	tempDate = tempDate & MonthName(Month(currDate))
	tempDate = tempDate & " "
	tempDate = tempDate & Year(currDate)
	SQLDate = tempDate
	tempDate = ""
End Function

Function SQLDateTime(currDate)
	'-----Format for this function:	
	Dim tempDate
							
	tempDate = fillDigit(Day(currDate))
	tempDate = tempDate & " "
	tempDate = tempDate & MonthName(Month(currDate))
	tempDate = tempDate & " "
	tempDate = tempDate & Year(currDate)
	SQLDateTime = tempDate & " " & FormatDateTime(currDate, 3)
	tempDate = ""
End Function
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
'- This function simply adds a zero to the front of a number if it is one character in length
'- This is of particular use when forming dates for SQL Strings
'-----------------------------------------------------------------------------------							
Function fillDigit(intDigit)
	Select Case len(intDigit)
		Case 1
			fillDigit = "0" & intDigit
		Case Else
			fillDigit = intDigit
	End Select
End Function
'----------------------------------------------------------------------------------
%>

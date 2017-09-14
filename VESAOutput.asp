<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="include/include.asp"-->
<!--#include file="include/functions.asp"--> 

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
	'- Constants ripped from adovbs.inc:
	Const adOpenStatic = 3
	Const adLockReadOnly = 1
	Const adCmdText = &H0001

	'- Our own constants:
	Const PAGE_SIZE = 300	' The size of our pages.

	'- Declare our variables... always good practice!
	Dim rstSearch			' ADO recordset
	Dim strSQL				' The SQL Query we build on the fly
	Dim strSearchFor		' The text being looked for

	Dim iPageCurrent		' The page we're currently on
	Dim iPageCount			' Number of pages of records
	Dim iRecordCount		' Count of the records returned
	Dim I					' Standard looping variable

	'- Retrieve page to show or default to the first
	If Request.QueryString("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = CInt(Request.QueryString("page"))
	End If

	'- Retreive the term being searched for.  I'm doing it on
	'- the QS since that allows people to bookmark results.
	'- You could just as easily have used the form collection.

	strActiveUnit = Request.Form("active")
	strInactiveUnit = Request.Form("inactive")

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		strSearch		= Request.Form("search")

		Select Case strSearch
			Case 1
				strSearch = "All"
			
			Case 2
				strSearch		= "RecipientID"
				strSearchFor	= Request.Form("searchForRecipientID")

			Case 3
				strSearch		= "Surname/Organisation"
				strSearchFor	= Request.Form("searchForSurname_Organisation")
	  
			Case 4
	 			strSearch		= "First Name"
				strSearchFor	= Request.Form("searchForFirstName")

			Case 5
				strSearch		= "Address"
				strSearchFor	= Request.Form("searchForAddress")
	  
			Case 6
				strSearch		= "Suburb"
				strSearchFor	= Request.Form("searchForSuburb")

			Case 7
				strSearch		= "Postcode"
				strSearchFor	= Request.Form("searchForPostcode")
	  
			Case 8
				strSearch		= "State"
				strSearchFor	= Request.Form("searchForState")

			Case 9
				strSearch		= "Membership Number"
				strSearchFor	= Request.Form("searchForMembershipNumber")

			Case 10
				strSearch		= "VESA Unit"
				strSearchFor	= Request.Form("searchForVESAUnit")
				
			Case 11
				strSearch		= "SES Region"
				strSearchFor	= Request.Form("searchForSESRegion")

			Case 12
				strSearch		= "VESA Unit"
				strUnit			= Request.Form("searchForVESAUnit")
				strSearchFor	= Replace(strUnit, ", ", "")

			Case 13
				strSearch		= "VESA Unit"
				strUnit			= Request.Form("searchForVESAUnit")
				strSearchFor	= Replace(strUnit, ", ", "")
			
			Case Else
				strSearch		= "Please Choose"
			End Select
	Else 
		strSearch		= Request.QueryString("search")
		strSearchFor	= Request.QueryString("searchFor")
	End If

	' -Start Query Here ----------------------------------------------------------------------------
	Select Case strSearch
		Case "RecipientID"
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying Member with a Recipient ID of " & strSearchFor)
				Call OutputPage()
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE RecipientID='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
    
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying Member with a Recipient ID of " & strSearchFor)
	 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
   
				CloseConnection()
			End If  
		'--------------------------------------------------------------------------------------------

		Case "Surname/Organisation"
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying All Members with a Surname/Organisation of " & strSearchFor)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE M.Surname_Organization='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
		
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying All Members with a Last Name of " & strSearchFor)
		 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
         End If
      '--------------------------------------------------------------------------------------------

		Case "First Name"
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying All Members with a First Name of " & strSearchFor)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE FirstName='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
    
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying All Members with a First Name of " & strSearchFor)
		 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
			End If
		'--------------------------------------------------------------------------------------------

		Case "Address"
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying Member with an Address of " & strSearchFor)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE M.Address='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
		
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying Member with an Address of " & strSearchFor)
		 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
			End If
		'--------------------------------------------------------------------------------------------

		Case "Suburb"
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying All Members in the Suburb of " & strSearchFor)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE Suburb='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
		
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying All Members in the Suburb of " & strSearchFor)
		 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
			End If  
		'--------------------------------------------------------------------------------------------

		Case "Postcode"
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying All Members in the Postcode of " & strSearchFor)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE Postcode='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
		
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying All Members in the Postcode of " & strSearchFor)
		 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
			End If  
		'--------------------------------------------------------------------------------------------

		Case "State"
			Dim strState
			Select Case strSearchFor
				Case 1
					strState = "Australian Capital Territory"
				Case 2
					strState = "New South Wales"
				Case 3
					strState = "Northern Territory"
				Case 4
					strState = "Queensland"
				Case 5
					strState = "South Australia"
				Case 6
					strState = "Tasmania"
				Case 7
					strState = "Victoria"
				Case 8
					strState = "Western Australia"
			End Select   
	     
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying All Members in the State of " & strState)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE S.StateID='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
		
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying All Members in the State of " & strState)
		 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
			End If  
		'--------------------------------------------------------------------------------------------

		Case "Membership Number"
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying Member with a Membership Number of " & strSearchFor)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE MembershipNumber='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
    
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying Member with a Membership Number of " & strSearchFor)
	 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
   
				CloseConnection()
			End If
		'--------------------------------------------------------------------------------------------

		Case "VESA Unit"
			EstablishConnection()

			If strSearchFor = "" Then
				Call showVESAUnit(Conn, rs, "VESA_tblUnit", strSearchFor, strVESAUnit)
				DisplayHeader("VESA Members Database : Displaying All Members in the VESA Unit of " & strVESAUnit)
				Call OutputPage()
      
			Else
				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE M.VESAUnitID='" & strSearchFor & "'"
				
				If strActiveUnit = "1" Then
					strSQL = strSQL & " AND U.IsActive='1'"
				
				ElseIf strInactiveUnit = "1" Then
					strSQL = strSQL & " AND U.IsActive='0'"
				End If 
				
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"

				' Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				' Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				' Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				Call showVESAUnit(Conn, rs, "VESA_tblUnit", strSearchFor, strVESAUnit)
				DisplayHeader("VESA Members Database : Displaying All Members in the VESA Unit of " & strVESAUnit)
		  
				Call OutputPage()

				' Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
			End If
		'--------------------------------------------------------------------------------------------

		Case "SES Region"
			Dim strSESRegion
			Select Case strSearchFor
				Case 1
					strSESRegion = "Central"
				Case 2
					strSESRegion = "South West"
				Case 3
					strSESRegion = "East"
				Case 4
					strSESRegion = "North East"
				Case 5
					strSESRegion = "Mid West"
				Case 6
					strSESRegion = "North West"
			End Select   
	     
			If strSearchFor = "" Then
				DisplayHeader("VESA Members Database : Displaying All Members in the SES Region of " & strSESRegion)
				Call OutputPage()
      
			Else
				EstablishConnection()

				'- Build our query based on the input.
				strSQL = "SELECT * FROM VESA_tblMembers M"
				strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
				strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
				strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
				strSQL = strSQL & " WHERE R.SESRegionID='" & strSearchFor & "'"
				strSQL = strSQL & " AND U.IsActive='1'"
				strSQL = strSQL & " ORDER BY Surname_Organization ASC"
		
				'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
				Set rstSearch = Server.CreateObject("ADODB.Recordset")
				rstSearch.PageSize  = PAGE_SIZE
				rstSearch.CacheSize = PAGE_SIZE

				'- Open our recordset
				rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

				'- Get a count of the number of records and pages for use in building the header and footer text.
				iRecordCount = rstSearch.RecordCount
				iPageCount   = rstSearch.PageCount

				DisplayHeader("VESA Members Database : Displaying All Members in the SES Region of " & strSESRegion)
		 
				Call OutputPage()

				'- Close our recordset and connection and dispose of the objects
				rstSearch.Close
				Set rstSearch = Nothing
	   
				CloseConnection()
			End If  
		'--------------------------------------------------------------------------------------------

		Case "All"
			EstablishConnection()
	     
			'- Build our query based on the input.
			strSQL = "SELECT * FROM VESA_tblMembers M"
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
			strSQL = strSQL & " LEFT JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE (U.IsActive = '1')"
			strSQL = strSQL & " ORDER BY Surname_Organization ASC"
			
			'- Execute our query using the connection object.  It automatically creates and returns a recordset which we store in our variable.
			Set rstSearch = Server.CreateObject("ADODB.Recordset")
			rstSearch.PageSize  = PAGE_SIZE
			rstSearch.CacheSize = PAGE_SIZE

			'- Open our recordset
			rstSearch.Open strSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

			'- Get a count of the number of records and pages for use in building the header and footer text.
			iRecordCount = rstSearch.RecordCount
			iPageCount   = rstSearch.PageCount

			DisplayHeader("VESA Members Database : Displaying All Members")
      
			Call OutputPage()

			'- Close our recordset and connection and dispose of the objects
			rstSearch.Close
			Set rstSearch = Nothing
   
			CloseConnection()
			'--------------------------------------------------------------------------------------------
			
		Case Else
			EstablishConnection()
			DisplayHeader("VESA Members Database : Error in your search")
			Call OutputPage()
			CloseConnection()
	End Select 
End If

Sub DisplayHeader(strMessage) %>
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title><%=strMessage%></title>
		<meta name="keywords" content="" />
		<meta name="VESA Members Database" content="" />
		<link href="css/databaseView.css" rel="stylesheet" type="text/css" media="screen" />
		<link rel="stylesheet" href="css/buttons.css">
		<link rel="stylesheet" href="css/forms.css">
		<link rel="stylesheet" href="css/base.css">
		<link rel="stylesheet" href="css/grids.css">
		<script language="javascript">
		<!--

		function stopSubmit() {
		   return false;
		}

		// Log out
		function logOut() 
		{
			<% If Session("AccessRights") = "Level 1" Then %>
				document.location.href = "AdminLogin.asp";
			<% Else %>
				document.location.href = "VESAUnitLogin.asp";
			<% End If %>
		}

		// Back to Main
		function goBack() {
		   document.location.href = "VESAMain.asp";
		}

		// Search Member
		function searchAgain() {
		   document.location.href = "VESASearch.asp";
		}

		// Add New Member 
		function addNewMember()
		{
		   if (<%=Session("VESAID")%> == 1) {
			  document.AddForm.submit();
		   }
		}

		// Show Member History
		function showHistory() 
		{
		   if (<%=Session("VESAID")%> == 1) {
			  document.ShowHistoryForm.submit();
		   }
		}

		//Member Selected
		function memberSelected(strUnit)
		{
		   if(<%=Session("VESAID")%> == 1) 
		   {
			  document.EditForm.RecipientID.value = strUnit;
			  document.EditForm.submit();
		   }
		}

		// Export Member List into an Excel File 
		function exportFile(){
		   document.ExportForm.submit();
		}

		// Delete Member
		function deleteSelected()
		{
		   var ctr;
		   
		   ctr = 0;
		   
		   // check for single checkbox by seeing if an array has been created
		   var cblength = document.forms['MultiDeleteForm'].elements['DoDelete'].length;
		   if(typeof cblength == "undefined")
		   {
			  if(document.forms['MultiDeleteForm'].elements['DoDelete'].checked == true) ctr++;
		   }
		   else
		   {
			  for(i = 0; i < document.forms['MultiDeleteForm'].elements['DoDelete'].length; i++)
			  {
				 if(document.forms['MultiDeleteForm'].elements['DoDelete'][i].checked) ctr++;
			  }
			}
						  
		   if (ctr == 1) 
		   {
			   var answer;
			   answer = confirm('Are you sure you want to delete this member?');
			   if (answer)
			   {
				  document.MultiDeleteForm.submit();
				  return false;   
			   }

			   //else {;}
			}
			
			else if (ctr > 1) 
			{
			   var answer;
			   answer = confirm('Are you sure you want to delete ' + ctr + ' members?');
			   if (answer)
			   {
				  document.MultiDeleteForm.submit();
				  return false;
			   }

			   //else {;}
			}
			
			else 
			{
			   confirm("No members selected for deletion");
			   return true;
			}
		}

		// Select all Members to be deleted
		checked=false;
		function checkedAll (MultiDeleteForm) 
		{
			var aa= document.getElementById('MultiDeleteForm');
			if (checked == false) {
				checked = true
			}
			else {
				checked = false
			}
			
			for (var i =0; i < aa.elements.length; i++) {
				aa.elements[i].checked = checked;
			}
		}

		// Activate Member
		function activateSelected()
		{
		   var ctr;
		   
		   ctr = 0;
		   
		   // check for single checkbox by seeing if an array has been created
		   var cblength = document.forms['ActivateForm'].elements['DoActivate'].length;
		   if(typeof cblength == "undefined")
		   {
			  if(document.forms['ActivateForm'].elements['DoActivate'].checked == true) ctr++;
		   }
		   else
		   {
			  for(i = 0; i < document.forms['ActivateForm'].elements['DoActivate'].length; i++)
			  {
				 if(document.forms['ActivateForm'].elements['DoActivate'][i].checked) ctr++;
			  }
			}
						  
		   if (ctr == 1) 
		   {
			   var answer;
			   answer = confirm('Are you sure you want to activate this member?');
			   if (answer)
			   {
				  document.ActivateForm.submit();
				  return false;   
			   }

			   //else {;}
			}
			
			else if (ctr > 1) 
			{
			   var answer;
			   answer = confirm('Are you sure you want to activate ' + ctr + ' members?');
			   if (answer)
			   {
				  document.ActivateForm.submit();
				  return false;
			   }

			   //else {;}
			}
			
			else 
			{
			   confirm("No members selected for deletion");
			   return true;
			}
		}
		//-->
		</script>
	</head>
<% End Sub

Sub OutputPage() %>
	<body>
	<div id="wrapper">
		<div id="menu">
			<ul id="main">
				<li><a href="VESAMain.asp">Home</a></li>
				<li><a href="http://www.vesa.com.au/">VESA Website</a></li>
				<li><a href="http://www.cwaustral.com.au/">Countrywide Austral</a></li>
				<li><a href="contactUs.html">Contact Us</a></li>
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
			<% Call viewMembers() %>
			<div style="clear: both;">&nbsp;</div>
		</div>
		<!-- end page -->
	</div>
	<div id="footer">
	  <p class="copyright">&copy;&nbsp;&nbsp;2008 -  <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
	</div>
	</body>
	</html>
<% End Sub

Sub DisplayTitle(strSearch)
	Select Case strSearch
		Case "RecipientID"
			If strSearchFor <> "" Then 
				Response.Write "Displaying Member with a Recipient ID of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the Recipient ID that you are searching for."
			End If

		Case "Surname/Organisation"
			If strSearchFor <> "" Then 
				Response.Write "Displaying All Members with a SURNAME/ORGANISATION of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. <br /> Please specify the SURNAME/ORGANISATION that you are searching for."
			End If

		Case "First Name"
			If strSearchFor <> "" Then 
				Response.Write "Displaying All Members with a FIRST NAME of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. <br /> Please specify the FIRST NAME that you are searching for."
			End If
					 
		Case "Address"
			If strSearchFor <> "" Then 
				Response.Write "Displaying Member with an ADDRESS of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the ADDRESS that you are searching for."
			End If
					 
		Case "Suburb"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Members in the SUBURB of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the SUBURB that you are searching for."
			End If

		Case "Postcode"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Members in the POSTCODE of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the POSTCODE that you are searching for."
			End If

		Case "State"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Members in the STATE of <font color=""#ff0000"">"
				Select Case strSearchFor
					Case 1
						Response.Write "Australian Capital Territory"
					Case 2
						Response.Write "New South Wales"
					Case 3
						Response.Write "Northern Territory"
					Case 4
						Response.Write "Queensland"
					Case 5
						Response.Write "South Australia"
					Case 6
						Response.Write "Tasmania"
					Case 7
						Response.Write "Victoria"
					Case 8
						Response.Write "Western Australia"
				 End Select
				 Response.Write "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the STATE that you are searching for."
			End If
					
		Case "Membership Number"
			If strSearchFor <> "" Then 
				Response.Write "Displaying Member with a MEMBERSHIP NUMBER of <font color=""#ff0000"">" & strSearchFor & "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the MEMBERSHIP NUMBER that you are searching for."
			End If

			 
		Case "VESA Unit"
			Call showVESAUnit(Conn, rs, "VESA_tblUnit", strSearchFor, strVESAUnit)
						
			If strSearchFor <> "" Then
			%>
				<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td>
					<%
					If strActiveUnit = "1" Then
						Response.Write "Displaying All Members in"
					
					ElseIf strInactiveUnit = "1" Then
						Response.Write "Displaying All Inactive Members in"
					End If 
					%>
					</td>
				</tr>				 
				<tr>
					<td>
						<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td>UNIT/DESIGNATION:</td>
							<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
							<%							 
							If strSearchFor = "-1" Or strSearchFor = "0" Or strSearchFor = "" Then
								Response.Write "<td><font color=""#ff0000"">No VESA Unit</font></td>" 
							Else
								Response.Write "<td><font color=""#ff0000"">" & UCase(strVESAUnit) & "</font></td>"
							End If 
							%>				 
						</tr>
						</table>
					</td>
				</tr>
				</table>
			<%  
			 
				Else 
					Response.Write "Error in your Search. Please specify the UNIT/DESIGNATION that you are searching for."
				End If

		Case "SES Region"
			If strSearchFor <> "" Then
				Response.Write "Displaying All Members in the SES REGION of <font color=""#ff0000"">"
				Select Case strSearchFor
					Case 1
						Response.Write "Central"
					Case 2
						Response.Write "South West"
					Case 3
						Response.Write "East"
					Case 4
						Response.Write "North East"
					Case 5
						Response.Write "Mid West"
					Case 6
						Response.Write "North West"
				End Select
						 Response.Write "</font>"
			Else 
				Response.Write "Error in your Search. Please specify the SES Region that you are searching for."
			End If
     			
		Case Else
			Response.Write "Displaying All Members"
	End Select
End Sub 

Sub viewMembers()
	If iRecordCount = 0 Then			
		Response.Write "<!-- start content -->"
		Response.Write "<div id=""content"">"
		Response.Write "<div class=""post"">"
					   
		'- Check page count to prevent bombing when zero results are returned!-----------------
		Response.Write "<!--/* header */-->" & vbCrLf
		Response.Write "<h1 class=""title""><a href=""#"">" & vbCrLf
		Call DisplayTitle(strSearch)	  
		Response.Write "</a></h1>" & vbCrLf
		Response.Write "<p class=""byline""><strong><font color=""#c40000"">No Record Found.</font></strong></p>" & vbCrLf
		Response.Write "<!--/* end of header */-->" & vbCrLf
		Response.Write "<div class=""entry"">" & vbCrLf
		%>
				<div id="sidebar3" class="sidebar4">
					<ul>
						<li> 
							<h2>No VESA member found!</h2>
							<ul>
								<li>There were no active VESA Member found.</li>
							</ul>
						</li>
					</ul>
				</div>
				
				<div style="padding-top:10px;">
					<table border="0" cellpadding="0" cellspacing="0" width="300">
						<tr>
							<td><button class="pure-button" type="button" onClick="searchAgain()">Back to search</button></td>
						</tr>
					</table>
				</div>
			</div>
		</div>
	</div>
		
	<%
	Else
		Response.Write "<!--/* Start Here */-->" &_
						"<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""1140"">" &_
						"<tr>" &_
						"<td>" &_
						"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" &_
						"<tr>" &_
						"<td>"
					
		Response.Write "<!--/* header */-->" 
		Response.Write "<h1 class=""title""><a href=""#"">"
		Call DisplayTitle(strSearch)	  
		Response.Write "</a></h1>" & vbCrLf

		rstSearch.AbsolutePage = iPageCurrent
		Response.Write "<p class=""byline""><strong><font color=""#c40000"">" & iRecordCount & " Records Found.</font></strong></p>" & vbCrLf
		Response.Write "<!--/* end of header */-->" &_
						"</td>" &_
						"</tr>"
	%>

					<tr height="10"><td><img src="images/spacer.gif" width="1" height="10" border="0"></td></tr>

					<tr>
						<td>
							<table border="0" cellpadding="0" cellspacing="0" width="1140">
							<tr>
								<td>
								<% If strActiveUnit = "1" Then %>
									<strong><font color="#ff0000">There are <%= iRecordCount %> active members.</font></strong></td>
								<% ElseIf strInactiveUnit = "1" Then %>
									<strong><font color="#ff0000">There are <%= iRecordCount %> inactive members.</font></strong></td>
								<% End If %>
								<td align="right">
								<span style="color: #ff0000; font-weight:bold">
								<% Call recordsDisplay() %>
								</span>
								</td>
							</tr>
							</table>
						</td>
					</tr>

					<tr><td><img src="images/spacer.gif" width="1" height="5" border="0"></td></tr>

					<!--/* Output Search */-->
					<tr>
						<td>
							<% Call listView() %>
						</td>
					</tr>
					<!--/* Output Search */-->
					</table>
					
					<form name="AddForm" id="AddForm" action="VESAAddNewMember.asp" method="post">
						<input type="hidden" name="search" value="<%=strSearch%>">
						<input type="hidden" name="Suburb" value="<%=Suburb%>">
						<input type="hidden" name="Postcode" value="<%=Postcode%>">
					</form>

					<form name="EditForm" id="EditForm" action="VESAEdit.asp" method="post">
						<input type="hidden" name="RecipientID" value="<%=RecipientID%>">
					</form>

					<form name="ShowHistoryForm" id="ShowHistoryForm" action="VESAAudit.asp" method="post">
						<input type="hidden" name="search" value="<%=strSearch%>">
						<input type="hidden" name="searchFor" value="<%=strSearchFor%>">
					</form>

					<form name="ExportForm" id="ExportForm" action="VESAExport.asp" method="post">
						<input type="hidden" name="search" value="<%=strSearch%>">
						<input type="hidden" name="searchFor" value="<%=strSearchFor%>">
					</form>
				<% End If %>
				</table>
			</td>
		</tr>
		</table>
		<!--/* End Here */-->
<% End Sub 

Sub listView()
%>
	<!--/* Start Table */-->
	<table cellspacing="0" cellpadding="0" border="0" style="width:1140px !important;">
	<tr>
		<td>Page <strong><%=iPageCurrent%></strong> of <strong><%=iPageCount%></strong></td>
		<td>
		<div align="right">					
			<div class="pages" style="width:600px !important;">
				<% Call databasePaging(iPageCurrent, iPageCount, strSearch, strSearchFor) %>
			</div>
		</div>
		</td>
	</tr>

	<tr><td colspan="2"><img src="images/spacer.gif" width="1" height="5" border="0"></td></tr>

	<tr>
		<td colspan="2">
		<% Call listButtons() %>
		</td>
	</tr>

	<tr><td valign="top" colspan="2"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>

	<tr>
		<td colspan="2" bgcolor="#eeeeee">
		<% If strActiveUnit = "1" Then %>
			<form name="MultiDeleteForm" id="MultiDeleteForm" action="VESASave.asp" method="post" onSubmit="return stopSubmit()">
			<input type="hidden" name="VESAID" value="<%=Session("VESAID")%>">
			<input type="hidden" name="ActionType" value="Delete">
		<% ElseIf strInactiveUnit = "1" Then %>
			<form name="ActivateForm" id="ActivateForm" action="VESASave.asp" method="post" onSubmit="return stopSubmit()">
			<input type="hidden" name="VESAID" value="<%=Session("VESAID")%>">
			<input type="hidden" name="ActionType" value="Activate">
		<% End If %>
					
		<table id="main_table" border="0" align="center" cellspacing="2" cellpadding="1" width="1140">
		<tr align="center" height="30">
			<td class="tab_header_cell"><font color="#0000a0"><b>Edit</b></font></td>
			
			<% If strActiveUnit = "1" Then %>
				<td class="tab_header_cell">
				<font color="#0000a0"><b>Delete</b></font><br />
				<input type='checkbox' name='checkall' onclick='checkedAll(MultiDeleteForm);'><br />
				<font size="1">(select / unselect all)</font>
				</td>
			<% ElseIf strInactiveUnit = "1" Then %>
				<td class="tab_header_cell"><b>Activate?</b></td>
			<% End If %>
			
			<td class="tab_header_cell"><b>Membership <br /> Number</b></td>
			<td class="tab_header_cell"><b>Name / Organisation</b></td>
			<td class="tab_header_cell"><b>Email Address</b></td>
			<td class="tab_header_cell"><b>Publication Assigned</b></td>
			<td class="tab_header_cell"><b>Unit / Designation</b></td>
			<td class="tab_header_cell"><b>SES Region</b></td>
			<%			
				Do While Not rstSearch.EOF And rstSearch.AbsolutePage = iPageCurrent
					RecipientID				= rstSearch("RecipientID") & ""
					IDArray					= CInt(rstSearch("RecipientID")) & ""
					Surname_Organization	= rstSearch("Surname_Organization") & ""
					FirstName				= rstSearch("FirstName") & ""
					Address					= rstSearch("Address") & ""
					Suburb					= rstSearch("Suburb") & ""
					Postcode				= rstSearch("Postcode") & ""
					State					= rstSearch("State_Name") & ""
					MembershipNumber		= rstSearch("MembershipNumber") & ""
					EmailAddress			= rstSearch("MemberEmailAddress") & ""
					PhoenixCopies			= rstSearch("PhoenixCopies") & ""
					VESAPocketDiary			= rstSearch("VESAPocketDiary") & ""
					VESAWallCalendar		= rstSearch("VESAWallCalendar") & ""
					VESAUnit				= rstSearch("VESAUnit") & ""
					SESRegion				= rstSearch("SESRegion") & ""
					
					j = j + 1
					Response.write "<tr height=""20"" class=""listTableText" & (j And 1) & """>"
				%>
						<td align="center"><a href="javascript:memberSelected(<%=RecipientID%>)"><img src="images/edit.gif" width="16" height="16" border="0" alt="<%=RecipientID%>" /></a></td>
						
						<% If strActiveUnit = "1" Then %>
							<td align="center"><input type="checkbox" id="DoDelete" name="DoDelete" value="<%=IDArray%>"></td>
						<% ElseIf strInactiveUnit = "1" Then %>
							<td align="center"><input type="checkbox" id="DoActivate" name="DoActivate" value="<%=IDArray%>"></td>
						<% End If %>
						
						<td align="center">
						<%
						If IsNull(rstSearch.Fields("MembershipNumber")) Then
							Response.Write "<font color=""#ff0000"">No Membership <br /> Number provided</font>"
						Else
							Response.Write rstSearch.Fields("MembershipNumber")
						End If			
						%>
						</td>
						
						<td>
						<div style="padding:5px !important;">
							<p><span style="color: #0000a0; font-weight:bold"><%=RecipientID%></span></p>
							<p>
							<%
							If IsNull(rstSearch.Fields("FirstName")) Or rstSearch.Fields("FirstName") = ""  Then
								Response.Write UCase(Surname_Organization)

								Else
									If Len(Surname_Organization) > 15 Then 
										Response.Write FirstName & "<br />" 
										Response.Write UCase(Surname_Organization)
									
									ElseIf InStr(1, FirstName, "C/-") > 0 Then
										Response.Write FirstName & "<br />" 
										Response.Write UCase(Surname_Organization)
									Else
										Response.Write "<strong>" & FirstName & "&nbsp;" & UCase(Surname_Organization) & "</strong>"
									End If 
								End If 
							%>
							</p>
							<p>
							<%=Address%> <br />
							<%=UCase(Suburb)%>&nbsp;<%=State%>&nbsp;<%=Postcode%>
							</p>
						</div>
						</td>

						<td align="center">
						<%
						If IsNull(rstSearch.Fields("MemberEmailAddress")) Then
							Response.Write "<font color=""#ff0000"">No Email Address provided</font>"
						Else
							Response.Write LCase(rstSearch.Fields("MemberEmailAddress"))
						End If			
						%>
						</td>
						
						<td>
						<div style="padding:5px !important; width:150px !important;">
							<div style="padding-bottom:15px !important;">
								<div style="width: 80%; float:left">Phoenix Copies:</div>
								<div style="width: 20%; float:right">
									<% 
									If PhoenixCopies > 1 Then 
										Response.Write "<strong>" & PhoenixCopies & "</strong>"
									Else
										Response.Write PhoenixCopies
									End If 
									%>
								</div>
							</div>

							<div style="padding-bottom:15px !important;">
								<div style="width: 80%; float:left">Pocket Diary:</div>
								<div style="width: 20%; float:right">
								<% 
									If VESAPocketDiary > 1 Then 
										Response.Write "<strong>" & VESAPocketDiary & "</strong>"
									Else
										Response.Write VESAPocketDiary
									End If 
								%>
								</div>
							</div>

							<div style="padding-bottom:15px !important;">
								<div style="width: 80%; float:left">Wall Calendar:</div>
								<div style="width: 20%; float:right">
									<% 
									If VESAWallCalendar > 1 Then 
										Response.Write "<strong>" & VESAWallCalendar & "</strong>"
									Else
										Response.Write VESAWallCalendar
									End If 
									%>
								</div>
							</div>
						</div>
						</td>
						
						<td align="center"><font color="#ff0000"><%=UCase(VESAUnit)%></font></td>
						<%
						If Session("AccessRights") = "Level 1" Then 
							If Not SESRegion = "168" Then 
								Response.Write "<td align=""center"">" & UCase(SESRegion) & "</td>"
								Response.Write "</tr>"
							Else
								Response.Write "<td>&nbsp;</td>"
								Response.Write "</tr>"
							End If 
						Else
								Response.Write "<td>&nbsp;</td>"
								Response.Write "</tr>"
						End If
							
					rstSearch.MoveNext
				Loop
				%>
			</table>
			</form>
		</td>
	</tr>

	<tr><td valign="top" colspan="2"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>

	<tr>
		<td>Page <strong><%=iPageCurrent%></strong> of <strong><%=iPageCount%></strong></td>
		<td>
		<div align="right">					
			<div class="pages" style="width:600px !important;">
				<% Call databasePaging(iPageCurrent, iPageCount, strSearch, strSearchFor) %>
			</div>
		</div>
		</td>
	</tr>

	<tr><td valign="top" colspan="2"><img src="images/spacer.gif" width="1" height="10" alt="" /></td></tr>

	<tr>
		<td colspan="2">
		<% Call listButtons() %>
		</td>
	</tr>
	</table>
	<!--/* End Table */-->
<% 
End Sub

Sub listButtons()
%>
	<% If Session("AccessRights") = "Level 1" Then %>
		<table border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td><button type="button" class="pure-button" onClick="goBack()">Back to Main</button></td>
			<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
				        
			<% If strActiveUnit = "1" Then %>
				<td><button type="button" class="pure-button" onClick="searchAgain()">Search</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
				<td><button type="button" class="pure-button" onClick="addNewMember()">Add</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
				<td><button type="button" class="pure-button" onClick="deleteSelected()">Delete</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
			
			<% ElseIf strInactiveUnit = "1" Then %>
				<td><button type="button" class="pure-button" onClick="activateSelected()">Activate</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
			<% End If %>
			
			<td><button type="button" class="pure-button" onClick="showHistory()">History</button></td>
			
			<% 
			Select Case strSearch
				Case "RecipientID", "Membership Number", "First Name", "Last Name", "Address" 
					Response.Write "</tr>"
				Case ""
					Response.Write "</tr>"
				Case Else
			%>
					<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
					<td><button type="button" class="pure-button" onClick="exportFile()">Export as Excel</button></td>
			<% End Select %>
					
			</tr>
			</table>

		<% Else %>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td><button type="button" class="pure-button" onClick="goBack()">Back to Main</button></td>
				<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
				
				<% If strActiveUnit = "1" Then %>
					<td><button type="button" class="pure-button" onClick="addNewMember()">Add</button></td>
					<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
					<td><button type="button" class="pure-button" onClick="deleteSelected()">Delete</button></td>
					<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>

				<% ElseIf strInactiveUnit = "1" Then %>
					<td><button type="button" class="pure-button" onClick="activateSelected()">Activate</button></td>
					<td><img src="images/spacer.gif" width="10" height="1" alt="" /></td>
				<% End If %>
				
				<td><button type="button" class="pure-button" onClick="exportFile()">Export as Excel</button></td>
			</tr>
			</table>
		<% End If %>
<%
End Sub 

Sub databasePaging(aPageCurrent, aPageCount, aSearch, aSearchFor)
	If aPageCurrent < 10 Then
      StartPage = 1
      EndPage = 10
	Else
		  StartPage = aPageCurrent - 5
		  EndPage = aPageCurrent + 4
		  If EndPage > aPageCount Then
				EndPage = aPageCount
				StartPage = EndPage - 9
		  End If                              
	End if
							
	If aSearch = "All" Then
		If aPageCurrent > 1 Then 
		%>
			<a href="VESAOutput.asp?search=<%=aSearch%>&page=<%=aPageCurrent - 1%>">&lt;&nbsp;Prev</a>
		<%
		Else
			Response.Write "<span class=""disabled"">&lt;&nbsp;Prev</span>"
		End If
		
		For I = StartPage To EndPage
			If I <> aPageCurrent Then
			%>
				<a href="VESAOutput.asp?search=<%=Server.URLEncode(aSearch)%>&page=<%=I%>"><%=I%></a>
			<%
			Else
				'The active page
				Response.Write "<span class=""current"">" & I & "</span>" & vbCrLf
			End If
			'Writes | as a separator if we're not at the last link
			'If I <> aPageCount Then Response.Write(" | ")                          
		Next
			If aPageCurrent < aPageCount Then
			%>
				<a href="VESAOutput.asp?search=<%=Server.URLEncode(aSearch)%>&page=<%=aPageCurrent + 1%>">Next&nbsp;&gt;</a>
			<%
															
			Else
				Response.Write "<span class=""disabled"">Next&nbsp;&gt;</span>"
			End If

	Else
		If aPageCurrent > 1 Then 
		%>
			<a href="VESAOutput.asp?search=<%=Server.URLEncode(aSearch)%>&searchFor=<%= Server.URLEncode(aSearchFor) %>&page=<%=aPageCurrent - 1%>">&lt;&nbsp;Prev</a>
		<%
		Else
			Response.Write "<span class=""disabled"">&lt;&nbsp;Prev</span>"
		End If
		
		For I = 1 To aPageCount
			If I <> aPageCurrent Then
			%>
				<a href="VESAOutput.asp?search=<%=Server.URLEncode(aSearch)%>&searchFor=<%=Server.URLEncode(aSearchFor)%>&page=<%=I%>"><%=I%></a>
			<%
			Else
				'The active page
				Response.Write "<span class=""current"">" & I & "</span>" & vbCrLf
			End If               
		Next
			If aPageCurrent < aPageCount Then
			%>
				<a href="VESAOutput.asp?search=<%=Server.URLEncode(aSearch)%>&searchFor=<%= Server.URLEncode(aSearchFor) %>&page=<%=aPageCurrent + 1%>">Next&nbsp;&gt;</a>
			<%
															
			Else
				Response.Write "<span class=""disabled"">Next&nbsp;&gt;</span>"
			End If
	End If 
End Sub

Sub recordsDisplay()
	If iRecordCount > PAGE_SIZE Then
		
		For k = 1 to iPageCount 
			If k = iPageCurrent Then 
				
				m = 1
				n = k - 1
				
				If k = 1 Then 
					Response.Write "Displaying Records " & m & " - " & (PAGE_SIZE * k) & " of " & iRecordCount & "&nbsp;&bull;&nbsp;"
				
				ElseIf k = iPageCount Then 
					Response.Write "Displaying Records " & (n * PAGE_SIZE + m) & " - " & iRecordCount & " of " & iRecordCount & "&nbsp;&bull;&nbsp;"
				
				Else
					Response.Write "Displaying Records " & (n * PAGE_SIZE + m) & " - " & (PAGE_SIZE * k) & " of " & iRecordCount & "&nbsp;&bull;&nbsp;" 
				End If 

				Response.Write "Page " & k & " of " & iPageCount & ":"
			End If 
		Next
	Else
		If iRecordCount <> 1 Then
			Response.Write "Displaying " & iRecordCount & " Records &nbsp;&bull;&nbsp; Page " & iPageCurrent & " of " & iPageCount & ":"
		Else
			Response.Write "Displaying " & iRecordCount & " Record &nbsp;&bull;&nbsp; Page " & iPageCurrent & " of " & iPageCount & ":"
		End If 
	End If
End Sub
%>


<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/include.asp"-->
<!--#INCLUDE FILE="include/functions.asp"-->
<%
Dim strSearch, strSearchFor

strSearch		= Request.Form("search")
strSearchFor	= Request.Form("searchFor")
strType			= Request.Form("type")

EstablishConnection()

Select Case strSearch 
	'Export Member by Title in an Excel Format
	Case "Surname/Organisation" 
		Call exportToExcel(Conn, rs, "VESA_tblMembers", "" & strSearchFor & "", "ORDER BY RecipientID DESC", "VESA Members Database - Surname/Organization Spreadsheet")

	'Export Member by Suburb in an Excel Format
	Case "Suburb" 
		Call exportToExcel(Conn, rs, "VESA_tblMembers", "" & strSearchFor & "", "ORDER BY RecipientID DESC", "VESA Members Database - Suburb Spreadsheet")

	'Export Member by Postcode in an Excel Format
	Case "Postcode" 
		Call exportToExcel(Conn, rs, "VESA_tblMembers", "" & strSearchFor & "", "AND IsActive='1' ORDER BY RecipientID DESC", "VESA Members Database - Postcode Spreadsheet")

	'Export Member by State in an Excel Format
	Case "State"
		Call exportToExcel(Conn, rs, "VESA_tblMembers", "" & strSearchFor & "", "AND IsActive='1' ORDER BY RecipientID DESC", "VESA Members Database - State Spreadsheet")

	'Export Member by VESA Unit in an Excel Format
	Case "VESA Unit"
		Call exportToExcel(Conn, rs, "VESA_tblMembers", "" & strSearchFor & "", "AND IsActive='1' ORDER BY RecipientID DESC", "VESA Members Database - Unit/Distribution Spreadsheet")

	'Export Member by VESA Unit in an Excel Format
	Case "SES Region"
		Call exportToExcel(Conn, rs, "VESA_tblMembers", "" & strSearchFor & "", "AND IsActive='1' ORDER BY RecipientID DESC", "VESA Members Database - SES Region Spreadsheet")
	    
	'Export VESA Unit in an Excel Format
	Case "All VESA Unit"
		Call exportToExcel(Conn, rs, "VESA_tblUnit", "1", "AND IsActive='1' ORDER BY VESAUnitID", "VESA Members Database - Unit/Distribution Spreadsheet")

	'Export VESA Unit with Password in an Excel Format
	Case "All VESA Unit Password"
		'Call exportToExcel(Conn, rs, "VESA_tblUnit", "1", "AND IsActive='1' ORDER BY VESAUnitID ASC", "VESA Members Database - All VESA Unit with Password Spreadsheet")
		Call exportToExcel(Conn, rs, "VESA_tblUnit", "" & strSearchFor & "", "WHERE IsActive='1' ORDER BY VESAUnitID ASC", "VESA Members Database - All VESA Unit with Password Spreadsheet")
		
	'Export Inactive Members 
	Case "Inactive"
		Call exportToExcel(Conn, rs, "VESA_tblDeletedMembers", "1", "AND IsActive='1' ORDER BY RecipientID DESC", "VESA Members Database - Inactive Members Spreadsheet")

	'Export Audit Members 
	Case "Audit"
		Call exportToExcel(Conn, rs, "VESA_tblAudit", "" & strSearchFor & "", "AND IsActive='1' ORDER BY AuditID", "VESA Members Database - Audit Members Spreadsheet")

	'Export Users 
	Case "Users"
		Call exportToExcel(Conn, rs, "CWM_tblUsers", "" & strSearchFor & "", "AND IsActive='1' ORDER BY UserID", "VESA Members Database - Username and Password Spreadsheet")

	Case Else
		If strType = "XLS" Then
			Call exportToExcel_AllMembers(Conn, rs, "VESA_tblMembers", "WHERE (U.IsActive='1') ORDER BY RecipientID ASC", "VESA Members Database Spreadsheet")
		Else 
			Call exportToCSV_AllMembers(Conn, rs, "VESA_tblMembers", "WHERE (U.IsActive='1') ORDER BY RecipientID ASC")
		End If 
End Select 

CloseConnection()
%>

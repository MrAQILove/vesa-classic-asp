<%
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'		VESAAddNewMember.asp and VESAEdit.asp
'		(Show dropdown list for State field)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub showSelectedValue(c, r, table, dbField1, dbField2, formField)
	
	'-- SQL Statement
	strSQL = "SELECT * FROM " & table
	
	If table <> "MembersDB_tblState" And table <> "VESA_tblSESRegion" And table <> "MembersDB_tblUserAccess" Then
		strSQL = strSQL & " WHERE IsActive = '1'"
	End If 
	
	strSQL = strSQL & " ORDER BY " & dbField1 & " ASC"
      
	'-- Execute our SQL statement and store the recordset
	Set r = c.Execute(strSQL)

	'-- START MAIN CODE BLOCK
	Response.Write "<select id=""" & formField & """ name=""" & formField & """ class=""pure-input-medium"">"
	Response.Write "<option value=""0"">Please Choose</option>"

	'-- loop and build each database entry as a selectable option
	While r.EOF = false
		Response.Write "<option value=""" & r.Fields("" & dbField1 & "").Value & """>" _ 
		& r.Fields("" & dbField2 & "").Value & "</option>" & vbCrLf

		'-- Move recordset to the next value
		r.movenext
	Wend
	'--END OF MAIN CODE BLOCK

	'-- close select/form tags
	Response.Write "</select>" & vbCrlf
End Sub

'-----------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'		admin/DeleteVESAUnit.asp
'		(Show dropdown list for VESA Unit field)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub showActiveVESAUnit(c, r, table, dbField1, dbField2, formField)

   '-- SQL Statement
   strSQL = "SELECT * FROM " & table & " WHERE IsUnitSES = '1' ORDER BY " & dbField1 & " ASC"
      
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   Response.Write "<select id=""" & formField & """ name=""" & formField & """ class=""pure-input-medium"">"
   Response.Write "<option value="""">Please Choose</option>"

   '-- loop and build each database entry as a selectable option
   While r.EOF = false
      Response.Write "<option value=""" & r.Fields("" & dbField1 & "").Value & """>" _ 
	  & r.Fields("" & dbField2 & "").Value & "</option>" & vbCrLf

      '-- Move recordset to the next value
      r.movenext
   Wend
   '--END OF MAIN CODE BLOCK

   '-- close select/form tags
   Response.Write "</select>" & vbCrlf
End Sub

'-----------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'		admin/DeleteVESAUnit.asp
'		(Show dropdown list for VESA Unit field)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub showInactiveVESAUnit(c, r, table, dbField1, dbField2, formField)

   '-- SQL Statement
   strSQL = "SELECT * FROM " & table & " WHERE IsUnitSES = '0' ORDER BY " & dbField1 & " ASC"
      
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   Response.Write "<select id=""" & formField & """ name=""" & formField & """ class=""pure-input-medium"">"
   Response.Write "<option value="""">Please Choose</option>"

   '-- loop and build each database entry as a selectable option
   While r.EOF = false
      Response.Write "<option value=""" & r.Fields("" & dbField1 & "").Value & """>" _ 
	  & r.Fields("" & dbField2 & "").Value & "</option>" & vbCrLf

      '-- Move recordset to the next value
      r.movenext
   Wend
   '--END OF MAIN CODE BLOCK

   '-- close select/form tags
   Response.Write "</select>" & vbCrlf
End Sub

'-----------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'			 VESAAddNewMember.asp and VESAEdit.asp
'			 (Show dropdown list for VESA Unit field with selected Unit)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub selectedVESAUnitList(c, r, table, sqlFieldName, selectName, sessionObj, strCSS)
   '-- SQL Statement
   strSQL = "SELECT * FROM " & table & " WHERE IsActive='1' ORDER BY " & sqlFieldName & " ASC"
   
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   '-- If we have records to return
   Response.Write "<select name=""" & selectName & """ id=""" & selectName & """ class=""" & strCSS & """>" & vbCrLf
%>
   <option <%If sessionObj = "" Then Response.Write "class=""selectedItem"" selected"%>>Please Choose</option>

<%
   While Not r.EOF
      VESAUnitID	= r("VESAUnitID") & ""
	  VESAUnit		= r("VESAUnit") & ""
%>
	  <option value="<%=VESAUnitID%>" <%If sessionObj = "" & VESAUnitID & "" Then Response.Write "class=""selectedItem"" selected"%>><%=VESAUnit%></option>
<%
	  r.MoveNext 
   Wend
   
   Response.Write "</select>" & vbCrLf
   '--END OF MAIN CODE BLOCK
End Sub
'-----------------------------------------------------------------------------------------------

'----- %%%%% ----------------------------------------------------------------------- %%%%% ----- 
'            VESACheckMembers.asp 
'            (Show selected value from position, organisation, etc.)
'----- %%%%% ----------------------------------------------------------------------- %%%%% -----
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub showdbSelectedValue(c, r, table, objValue, objCondition, strSearch)

   '-- SQL Statement
   strSQL = "SELECT DISTINCT " & objValue & " FROM " & table & " WHERE " & objCondition & "='" & strSearch & "'"
   
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   '-- If we have records to return
   If Not r.EOF Then
	  Response.Write r.Fields("" & objValue & "").Value 
   End If
   
   r.Close
   Set r = Nothing

   '--END OF MAIN CODE BLOCK

End Sub
'--------------------------------------------------------------------------------------------------

'----- %%%%% ----------------------------------------------------------------------- %%%%% --------
'			 VESAMain.asp
'            (Display welcome page according to user log in)
'----- %%%%% ----------------------------------------------------------------------- %%%%% --------
Sub displayWelcome(objUser)  
   Select Case objUser
      Case "Webmaster", "Administrator", "Editor", "VESA Shop"
	     Response.Write "<h1 class=""title""><a href=""#"">Welcome to the VESA Administrator's area</a></h1>"
	  
	  Case "Chaplain", "Committee2", "Committee3", "President", "Vice President", "Secretary", "Treasurer"   
	     Response.Write "<h1 class=""title""><a href=""#"">Welcome to the " & objUser & "VESA Administrator's area.</a></h1>" & vbCrLf

	  Case "Unit"
		 Dim strVESAUnit
		 Call showVESAUnit(Conn, rs, "VESA_tblUnit", Session("VESAUnitID"), strVESAUnit)
	     
		 Response.Write "<h1 class=""title""><a href=""#"">Welcome to the " & UCase(strVESAUnit) & " VESA Unit</a></h1>" & vbCrLf
    End Select
	
	Response.Write "<p class=""byline"">Choose a menu from the left navigation to get started.</p>"
End Sub

Sub VESAUnitMainMenu()
%>
	<div id="sidebar1" class="sidebar">
		<ul>
			<li> 
				<h2>
				<% 
				Dim strVESAUnit
				Call showVESAUnit(Conn, rs, "VESA_tblUnit", Session("VESAUnitID"), strVESAUnit)
				  
				Response.Write strVESAUnit & " Members" 
				%> 
				</h2>
				<ul>
					<% Call displaySelectedMenu(Request.ServerVariables("SCRIPT_NAME")) %>
					<li><a href="VESAUnitLogin.asp">Log Out</a></li>
				</ul>
			</li>
		</ul>
	</div>
<% 
End Sub

Sub displaySelectedMenu(strCode)
   Select Case strCode
      Case "/database/vesa/Members/VESAMain.asp" %>
		 <li class="hover"><a href="#">Home</a></li>
		 <li><a href="#" onClick="viewMembers()">View All Members</a></li>
         <li><a href="#" onClick="addMember()">Add a New Member</a></li>
         <li><a href="#" onClick="deleteMember()">Delete a Member</a></li>

	  <% Case "/database/vesa/Members/VESAViewAllMembers.asp" %>
		 <li><a href="VESAMain.asp">Home</a></li>
		 <li class="hover"><a href="#">View All Members</a></li>
         <li><a href="#" onClick="addMember()">Add a New Member</a></li>
         <li><a href="#" onClick="deleteMember()">Delete a Member</a></li>

      <% Case "/database/vesa/Members/VESAAddNewMember.asp" %>
		 <li><a href="VESAMain.asp">Home</a></li>
		 <li><a href="#" onClick="viewMembers()">View All Members</a></li>
         <li class="hover"><a href="#">Add a New Member</a></li>
         <li><a href="#" onClick="deleteMember()">Delete a Member</a></li>
		 
	  <% Case "/database/vesa/Members/VESADeleteMember.asp" %>
		 <li><a href="VESAMain.asp">Home</a></li>
		 <li><a href="#" onClick="viewMembers()">View All Members</a></li>
         <li><a href="#" onClick="addMember()">Add a New Member</a></li>
         <li class="hover"><a href="#">Delete a Member</a></li>

	  <% Case "/database/vesa/Members/VESAEdit.asp" %>
		 <li><a href="VESAMain.asp">Home</a></li>
		 <li><a href="#" onClick="viewMembers()">View All Members</a></li>
         <li><a href="#" onClick="addMember()">Add a New Member</a></li>
		 <li class="hover"><a href="#">Edit a New Member</a></li>
         <li><a href="#" onClick="deleteMember()">Delete a Member</a></li>

	  <% Case "/database/vesa/Members/VESASave.asp" %>
		 <li><a href="VESAMain.asp">Home</a></li>
		 <li><a href="#" onClick="viewMembers()">View All Members</a></li>
         <li><a href="#" onClick="addMember()">Add a New Member</a></li>
         <li><a href="#" onClick="deleteMember()">Delete a Member</a></li>

	  <% Case "/database/vesa/Members/VESACheckMembers.asp" %>
		 <li><a href="VESAMain.asp">Home</a></li>
		 <li><a href="#" onClick="viewMembers()">View All Members</a></li>
         <li><a href="#" onClick="addMember()">Add a New Member</a></li>
         <li><a href="#" onClick="deleteMember()">Delete a Member</a></li>
<%
   End Select 
End Sub

Sub displayFORMLinks()
%>
	<form name="viewMembersForm" id="viewMembersForm" action="VESAOutput.asp" method="post">
		<input type="hidden" name="search" value="10">
		<input type="hidden" name="searchForVESAUnit" value="<%=Session("VESAUnitID")%>">
	</form>

	<form name="addMemberForm" id="addMemberForm" action="VESAAddNewMember.asp" method="post">
		<input type="hidden" name="VESAUnitID" value="<%=Session("VESAUnitID")%>">
	</form>

	<form name="deleteMemberForm" id="deleteMemberForm" action="VESADeleteMember.asp" method="post">
		<input type="hidden" name="VESAUnitID" value="<%=Session("VESAUnitID")%>">
	</form>
<%
End Sub

'--------------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% -----------
'			 VESADeleteMember.asp
'            (Show the Recipient Name)
'----- %%%%% -------------------------------------------------------------------- %%%%% -----------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub showDeleteMember(c, r, table)

   '-- SQL Statement
   strSQL = "SELECT * FROM " & table & " M"
   strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
   strSQL = strSQL & " WHERE IsActive = '1' ORDER BY RecipientID ASC"
   
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   '-- If we have records to return
   
   Response.Write "<select id=""DoDelete"" name=""DoDelete"" class=""pure-input-medium"">"
   Response.Write "<option>Please Choose</option>"  & vbCrLf

   '-- loop and build each database entry as a selectable option
   While r.EOF = false
      Dim Recipient_ID, Surname_Organization, VESAUnitID

	  Recipient_ID			= r("RecipientID") & ""
	  Surname_Organization	= r("Surname_Organization") & ""
	  FirstName				= r("FirstName") & ""
	  VESAUnitID			= r("VESAUnitID") & ""
	  
	  Select Case VESAUnitID
		Case 1, 2, 3, 5, 6, 7, 8, 13, 14, 15, 16
			If IsNull(r.Fields(2)) Then
				Response.Write "<option value=""" & Recipient_ID & """>"
				Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & Surname_Organization 
				Response.Write "</option>" & vbCrLf
			Else
				Response.Write "<option value=""" & Recipient_ID & """>"
				Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & FirstName & "&nbsp;-&nbsp;" & Surname_Organization
				Response.Write "</option>" & vbCrLf
			End If 
		
		Case 4, 9, 10, 11, 12
			Response.Write "<option value=""" & Recipient_ID & """>"
			Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & FirstName & "&nbsp;" & Surname_Organization 
			Response.Write "</option>" & vbCrLf

		Case Else
			Response.Write "<option value=""" & Recipient_ID & """>"
			Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & FirstName & "&nbsp;" & Surname_Organization 
			Response.Write "</option>" & vbCrLf
		End select

      '-- Move recordset to the next value
      r.movenext
   Wend
   '--END OF MAIN CODE BLOCK

   '-- close select/form tags
   Response.Write "</select>" & vbCrLf

End Sub
'-----------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'			VESADeleteMember.asp
'           (Branch section - Show the Addressee)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub showDeleteMemberByVESAUnit(c, r, table, strVESAUnitID)

   '-- SQL Statement
   strSQL = "SELECT * FROM " & table & " M"
   strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
   strSQL = strSQL & " AND M.VESAUnitID='" & strVESAUnitID & "'"
   strSQL = strSQL & " ORDER BY RecipientID ASC"
   
   '-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)

   '-- START MAIN CODE BLOCK
   '-- If we have records to return
   
   Response.Write "<select id=""DoDelete"" name=""DoDelete"" class=""inputSelection"">"
   Response.Write "<option>Please Choose</option>"  & vbCrLf

   '-- loop and build each database entry as a selectable option
   While r.EOF = false
      Dim Recipient_ID, Surname_Organization, VESAUnitID

	  Recipient_ID			= r("RecipientID") & ""
	  Surname_Organization	= r("Surname_Organization") & ""
	  FirstName				= r("FirstName") & ""
	  VESAUnitID			= r("VESAUnitID") & ""
	  
	  Select Case VESAUnitID
		Case 1, 2, 3, 5, 6, 7, 8, 13, 14, 15, 16
			If IsNull(r.Fields(2)) Then
				Response.Write "<option value=""" & Recipient_ID & """>"
				Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & Surname_Organization 
				Response.Write "</option>" & vbCrLf
			Else
				Response.Write "<option value=""" & Recipient_ID & """>"
				Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & FirstName & "&nbsp;-&nbsp;" & Surname_Organization
				Response.Write "</option>" & vbCrLf
			End If 
		
		Case 4, 9, 10, 11, 12
			Response.Write "<option value=""" & Recipient_ID & """>"
			Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & FirstName & "&nbsp;" & Surname_Organization 
			Response.Write "</option>" & vbCrLf

		Case Else
			Response.Write "<option value=""" & Recipient_ID & """>"
			Response.Write "(" & Recipient_ID & ")" & "&nbsp;" & FirstName & "&nbsp;" & Surname_Organization 
			Response.Write "</option>" & vbCrLf
		End select

      '-- Move recordset to the next value
      r.movenext
   Wend
   '--END OF MAIN CODE BLOCK

   '-- close select/form tags
   Response.Write "</select>" & vbCrLf

End Sub
'-----------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'			 VESAMain.asp and VESAOutput.asp
'            (Show the VESA Unit Name)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r and table are passed when the sub is built.
Sub showVESAUnit(c, r, table, strVESAUnitID, str)
      
	  '-- SQL Statement
      strSQL = "SELECT * FROM " & table
	  strSQL = strSQL & " WHERE VESAUnitID ='" & strVESAUnitID & "'"
	  strSQL = strSQL & " ORDER BY VESAUnitID"
      
	  '-- Execute our SQL statement and store the recordset
      Set r = c.Execute(strSQL)

      '-- START MAIN CODE BLOCK
      
	  If Not r.EOF Then
	     str = r.Fields("VESAUnit").Value
	  End If 
	  
	  '--END OF MAIN CODE BLOCK
End Sub
'-----------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'			VESAExport.asp
'            (Export the searched list as a Excel file)
'----- %%%%% -------------------------------------------------------------------- %%%%% --------
'-- Sub Procedure that builds the dropdown list.
'-- Parameters c, r table and condition are passed when the sub is built.
Sub exportToExcel_AllMembers(c, r, table, condition, u_title)
	'Look for All Members
	strSQL = "SELECT * FROM " & table & " M"
	strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
	strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
	strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
	strSQL = strSQL & " " & condition

	'-- Execute our SQL statement and store the recordset
    Set r = c.Execute(strSQL)
	  
    '-- START MAIN CODE BLOCK
    Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=VESA_Members_List.xls"

	Response.Write "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
	Response.Write "<head>"
    Response.Write "<!--[if gte mso 9]><xml>"
    Response.Write "<x:ExcelWorkbook>"
    Response.Write "<x:ExcelWorksheets>"
    Response.Write "<x:ExcelWorksheet>"
    Response.Write "<x:Name>"& u_title &"</x:Name>"
    Response.Write "<x:WorksheetOptions>"
    Response.Write "<x:Print>"
    Response.Write "<x:ValidPrinterInfo/>"
    Response.Write "</x:Print>"
    Response.Write "</x:WorksheetOptions>"
    Response.Write "</x:ExcelWorksheet>"
    Response.Write "</x:ExcelWorksheets>"
    Response.Write "</x:ExcelWorkbook>"
    Response.Write "</xml>"
    Response.Write "<![endif]--> "
    Response.Write "</head>"
    Response.Write "<body>"

	Response.Write "<table border=""1"">" 
	Response.Write "<tr>"
	
	'Loop through Fields Names and print out the Field Names
	j = 2 'row counter
      
	Response.Write "<th><b>Recipient ID</b></th>"
	Response.Write "<th><b>Surname / Organisation</b></th>"
	Response.Write "<th><b>First Name</b></th>"
	Response.Write "<th><b>Address</b></th>"
	Response.Write "<th><b>Suburb</b></th>"
	Response.Write "<th><b>Postcode</b></th>"
	Response.Write "<th><b>State</b></th>"
	Response.Write "<th><b>Membership Number</b></th>"
	Response.Write "<th><b>Email Address</b></th>"
	Response.Write "<th><b>Phoenix Copies</b></th>"
	Response.Write "<th><b>Pocket Diary</b></th>"
	Response.Write "<th><b>Wall Calender</b></th>"
	Response.Write "<th><b>Unit / Designation</b></th>"
	Response.Write "<th><b>SES Region</b></th>"
	Response.Write "</tr>"
	 		
	'Loop through rows, displaying each field
	Do While Not r.EOF
		RecipientID				= r("RecipientID") & ""
		Surname_Organization	= r("Surname_Organization") & ""
		FirstName				= r("FirstName") & ""
		Address					= r("Address") & ""
		Suburb					= r("Suburb") & ""
		Postcode				= r("Postcode") & ""
		State					= r("State_Name") & ""
		MembershipNumber		= r("MembershipNumber") & ""
		MemberEmailAddress		= r("MemberEmailAddress") & ""
		PhoenixCopies			= r("PhoenixCopies") & ""
		VESAPocketDiary			= r("VESAPocketDiary") & ""
		VESAWallCalendar		= r("VESAWallCalendar") & ""
		VESAUnit				= r("VESAUnit") & ""
		SESRegion				= r("SESRegion") & ""

		Response.Write "<tr>" 
		Response.Write "<td align=""center"">" & RecipientID & "</td>"
		Response.Write "<td align=""center"">" & Surname_Organization & "</td>"
		Response.Write "<td align=""center"">"
		If Not IsNull(r.Fields(2)) Then
			Response.Write FirstName
		Else
			Response.Write "&nbsp;"
		End If 
			Response.Write "</td>"

		Response.Write "<td align=""center"">" & Address & "</td>"
		Response.Write "<td align=""center"">" & UCase(Suburb) & "</td>"
		Response.Write "<td align=""center"">" & Postcode & "</td>"
		Response.Write "<td align=""center"">" & State & "</td>"
		Response.Write "<td align=""center"">" & MembershipNumber & "</td>"
		Response.Write "<td align=""center"">"
		If Not IsNull(r.Fields(8)) Then
			Response.Write MemberEmailAddress
		Else
			Response.Write "&nbsp;"
		End If 
		Response.Write "</td>"
		Response.Write "<td align=""center"">" & PhoenixCopies & "</td>"
		Response.Write "<td align=""center"">" & VESAPocketDiary & "</td>"
		Response.Write "<td align=""center"">" & VESAWallCalendar & "</td>"
		Response.Write "<td align=""center"">" & UCase(VESAUnit) & "</td>"
		Response.Write "<td align=""center"">" & SESRegion & "</td>"
		 
		r.MoveNext
		j = j + 1
	Loop
	
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</body>"
    Response.Write "</html>"
	'--END OF MAIN CODE BLOCK

	'Make sure to close the Result Set and the Connection object
	r.Close
    '-- close select/form tags
End Sub


Sub exportToCSV_AllMembers(c, r, table, condition)
	'Look for All Members
	strSQL = "SELECT M.RecipientID, M.MembershipNumber, M.Surname_Organization, M.FirstName, M.Address, M.Suburb, M.Postcode, S.State_Name," 
    strSQL = strSQL & " M.MemberEmailAddress, M.PhoenixCopies, M.VESAPocketDiary, M.VESAWallCalendar, U.VESAUnit, R.SESRegion"
	strSQL = strSQL & " FROM " & table & " M INNER JOIN"
    strSQL = strSQL & " MembersDB_tblState S ON M.StateID = S.StateID INNER JOIN"
    strSQL = strSQL & " VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID LEFT OUTER JOIN"
    strSQL = strSQL & " VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
	strSQL = strSQL & " " & condition

	'-- Execute our SQL statement and store the recordset
    Set r = c.Execute(strSQL)
	  
    '-- START MAIN CODE BLOCK
    Response.ContentType = "text/csv"
	Response.AddHeader "Content-Disposition", "attachment;filename=VESA_Members_List.csv"

	Call Write_CSV_From_Recordset (r)
End Sub 

Sub Write_CSV_From_Recordset(RS)
	'
    ' Export Recordset to CSV
    ' http://911-need-code-help.blogspot.com/2009/07/export-recordset-data-to-csv-using.html
    '
    ' This sub-routine Response.Writes the content of an ADODB.RECORDSET in CSV format
    ' The function closely follows the recommendations described in RFC 4180:
    ' Common Format and MIME Type for Comma-Separated Values (CSV) Files
    ' http://tools.ietf.org/html/rfc4180
    '
    ' @RS: A reference to an open ADODB.RECORDSET object
    '

    If RS.EOF Then 
        '
        ' There is no data to be written
        '
        Exit Sub  
    End If 

    Dim RX
    Set RX = New RegExp
        RX.Pattern = "\r|\n|,|"""

    Dim i 
    Dim Field
    Dim Separator

    '
    ' Writing the header row (header row contains field names)
    '
    Separator = ""
    For i = 0 To RS.Fields.Count - 1
        Field = RS.Fields(i).Name
        If RX.Test(Field) Then 
            '
            ' According to recommendations:
            ' - Fields that contain CR/LF, Comma or Double-quote should be enclosed in double-quotes
            ' - Double-quote itself must be escaped by preceeding with another double-quote
            '
            Field = """" & Replace(Field, """", """""") & """"
        End If 
        Response.Write Separator & Field
        Separator = ","
    Next 
    Response.Write vbNewLine

    '
    ' Writing the data rows
    '
    Do Until RS.EOF
        Separator = ""
        For i = 0 To RS.Fields.Count - 1
            '
            ' Note the concatenation with empty string below
            ' This assures that NULL values are converted to empty string
            '
            Field = RS.Fields(i).Value & ""
            If RX.Test(Field) Then 
                Field = """" & Replace(Field, """", """""") & """"
            End If 
            Response.Write Separator & Field
            Separator = ","
        Next 
        Response.Write vbNewLine
        RS.MoveNext
    Loop 

End Sub 

'
' EXAMPLE USAGE
'
' - Open a RECORDSET object (forward-only, read-only recommended)
' - Send appropriate response headers
' - Call the function
'
'Dim RS1
'Set RS1 = Server.CreateObject("ADODB.RECORDSET")
'    RS1.Open "SELECT * FROM TABLE_NAME_HERE", "CONNECTION_STRING_HERE", 0, 1

'Response.ContentType = "text/csv"
'Response.AddHeader "Content-Disposition", "attachment;filename=export.csv"
'Write_CSV_From_Recordset RS1
'

Sub exportToExcel(c, r, table, strSearchFor, condition, u_title)      
	  '-- SQL Statement
	  Select Case strSearch
	     'Look for Members by Surname/Organisation
		 Case "Surname/Organisation"
		    strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
            strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE Surname_Organization='" & strSearchFor & "'" & condition

		 'Look for Members by Suburb
		 Case "Suburb"
		    strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
            strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE Suburb='" & strSearchFor & "'" & condition
	     
		 'Look for Members by Postcode
		 Case "Postcode"
		    strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
            strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE Postcode='" & strSearchFor & "'" & condition

		'Look for Members by State
		 Case "State"
		    strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
            strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE S.StateID='" & strSearchFor & "'" & condition
	     
		 'Look for Members by VESA Unit
		 Case "VESA Unit"
			strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
			strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE M.VESAUnitID='" & strSearchFor & "'" & condition

		'Look for Members by SES Region
		 Case "SES Region"
			strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
			strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE R.SESRegionID='" & strSearchFor & "'" & condition
			
		 'Look for All VESA Unit
		 Case "All VESA Unit"
			strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON M.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE IsUnitSES='" & strSearchFor & "'" & condition

		Case "All VESA Unit Password"
			strSQL = "SELECT * FROM " & table & " M"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON M.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " " & condition
	
	     'Look for Members that are Inactive
		 Case "Inactive" 
		    strSQL = "SELECT * FROM " & table & " M" 
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
			strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE IsDeleteMember='" & strSearchFor & "'" & condition

	     'Look for Members that are Audited
		 Case "Audit" 
		    strSQL = "SELECT * FROM " & table & " M" 
			strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
			strSQL = strSQL & " INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
			strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
			strSQL = strSQL & " WHERE M.IsAuditMember='" & strSearchFor & "'" & condition

	     'Look for Members that are Inactive
		 Case "Users" 
		    strSQL = "SELECT * FROM " & table & " WHERE UserActive='" & strSearchFor & "'" & condition
	  End Select

	  '-- Execute our SQL statement and store the recordset
      Set r = c.Execute(strSQL)
	  
      '-- START MAIN CODE BLOCK
      Response.ContentType = "application/vnd.ms-excel"
		  
		  Select Case strSearch
			Case "Surname/Organisation"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchBySurname_Organisation_List.xls"
			
			Case "Suburb"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchBySuburb_List.xls"
			
			Case "Postcode"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchByPostcode_List.xls"
			
			Case "State"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchByState_List.xls"
			
			Case "VESA Unit"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchByVESAUnit_List.xls"
			
			Case "SES Region"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchBySESRegion_List.xls"
			
			Case "All VESA Unit"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchAllVESAUnit_List.xls"
			
			Case "All VESA Unit Password"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_UnitWithPassword.xls"
			
			Case "Inactive"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchByInactiveMembers_List.xls"
			
			Case "Audit"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchByAudit_List.xls"
			
			Case "Users"
				Response.AddHeader "Content-Disposition", "attachment; filename=VESA_SearchByAdministrationUsers_List.xls"
		  End Select 

	  Response.Write "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
      Response.Write "<head>"
      Response.Write "<!--[if gte mso 9]><xml>"
      Response.Write "<x:ExcelWorkbook>"
      Response.Write "<x:ExcelWorksheets>"
      Response.Write "<x:ExcelWorksheet>"
      Response.Write "<x:Name>"& u_title &"</x:Name>"
      Response.Write "<x:WorksheetOptions>"
      Response.Write "<x:Print>"
      Response.Write "<x:ValidPrinterInfo/>"
      Response.Write "</x:Print>"
      Response.Write "</x:WorksheetOptions>"
      Response.Write "</x:ExcelWorksheet>"
      Response.Write "</x:ExcelWorksheets>"
      Response.Write "</x:ExcelWorkbook>"
      Response.Write "</xml>"
      Response.Write "<![endif]--> "
      Response.Write "</head>"
      Response.Write "<body>"

	  Response.Write "<table border=""1"">" 
	  Response.Write "<tr>"

	  Select Case strSearch
	     Case "Users"
	        'Loop through Fields Names and print out the Field Names
            j = 2 'row counter
      
	        Response.Write "<th><b>User ID</b></th>"
	        Response.Write "<th><b>Name</b></th>"
	        Response.Write "<th><b>Email Address</b></th>"
	        Response.Write "<th><b>Admin Type</b></th>"
	        Response.Write "<th><b>Username</b></th>"
	        Response.Write "<th><b>Password</b></th>"
		    Response.Write "<th><b>Registration Date</b></th>"
	        Response.Write "</tr>"
	  
	        'Loop through rows, displaying each field
	        Do While Not r.EOF
		       UserID				= r("UserID") & ""
		       Name					= r("FirstName") & "" & "&nbsp;" & r("Surname") & ""
			   Email				= r("Email") & ""
			   AdminType			= r("AdminType") & "" 
		       Username				= r("UserName") & ""
		       Password				= r("Password") & ""
		       RegistrationDate		= r("RegistrationDate") & ""

		       Response.Write "<tr>"
		       Response.Write "<td valign=""top"">" & UserID & "</td>"
		       Response.Write "<td valign=""top"">" & Name & "</td>"
		       Response.Write "<td valign=""top"">" & Email & "</td>"
		       Response.Write "<td valign=""top"">" & AdminType & "</td>"
		       Response.Write "<td valign=""top"">" & Username & "</td>"
		       Response.Write "<td valign=""top"">" & Password & "</td>"
		       Response.Write "<td valign=""top"">" & RegistrationDate & "</td>"
		 
		       r.MoveNext
               j = j + 1
            Loop
	     
		 '-----------------------------------------------------------------------------------------
		 'Build Excel Spreadsheet for Inactive Members
		 '-----------------------------------------------------------------------------------------
		 Case "Inactive"
	        'Loop through Fields Names and print out the Field Names
            j = 2 'row counter
      
	        Response.Write "<th><b>Recipient ID</b></th>"
			Response.Write "<th><b>Surname / Organisation</b></th>"
			Response.Write "<th><b>First Name</b></th>"
			Response.Write "<th><b>Address</b></th>"
			Response.Write "<th><b>Suburb</b></th>"
			Response.Write "<th><b>Postcode</b></th>"
			Response.Write "<th><b>State</b></th>"
			Response.Write "<th><b>Membership Number</b></th>"
			Response.Write "<th><b>Email Address</b></th>"
			Response.Write "<th><b>Phoenix Copies</b></th>"
			Response.Write "<th><b>Pocket Diary</b></th>"
			Response.Write "<th><b>Wall Calender</b></th>"
			Response.Write "<th><b>Unit / Designation</b></th>"
			Response.Write "<th><b>SES Region</b></th>"
			Response.Write "<th><b>Reason for Deletion</b></th>"
			Response.Write "<th><b>Specify Reason</b></th>"
			Response.Write "<th><b>Date Deleted</b></th>"
			Response.Write "</tr>"
	  
	        'Loop through rows, displaying each field
	        Do While Not r.EOF
		       RecipientID				= r("RecipientID") & ""
			   Surname_Organization		= r("Surname_Organization") & ""
			   FirstName				= r("FirstName") & ""
			   Address					= r("Address") & ""
			   Suburb					= r("Suburb") & ""
			   Postcode					= r("Postcode") & ""
			   State					= r("State_Name") & ""
			   MembershipNumber			= r("MembershipNumber") & ""
			   MemberEmailAddress		= r("MemberEmailAddress") & ""
			   PhoenixCopies			= r("PhoenixCopies") & ""
			   VESAPocketDiary			= r("VESAPocketDiary") & ""
			   VESAWallCalendar			= r("VESAWallCalendar") & ""
			   VESAUnit					= r("VESAUnit") & ""
			   SESRegion				= r("SESRegion") & ""
			   WhyDelete				= r("WhyDelete") & ""
			   SpecifyReason			= r("SpecifyReason") & ""
			   DateDeleted				= r("DateDeleted") & ""

		       Response.Write "<tr>" 
			   Response.Write "<td align=""center"">" & RecipientID & "</td>"
			   Response.Write "<td align=""center"">" & Surname_Organization & "</td>"
			   
			   Response.Write "<td align=""center"">"
			   If Not IsNull(r.Fields(3)) Then
			      Response.Write FirstName
			   Else
			      Response.Write "&nbsp;"
		       End If 
			   Response.Write "</td>"

			   Response.Write "<td align=""center"">" & Address & "</td>"
			   Response.Write "<td align=""center"">" & UCase(Suburb) & "</td>"
			   Response.Write "<td align=""center"">" & Postcode & "</td>"
			   Response.Write "<td align=""center"">" & State & "</td>"
			   Response.Write "<td align=""center"">" & MembershipNumber & "</td>"
			   Response.Write "<td align=""center"">"
			   If Not IsNull(r.Fields(9)) Then
			      Response.Write MemberEmailAddress
			   Else
			      Response.Write "&nbsp;"
			   End If 
			   Response.Write "</td>"
			   Response.Write "<td align=""center"">" & PhoenixCopies & "</td>"
			   Response.Write "<td align=""center"">" & VESAPocketDiary & "</td>"
			   Response.Write "<td align=""center"">" & VESAWallCalendar & "</td>"
			   Response.Write "<td align=""center"">" & VESAUnit & "</td>"
			   Response.Write "<td align=""center"">" & SESRegion & "</td>"
	           
			   Response.Write "<td align=""center"">"
	           If Not IsNull(r.Fields(16)) Then
			      Response.Write WhyDelete
			   Else
			      Response.Write "&nbsp;"
			   End If 
			   Response.Write "</td>"

			   Response.Write "<td align=""center"">"
	           If Not IsNull(r.Fields(17)) Then
			      Response.Write SpecifyReason
			   Else
			      Response.Write "&nbsp;"
			   End If 
			   Response.Write "</td>"

			   Response.Write "<td align=""center"">" & DateDeleted & "</td>"
		 
		       r.MoveNext
               j = j + 1
            Loop

		 '-----------------------------------------------------------------------------------------
		 'Build Excel Spreadsheet for Audit Members
		 '-----------------------------------------------------------------------------------------
		 Case "Audit"
	        'Loop through Fields Names and print out the Field Names
            j = 2 'row counter
      
	        Response.Write "<th><b>Change</b></th>"
			Response.Write "<th><b>Change By</b></th>"
			Response.Write "<th><b>Change Date</b></th>"
			Response.Write "<th><b>Recipient ID</b></th>"
			Response.Write "<th><b>Surname / Organisation</b></th>"
			Response.Write "<th><b>First Name</b></th>"
			Response.Write "<th><b>Address</b></th>"
			Response.Write "<th><b>Suburb</b></th>"
			Response.Write "<th><b>Postcode</b></th>"
			Response.Write "<th><b>State</b></th>"
			Response.Write "<th><b>Membership Number</b></th>"
			Response.Write "<th><b>Email Address</b></th>"
			Response.Write "<th><b>Phoenix Copies</b></th>"
			Response.Write "<th><b>Pocket Diary</b></th>"
			Response.Write "<th><b>Wall Calender</b></th>"
			Response.Write "<th><b>Unit / Designation</b></th>"
			Response.Write "<th><b>SES Region</b></th>"
			Response.Write "</tr>"
	  
	        'Loop through rows, displaying each field
	        Do While Not r.EOF
		       Change					= r("ActionType") & ""
			   ChangeBy					= r("ChangedBy") & ""
			   ChangeDate				= r("ActionDateTime") & ""
			   RecipientID				= r("RecipientID") & ""
			   Surname_Organization		= r("Surname_Organization") & ""
			   FirstName				= r("FirstName") & ""
			   Address					= r("Address") & ""
			   Suburb					= r("Suburb") & ""
			   Postcode					= r("Postcode") & ""
			   State					= r("State_Name") & ""
			   MembershipNumber			= r("MembershipNumber") & ""
			   MemberEmailAddress		= r("MemberEmailAddress") & ""
			   PhoenixCopies			= r("PhoenixCopies") & ""
			   VESAPocketDiary			= r("VESAPocketDiary") & ""
			   VESAWallCalendar			= r("VESAWallCalendar") & ""
			   VESAUnit					= r("VESAUnit") & ""
			   SESRegion				= r("SESRegion") & ""

		       Response.Write "<tr>" 
			   Response.Write "<td align=""center"">" & Change & "</td>"
			   Response.Write "<td align=""center"">" & ChangeBy & "</td>"
			   Response.Write "<td align=""center"">" & ChangeDate & "</td>"
			   Response.Write "<td align=""center"">" & RecipientID & "</td>"
			   Response.Write "<td align=""center"">" & Surname_Organization & "</td>"
			   Response.Write "<td align=""center"">"
			   If Not IsNull(r.Fields(6)) Then
			      Response.Write FirstName
			   Else
			      Response.Write "&nbsp;"
			   End If 
			   Response.Write "</td>"
			   Response.Write "<td align=""center"">" & Address & "</td>"
			   Response.Write "<td align=""center"">" & UCase(Suburb) & "</td>"
			   Response.Write "<td align=""center"">" & Postcode & "</td>"
			   Response.Write "<td align=""center"">" & State & "</td>"
			   Response.Write "<td align=""center"">" & MembershipNumber & "</td>"
			   Response.Write "<td align=""center"">"
			   If Not IsNull(r.Fields(12)) Then
			      Response.Write MemberEmailAddress
			   Else
			      Response.Write "&nbsp;"
			   End If 
			   Response.Write "</td>"
			   Response.Write "<td align=""center"">" & PhoenixCopies & "</td>"
			   Response.Write "<td align=""center"">" & VESAPocketDiary & "</td>"
			   Response.Write "<td align=""center"">" & VESAWallCalendar & "</td>"
			   Response.Write "<td align=""center"">" & VESAUnit & "</td>"
			   Response.Write "<td align=""center"">" & SESRegion & "</td>"
		       
			   r.MoveNext
               j = j + 1
            Loop

		 Case "All VESA Unit"
			'Loop through Fields Names and print out the Field Names
			j = 2 'row counter
      
			Response.Write "<th><b>VESA Unit ID</b></th>"
			Response.Write "<th><b>VESA Unit</b></th>"
			Response.Write "<th><b>Password</b></th>"
			Response.Write "<th><b>Email Address</b></th>"
			Response.Write "<th><b>SES Region</b></th>"
			Response.Write "</tr>"
	  
			'Loop through rows, displaying each field
			Do While Not r.EOF
				VESAUnitID		= r("VESAUnitID") & ""
				VESAUnit		= r("VESAUnit") & ""
				Password		= r("Password") & ""
				EmailAddress	= r("EmailAddress") & ""
				SESRegion		= r("SESRegion") & ""

				Response.Write "<tr>" 
				Response.Write "<td align=""center"">" & VESAUnitID & "</td>"
				Response.Write "<td align=""center"">" & VESAUnit & "</td>"
				Response.Write "<td align=""center"">" & Password & "</td>"
				Response.Write "<td align=""center"">"
				If EmailAddress <> "" Then
					Response.Write EmailAddress
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"
				Response.Write "<td align=""center"">"
				If SESRegion <> "" Then
					Response.Write SESRegion
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"
		 
				r.MoveNext
				j = j + 1
			Loop

		Case "All VESA Unit Password"
			'Loop through Fields Names and print out the Field Names
			j = 2 'row counter
      
			Response.Write "<th><b>VESA Unit ID</b></th>"
			Response.Write "<th><b>VESA Unit</b></th>"
			Response.Write "<th><b>Password</b></th>"
			Response.Write "<th><b>No. Of Members</b></th>"
			Response.Write "<th><b>Email Address</b></th>"
			Response.Write "<th><b>Is Financial</b></th>"
			Response.Write "<th><b>Date Financial Change</b></th>"
			Response.Write "<th><b>SES Region ID</b></th>"
			Response.Write "<th><b>SES Region</b></th>"
			Response.Write "<th><b>Is SES Unit?</b></th>"
			Response.Write "</tr>"
	  
			'Loop through rows, displaying each field
			Do While Not r.EOF
				VESAUnitID			= r("VESAUnitID") & ""
				VESAUnit			= r("VESAUnit") & ""
				Password			= r("Password") & ""
				EmailAddress		= r("EmailAddress") & ""
				SESRegion			= r("SESRegion") & ""
				SESRegionID			= r("SESRegionID") & ""
				IsUnitSES			= r("IsUnitSES") & ""
				IsFinancial			= r("IsFinancial") & ""
				DateFinancialChange	= r("DateFinancialChange") & ""
				
				Response.Write "<tr>" 
				Response.Write "<td align=""center"">" & VESAUnitID & "</td>"
				Response.Write "<td align=""center"">" & VESAUnit & "</td>"
				Response.Write "<td align=""center"">" & Password & "</td>"
				Response.Write "<td align=""center"">" 
				Call noOfMembers(VESAUnitID) 
				Response.Write "</td>"
				
				Response.Write "<td align=""center"">"
				If EmailAddress <> "" Then
					Response.Write EmailAddress
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"

				Response.Write "<td align=""center"">"
				If IsFinancial <> "" Then
					Response.Write IsFinancial
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"

				Response.Write "<td align=""center"">"
				If DateFinancialChange <> "" Then
					Response.Write DateFinancialChange
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"

				Response.Write "<td align=""center"">"
				If SESRegionID <> "" Then
					Response.Write SESRegionID
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"

				Response.Write "<td align=""center"">"
				If SESRegion <> "" Then
					Response.Write SESRegion
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"

				Response.Write "<td align=""center"">"
				If IsUnitSES <> "" Then
					Response.Write IsUnitSES
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"
		 
				r.MoveNext
				j = j + 1
			Loop
		 
		 Case Else
			'Loop through Fields Names and print out the Field Names
			j = 2 'row counter
      
			Response.Write "<th><b>Recipient ID</b></th>"
			Response.Write "<th><b>Surname / Organisation</b></th>"
			Response.Write "<th><b>First Name</b></th>"
			Response.Write "<th><b>Address</b></th>"
			Response.Write "<th><b>Suburb</b></th>"
			Response.Write "<th><b>Postcode</b></th>"
			Response.Write "<th><b>State</b></th>"
			Response.Write "<th><b>Membership Number</b></th>"
			Response.Write "<th><b>Email Address</b></th>"
			Response.Write "<th><b>Phoenix Copies</b></th>"
			Response.Write "<th><b>Pocket Diary</b></th>"
			Response.Write "<th><b>Wall Calender</b></th>"
			Response.Write "<th><b>Unit / Designation</b></th>"
			Response.Write "<th><b>SES Region</b></th>"
			Response.Write "</tr>"
	  		
			'Loop through rows, displaying each field
			Do While Not r.EOF
				RecipientID				= r("RecipientID") & ""
				Surname_Organization	= r("Surname_Organization") & ""
				FirstName				= r("FirstName") & ""
				Address					= r("Address") & ""
				Suburb					= r("Suburb") & ""
				Postcode				= r("Postcode") & ""
				State					= r("State_Name") & ""
				MembershipNumber		= r("MembershipNumber") & ""
				MemberEmailAddress		= r("MemberEmailAddress") & ""
				PhoenixCopies			= r("PhoenixCopies") & ""
				VESAPocketDiary			= r("VESAPocketDiary") & ""
				VESAWallCalendar		= r("VESAWallCalendar") & ""
				VESAUnit				= r("VESAUnit") & ""
				SESRegion				= r("SESRegion") & ""

				Response.Write "<tr>" 
				Response.Write "<td align=""center"">" & RecipientID & "</td>"
				Response.Write "<td align=""center"">" & Surname_Organization & "</td>"
				Response.Write "<td align=""center"">"
				If Not IsNull(r.Fields(2)) Then
					Response.Write FirstName
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"

				Response.Write "<td align=""center"">" & Address & "</td>"
				Response.Write "<td align=""center"">" & UCase(Suburb) & "</td>"
				Response.Write "<td align=""center"">" & Postcode & "</td>"
				Response.Write "<td align=""center"">" & State & "</td>"
				Response.Write "<td align=""center"">" & MembershipNumber & "</td>"
				Response.Write "<td align=""center"">"
			    If Not IsNull(r.Fields(8)) Then
					Response.Write MemberEmailAddress
				Else
					Response.Write "&nbsp;"
				End If 
				Response.Write "</td>"
				Response.Write "<td align=""center"">" & PhoenixCopies & "</td>"
				Response.Write "<td align=""center"">" & VESAPocketDiary & "</td>"
				Response.Write "<td align=""center"">" & VESAWallCalendar & "</td>"
				Response.Write "<td align=""center"">" & VESAUnit & "</td>"
				Response.Write "<td align=""center"">" & SESRegion & "</td>"
		 
				r.MoveNext
				j = j + 1
			Loop
	End Select 

	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</body>"
    Response.Write "</html>"
	'--END OF MAIN CODE BLOCK

	'Make sure to close the Result Set and the Connection object
	r.Close
	'-- close select/form tags
End Sub
'--------------------------------------------------------------------------------------------------


'----- %%%%% -------------------------------------------------------------------------- %%%%% -----
'            VESAAudit.asp and BlueLightViewHistory (Format the Date and Time)
'----- %%%%% -------------------------------------------------------------------------- %%%%% -----
Function FormatAuditDate(DateValue)
   Dim strYYYY
   Dim strMM
   Dim strDD
   Dim strHH
   Dim strMin
   Dim strAMPM

   strYYYY = CStr(DatePart("yyyy", DateValue))

   strMM = CStr(DatePart("m", DateValue))
   If Len(strMM) = 1 Then strMM = "0" & strMM

   strDD = CStr(DatePart("d", DateValue))
   If Len(strDD) = 1 Then strDD = "0" & strDD

   If DatePart("h", DateValue) > 12 Then
      strHH = CStr(DatePart("h", DateValue) - 12)
	  strAMPM = "PM"
   
   ElseIf DatePart("h", DateValue) = 0 Then
      strHH = "12"
      strAMPM = "AM"
   
   ElseIf DatePart("h", DateValue) = 12 Then
      strHH = "12"
      strAMPM = "PM"
   Else
      strHH = CStr(DatePart("h", DateValue))
      strAMPM = "AM"
   End If

   strMin = CStr(DatePart("n", DateValue))
   If Len(strMin) = 1 Then strMin = "0" & strMin

      FormatAuditDate = strDD & "/" & strMM & "/" & strYYYY & " " & strHH & ":" & strMin & " " & strAMPM
End Function
'--------------------------------------------------------------------------------------------------

'----- %%%%% -------------------------------------------------------------------------- %%%%% -----
'            VESAEditAdminUser.asp 
'			 CHECK PASSWORD PATTERN and CONVERT it to (*****)
'----- %%%%% -------------------------------------------------------------------------- %%%%% -----
Function kCheckRegExp(vPattern, vStr)
	Dim oRegExp
	Set oRegExp = New RegExp
	oRegExp.Pattern = vPattern
	oRegExp.IgnoreCase = False
	kCheckRegExp = oRegExp.Test(vStr)
	Set oRegExp = Nothing
End Function

Function kLeachRegExp(vStr1, vPattern, vStr2)
	Dim oRegExp
	Set oRegExp = New RegExp
	oRegExp.Pattern = vPattern
	oRegExp.IgnoreCase = True
	oRegExp.Global = True
	kLeachRegExp = oRegExp.Replace(vStr1, vStr2)
	Set oRegExp = Nothing
End Function
'--------------------------------------------------------------------------------------------------

Sub bulkInsert(c, r, table, filePath)

	strSQL = "BULK INSERT " & table & " FROM '" & filePath & "' " & _
            "WITH ( FIELDTERMINATOR = ',', ROWTERMINATOR = '\n' )"

	'-- Execute our SQL statement and store the recordset
   Set r = c.Execute(strSQL)
   
	r.Close
	Set r = Nothing
End Sub

Sub noOfMembers(strVESAUnitID)
	'-----Count All Members with corresponding VESAUnit
	Set ObjRs = Server.CreateObject("ADODB.Recordset")
	ObjRs.Open "SELECT * FROM VESA_tblMembers M INNER JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID WHERE M.VESAUnitID = '" & strVESAUnitID & "'", Conn, 1, 1
	ObjRsCount = ObjRs.RecordCount
	ObjRs.Close

	Response.Write ObjRsCount

End Sub
%>
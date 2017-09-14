<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="include/JSON_2.0.4.asp"-->
<!--#include file="include/include.asp"-->

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
  EstablishConnection()                                  

  '- Build our query based on the input.
  strSQL = "SELECT * FROM VESA_tblMembers M"
  strSQL = strSQL & " INNER JOIN MembersDB_tblState S ON M.StateID = S.StateID"
  strSQL = strSQL & " LEFT JOIN VESA_tblUnit U ON M.VESAUnitID = U.VESAUnitID"
  strSQL = strSQL & " LEFT JOIN VESA_tblSESRegion R ON U.SESRegionID = R.SESRegionID"
  strSQL = strSQL & " WHERE (U.IsActive = '1')"
  strSQL = strSQL & " AND U.VESAUnitID = ?"
  strSQL = strSQL & " ORDER BY M.RecipientID DESC"
                             
  VESAUnitID = Request.QueryString("VESAUnitID")

  arParams = array(VESAUnitID)

  Set cmd = Server.CreateObject("ADODB.Command")

  cmd.CommandText = strSQL

  Set cmd.ActiveConnection = Conn

  QueryToJSON(cmd, arParams).Flush
                             
  CloseConnection()
End If

 Function QueryToJSON(dbcomm, params)
    Dim rs, jsa
    Set rs = dbcomm.Execute(,params,1)
    Set jsa = jsArray()
    
    Do While Not (rs.EOF Or rs.BOF)
      Set jsa(Null) = jsObject()
      
      For Each col In rs.Fields
        jsa(Null)(col.Name) = col.Value
      
      Next
      rs.MoveNext
    Loop
    
    Set QueryToJSON = jsa
    rs.Close
  End Function 
                           
 %> 
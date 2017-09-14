<%
Dim Conn

Sub EstablishConnection()
   on error resume next

   Set Conn = Server.CreateObject("ADODB.Connection")
   Conn.ConnectionString= "Provider=SQLOLEDB; Data Source=mssql2012.netregistry.net; Initial Catalog=database1_cwmedia_com_au; User ID=data1051; Password=ai2JDT0rxU;"
   Conn.Open
End Sub

Sub CloseConnection()
   Conn.close
   Set Conn = nothing
End Sub

%>

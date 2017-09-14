<%@LANGUAGE="VBSCRIPT"%>
<!--#INCLUDE FILE="include/include.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>NZ Blue Light Members Database : Branch Login</title>
<meta name="keywords" content="" />
<meta name="NZ Blue Light Members Database" content="" />
<link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript" src="javascript/resetButton.js"></script>
</head>
<body>
<div id="wrapper">
  <div id="menu">
		<ul id="main">
			<li><a href="AdminLogin.asp">Home</a></li>
			<li><a href="VESAUnitLogin.asp">Unit Login</a></li>
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
    <div id="sidebar1" class="sidebar">
      <ul>
        <li> <br />
		  <img src="images/NZBL-Logo.jpg" width="222" height="80" alt="" />
          <!--<ul>
            <li>&nbsp;</li>
          </ul>-->
        </li>
      </ul>
    </div>
    <!-- start content -->
    <div id="content">
      <div class="post">
        <h1 class="title"><a href="#">Welcome to the NZ Blue Light Branch Log-In area</a></h1>
        <p class="byline">Please select your branch and enter password below:</p>
        <div class="entry">
			<form name="AdminForm" action="NZBLLogin.asp" method="post">
			<input type="hidden" name="Logintype" value="BranchLogin">

			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
			<td>
			   <table border="0" cellspacing="1" cellpadding="0">
			   <tr>
			   <td class="blacktext"><div align="left"><b><label for="Branch">Branch:</label></b></div></td>
			   <td><img src="../../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td valign="top">
			   <div align="left">
			   <select name="BranchID" id="BranchID" class="inputTextField">
			   <option value="0">Please Choose</option>
			   <%
			   Dim rsNZBL
			   Dim strSQL
							  
			   EstablishConnection()

			   strSQL = "SELECT DISTINCT BranchID, Branch FROM NZBlueLight_tblBranch WHERE Branch IS NOT NULL ORDER BY BranchID"
			   
			   Set rsNZBL = Server.CreateObject("ADODB.Recordset")
			   
			   rsNZBL.Open strSQL,Conn
			   rsNZBL.MoveFirst
							  
			   Do Until rsNZBL.EOF %>
			      <option value="<%=rsNZBL.Fields("BranchID")%>"><%=rsNZBL.Fields("Branch")%></option>
			   <%
			      rsNZBL.MoveNext
			   Loop
			   
			   rsNZBL.Close
			   Set rsNZBL = Nothing
			   
			   CloseConnection()
			   %>
			   </select>
			   </div>
			   </td>
			   </tr>

			   <tr>
			   <td class="blacktext"><div align="left"><b><label for="Password">Password:</label></b></div></td>
			   <td><img src="../../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td><div align="left"><input type="password" id="Password" name="Password" size="30" maxlength="30" class="inputTextField2"></div></td>
			   </tr>
			   
			   <tr>
			   <td>&nbsp;</td>
			   <td><img src="../../images/spacer.gif" width="5" height="1" alt="" /></td>
			   <td align="left">
			      <table border=0 cellspacing=0 cellpadding=0>
				  <tr>
				  <td>
				  <input type="image" name="login" class="login-btn" src="http://www.roscripts.com/images/btn.gif" alt="login" title="login" />
				  <input type="hidden" name="login" value="true">
				  </td>
				  <td><img src="../../images/spacer.gif" width="5" height="1" alt="" /></td>
				  <td valign="top">
				  <script type="text/javascript">
				  var ri = new resetimage("images/button/reset_off.gif");
				  ri.name = "resetter";
				  ri.rollover = "images/button/reset_on.gif";
				  ri.write();
				  </script>
				  <noscript><input type="reset"></noscript>
				  </td>
				  </tr>
				  </table>
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



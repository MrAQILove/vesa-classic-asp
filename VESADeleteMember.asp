<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="include/include.asp"-->

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
  Call DeleteMember()
End If

Sub DeleteMember() 
%>
    <!DOCTYPE html>

    <html lang="en">
    <head>
    <meta charset="utf-8" />
    <link href="css/default.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="css/font-awesome.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="css/buttons.css">
    <link rel="stylesheet" href="css/forms.css">
    <link rel="stylesheet" href="css/base.css">
    <title>Delete Members</title>
        <script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
        <script type="text/javascript">
           // Replaces null values in JSON
          function replaceEmpty(json, defaultStr)
          {
            return json.map(function (el) 
            {
              Object.keys(el).forEach(function(key)
              {
                el[key] = el[key] || defaultStr;
              });
              
              return el;
            });
          }

          // Capitalize the first letter of a String
          function jsUcfirst(string) {
            return string.charAt(0).toUpperCase() + string.slice(1);
          }
            
          $(function() {
            $('#VESAUnitID').change(function() {
              $.getJSON('VESAMembersJSON.asp?VESAUnitID=' + $('#VESAUnitID').val(), function(member) 
              {
                // Replaces null with "Not Available" in JSON
                member = replaceEmpty(member,"Not Available");
                      
                $('#MembersDetails').empty();

                $('#MembersDetails').append('<tr align="center" height="30"><td class="tab_header_cell"><b>Delete</b></td><td class="tab_header_cell"><b>Membership <br /> Number</b></td><td class="tab_header_cell"><b>Name / Organisation</b></td><td class="tab_header_cell"><b>Email Address</b></td><td class="tab_header_cell"><b>Publication Assigned</b></td><td class="tab_header_cell"><b>Unit / Designation</b></td><td class="tab_header_cell"><b>SES Region</b></td></tr>');
                      
                for (var i=0; i < member.length; i++ )
                {
                  if (i % 2 != 0) {
                    var class1 = "listTableText0" 
                  }

                  else {
                    class1 = "listTableText1"
                  }

                  var suburb = member[i].Suburb.toUpperCase();
                  var unit = member[i].VESAUnit.toUpperCase();
                  
                  // Replace '7' with VIC in the State field
                  if ( member[i].StateID == 7 ) {
                    var state = "VIC";
                  }

                  $('#MembersDetails').append('<tr height="20" class="'+ class1 +'">' + 
                          '<td align="center"><input type="checkbox" id="DoDelete" name="DoDelete" value="' + member[i].RecipientID + '"></td>' +
                          '<td align="center" style="color: #0000a0; font-weight:bold">' + member[i].MembershipNumber + '</td>' +
                          '<td><div style="padding:5px !important;">' + 
                          '<p><strong>' + jsUcfirst (member[i].FirstName) + '&nbsp;' + member[i].Surname_Organization.toUpperCase() + '</strong></p>' +
                          '<p>' + member[i].Address + '<br />' + suburb + '&nbsp;' + state + '&nbsp;' + member[i].Postcode +'</p>' +
                          '</div></td>' +
                          '<td align="center">' + member[i].MemberEmailAddress + '</td>' +
                          '<td><div style="padding:5px !important; width:150px !important;">' +
                          '<div style="padding-bottom:15px !important;">' + 
                          '<div style="width: 80%; float:left">Phoenix Copies:</div>' +
                          '<div style="width: 20%; float:right"><strong><font color=""#ff0000"">' + member[i].PhoenixCopies + '</font></strong></div></div>' +
                          '<div style="padding-bottom:15px !important;">' + 
                          '<div style="width: 80%; float:left">Wall Calendar:</div>' +
                          '<div style="width: 20%; float:right"><strong><font color=""#ff0000"">' + member[i].VESAWallCalendar + '</font></strong></div></div>' +
                          '</div></td>' +
                          '<td align="center"><font color="#ff0000">' + unit + '</font></td>' +
                          '<td align="center">' + member[i].SESRegion + '</td></tr>');
                }            
              });
            });
          });
        </script>
        <script type="text/javascript">
        
        function stopSubmit() {
          return false;
        }
        
        // Delete Member
        function deleteSelected()
        {
          var ctr;
       
          ctr = 0;
       
          // check for single checkbox by seeing if an array has been created
          var cblength = document.forms['MultiDeleteForm'].elements['DoDelete'].length;
          
          if(typeof cblength == "undefined") {
            if(document.forms['MultiDeleteForm'].elements['DoDelete'].checked == true) ctr++;
          }
          
          else
          {
            for(i = 0; i < document.forms['MultiDeleteForm'].elements['DoDelete'].length; i++) {
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
        checked = false;
        function checkedAll (MultiDeleteForm) 
        {
          var aa = document.getElementById('MultiDeleteForm');
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
      </script> 
    </head>
     
    <body>
      <div id="wrapper">
        <nav id="menu">
          <ul id="main">
            <li><a href="VESAMain.asp">Home</a></li>
            <li><a href="VESAUnitLogin.asp">Unit Login</a></li>
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
            EstablishConnection()

            '- Build our query based on the input.
            strSQL = "SELECT * FROM VESA_tblUnit"
            strSQL = strSQL & " WHERE IsUnitSES = '1'"
            strSQL = strSQL & " AND IsActive = '1'"
            strSQL = strSQL & " ORDER BY VESAUnit ASC"

            '-- Execute our SQL statement and store the recordset
            Set rs = Conn.Execute(strSQL)

            If Not rs.EOF Then
              arrUnits = rs.GetRows
              rs.Close
              Set rs = Nothing
              CloseConnection()
            %>
               <!--/* Start Here */-->
                <table border="0" cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td align="center">
                      <table border="0" align="center" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                          <td>
                            <!--/* header */-->
                            <h1 class="title"><a href="#">Delete a Member</a></h1>
                            <p class="byline"><b>Please select the Unit of the Member that you want to delete.</b></p>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                  
                  <tr height="10"><td>&nbsp;</td></tr>

                  <tr height="10">
                    <td>
                      <form class="pure-form pure-form-aligned" name="DeleteForm" method="get">
                        <fieldset>
                          <select name="VESAUnitID" id="VESAUnitID" class="pure-input-medium" required>
                            <option> -- Select VESA Unit -- </option>
                            <%   
                            For i = 0 To Ubound(arrUnits,2)
                              Response.Write "<option value=""" & arrUnits(0,i) & """>"
                              Response.Write arrUnits(1,i) & "</option>" & VbCrLf
                            Next
                            %>
                          </select>
                        <fieldset>
                      </form>    
                    </td>
                  </tr>

                  <tr height="10"><td>&nbsp;</td></tr>
                  
                  <tr>
                    <td>
                      <form name="MultiDeleteForm" id="MultiDeleteForm" action="VESASave.asp" method="post" onSubmit="return stopSubmit()">
                        <input type="hidden" name="VESAID" value="<%=Session("VESAID")%>">
                        <input type="hidden" name="ActionType" value="Delete">
                        
                        <div style="background-color:#eeeeee">
                          <table id="MembersDetails" border="0" cellspacing="2" cellpadding="1" width="100%"></table>
                        </div>
                        <div style="margin-top:10px; margin-bottom:10px;"><button type="button" class="pure-button" onClick="deleteSelected()">Delete</button></div>
                        </form>
                    </td>
                  </tr>
                </table>
            <% 
            Else 
              rs.Close
              Set rs = Nothing
              CloseConnection()
              Response.Write "<p>Something bad went wrong</p>"
            End If 
          %> 
        </section>
        <!-- end section -->
        
        <footer id="footer">
          <p class="copyright">&copy;&nbsp;&nbsp;2008 -  <%=Year(Date)%> All Rights Reserved &nbsp;&bull;&nbsp; Design and Developed by <a href="http://www.cwaustral.com.au">Countrywide Austral Pty Ltd</a>.</p>
        </footer>
      </body> 
    </html> 
<% End Sub %>
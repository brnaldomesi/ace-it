<% Option Explicit %>
<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->


<%
'--------------------------------------------
Sub PostActivity()
  Dim oDataActivity
  
  If mUserID <> 0 Then
    Set oDataActivity = New  clsACEDataConn
    oDataActivity.ACEMemberActivityInsert2 mUserID, 10, Request.ServerVariables("SCRIPT_NAME") & " " & Request.ServerVariables("QUERY_STRING")
    Set oDataActivity = Nothing
  End If
End Sub
'--------------------------------------------

  PostActivity
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ACE - Received Files List</title>

<script language="JavaScript">
<!--
ns4 = (document.layers)? true:false
ie4 = (document.all)? true:false
  
function NewWin(url, width, height) {
	var win;
	var windowName;
	var params;
	windowName  = "ACE";
	params = "toolbar=0,";
	params += "location=0,";
	params += "directories=0,";
	params += "status=0,";
	params += "menubar=0,";
	params += "scrollbars=1,";
	params += "resizable=1,";
	params += "top=10,";
	params += "left=10,"; 
	params += "width="+(screen.width - 50) +",";
	params += "height="+(screen.height - 75);
	//params += "width="+width+",";
	//params += "height="+height;
	win = window.open(url, windowName, params);
	if (win == null) {
	  alert("You must have popups enabled to use this site. \n If you have an add-in such as PopThis! goto PopThis! Options and disable protection for this site. \n If you need further assistance please contact support by clicking the support tab on this page. \n Thank you.");
	} else {
      win.focus();
	}
}

-->
</Script>
</head>

<body>

<p align="center"><img border="0" src="../Images/ace_logo.gif" width="149" height="71"><br>
<img border="0" src="../Images/LineGradBlue.jpg" width="422" height="11"></p>
<div id="divProcessing" >
  <p align="center"><font color="#808000" size="4">... Your request is being processed please wait ...</font></p>
</div>


<table border="0" cellspacing="1" width="100%">
  <tr>
    <td width="100%" align="center">
      <p align="center"><br>
      The list below contains the  last 300 files <b>received</b> by <b><i>ACE</i></b> for your Office/Company.<font color="#808080" size="1"><br>
      </font><font color="#808080" size="2">Please allow up to two hours from the time you submit a claim for it to appear here. If submitted after 4:30pm ET it may not appear until the next 
      business day.</font><br>
      <font size="2" color="#800080">The first Column in the list below contains the status code: <b>C</b> = Closed, <b>P</b> = Pending - the file has been 
      received and is in process.</font></p>
<script language="JavaScript">strCode = '';</script>

<%
Select Case mUserTypeID 
      Case cMemberType_Administrator, cMemberType_ACEEmployee, cMemberType_ACEAdmin, cMemberType_CoMgr
%>
<script language="JavaScript">jsparam = 'en/<%=EncStr("Received/" & mUserInsCoName & "/", 1)%>/';</script>
<%
      Case Else 'cMemberType_CoOfficeMgr, rep
%>      
<script language="JavaScript">jsparam = 'en/<%=EncStr("Received/" & mUserInsCoName & "/" & mUserACEInsCoOfficeName, 1)%>/';</script>
<%
    End Select
%>
        <p align="center">
        <select size="8" style="color: #000080" name="cmbFilesReceived" >
                                  <!------- Include code ------------->
<script src="../Inc/jsPullExtranet.js"></script>
<script language="JavaScript">document.write(strCode);</script>
                                  <!------- End Include code --------------->                                     
        </select>
        </p>
         
      <p align="center"><b><font color="#000080">Select a  file from the list above and then 
      &lt;&lt; <a href="Javascript:if (cmbFilesReceived.options[cmbFilesReceived.selectedIndex].value != '0') {NewWin(cmbFilesReceived.options[cmbFilesReceived.selectedIndex].value);}">click here</a> 
      &gt;&gt; to view the audit 
      or status for that 
      file.</font></b></p>
      <p>&nbsp;</p>
      <p align="center"><font color="#0000FF"><b>** Search Complete **</b> </font></p>
    </td>
  </tr>
</table>

<!--  onchange="if (this.options[this.selectedIndex].value != '0') {NewWin(this.options[this.selectedIndex].value);}" -->
<script language="JavaScript">
<!--
  if (ns4) {
    //divProcessing.visibility = "hide"
    document.layers['divProcessing'].visibility = "hide"
  } else if (ie4) {
    //divProcessing.visibility = "hidden"
    document.all['divProcessing'].style.visibility = "hidden"
  }  
-->
</Script>

</body>

</html>
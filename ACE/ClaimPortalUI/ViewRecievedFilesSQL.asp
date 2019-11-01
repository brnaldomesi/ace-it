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
  
  
Function SQLClaimsReceived(sCo, sOffice, lDays)
  Dim TopN
      
  SQLClaimsReceived = _
    "SELECT A.ClaimNo, A.EstimateIDPK, A.ClmLastName, " & _
            "CASE WHEN (ISNULL(E.Approved,'1/1/90') > ISNULL(E.LastProcessed,'1/1/90')) THEN 'C' ELSE 'P' END AS RStatus, " & _
            "A.ReceivedDate, A.InsuranceCompanyID " & _
    "FROM   tblAssignments A INNER JOIN " & _
           "tblInsuranceCompanies I ON A.InsuranceCompanyID = I.InsuranceCompanyIDPK LEFT OUTER JOIN " & _
           "tblAssignmentsEx E ON A.EstimateIDPK = E.EstimateIDPK " & _
    "WHERE A.ReceivedDate >= '" & Date - 7 & "' AND (I.InsuranceCompanyName = '" & sCo & "') "


  If Len(sOffice) > 0 Then SQLClaimsReceived = SQLClaimsReceived & _
    "AND (I.InsuranceCompanyOfficeLocation = '" & sOffice & "') "
    
  SQLClaimsReceived = SQLClaimsReceived & _
    "ORDER BY A.ReceivedDate DESC;"
End Function

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

<p align="center">

<font color="#808080" size="2">Please allow up to two hours from the time you 
submit a claim for it to appear here. If submitted after 4:30pm ET it may not 
appear until the next business day.</font><font color="#808080"> </font></p>
<table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="0">
  <tr>
    <td width="75%" bgcolor="#808080"><b><font color="#FFFFFF" size="4" face="Arial">&nbsp;Files 
    Received in the last 7 days: </font></b></td>
    <td width="25%" bgcolor="#808080" align="right">
      <b><a accesskey="R" href="JavaScript:window.location.reload();"><font color="#FFFF99">Refresh</font></a>&nbsp;</b>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">&nbsp; <b><font size="2" face="Arial"> Find a Claim:</font></b>
<input type="text" name="txtClaimNum" size="30" tabindex="2"> <input type="submit" value="Find..." name="cmdView" 
            tabindex="3">
      
<script language="JavaScript">strCode = '';</script>

<%
Select Case mUserTypeID 
      Case cMemberType_Administrator, cMemberType_ACEEmployee, cMemberType_ACEAdmin, cMemberType_CoMgr
%>
<script language="JavaScript">jsparam = 'en/<%=EncStr("Received/" & mUserInsCoName & "/", 1)%>/';</script>
<%
      Case Else 'cMemberType_CoOfficeMgr, rep
%><font face="Verdana" size="1" color="#808080"><b>Enter all or the first part 
    of the claim number in the box and click find. Searches all submitted claims 
    for a match.</b></font></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
    <p align="center">&nbsp; <b><font size="2" color="#800080">Click a closed file to 
      view the audit, clicking a pending file will show the file status.</font></b></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">&nbsp;</td>
  </tr>
</table>

<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <th width="5%" bgcolor="#C0C0C0">Status</th>
    <th width="35%" bgcolor="#C0C0C0">Claim #</th>
    <th width="35%" bgcolor="#C0C0C0">Insured</th>
    <th width="25%" bgcolor="#C0C0C0">Date Rcv.</th>
  </tr>
<%
  Dim sBColor, sFont, EvenRow
  Dim sCoURL, sPersonURL, sEditURL
  
  sFont = " <font face=""Arial"" size=""2"">"
  EvenRow = False
  
  'If Not rs.EOF Then rs.MoveLast: rs.MoveFirst 'don't need for disconnected recordset
  Do While Not rs.EOF
    If EvenRow Then sBColor = " bgcolor=""#E0E0E0"" " Else sBColor = " bgcolor=""#FFFFFF"" "
'    Response.Write "<tr>" & _
'                     "<td width=""25%""" & sBColor & ">" & sFont & "<a target=""_blank"" href=""frmContacts_Office_Entry.asp?ID=" & rs("ContactIDPK") & """>" & rs("Name") & "</a></font></td>" & _
'                     "<td width=""25%""" & sBColor & ">" & sFont & "" & rs("PrimaryContact") & "</font></td>" & _
'                     "<td width=""25%""" & sBColor & ">" & sFont & "" & rs("Phone") & "</font></td>" & _
'                     "<td width=""25%""" & sBColor & ">" & sFont & "" & rs("Fax") & "</font></td>" & _
'                   "</tr>"

'    Response.Write _
'                   "<tr onMouseOver=""this.bgColor = '#FFFF00'"" onMouseOut =""this.bgColor = '" & Mid(sBColor, 11, 7) & "'"" " & sBColor & ">" & _
'                     "<td width=""25%"">" & sFont & "<a target=""_blank"" href=""frmContacts_Office_Entry.asp?ID=" & rs("ContactIDPK") & """>" & rs("Name") & "</a></font></td>" & _
'                     "<td width=""25%"">" & sFont & "" & rs("PrimaryContact") & "</font></td>" & _
'                     "<td width=""25%"">" & sFont & "" & rs("Phone") & "</font></td>" & _
'                     "<td width=""25%"">" & sFont & "" & rs("Fax") & "</font></td>" & _
'                   "</tr>"

'    Temp = "frmContacts_Office_Entry.asp?ID=" & rs("ContactIDPK") & "&ref=" & Server.URLEncode(sURL)
'    Response.Write _
'                   "<tr OnClick=""window.location = '" & Temp & "';"" onMouseOver=""this.bgColor = '#FFFF00'"" onMouseOut =""this.bgColor = '" & Mid(sBColor, 11, 7) & "'"" " & sBColor & ">" & _
'                     "<td width=""25%"">" & sFont & "<a href=""" & Temp & """>" & rs("Name") & "</a></font></td>" & _
'                     "<td width=""25%"">" & sFont & "" & rs("PrimaryContact") & "</font></td>" & _
'                     "<td width=""25%"">" & sFont & "" & rs("Phone") & "</font></td>" & _
'                     "<td width=""25%"">" & sFont & "" & rs("Fax") & "</font></td>" & _
'                   "</tr>"
'    EvenRow = Not EvenRow
'    rs.MoveNext


    sCoURL = "frmContacts_Office_Entry.asp?ID=" & rs("ContactIDPK") & "&ref=" & Server.URLEncode(sURL)
    sPersonURL = "frmContacts_Children_Entry.asp?ID=" & rs("ContactIDPK") & "&ref=" & Server.URLEncode(sURL)
    
    sEditURL = sCoURL
    If rs("ContactTypeID") = 26 Then sEditURL = sPersonURL '& " target=""_blank"" "
    Response.Write _
                   "<tr OnClick=""window.location = '" & sEditURL & "';"" onMouseOver=""this.bgColor = '#FFFF00'"" onMouseOut =""this.bgColor = '" & Mid(sBColor, 11, 7) & "'"" " & sBColor & ">" & _
                     "<td width=""25%"">" & sFont & "<a href=""" & sEditURL & """>" & rs("Name") & "</a></font></td>" & _
                     "<td width=""25%"">" & sFont & "" & rs("PrimaryContact") & "</font></td>" & _
                     "<td width=""25%"">" & sFont & "" & rs("Phone") & "</font></td>" & _
                     "<td width=""25%"">" & sFont & "" & rs("Fax") & "</font></td>" & _
                   "</tr>"
    EvenRow = Not EvenRow
    rs.MoveNext


  Loop
  rs.close
%>
</table>

<table border="0" cellspacing="1" width="100%">
  <tr>
    <td width="100%" align="center">
      <p align="center"><br>
      &nbsp;</p>
        <p align="center">
        <select size="8" style="color: #000080" name="cmbFilesReceived" >
                                  <!------- Include code ------------->
<script src="../Inc/jsPullExtranet.js"></script>
<script language="JavaScript">document.write(strCode);</script>
                                  <!------- End Include code --------------->                                     
        </select>
        </p>
        <font size="2" color="#800080">The first Column in the list above contains the status code: <b>C</b> = Closed, <b>P</b> = Pending - the file has been 
      received and is in process.</font> 
      <p align="center"><b><font color="#000080">Select a CLOSED file from the list above and then <a href="Javascript:if (cmbFilesReceived.options[cmbFilesReceived.selectedIndex].value != '0') {NewWin(cmbFilesReceived.options[cmbFilesReceived.selectedIndex].value);}">click here</a> to view the audit for that 
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
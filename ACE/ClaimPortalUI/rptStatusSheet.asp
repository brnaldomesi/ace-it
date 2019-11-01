<% Option Explicit %>
<!-- #INCLUDE Virtual="/ACE/Login/Security.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsACEDataConn.asp" --><%
  Dim oData
  Dim rs
  Dim RecCount
  
  Dim AuditID
  
  AuditID = Request.Querystring("EstID") 
  If AuditID & "" = "" Then AuditID = "1101870" '"1101275" '"1107602"
  
  Set oData = New clsACEDataConn
    
  '--- Populate recordset ---
  Set rs = oData.rsCallStoredProc("ACE_StatusSheet", AuditID, "", "")
  RecCount = rs.RecordCount
%>
<html>

<head>
<meta http-equiv="expires" content="Tue, 20 Aug 1996 01:00:00 GMT">
<meta http-equiv="Pragma" content="no-cache">
<meta name="Author" content="Bob Puppo">
<meta name="Copyright" content="(C) Bob Puppo, 1999-2008, All rights reserved">
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ACE Status Sheet</title>
</head>

<body>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="700" id="AutoNumber1">
  <tr>
    <td width="100%" colspan="4">
    <p align="center"><img border="0" src="../Images/ace_logo.gif"><b><font size="6"><br>
    </font></b>1155 Phoenixville Pike, Suite 109<br>
    West Chester, PA 19380<br>
    (888) 816-2436<br>
&nbsp;</p>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="4">
    <p align="center"><b>AUDIT STATUS SHEET</b></p>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="4">&nbsp;</td>
  </tr>
  <tr>
    <td width="12%" align="left">To:</td>
    <td width="38%"><%=rs("InsuranceCompanyName")%></td>
    <td width="18%" align="left">File Number:</td>
    <td width="32%"><%=rs("FileNumber")%></td>
  </tr>
  <tr>
    <td width="12%" align="left">&nbsp;</td>
    <td width="38%" rowspan="3" valign="top"><%=Replace(rs("InsCoAddr"), vbCrLf, "<br/>")%></td>
    <td width="18%" align="left">Policy Number:</td>
    <td width="32%"><%=rs("PolicyNo")%></td>
  </tr>
  <tr>
    <td width="12%" align="left">&nbsp;</td>
    <td width="18%" align="left">Claim Number:</td>
    <td width="32%"><%=rs("ClaimNo")%></td>
  </tr>
  <tr>
    <td width="12%" align="left">&nbsp;</td>
    <td width="18%" align="left">Date of Loss:</td>
    <td width="32%"><%=rs("LossDate")%></td>
  </tr>
  <tr>
    <td width="12%" align="left">Insured:</td>
    <td width="38%"><%=rs("InsuredName")%></td>
    <td width="18%" align="left">Date of Billing:</td>
    <td width="32%"><%=rs("BillingDate")%></td>
  </tr>
  <tr>
    <td width="12%" align="left">Claimant:</td>
    <td width="38%"><%=rs("ClmName")%></td>
    <td width="18%" align="left"><!--Federal ID:--></td>
    <td width="32%"><!--fldFederalID--></td>
  </tr>
</table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-top-style:solid; border-top-width:1; border-bottom-style:solid; border-bottom-width:1; padding-left:0; padding-right:0; padding-top:1; padding-bottom:1" bordercolor="#111111" width="700">
  <tr>
    <td width="175">Status Date</td>
    <td width="10">&nbsp;</td>    
    <td width="515">Description</td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="700">
  <%
  Do While (Not rs.eof) And Not(IsNull(rs("StatusDTM")))
  %>
  <tr>
    <td width="175" valign="top" nowrap><%=rs("StatusDTM")%>&nbsp;</td>
    <td width="10">&nbsp;</td>    
    <td width="515"><%=Server.HTMLEncode(rs("StatusDesc") & "") & "<br>" & Server.HTMLEncode(rs("StatusNote") & "")%>&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" colspan="3">&nbsp;</td>
  </tr>
  <%
    rs.MoveNext
  Loop
  %>
  <tr>
    <td width="100%" colspan="3" align="center">
    <%
  If RecCount = 0 Or (IsNull(rs("StatusDTM"))) Then
    Response.Write "... No Status Entries For This File ..."
  End If
   %>&nbsp;
   </td>
  </tr>
</table>

</body>

</html>
<% Option Explicit %>
<!-- #INCLUDE Virtual="/ACE/Login/Security.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsACEDataConn.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsSQLStmt.asp" --><%
  Dim oData, oSQL
  Dim rs
  Dim RecCount
  
  Dim sSQL
  
  Set oData = New clsACEDataConn
  'Set oSQL = New clsSQLStmt
    
  Dim sEstID
        
  sEstID = Request.Querystring("EstID")
    
  sSQL = "SELECT a.ClaimNo, p.PhotoFileSpec " & _
     "FROM tblAssignments a " & _
     "Inner Join tblPhotos p On p.EstimateID = a.EstimateIDPK " & _
     "WHERE a.EstimateIDPK = " & sEstID     

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient

    rs.Open sSQL, oData.sAdminConn, 3, 4
    If Not rs.Eof Then rs.MoveLast: rs.MoveFirst
    Set rs.ActiveConnection = Nothing
  
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ACE - Claim Status Activity</title>
<meta name="Microsoft Border" content="b">
</head>

<body><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">

<table width="100%" style="border-collapse: collapse" bordercolor="#111111" cellpadding="1" cellspacing="1" border="0">
  <tr>
    <td width="50%">&nbsp; </td>
    <td><img border="0" src="../Images/ace_logo.gif"></td>
    <td width="50%">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3" align="center"><img border="0" src="../Images/spacer.gif" width="1" height="2"></td>
  </tr>
</table>
<table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" cellpadding="0">
  <tr>
    <td width="10%" bgcolor="#1A74A1" align="Left" nowrap></td>
    <td width="80%" bgcolor="#1A74A1" align="center" nowrap>
    <p><b><font color="#FFFFFF" size="3" face="Arial">Photos for Claim Number: <%=rs("ClaimNo")%></font></b></p>
    </td>
    <td width="10%" bgcolor="#1A74A1" align="right" nowrap><font color="#FFFF99" size="2" face="Arial">Count: <%=rs.RecordCount%></font> </td>
  </tr>
  <tr>
    <td width="100%" colspan="3"><img border="0" src="../Images/spacer.gif" width="1" height="2"></td>
  </tr>
  <tr>
    <td width="100%" colspan="3"><img border="0" src="../Images/spacer.gif" width="1" height="2"></td>
  </tr>
</table>

<div>
<%
  Dim i
  i = 1 
  Do While Not rs.EOF 
%> 
<table style="float: left; border-collapse: collapse" cellpadding="0" cellspacing="0">
  <caption align="bottom" valign="bottom"><font face="Arial" size="2">
  <a style="text-decoration: none;" title="Click to download this photo to your system" href="http://webportal.staging-portal.ace-it.com/photo/download?file=<%=Server.URLEncode(rs("PhotoFileSpec"))%>"><img border="0" src="../Images/dwn.GIF"></a> Photo <%=i%></font></caption>
  <tr><td><a href="#Photo<%=i%>"><img border="3" src="http://webportal.staging-portal.ace-it.com/photo?file=<%=Server.URLEncode(rs("PhotoFileSpec"))%>" width="100" hspace="4" vspace="4"></a></td></tr>
</table>
<%  rs.MoveNext
    i=i+1
  Loop
%>
</div>
<br clear="all"> <!-- stop the floating -->
<hr>

<%
  i = 1 
  If rs.RecordCount > 0 Then rs.MoveFirst
  Do While Not rs.EOF 
  'Response.Write rs("PhotoFileSpec")
%> <br>
<a name="Photo<%=i%>"></a>Photo <%=i%>: <a style="text-decoration: none;" title="Click to download this photo to your system" href="http://webportal.staging-portal.ace-it.com/photo/download?file=<%=Server.URLEncode(rs("PhotoFileSpec"))%>"><img border="0" src="../Images/dwn.GIF"></a> <br>
<img border="0" src="http://webportal.staging-portal.ace-it.com/photo?file=<%=Server.URLEncode(rs("PhotoFileSpec"))%>"> <hr>
<%  rs.MoveNext
    i=i+1
  Loop
%>

<!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<hr>
<p><font size="1" color="#666633">Unless otherwise noted all content �
<a href="http://www.staging-portal.ace-it.com">AMERICAN COMPUTER ESTIMATING, INC.</a><br>
Website design, programs and web code � <a href="http://www.puppozone.com/">PUPPO 
SOFTWARE SOLUTIONS, LLC.</a><br>
All Rights Reserved. <br>
Last modified:
January 09, 2008</font> </p>
</font>

<script src="https://www.google-analytics.com/urchin.js" type="text/javascript"></script>
<script type="text/javascript">_uacct = "UA-1644909-1"; urchinTracker(); </script>

</td></tr><!--msnavigation--></table></body>

</html>
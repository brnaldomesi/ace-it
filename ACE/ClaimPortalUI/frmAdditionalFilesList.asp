<% Option Explicit %>
<html>

<head>


<!-- #INCLUDE Virtual="/ACE/Login/Security.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsACEDataConn.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsSQLStmt.asp" -->

<%
	Dim oData, oSQL
	Dim rs
	Dim RecCount

	Dim sSQL
	Dim sEstID
	Dim sFileName
	
	'If Request.Form("fileName") & "" <> "" Then
		'Response.Write("post")
		'Response.End
	'End If

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		If Request.Form("txtState") = "FileList" Then
			ShowListOfFiles
		ElseIf Request.Form("txtState") = "Show" Then
			ShowFileContent
		End If
	End If
		
Sub ShowFileContent()		
		sEstID = Request.Form("txtEstID")
		sFileName = Request.Form("txtFileName")
		
		sSQL = "Select FileContent = dbo.fn_OriginalAssignmentReadFile(" & sEstID & " ,'" & Replace(sFileName, "'", "''") & "')"

		Set rs = Server.CreateObject("ADODB.Recordset")
		' Setting the cursor location to client side is important to get a disconnected recordset.
		rs.CursorLocation = 3 ' adUseClient
		
		Set oData = New clsACEDataConn
		rs.Open sSQL, oData.mConnStr, 3, 4
		If Not rs.Eof Then rs.MoveLast: rs.MoveFirst
		
		dim ContentType, FileExt, i
		ContentType = "application/pdf"

		i = InStrRev(sFileName,".")
		'FileExt = LCase(Right(sFileName, i))
		FileExt = LCase(Mid(sFileName, i, 99))
		ContentType = ContentTypeFromExtension(FileExt)
		
		Response.Clear

		Response.AddHeader "Expires", "0"
		Response.AddHeader "Content-Transfer-Encoding", "binary" 
		Response.AddHeader "Content-Description", "File Transfer" 
		If Request.Form("txtAction") = "download" Then
			Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName
		End If
		'Response.AddHeader "Content-Length", objStream.Size 
		Response.AddHeader "Content-Type", ContentType '** added this line **
		Response.CharSet = "UTF-8"
		Response.ContentType = ContentType
		
		Response.CodePage = 65001

		Response.BinaryWrite rs("FileContent")
		
		'Response.Write("posting " & sSQL)
		Response.End
End Sub

Sub ShowListOfFiles()
	Set oData = New clsACEDataConn
	'Set oSQL = New clsSQLStmt

	'sEstID = Request.Querystring("EstID") '2017-09-16
	sEstID = Request.Form("txtEstID")
 	
	sSQL = "Portal_FileAttachementsList " & mUserID & "," & sEstID
	 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 ' adUseClient

	rs.Open sSQL, oData.mConnStr, 3, 4
	If Not rs.Eof Then rs.MoveLast: rs.MoveFirst
	Set rs.ActiveConnection = Nothing
	
	If rs.Eof Then 
		Response.Write("...No file attachments available for this file...")
		Response.End
	End If
End Sub

'<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
'<meta name="ProgId" content="FrontPage.Editor.Document">
'<meta name="Microsoft Border" content="b">


%>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ACE - File Attachements</title>
<!--<script src="..\inc\download2.js" type="text/javascript"></script>
<script type="text/javascript">
	function readFile(){
		var x=new XMLHttpRequest();
		x.open("GET", "http://danml.com/wave2.gif", true);
		x.responseType = 'blob';
		x.onload=function(e){download(x.response, "dlBinAjax.gif", "image/gif" ); }
		x.send();
	}
</script>
-->

<!-- Includes --------------------
------------------------------ -->

</head>

<body>
<!-- <form id="frmViewFile" name="frmViewFile" method="POST" action="" target="_self"> -->
<form id="frmViewFile" name="frmViewFile" method="POST" action="" target="_blank">
	<INPUT type="hidden" name="txtState" value="Show" >
	<INPUT type="hidden" name="txtEstID" value=<%=sEstID%> >
	<INPUT type="hidden" name="txtFileName" value="">
	<INPUT type="hidden" name="txtAction" value="download">
</form>

<table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>

<td valign="top">


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
    <!-- <p><b><font color="#FFFFFF" size="3" face="Arial">Attachments for Claim Number: <%=rs("ClaimNo")%></font></b></p> -->
    <p><b><font color="#FFFFFF" size="3" face="Arial">Claim file attachements</font></b></p>
    </td>
    <td width="10%" bgcolor="#1A74A1" align="right" nowrap><!--<font color="#FFFF99" size="2" face="Arial">Count: <%=rs.RecordCount%> &nbsp;&nbsp;&nbsp;</font> -->
	</td>
  </tr>
  <tr>
    <td width="100%" colspan="3"><img border="0" src="../Images/spacer.gif" width="1" height="2"></td>
  </tr>
  <tr>
    <td width="100%" colspan="3"><img border="0" src="../Images/spacer.gif" width="1" height="2"></td>
  </tr>
</table>



<SCRIPT LANGUAGE="JavaScript">
	function ShowFile(fileName, action) {
		document.frmViewFile.txtFileName.value = fileName;
		document.frmViewFile.txtAction.value = action;
		document.frmViewFile.submit();
	}
</Script>

<style type="text/css">
table.paleBlueRows {
  font-family: Arial,Helvetica,sans-serif;
  border: 2px solid #AEC8FF;
  background-color: #EEEEEE;
  width: 100%;
  text-align: left;
  border-collapse: collapse;
}
table.paleBlueRows td, table.paleBlueRows th {
  border: 0px solid #FFFFFF;
  padding: 6px 6px;
}
table.paleBlueRows tbody td {
  font-size: 13px;
}
table.paleBlueRows tr:nth-child(even) {
  background: #D0E4F5;
}
table.paleBlueRows thead {
  background: #4D6794;
}
table.paleBlueRows thead th {
  font-size: 17px;
  font-weight: normal;
  color: #FFFFFF;
  text-align: left;
}
table.paleBlueRows tfoot td {
  font-size: 14px;
}
</Style>


<table class="paleBlueRows">
	<thead>
		<tr>
			<th colspan="3">Attachments for Claim Number: <font color="#FFFF99"><%=rs("ClaimNo")%></font> &nbsp;&nbsp;&nbsp;Count: <%=rs.RecordCount%> &nbsp;</th>
			<!--<th colspan="2"></th>-->
		</tr>
	</thead>
	<tbody>

<%
Dim sRowTemplate
Dim sRowWorkTemplate
sRowTemplate = _
		"<tr>" & _
			"<td><button onclick=""ShowFile('~fileNameEscaped~', 'view');"">View</button></td>" & _
			"<td><button onclick=""ShowFile('~fileNameEscaped~', 'download');"">Download</button></td>" & _
			"<td width=100% >~fileName~</td>" & _
		"</tr>"
		
'href="#" onclick="document.getElementById('form-id').submit();"
  Do While Not rs.EOF
	sRowWorkTemplate = sRowTemplate
	sRowWorkTemplate = Replace(sRowWorkTemplate, "~fileName~", rs("Name"))
	sRowWorkTemplate = Replace(sRowWorkTemplate, "~fileNameEscaped~", Replace(rs("Name"), "'", "\'"))
	sRowWorkTemplate = Replace(sRowWorkTemplate, "~fileName~", rs("Name"))
	Response.Write sRowWorkTemplate
	rs.MoveNext
  Loop
%>

	</tbody>
</table>



<!--
	<ul>
<%
'href="#" onclick="document.getElementById('form-id').submit();"
'  Do While Not rs.EOF 
'	'Response.Write "<li>" & "<a href=""CommingSoon.aspx"">" & rs("Name") & "</a>" & "</li>"
'	Response.Write "<li>" & "<a href=""#"" onclick=""ShowFile('" & Replace(rs("Name"), "'", "\'") & "');"">" & rs("Name") & "</a>" & "</li>"
'	rs.MoveNext
'  Loop
%> 
	</ul>
-->
	

</td>
</tr>

</table>


<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr><td>

<hr>
<p><font size="1" color="#666633">Unless otherwise noted all content �<a href="http://wwace-it.comom">AMERICAN COMPUTER ESTIMATING, INC.</a><br>
Website design, programs and web code �<a href="http://www.puppozone.com/">Anthony Puppo, LLC.</a><br>
All Rights Reserved. <br>
Last modified:
8/2017</font> </p>
</font>

<script src="https://www.google-analytics.com/urchin.js" type="text/javascript"></script>
<script type="text/javascript">_uacct = "UA-1644909-1"; urchinTracker(); </script>

</td></tr>
</table>

</body>

</html>

<%
function ContentTypeFromExtension(FileExt)
	Dim ContentType
	If FileExt = ".323" Then
		ContentType = "text/h323"
	ElseIf FileExt = ".jpe" Then
		ContentType = "image/jpeg"
	ElseIf FileExt = ".jpeg" Then
		ContentType = "image/jpeg"
	ElseIf FileExt = ".jpg" Then
		ContentType = "image/jpeg"
	ElseIf FileExt = ".pdf" Then
		ContentType = "application/pdf"
	ElseIf FileExt = ".pic" Then
		ContentType = "image/pict"
	ElseIf FileExt = ".pict" Then
		ContentType = "image/pict"
	ElseIf FileExt = ".png" Then
		ContentType = "image/png"
	ElseIf FileExt = ".svg" Then
		ContentType = "image/svg+xml"
	ElseIf FileExt = ".tif" Then
		ContentType = "image/tiff"
	ElseIf FileExt = ".tiff" Then
		ContentType = "image/tiff"
	ElseIf FileExt = ".tlh" Then
		ContentType = "text/plain"
	ElseIf FileExt = ".tli" Then
		ContentType = "text/plain"
	ElseIf FileExt = ".toc" Then
		ContentType = "application/octet-stream"
	ElseIf FileExt = ".tr" Then
		ContentType = "application/x-troff"
	ElseIf FileExt = ".trm" Then
		ContentType = "application/x-msterminal"
	ElseIf FileExt = ".trx" Then
		ContentType = "application/xml"
	ElseIf FileExt = ".ts" Then
		ContentType = "video/vnd.dlna.mpeg-tts"
	ElseIf FileExt = ".tsv" Then
		ContentType = "text/tab-separated-values"
	ElseIf FileExt = ".ttf" Then
		ContentType = "application/font-sfnt"
	ElseIf FileExt = ".tts" Then
		ContentType = "video/vnd.dlna.mpeg-tts"
	ElseIf FileExt = ".txt" Then
		ContentType = "text/plain"
	ElseIf FileExt = ".xps" Then
		ContentType = "application/vnd.ms-xpsdocument"	
	Else
		ContentType = "application/zip"
	End If

	ContentTypeFromExtension = ContentType

End function
%>
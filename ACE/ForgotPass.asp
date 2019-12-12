<!-- #INCLUDE Virtual="\clsFormMailer.asp" -->
<!-- #INCLUDE Virtual="\ACE\inc\clsACEDataConn.asp" -->

<!DOCTYPE html>
<html lang="en">

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title></title>
<meta name="Microsoft Border" content="t">
</head>

<body><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<p align="center"><a href="/index.aspx"><img border="0" src="Images/ace_logo.gif" width="149" height="71"></a><br>
<img border="0" src="Images/LineGradBlue.jpg" width="422" height="11"><br>
</p>

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">
<p align="center">
<br>
<%
  Dim mrs 
  Dim oData
  Dim oFormMailer
  
  If Request.QueryString("txtState") = "1" Then

    Set oDataConn = New clsACEDataConn
    Set mrs = oDataConn.ACEMemberGet(0, Request.Form("txtUserName"),"")

    If mrs.eof Then
      response.Write "Sorry, your User Name was not found in the membership database. Please sign up to become a member."
    Else
		Dim rsClaim 
		Dim sSQL
		
		sSQL = "SELECT * FROM tblClaimsSubmitted WHERE ClaimIDPK = 0"
    
		Set rsClaim = Server.CreateObject("ADODB.Recordset")
		rsClaim.Open sSQL, oDataConn.ConnStr, 3, adLockOptimistic     'adOpenDynamic 3, adLockOptimistic
		rsClaim.AddNew
		rsClaim("ClaimBody") = "Your ACE-IT password: " & mrs("MemberPassword")
		rsClaim("EmailTo") = mrs("MemberEmail")
		rsClaim.Update
		rsClaim.MoveFirst
		rsClaim.Update
		rsClaim.Close
		
      response.Write "Your email has been sent, please allow up to 15 minutes for it to arrive."
    End If
    mrs.close
  End If
%> </p>


<!--msnavigation--></td></tr><!--msnavigation--></table></body>

</html>
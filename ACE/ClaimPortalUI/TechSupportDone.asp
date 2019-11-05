<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
<!-- #INCLUDE VIRTUAL="/clsFormMailer.asp" -->

<%

  On Error Resume Next
  mUserName=request.form("txtUserName")
  mPassword=request.form("txtPassword")
  mState=Request.QueryString("txtState")
  
  '--- Email the form to the admin ---
  dim oFormMailer

  Set oFormMailer = New clsFormMailer

  oFormMailer.sHeader = "Form: ACE Claim Rep Portal WEB Site Support"
  oFormMailer.sFooter = "Form mailer created by PUPPO ENGINEERING, (C) 2001 - 2008"

  oFormMailer.sFrom = "aceitweb@comcast.net" '"ClaimRepPortal@ace-it.com" '<--- no spaces
  oFormMailer.sTo = "aceitweb@comcast.net" '"bob@puppozone.com" '"bpuppo@voicenet.com"
  oFormMailer.sSubject = "Claim Portal WEB Site Support Issue"
  'oFormMailer.sImportance = 2 'high

  oFormMailer.sBody = "UserID: " & mUserID & vbcrlf & vbcrlf
  oFormMailer.sBody = oFormMailer.sBody & "User Type: " & mUserTypeID & vbcrlf & vbcrlf
  oFormMailer.sBody = oFormMailer.sBody & "User Security Level: " & mUserSecurityL & vbcrlf & vbcrlf
  oFormMailer.sBody = oFormMailer.sBody & "Name: " & mName & vbcrlf & vbcrlf
  oFormMailer.sBody = oFormMailer.sBody & "Rep Name: " & mRepName & vbcrlf & vbcrlf
  oFormMailer.sBody = oFormMailer.sBody & "Co Name: " & mUserInsCoName & vbcrlf & vbcrlf
  
  'oFormMailer.sBody = "First Name: " & request.form(txtFirstName) & vbcrlf & vbcrlf
  'oFormMailer.sBody = oFormMailer.sBody & "Last Name: " & request.form(txtLastName) & vbcrlf & vbcrlf
  'oFormMailer.sBody = oFormMailer.sBody & "Email Addr.: " & request.form(txtEmail) & vbcrlf & vbcrlf
  'oFormMailer.sBody = oFormMailer.sBody & "Comments: " & request.form(txtComments) & vbcrlf & vbcrlf

  oFormMailer.BuildBody
  oFormMailer.Send
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>ACE - Support</title>
<meta name="Microsoft Border" content="t">
</head>

<body><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<p align="center"><a href="../ACE-IT/index.htm"><img border="0" src="../Images/ace_logo.gif" width="149" height="71"></a><br>
<img border="0" src="../Images/LineGradBlue.jpg" width="422" height="11"><br>
</p>

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">

<p>&nbsp;</p>
<p align="center">Your message has been sent. Thank You.
<p align="center">&nbsp;<p align="center"><a href="frmMain.asp">Return to Portal</a><!--msnavigation--></td></tr><!--msnavigation--></table></body></html>
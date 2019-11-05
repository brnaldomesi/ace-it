<% Option Explicit %>

<!-- #INCLUDE FILE="Security.asp" -->
<!-- #INCLUDE FILE="..\..\clsFormMailer.asp" -->
<!-- #INCLUDE FILE="..\inc\clsACEDataConn.asp" -->

<%

  Dim Msg
  Dim mNewUserName, mNewPassword
  Dim mrs
  Dim bResult
  Dim mState
  Dim oDataConn

  On Error Resume Next
  mNewUserName=request.form("txtUserName")
  mNewPassword=request.form("txtPassword")
  mState=Request.QueryString("txtState")
  On Error Goto 0

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>ACE Member Edit</title>
<meta name="Microsoft Border" content="tb">
</head>

<body><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<p align="center"><a href="../ACE-IT/index.htm"><img border="0" src="../Images/ace_logo.gif" width="149" height="71"></a><br>
<img border="0" src="../Images/LineGradBlue.jpg" width="422" height="11"><br>
</p>

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<%
  If mState = 0 Then
  Set oDataConn = New clsACEDataConn
  Set mrs = oDataConn.ACEMemberGet(mUserID,"","")
%> <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.txtCoName.value == "")
  {
    alert("Please enter a value for the \"Company Name\" field.");
    theForm.txtCoName.focus();
    return (false);
  }

  if (theForm.txtCoName.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Company Name\" field.");
    theForm.txtCoName.focus();
    return (false);
  }

  if (theForm.txtCoName.value.length > 30)
  {
    alert("Please enter at most 30 characters in the \"Company Name\" field.");
    theForm.txtCoName.focus();
    return (false);
  }

  if (theForm.txtFirstName.value == "")
  {
    alert("Please enter a value for the \"First Name\" field.");
    theForm.txtFirstName.focus();
    return (false);
  }

  if (theForm.txtFirstName.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"First Name\" field.");
    theForm.txtFirstName.focus();
    return (false);
  }

  if (theForm.txtFirstName.value.length > 20)
  {
    alert("Please enter at most 20 characters in the \"First Name\" field.");
    theForm.txtFirstName.focus();
    return (false);
  }

  if (theForm.txtLastName.value == "")
  {
    alert("Please enter a value for the \"Last Name\" field.");
    theForm.txtLastName.focus();
    return (false);
  }

  if (theForm.txtLastName.value.length < 2)
  {
    alert("Please enter at least 2 characters in the \"Last Name\" field.");
    theForm.txtLastName.focus();
    return (false);
  }

  if (theForm.txtLastName.value.length > 20)
  {
    alert("Please enter at most 20 characters in the \"Last Name\" field.");
    theForm.txtLastName.focus();
    return (false);
  }

  if (theForm.txtEmail.value == "")
  {
    alert("Please enter a value for the \"E-mail Address\" field.");
    theForm.txtEmail.focus();
    return (false);
  }

  if (theForm.txtEmail.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"E-mail Address\" field.");
    theForm.txtEmail.focus();
    return (false);
  }

  if (theForm.txtUserName.value == "")
  {
    alert("Please enter a value for the \"User name\" field.");
    theForm.txtUserName.focus();
    return (false);
  }

  if (theForm.txtUserName.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"User name\" field.");
    theForm.txtUserName.focus();
    return (false);
  }

  if (theForm.txtUserName.value.length > 20)
  {
    alert("Please enter at most 20 characters in the \"User name\" field.");
    theForm.txtUserName.focus();
    return (false);
  }

  if (theForm.txtPassword.value == "")
  {
    alert("Please enter a value for the \"Password\" field.");
    theForm.txtPassword.focus();
    return (false);
  }

  if (theForm.txtPassword.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Password\" field.");
    theForm.txtPassword.focus();
    return (false);
  }

  if (theForm.txtPassword.value.length > 20)
  {
    alert("Please enter at most 20 characters in the \"Password\" field.");
    theForm.txtPassword.focus();
    return (false);
  }

  if (theForm.txtPasswordVerify.value == "")
  {
    alert("Please enter a value for the \"Password Verification\" field.");
    theForm.txtPasswordVerify.focus();
    return (false);
  }

  if (theForm.txtPasswordVerify.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Password Verification\" field.");
    theForm.txtPasswordVerify.focus();
    return (false);
  }

  if (theForm.txtPasswordVerify.value.length > 20)
  {
    alert("Please enter at most 20 characters in the \"Password Verification\" field.");
    theForm.txtPasswordVerify.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name="FrontPage_Form1" method="post" action="MemberEdit.asp?txtState=2" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
  <div align="center">
    <table border="0" cellpadding="1" bgcolor="#CD0606" width="425" cellspacing="1">
      <tr>
        <td width="562" align="center"><font face="Arial" size="3" color="#FFFFFF"><b>Edit Your Profile<br>
          </b></font><font face="Arial" size="1" color="#C0C0C0">(*Red label names denote required fields)</font></td>
      </tr>
      <tr>
        <td width="562">
          <table border="0" cellpadding="15" cellspacing="0" bgcolor="#FFFFFF">
            <tr>
              <td width="530">
                <table border="0" cellpadding="1" width="447" bgcolor="#FFFFCC">
                  <tr>
                    <td bgcolor="#356C91" align="center" valign="middle" colspan="3" width="432"><b><font size="1" face="verdana, geneva, arial" color="#FFFF00">Company
                      Information</font></b></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial" color="#FF0000">Company</font></td>
                    <td colspan="2" width="373">
                    <!--webbot bot="Validation" s-display-name="Company Name" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="30" --><input name="txtCoName" size="40" maxlength="30" value="<%=CStr(mrs("InsCoName"))%>"></td>
                  </tr>
                  <tr>
                    <td colspan="3" align="center" bgcolor="#356C91" width="432"><b><font size="1" face="verdana, geneva, arial" color="#FFFF00">Personal
                      Information</font></b></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" valign="middle" width="75" align="right"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> <font color="#FF0000">First
                      &amp; Last Name</font></font></td>
                    <td colspan="2" width="373">
                    <!--webbot bot="Validation" s-display-name="First Name" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input name="txtFirstName" size="20" maxlength="20" value="<%=mrs("MemberFirstName")%>">&nbsp;&nbsp;
                    <!--webbot bot="Validation" s-display-name="Last Name" s-data-type="String" b-value-required="TRUE" i-minimum-length="2" i-maximum-length="20" --><input name="txtLastName" size="20" maxlength="20" value="<%=mrs("MemberLastName")%>"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font>&nbsp;
                      Work Phone</font></td>
                    <td colspan="2" width="373"><b>(</b><input name="txtPhoneAc" size="3" maxlength="3" value="<%=Left(mrs("MemberPhone"),3)%>"><b>)</b>&nbsp;<input name="txtPhoneExch" size="3" maxlength="3" value="<%=mid(mrs("MemberPhone"),4,3)%>">&nbsp;<b>-</b>&nbsp;<input name="txtPhoneL4" size="4" maxlength="4" value="<%=right(mrs("MemberPhone"),4)%>">&nbsp;<font size="1">Extension</font>&nbsp;<input name="phoneextn" size="4" maxlength="4" value="<%=mrs("MemberPhoneEx")%>"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> <font color="#FF0000">E-mail
                      Address</font></font></td>
                    <td colspan="2" width="373">
                    <!--webbot bot="Validation" s-display-name="E-mail Address" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" --><input name="txtEmail" size="40" value="<%=mrs("MemberEmail")%>"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial">Gender</font></td>
                    <td colspan="2" width="373"><input type="radio" name="grpGender" value="1"><font face="geneva,arial" size="-1"> Male&nbsp; </font><input type="radio" name="grpGender" value="2"><font face="geneva,arial" size="-1">
                      Female</font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#356C91" align="center" valign="middle" colspan="3" width="432"><font size="1" face="verdana, geneva, arial" color="#FFFF00"><b>Choose
                      Your Sign In Information</b></font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font>&nbsp;
                      <font color="#FF0000">User Name</font></font></td>
                    <td width="150">
                    <!--webbot bot="Validation" s-display-name="User name" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input name="txtUserName" size="20" maxlength="20" value="<%=mrs("MemberLogonName")%>" readonly></td>
                    <td width="217"><font size="1" face="verdana, geneva, arial">You may not change your user name</font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial" color="#FF0000">Password</font></td>
                    <td colspan="2" width="373">
                    <!--webbot bot="Validation" s-display-name="Password" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input type="password" name="txtPassword" size="20" maxlength="20" value="<%=mrs("MemberPassword")%>">--
                      <em>keep this private!</em></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial" color="#FF0000">Re-enter
                      Password</font></td>
                    <td colspan="2" width="373">
                    <!--webbot bot="Validation" s-display-name="Password Verification" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input type="password" name="txtPasswordVerify" size="20" maxlength="20" value="<%=mrs("MemberPassword")%>">--
                      <em>for verification</em></td>
                  </tr>
                  <tr bgcolor="#EEEEEE">
                    <td colspan="3" align="CENTER" width="432">
                      <p align="justify"><b><font face="Arial" size="1"><br>
                      </font></b><font size="1" face="verdana, geneva, arial"><b>Make sure your Email address is correct.</b> You must respond to a confirmation
                      Email in order for your account to be updated.</font></p>
                      <p><input type="SUBMIT" value="Submit"></p>
                      <p><font face="geneva,arial" size="-2">By clicking on the above &quot;Submit&quot; button, you are agreeing to our Terms of Use.<br>
                      To get information about the conditions of using ACE, Inc. services,<br>
                      read the ACE, Inc. <a href="../UnderConstruction.htm">Terms of Service Agreement</a>.</font></p>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
  <input type="hidden" name="txtState" value="0">
</form>
<p>&nbsp;</p>
&nbsp;

<%
  mrs.close
  Set mrs=Nothing
  Set oDataConn=Nothing
  
  Else '<--- Submit Changes
%>
<p align="center"><font color="#808000">Your request is being processed - Please wait.</font></p>
<%
  '--- Email the form to the admin ---
  dim oFormMailer

  Set oFormMailer = New clsFormMailer

  oFormMailer.sHeader = "Form: ACE Member Edit - Member ID=" & mUserID
  oFormMailer.sFooter = "Form mailer created by PUPPO ENGINEERING, 2001"

  oFormMailer.sFrom = "ACEMemberSignup@ace-it.com" '<--- no spaces
  oFormMailer.sTo = "aceitweb@comcast.net" '"bob@puppozone.com" '"bpuppo@voicenet.com"
  oFormMailer.sSubject = "ACE Member Edit"
  'oFormMailer.sImportance = 2 'high

  oFormMailer.BuildBody
  'oFormMailer.Send
%>
<p align="center"><b>Processing Complete.</b> You will receive a confirmation email shortly.</p>
<%
  End If
%>
<p>&nbsp;</p>
&nbsp;
<p>&nbsp;</p>
&nbsp; <!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<hr>
<p><font size="1" color="#666633">Unless otherwise noted all content �
<a href="http://www.ace-it.com">AMERICAN COMPUTER ESTIMATING, INC.</a><br>
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

<%
  Private Sub CleanUp()
    On Error Resume Next
    mrs.Close
    Set mrs = Nothing
    Set oDataConn = Nothing
  End Sub
%>
<% Option Explicit %>

<!-- #INCLUDE Virtual="\clsFormMailer.asp" -->

<!-- #INCLUDE Virtual="\ACE\inc\clsACEDataConn.asp" -->

<%
'--------------------------------------------------------------
' Send Email
'--------------------------------------------------------------
%>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
<%
SUB sendmail(sFrom, sTo, sCC, sSubject, sTextBody, sHTMLBody, sURL, sAttach1, sAttach2, sAttach3, sAttach4, sAttach5)
  Dim objCDO
  Dim iConf
  Dim Flds

  Const cdoSendUsingPort = 2

  Set objCDO = Server.CreateObject("CDO.Message")
  Set iConf = Server.CreateObject("CDO.Configuration")

  Set Flds = iConf.Fields
  With Flds
    .Item(cdoSendUsingMethod) = cdoSendUsingPort
    .Item(cdoSMTPServer) = "mail-fwd"
    .Item(cdoSMTPServerPort) = 25
    .Item(cdoSMTPconnectiontimeout) = 10
    .Update
  End With

  Set objCDO.Configuration = iConf
  
  With objCDO
  
    .From = sFrom
    .To = sTo
    .CC = sCC
    .Subject = sSubject
    .TextBody = sTextBody
    If Len(sHTMLBody) <> 0 Then .HTMLBody = sHTMLBody
    
    If Len(sURL) <> 0 Then 
      'iMsg.CreateMHTMLBody "http://server.example.com", cdoSuppressAll, "domain\username", "password"
      .CreateMHTMLBody sURL, cdoSuppressNone, "", ""
    End If
  
    '.AddAttachment "http://server.example.com/picture.gif"
    '.AddAttachment "file://d:/temp/test.doc"
    '.AddAttachment "C:\files\another.doc"

    If Len(sAttach1) > 0 then .AddAttachment sAttach1
    If Len(sAttach2) > 0 then .AddAttachment sAttach2
    If Len(sAttach3) > 0 then .AddAttachment sAttach3
    If Len(sAttach4) > 0 then .AddAttachment sAttach4
    If Len(sAttach5) > 0 then .AddAttachment sAttach5
     
    .Send
    
  End With

  'Cleanup
  Set objCDO = Nothing
  Set iConf = Nothing
  Set Flds = Nothing
END SUB
%>

<%

  Dim Msg
  Dim mUserName, mPassword, mRememberMe
  Dim mrs
  Dim bResult
  Dim Result
  Dim mState
  Dim oDataConn

  On Error Resume Next
  mUserName=request.form("txtUserName")
  mPassword=request.form("txtPassword")
  mRememberMe = request.form("chkRememberMe")
  If mRememberMe = "" Then mRememberMe = "0" Else mRememberMe = "1"  
  mState=Request.QueryString("txtState")(1)
  On Error Goto 0
  'If mState = "" Then mState = 0

  IF mUserName <> "" Then '--- user logging in ? ...
  
    '--- handle cookies ---
    If mRememberMe = "1" Then
      '--- Set Remember Me Cookies ---
      Response.Cookies("chkRememberMe") = mRememberMe
      Response.Cookies("txtUserName") = mUserName
      Response.Cookies("txtPassword") = mPassword 
    Else
      Response.Cookies("chkRememberMe") = mRememberMe
      Response.Cookies("txtUserName") = ""
      Response.Cookies("txtPassword") = ""    
    End If
    
    Response.Cookies("chkRememberMe").expires = Date() + 90
    Response.Cookies("txtUserName").expires = Date() + 90
    Response.Cookies("txtPassword").expires = Date() + 90

    Response.Cookies("chkRememberMe").Path = "/"            
    Response.Cookies("txtUserName").Path = "/"            
    Response.Cookies("txtPassword").Path = "/"            
    'Response.Cookies("name").Domain = ".ace-it.com"
    '--- end cookies ---
        
    Set oDataConn = New clsACEDataConn

    Set mrs = oDataConn.ACEMemberGet(0, mUserName,"")

    If Not mrs.eof Then
      If UCase(mrs("MemberPassword")) = UCase(mPassword) Then 'OK we validated the user!!
        If mrs("MemberAccountEnabled") = True Then
          Session("MemberID")=mrs("MemberIDPK")
          Session("MemberUserName") = mUserName
          oDataConn.ACEMemberActivityInsert (mrs.fields("MemberIDPK")), Activity_Login	'Log activity
          Session("MemberName")=mrs("MemberFirstName") & " " & mrs("MemberLastName")
          Session("MemberTypeID")=mrs("MemberTypeID")
          Session("MemberSecurityLevel")=mrs("MemberSecurityLevel")
          'Session("InsCompanyID")=mrs("InsCompanyID") 'Name of Ins Co. as listed in ACE database
          Session("RepName")=mrs("ACERepName") 'Full rep name
          Session("InsCompanyName")=mrs("InsCoName") 'Name of company as supplied by member
          Session("ACEInsCoOfficeID")=mrs("ACEInsCoOfficeID") 'Actual ACE ID A51B
          Session("ACEInsCoOfficeName")=mrs("ACEInsCoOfficeName") '
          Session("MemberGender")=mrs("MemberGender")
          Session("ShopNetworkAccess")=mrs("ShopNetworkAccess")
          Session("ShopDemoAccess")=mrs("ShopDemoAccess")


		  '--- The following are for submission forms ---
		  Session("UserInsCoOfficeID_AutoForm") = mrs("Auto_OfficeID")
		  Session("UserInsCoOfficeID_SubrogationForm") = mrs("Subro_OfficeID")
		  Session("UserInsCoOfficeID_HeavySpecialtyForm") = mrs("HeavySpecialty_OfficeID")
		  Session("UserInsCoOfficeID_PropertyForm") = mrs("Property_OfficeID")
		  Session("UserInsCoOfficeID_AutoTotalLossForm") = mrs("AutoTotalLoss_OfficeID")
		  Session("UserInsCoOfficeID_PhotoQuoteForm") = mrs("PQ_OfficeID")
  		  
		  
          '--- The following are for the message system ---
          Session("logged") = "true"
          Session("uname") = mUserName
			'---

          'mIDEncrypt="AC" & HEX(mrs("MemberIDPK"))
          'response.redirect "../cgi-bin/CGIRepLogin.exe?ID="& mIDEncrypt
          'response.redirect "../cgi-bin/CGIRepLogin.exe?ID="& mrs("MemberIDPK")
          CleanUp
          'response.redirect "../RepMenu2.asp"
          
          'If Instr(Request.ServerVariables("HTTP_REFERER"), "ClaimPortalUI/") Then
          '  Response.Redirect Request.ServerVariables("HTTP_REFERER")
          'Else
          '  Response.Redirect "../ClaimPortalUI/frmMain.asp"
          'End If
          
          If (Session("RequestedURL") & "") <> "" Then
            Result = Session("RequestedURL")
            Session("RequestedURL") = ""
            Response.Redirect Result
          Else
            Response.Redirect "../ClaimPortalUI/frmMain.asp"
          End If
          
        ElseIf mrs("MemberConfEmailRespRecvd") = False Then
          Msg=3  'Confirmation Email not received
        Else
          Msg=4  'Account disabled
        End If
      Else
        Msg=2 'Incorrect Password
      End If
    ELSE
      Msg=1 'Username not found
    END IF
    CleanUp
  END IF  
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Content-Language" content="en-us">
<meta name="Author" content="Bob Puppo">
<meta name="Copyright" content="(C) Bob Puppo, 1999-2008, All rights reserved">
<title>ACE Logon</title>
<script language="JavaScript" src="../JavaScript/RememberMe.js"></script>
<SCRIPT LANGUAGE="JavaScript" src="../../javascript/FormsPlaceFocus.js"></script>
<SCRIPT LANGUAGE="JavaScript">
function ValidatePassword() {
    If (document.FrontPage_Form2.txtPasswordVerify.value != document.FrontPage_Form2.txtPassword.value)
    {
      Alert("Your passwords do not match please re-enter them.");
    }
  }
</Script>
<SCRIPT LANGUAGE="JavaScript" src="../javascript/jsCookie.js"></script>
<SCRIPT LANGUAGE="JavaScript">
  setCookie("test", "OK");
  //if (getCookie("test") != "OK") window.location = "NoCookies.htm";
  if (getCookie("test") != "OK") window.location.replace("NoCookies.htm");
</SCRIPT>

<meta name="Microsoft Border" content="none">
</head>

<body onload="placeFocus()">

<p style="margin-top: 0; margin-bottom: 0" align="center"><img border="0" src="../Images/ace_logo.gif"><br>
<img border="0" src="../Images/LineGradBlue.jpg"></p>

<%
  If mState=0 Then
    Result = Session("RequestedURL")
    session.abandon '--- doesn't take effect until page ends
    Session("RequestedURL") = Result
%>
&nbsp;
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.txtUserName.value == "")
  {
    alert("Please enter a value for the \"User name\" field.");
    theForm.txtUserName.focus();
    return (false);
  }

  if (theForm.txtUserName.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"User name\" field.");
    theForm.txtUserName.focus();
    return (false);
  }

  if (theForm.txtUserName.value.length > 30)
  {
    alert("Please enter at most 30 characters in the \"User name\" field.");
    theForm.txtUserName.focus();
    return (false);
  }

  if (theForm.txtPassword.value == "")
  {
    alert("Please enter a value for the \"Password \" field.");
    theForm.txtPassword.focus();
    return (false);
  }

  if (theForm.txtPassword.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Password \" field.");
    theForm.txtPassword.focus();
    return (false);
  }

  if (theForm.txtPassword.value.length > 30)
  {
    alert("Please enter at most 30 characters in the \"Password \" field.");
    theForm.txtPassword.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name="FrontPage_Form1" method="POST" action="ACELogin.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
  <div align="center">
    <center>
          <table cellpadding="1" cellspacing="0" border="0" width="250" bgcolor="#CCCCCC">
            <tr>
              <td>
                <table cellpadding="1" cellspacing="0" border="0" width="100%" bgcolor="ffffcc">
                  <tr>
                    <td bgcolor="ffcc00">&nbsp;<font face="arial,geneva"><b>New User</b></font></td>
                  </tr>
                  <tr>
                    <td>
                      <table width="100%" cellpadding="3" cellspacing="3" border="0">
                        <tr>
                          <!--<td align="center"><font face="helvetica,geneva"><a href="ACELogin.asp?txtState=1"><b>Sign me up!</b></a></font></td> -->
                          <td align="center"><font face="helvetica,geneva"><a href="/enroll.aspx"><b>Sign me up!</b></a></font></td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                  <tr>
                    <td bgcolor="ffcc00">&nbsp;<font face="arial,geneva"><b>ACE Member</b></font></td>
                  </tr>
                  <tr>
                    <td>
                      <table width="100%" cellpadding="3" cellspacing="0" border="0">
                        <tr>
                          <td>
                            <% If 0 Then '(Session("RequestedURL") & "") <> "" Then %>
                            <p style="margin-top: 0; margin-bottom: 0"><font size="2"><b><font #800000"" color="#800000">Your Session has expired, this occurs 
                            automatically after 20 minutes of inactivity.</font> Please log in again. </b></font></p>
                            <% End If %>    
                            <%
    Select Case Msg
      Case 1
                            %>
                            <p style="margin-top: 0; margin-bottom: 0"><font size="2"><b>The<font color="#800000"> User Name</font></b><font color="#800000"> </font><b>you
                            entered was not found please reenter it</b></font></p>
                            <%
      Case 2
                            %>
                            <p style="margin-top: 0; margin-bottom: 0"><font size="2"><b>The<font color #800000""> </font><font #800000"" color="#800000">Password</font></b><font color #800000"">
                            </font><b>you entered was incorrect please reenter it</b></font></p>
                            <%
      Case 3
                            %>
                            <p style="margin-top: 0; margin-bottom: 0"><font size="2"><b><font #800000"" color="#800000">Your information is on file but your
                            account is not activated.</font></b><font color #800000""> </font><b>You must click the link in your confirmation Email to enable
                            your account. If it has been more than 48 hours since you signed up and you have not received your confirmation Email please contact
                            <a href="mailto:support@ace-it.com">Support@ace-it.com</a></b></font></p>
                            <%
      Case 4
                            %>
                            <p style="margin-top: 0; margin-bottom: 0"><font size="2"><b><font #800000"" color="#800000">Your information is on file but your
                            account is disabled</font><font #800000""><font color="#800000">.</font></font> Please contact <a href="mailto:support@ace-it.com">Support@ace-it.com</a>
                            if you have any questions.</b></font></p>
                            <%
    End Select
                            %>
                          </td>
                        </tr>
                      </table>
                      <table width="100%" cellpadding="3" cellspacing="0" border="0">
                        <tr>
                          <td align="right" valign="top"><font face="geneva,arial" size="-1"><b>Member Name:</b></font></td>
                          <td valign="top">
                          <!--webbot bot="Validation" s-display-name="User name" s-data-type="String" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="30" --><input type="text" name="txtUserName" size="15" tabindex="1" maxlength="30"></td>
                        </tr>
                        <tr>
                          <td align="right" valign="center"><font face="geneva,arial" size="-1"><b>Password:</b></font></td>
                          <td valign="center">
                          <!--webbot bot="Validation" s-display-name="Password " s-data-type="String" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="30" --><input type="password" name="txtPassword" size="15" tabindex="2" maxlength="30"></td>
                        </tr>
                        <tr>
                          <td colspan="2" valign="top" align="right"><input type="checkbox" name="chkRememberMe" value="ON"><font size="2">Remember me&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
                          </font>&nbsp;</td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                          <td><input border="0" src="../images/btn_login_blue.gif" name="Login" type="image" tabindex="3"></td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                  <tr bgcolor="ffffcc">
                    <td>&nbsp;</td>
                  </tr>
                  <tr bgcolor="ffcc00">
                    <td bgcolor="ffcc00">&nbsp;<font face="arial,geneva"><b>Help</b></font></td>
                  </tr>
                  <tr bgcolor="ffffcc">
                    <td>
                      <table width="100%" cellpadding="3" cellspacing="0" border="0">
                        <tr>
                          <td><font face="geneva,arial" size="-1"><a href="/contact.aspx">Get help with sign in</a><br>
                            <a href="/contact.aspx">Member Services</a></font><br>
                            &nbsp;</td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
    </table>
    </center>
  </div>
  <input type="hidden" name="txtState" value="0">
</form>
<script>
  RetrievMe(document.forms[0]); //retriev login settings if user wanted them remembered
</script>


  <%
  ElseIf mState=1 Then  'Signup Form
  %> <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form2_Validator(theForm)
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

  if (theForm.txtCoOffice.value == "")
  {
    alert("Please enter a value for the \"Company Name\" field.");
    theForm.txtCoOffice.focus();
    return (false);
  }

  if (theForm.txtCoOffice.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Company Name\" field.");
    theForm.txtCoOffice.focus();
    return (false);
  }

  if (theForm.txtCoOffice.value.length > 30)
  {
    alert("Please enter at most 30 characters in the \"Company Name\" field.");
    theForm.txtCoOffice.focus();
    return (false);
  }

  if (theForm.txtCoCity.value == "")
  {
    alert("Please enter a value for the \"Company City \" field.");
    theForm.txtCoCity.focus();
    return (false);
  }

  if (theForm.txtCoCity.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Company City \" field.");
    theForm.txtCoCity.focus();
    return (false);
  }

  if (theForm.cmbCoState.selectedIndex < 0)
  {
    alert("Please select one of the \"Company State\" options.");
    theForm.cmbCoState.focus();
    return (false);
  }

  if (theForm.cmbCoState.selectedIndex == 0)
  {
    alert("The first \"Company State\" option is not a valid selection.  Please choose one of the other options.");
    theForm.cmbCoState.focus();
    return (false);
  }

  if (theForm.txtCoPhoneAc.value == "")
  {
    alert("Please enter a value for the \"Company Phone Area Code\" field.");
    theForm.txtCoPhoneAc.focus();
    return (false);
  }

  if (theForm.txtCoPhoneAc.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Company Phone Area Code\" field.");
    theForm.txtCoPhoneAc.focus();
    return (false);
  }

  if (theForm.txtCoPhoneAc.value.length > 3)
  {
    alert("Please enter at most 3 characters in the \"Company Phone Area Code\" field.");
    theForm.txtCoPhoneAc.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.txtCoPhoneAc.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Company Phone Area Code\" field.");
    theForm.txtCoPhoneAc.focus();
    return (false);
  }

  if (theForm.txtCoPhoneExch.value == "")
  {
    alert("Please enter a value for the \"Company Phone Exchange\" field.");
    theForm.txtCoPhoneExch.focus();
    return (false);
  }

  if (theForm.txtCoPhoneExch.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Company Phone Exchange\" field.");
    theForm.txtCoPhoneExch.focus();
    return (false);
  }

  if (theForm.txtCoPhoneExch.value.length > 3)
  {
    alert("Please enter at most 3 characters in the \"Company Phone Exchange\" field.");
    theForm.txtCoPhoneExch.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.txtCoPhoneExch.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Company Phone Exchange\" field.");
    theForm.txtCoPhoneExch.focus();
    return (false);
  }

  if (theForm.txtCoPhoneL4.value == "")
  {
    alert("Please enter a value for the \"Company Phone Last 4 digits\" field.");
    theForm.txtCoPhoneL4.focus();
    return (false);
  }

  if (theForm.txtCoPhoneL4.value.length < 4)
  {
    alert("Please enter at least 4 characters in the \"Company Phone Last 4 digits\" field.");
    theForm.txtCoPhoneL4.focus();
    return (false);
  }

  if (theForm.txtCoPhoneL4.value.length > 4)
  {
    alert("Please enter at most 4 characters in the \"Company Phone Last 4 digits\" field.");
    theForm.txtCoPhoneL4.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.txtCoPhoneL4.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Company Phone Last 4 digits\" field.");
    theForm.txtCoPhoneL4.focus();
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
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name="FrontPage_Form2" method="post" action="ACELogin.asp?txtState=2" onsubmit="return FrontPage_Form2_Validator(this)" language="JavaScript">
  <input type="hidden" name="site" value="7"><input type="hidden" name="action" value="submit">
  <div align="center">
    <table border="0" cellpadding="1" bgcolor="#356C91" width="425" cellspacing="1">
      <tr>
        <td width="562" align="center"><b><font face="arial,geneva" color="#FFFFFF">Member Signup</font><font face="Arial" size="3" color="#FFFFFF"><br>
          </font></b><font face="Arial" size="1" color="#C0C0C0">(*Red label names denote required fields)</font></td>
      </tr>
      <tr>
        <td width="562">
          <table border="0" cellpadding="15" cellspacing="0" bgcolor="#FFFFFF">
            <tr>
              <td width="530">
                <table border="0" cellpadding="1" width="447" bgcolor="#FFFFCC">
                  <tr>
                    <td bgcolor="#356C91" align="center" valign="middle" colspan="5" width="432"><b><font size="1" face="verdana, geneva, arial" color="#FFFF00">Company
                      Information</font></b></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial" color="#FF0000">Company</font></td>
                    <td colspan="4" width="373">
                    <!--webbot bot="Validation" s-display-name="Company Name" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="30" --><input type="TEXT" name="txtCoName" size="40" maxlength="30"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial" color="#FF0000">Office</font></td>
                    <td colspan="4" width="373">
                    <!--webbot bot="Validation" s-display-name="Company Name" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="30" --><input name="txtCoOffice" size="40" maxlength="30"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font>
                      Street Address</font></td>
                    <td colspan="4" width="373"><input type="TEXT" name="txtCoAddress" size="40"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> <font color="#FF0000">City</font></font></td>
                    <td colspan="4" width="373">
                    <!--webbot bot="Validation" s-display-name="Company City " s-data-type="String" b-value-required="TRUE" i-minimum-length="3" --><input name="txtCoCity" size="40"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> <font color="#FF0000">State</font></font></td>
                    <td width="113">
                    <!--webbot bot="Validation" s-display-name="Company State" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select name="cmbCoState" size="1">
                        <option>
                        <option>AB
                        <option>AK
                        <option>AL
                        <option>AR
                        <option>AZ
                        <option>BC
                        <option>CA
                        <option>CO
                        <option>CT
                        <option>DC
                        <option>DE
                        <option>FL
                        <option>GA
                        <option>HI
                        <option>IA
                        <option>ID
                        <option>IL
                        <option>IN
                        <option>KS
                        <option>KY
                        <option>LA
                        <option>MA
                        <option>MB
                        <option>MD
                        <option>ME
                        <option>MI
                        <option>MN
                        <option>MO
                        <option>MS
                        <option>MT
                        <option>NB
                        <option>NC
                        <option>ND
                        <option>NE
                        <option>NF
                        <option>NH
                        <option>NJ
                        <option>NM
                        <option>NS
                        <option>NT
                        <option>NV
                        <option>NY
                        <option>OH
                        <option>OK
                        <option>ON
                        <option>OR
                        <option>PA
                        <option>PE
                        <option>PQ
                        <option>RI
                        <option>SC
                        <option>SD
                        <option>SK
                        <option>TN
                        <option>TX
                        <option>UT
                        <option>VA
                        <option>VT
                        <option>WA
                        <option>WI
                        <option>WV
                        <option>WY
                        <option>YU
                        <option>OTHER
                      </select></td>
                    <td align="right" bgcolor="#FFFFCC" colspan="2" width="77"><font size="1" face="verdana, geneva, arial">Zip Code</font></td>
                    <td width="171"><input type="TEXT" name="txtCoZip" size="14"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> <font color="#FF0000">Phone</font></font></td>
                    <td colspan="4" width="373"><b>(</b><!--webbot bot="Validation" s-display-name="Company Phone Area Code" s-data-type="String" b-allow-digits="TRUE" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="3" --><input type="TEXT" name="txtCoPhoneAc" size="3" maxlength="3"><b>)</b>&nbsp;<!--webbot bot="Validation" s-display-name="Company Phone Exchange" s-data-type="String" b-allow-digits="TRUE" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="3" --><input type="TEXT" name="txtCoPhoneExch" size="3" maxlength="3">&nbsp;<b>-</b>&nbsp;<!--webbot bot="Validation" s-display-name="Company Phone Last 4 digits" s-data-type="String" b-allow-digits="TRUE" b-value-required="TRUE" i-minimum-length="4" i-maximum-length="4" --><input type="TEXT" name="txtCoPhoneL4" size="4" maxlength="4">&nbsp;</td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Fax</font></td>
                    <td colspan="4" width="373"><b>(</b><input type="TEXT" name="txtCoFaxAc" size="3" maxlength="3"><b>)</b>&nbsp;<input type="TEXT" name="txtCoFaxExch" size="3" maxlength="3">&nbsp;<b>-</b>&nbsp;<input type="TEXT" name="txtCoFaxL4" size="4" maxlength="4"></td>
                  </tr>
                  <tr>
                    <td colspan="5" align="center" bgcolor="#356C91" width="432"><b><font size="1" face="verdana, geneva, arial" color="#FFFF00">Personal
                      Information</font></b></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" valign="middle" width="75" align="right"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> <font color="#FF0000">First
                      Name</font></font></td>
                    <td colspan="2" width="187">
                    <!--webbot bot="Validation" s-display-name="First Name" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input name="txtFirstName" size="20" maxlength="20">&nbsp;&nbsp;</td>
                    <td width="60" align="right"><font color="#FF0000" size="1" face="verdana, geneva, arial">Last<br>
                      Name</font></td>
                    <td width="125">
                    <!--webbot bot="Validation" s-display-name="Last Name" s-data-type="String" b-value-required="TRUE" i-minimum-length="2" i-maximum-length="20" --><input name="txtLastName" size="20" maxlength="20"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font>&nbsp;
                      Work Phone</font></td>
                    <td colspan="4" width="373"><b>(</b><input name="txtPhoneAc" size="3" maxlength="3"><b>)</b>&nbsp;<input type="TEXT" name="txtPhoneExch" size="3" maxlength="3">&nbsp;<b>-</b>&nbsp;<input type="TEXT" name="txtPhoneL4" size="4" maxlength="4">&nbsp;<font size="1">Extension</font>&nbsp;<input type="TEXT" name="phoneextn" size="4" maxlength="4"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> <font color="#FF0000">E-mail
                      Address</font></font></td>
                    <td colspan="4" width="373">
                    <!--webbot bot="Validation" s-display-name="E-mail Address" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" --><input name="txtEmail" size="40"></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial">Gender</font></td>
                    <td colspan="4" width="373"><input type="radio" name="grpGender" value="1"><font face="geneva,arial" size="-1"> Male&nbsp; </font><input type="radio" name="grpGender" value="2"><font face="geneva,arial" size="-1">
                      Female</font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#356C91" align="center" valign="middle" colspan="5" width="432"><font size="1" face="verdana, geneva, arial" color="#FFFF00"><b>Choose
                      Your Sign In Information</b></font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font>&nbsp;
                      <font color="#FF0000">User Name</font></font></td>
                    <td colspan="2" width="150">
                    <!--webbot bot="Validation" s-display-name="User name" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input name="txtUserName" size="20" maxlength="20"></td>
                    <td colspan="2" width="217"><font size="1" face="verdana, geneva, arial">Your <i>User Name</i> must be between <b>3 and 20 alphanumerical</b>
                      characters in length.</font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial" color="#FF0000">Password</font></td>
                    <td colspan="4" width="373">
                    <!--webbot bot="Validation" s-display-name="Password" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input type="password" name="txtPassword" size="20" maxlength="20">-- <em>keep this private!</em></td>
                  </tr>
                  <tr>
                    <td bgcolor="#FFFFCC" align="right" valign="middle" width="75"><font size="1" face="verdana, geneva, arial" color="#FF0000">Re-enter
                      Password</font></td>
                    <td colspan="4" width="373">
                    <!--webbot bot="Validation" s-display-name="Password Verification" s-data-type="String" b-value-required="TRUE" i-minimum-length="3" i-maximum-length="20" --><input type="password" name="txtPasswordVerify" size="20" maxlength="20">-- <em>for
                      verification</em></td>
                  </tr>
                  <tr bgcolor="#EEEEEE">
                    <td colspan="5" align="CENTER" width="432">
                      <p align="justify"><b><font face="Arial" size="1"><br>
                      </font></b><font size="1" face="verdana, geneva, arial"><b>Make sure your Email address is correct.</b> You will be notified via email
                      when your account is setup. Your account can only be activated through a link in this email. We validate each user manually to insure
                      security. Your account may take up to 48 hours to setup. Thank You.</font></p>
                      <center>
                      <p><input type="SUBMIT" value="Submit"></p>
                      <p><font face="geneva,arial" size="-2">By clicking on the above &quot;Submit&quot; button, you are agreeing to our Terms of Use.<br>
                      To get information about the conditions of using ACE, Inc. services,<br>
                      read the ACE, Inc. <a href="../UnderConstruction.htm">Terms of Service Agreement</a>.</font></p>
                      </center></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
</form>
  <%
  ElseIf mState=2 Then '--- Email the form
  %>
<p align="center"><font color="#808000">Your request is being processed. Please wait</font>.</p>
<p align="center">&nbsp;</p>
  <%
  If Response.Buffer=True Then Response.Flush

  Set oDataConn = New clsACEDataConn

  '--- Check to see if the member login name already exists ---
  Set mrs = oDataConn.ACEMemberGet(0, mUserName,"")

  If Not mrs.eof Then 'If user name already exists
  %>
<div align="center">
  <center>
  <table border="1" cellspacing="0" width="90%" bordercolor="#800000">
    <tr>
      <td width="100%" align="center" bgcolor="#800000" colspan="2"><b><font color="#FFFFFF">There was a problem with your submission</font></b></td>
    </tr>
    <tr>
      <td width="70%" align="center" rowspan="2"><b>The User Name you selected is already in use</b>&nbsp;
        <p>Please click your browsers back button and select another user name.<br>
        <font size="1">Note: you must also reenter your password</font></td>
      <td width="25%" align="center" bgcolor="#356C91"><b><font color="#000000">Tip</font></b></td>
    </tr>
    <tr>
      <td width="30%" align="center"><font color="#000080"><b>&nbsp;</b></font>Your email address is a good candidate for a user name as it's easy to remember
        and guaranteed to be unique. Example:<br>
        JohnDoe@InsuranceCo.com.</td>
    </tr>
  </table>
  </center>
</div>
<%
    CleanUp
    Response.End
  End If

'  On error Goto SubmitErr

  '--- Enter Form Data into database ---
  oDataConn.ACEMemberInsert(mrs)

  '--- A97 Autonumber Hack ---
  mrs.close
  Set mrs=Nothing
  Set mrs = oDataConn.ACEMemberGet(0, mUserName,"")
  '--- End A97 Autonumber Hack ---

  mrs.MoveFirst 'I don't know why I'm moving first??



  '--- Email the form to the admin ---
  dim oFormMailer

  Set oFormMailer = New clsFormMailer

  oFormMailer.sHeader = "Form: ACE Member Signup - New Member ID=" & mrs("MemberIDPK")
  oFormMailer.sFooter = "Form mailer created by PUPPO ENGINEERING, 2001 - 2008"

  oFormMailer.sFrom = "ACEMemberSignup@ace-it.com" '<--- no spaces
  oFormMailer.sTo = "aceitweb@comcast.net" '"support@ace-it.com" '"bob@puppozone.com" '"bpuppo@voicenet.com"
  oFormMailer.sSubject = "ACE Member Signup"
  'oFormMailer.sImportance = 2 'high

  oFormMailer.BuildBody

  '--- >>>>> Save Body to database <<<<< ---
  mrs("MemberEmailBody") = oFormMailer.sHeader & VBCrLf & oFormMailer.sBody & VBCrLf & oFormMailer.sFooter
  oDataConn.ACEMemberUpdate mrs
  mrs.movefirst

  on error resume next
  
  sendmail "ACEMemberSignup@ace-it.com", "aceitweb@comcast.net", "", "ACE Member Signup", _
            oFormMailer.sHeader & VBCrLf & oFormMailer.sBody & VBCrLf & oFormMailer.sFooter, _
            "", "", "", "", "", "", ""

  'oFormMailer.Send

If False Then
  '--- Send Email to applicant ---
  Dim EncodedMemberID

  EncodedMemberID = (int(rnd*70000)+11113) & "~" & mrs("MemberIDPK") & "_" & (int(rnd*70000)+10007)


  oFormMailer.sHeader = "ACE Member Signup"
  oFormMailer.sFooter = "Account Validator V1.01 created by PUPPO ENGINEERING www.PUPPOZone.com"

  oFormMailer.sFrom = "ACEMemberSignup@ace-it.com" '<--- no spaces
'  oFormMailer.sTo = request.form("txtEmail")
  oFormMailer.sTo = "aceitweb@comcast.net" '"bob@puppozone.com"  '<---- Send to admin
  oFormMailer.sSubject = "ACE Member Signup"
  'oFormMailer.sImportance = 2 'high
  oFormMailer.sBody = "Please click the link below to activate your Free membership with ACE Internet services. " & vbcrlf & vbcrlf
  oFormMailer.sBody = oFormMailer.sBody & "http://ACE.PUPPOZone.com/Login/MemberActivation.asp?ID=" & EncodedMemberID & vbcrlf & vbcrlf
  oFormMailer.sBody = oFormMailer.sBody & "Please to not relpy to this email just click the link above to activate your membership."
  oFormMailer.Send
End IF

  oDataConn.ACEMemberActivityInsert (mrs.fields("MemberIDPK")), Activity_WelcomeEmail	'Log activity


'SubmitErr:
  Cleanup
  Set oFormMailer = Nothing
%>
<div align="center">
  <center>
  <table border="0" cellspacing="1" width="56%">
    <tr>
      <td width="100%" align="center">
        <p align="center"><font color="#008000"><b>Processing Complete.</b></font><br>
        Your information has been submitted.</p>
        <p align="center">As soon as your information is verified you will receive an activation Email.</p>
        <p align="center"><a href="../ACE-IT/index.aspx">Home</a></p>
        <p align="center">&nbsp;</p>
        <p align="center">Thank You.</p>
        <p>&nbsp;</td>
    </tr>
  </table>
  </center>
</div>
        <%
  End IF
        %>
<hr>
<font size="1" color="#666633">Unless otherwise noted all content <font face="arial, helvetica" color="#666666" size="1">� American Computer Estimating, Inc.</font></font><font face="arial, helvetica" color="#666666" size="1"><br>
</font><font size="1" color="#666633">Unless otherwise noted all program and web code <font face="arial, helvetica" color="#666666" size="1">�</font></font><font face="arial, helvetica" color="#666666" size="1">
PUPPO ENGINEERING</font><font size="1" color="#666633"><br>
Web Site Programming and Design By</font><font face="Arial" size="1"><font color="#666633"> </font><b><a href="http://www.puppozone.com/" target="_top">PUPPO
ENGINEERING</a></b></font><font color="#666633"><b><font size="1">. </font></b><font size="1">Send</font></font><font size="1" color="#666633"> mail to <a href="mailto:WebMaster@PUPPOZone.com">
<!--webbot
bot="Substitution" s-variable="CompanyWebmaster" startspan -->Webmaster@PuppoZone.com<!--webbot bot="Substitution" i-checksum="314" endspan --></a> with questions or comments about this web site.</font>

  <%
  Private Sub CleanUp()
    On Error Resume Next
    mrs.Close
    Set mrs = Nothing
    Set oDataConn = Nothing
  End Sub
  %>

</body>
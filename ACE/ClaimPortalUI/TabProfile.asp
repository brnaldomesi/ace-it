<style>
<!--
  td.styLabelCell        { font-family: verdana, geneva, arial; font-size: 8pt; font-weight:bold; /*background-color:#F9F9E3 #FFFFCC*/ }
-->
</style>
<script language="JavaScript">
function profileValidate(frm) {
  if (frm.txtPassword.value == frm.txtPasswordVerify.value) {
  	return true;
  } else {
    alert('Your paswords do not match. Please retype them.');
    return false;
  }
}
</script>

<% If mActiveTab = "" Then response.redirect "http://www.staging-portal.ace-it.com/" %> 
<%
  Dim CoName, SubmittedBy, OfficeID, OfficeName
  Dim PhoneAc, PhoneExchg, PhoneL4, PhoneEx
  Dim FaxAc, FaxExchg, FaxL4
  Dim sSex 
  Dim EmailAddr
  Dim oData, mrs
  Dim bResult

  Dim oFormMailer

  Dim oUpload, oFile, oItem
  
  Set oData = New clsACEDataConn
  Set mrs = oData.ACEMemberGet(mUserID, "","")

  If Not mrs.eof Then
    EmailAddr = mrs("MemberEmail") & ""
    PhoneEx = mrs("MemberPhone") & ""
    PhoneAc = Left(PhoneEx, 3)
    PhoneExchg = Mid(PhoneEx,4,3)
    PhoneL4 = Mid(PhoneEx,7)
    PhoneEx = mrs("MemberPhoneEx") & ""
    FaxL4 = mrs("MemberFax") & ""
    FaxAc = Left(FaxL4, 3)
    FaxExchg = Mid(FaxL4, 4, 3)
    FaxL4 = Mid(FaxL4, 7)
    If mUserGender = 1 Then
      sSex = "Male"
    Else
      sSex = "Female"
    End If
  End If
  mrs.Close
  
  'OfficeID = mUserACEInsCoOfficeID
  
  bResult = 0
  'If Request.Form("txtState") <> "1" Then
  If Request.Querystring("txtState") = "1" Then

    bResult = ((Request.Form("txtCompany") <> mUserInsCoName) OR bResult)
    bResult = ((Request.Form("txtOfficeName") <> mUserACEInsCoOfficeName) OR bResult)
    bResult = ((Request.Form("txtName") <> mRepName) OR bResult)
    bResult = ((Request.Form("txtPhoneAc") <> PhoneAc) OR bResult)
    bResult = ((Request.Form("txtPhoneExchg") <> PhoneExchg) OR bResult)
    bResult = ((Request.Form("txtPhoneL4") <> PhoneL4) OR bResult)
    bResult = ((Request.Form("txtPhoneEx") <> PhoneEx) OR bResult)
    bResult = ((Request.Form("txtFaxAc") <> FaxAc) OR bResult)
    bResult = ((Request.Form("txtFaxExchg") <> FaxExchg) OR bResult)
    bResult = ((Request.Form("txtFaxL4") <> FaxL4) OR bResult)
    bResult = ((Request.Form("txtEmail") <> EmailAddr) OR bResult)
    bResult = ((Request.Form("cmbSex") <> sSex) OR bResult)
    bResult = ((Request.Form("txtPassword") <> "") OR bResult)

	Dim sMsgChanged
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtCompany") <> mUserInsCoName), ", Company name", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtOfficeName") <> mUserACEInsCoOfficeName), ", Office name", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtName") <> mRepName), ", Rep name", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtPhoneAc") <> PhoneAc), ", Phone area code", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtPhoneExchg") <> PhoneExchg), ", Phone exchange", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtPhoneL4") <> PhoneL4), ", Phone", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtPhoneEx") <> PhoneEx), ", Phone extension", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtFaxAc") <> FaxAc), ", Fax area code", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtFaxExchg") <> FaxExchg), ", Fax exchange", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtFaxL4") <> FaxL4), ", Fax", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtEmail") <> EmailAddr), ", Email", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("cmbSex") <> sSex), ", Sex", "")
    sMsgChanged = sMsgChanged & IIF((Request.Form("txtPassword") <> ""), ", Password", "")
    
    If sMsgChanged <> "" Then sMsgChanged = "--> " & Mid(sMsgChanged, 3)

    If bResult <> 0 Then
		'--- Email ---
	  
		Dim oDataConn
		Dim rsClaim 
		Dim sSQL
		
		sSQL = "SELECT * FROM tblClaimsSubmitted WHERE ClaimIDPK = 0"
		
		Set oFormMailer = New clsFormMailer
		Dim sBody
		sBody = oFormMailer.fBuildBody
		set oFormMailer = Nothing

    
		Set oDataConn = New clsACEDataConn
		Set rsClaim = Server.CreateObject("ADODB.Recordset")
		rsClaim.Open sSQL, oDataConn.ConnStr, 3, adLockOptimistic     'adOpenDynamic 3, adLockOptimistic
		rsClaim.AddNew
		rsClaim("ClaimBody") = "ACE-IT Profile change request." & vbcrlf & _
      	  		 "--> Member Record ID = " & mUserID & " <--    --> User Name = " & mName & " <--" & vbcrlf & _
      	  		 "This member has requested the following changes:" & vbcrlf & _
      	  		 sMsgChanged & vbcrlf & vbcrlf & _
				 sBody
		rsClaim("EmailTo") = "aceitweb@comcast.net"
		rsClaim.Update  '--- Do update here to make sure we get record in case there's an issue with the attachments
		rsClaim.MoveFirst
		rsClaim.Update
		rsClaim.Close
	  
	  
	  
      'Set oFormMailer = New clsFormMailer
      'With oFormMailer
        '.sHeader = "ACE-IT Profile change request." & vbcrlf & _
      	  		 '"--> Member Record ID = " & mUserID & " <--    --> User Name = " & mName & " <--" & vbcrlf & _
      	  		 '"This member has requested the following changes:" & vbcrlf & _
      	  		 'sMsgChanged & vbcrlf
        '.sFooter = ""
        '.sFrom = "aceitweb@comcast.net" '"ACEPortal@staging-portal.ace-it.com" '<--- no spaces
        '.sTo = "aceitweb@comcast.net"
        '.sSubject = "ACE-IT Portal, Profile Change Request"
        '.BuildBody
        '.Send
      'End With
%>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center">Your request has been sent.</p>
<p align="center">You will be notified via email when your account has been updated.</p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<%
    Else
%>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center">Request canceled: Your did not change any information. </p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<%
    End If
  Else
  
%>
<!--<table width="100%" border="1" cellspacing="3" cellpadding="0" style="border-collapse: collapse">-->
<!--<table width="100%" border="0" cellspacing="3" cellpadding="0" style="border-collapse: collapse">-->
<table border="0" cellspacing="3" cellpadding="0" style="border-collapse: collapse">
  <tr>
    <td align="center" width="100%">
    <table border="0" cellpadding="3" cellspacing="3">
      <form method="POST" name="frmProfile" action="frmMain.asp?t=<%=mActiveTab%>&txtState=1"
            onsubmit="return profileValidate(this);">
        <tr>
          <td align="right" class="styLabelCell">Company Name </td>
          <td><input type="text" name="txtCompany" size="40" value="<%=mUserInsCoName%>"></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell">Office </td>
          <td><input type="text" name="txtOfficeName" size="40" value="<%=mUserACEInsCoOfficeName%>"></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell">Your Name </td>
          <td><input type="text" name="txtName" size="40" value="<%=mRepName%>"></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell">Your Phone Number </td>
          <td><b>(</b><input type="text" name="txtPhoneAc" size="3" maxlength="3" value="<%=PhoneAc%>"><b>)</b> <input type="text" name="txtPhoneExchg" size="3" 
          maxlength="3" value="<%=PhoneExchg%>"> <b>-</b> <input type="text" name="txtPhoneL4" size="4" maxlength="4" value="<%=PhoneL4%>"> <font size="1">&nbsp;Ext</font>
          <input type="text" name="txtPhoneEx" size="5" maxlength="5" value="<%=PhoneEx%>"></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell">Your Fax Number </td>
          <td><b>(</b><input type="text" name="txtFaxAc" size="3" maxlength="3" value="<%=FaxAc%>"><b>)</b> <input type="text" name="txtFaxExchg" size="3" 
          maxlength="3" value="<%=FaxExchg%>"> <b>-</b> <input type="text" name="txtFaxL4" size="4" maxlength="4" value="<%=FaxL4%>"></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell">Your Email Address </td>
          <td><input type="text" name="txtEmail" size="40" value="<%=EmailAddr%>"></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell">Sex </td>
          <td><select size="1" name="cmbSex">
          <option <%IF mUserGender=2 Then Response.Write "selected"%> value="Female">Female</option>
          <option <%IF mUserGender=1 Then Response.Write "selected"%> value="Male">Male</option>
          </select></font></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell"><font color="#FF0000">Change Password To&nbsp; </font></td>
          <td><input type="password" name="txtPassword" size="20" maxlength="20">-- <em>keep this private!</em><em style="font-style: normal"><font size="1"><br>
&nbsp;</font><b><font size="2">Leave blank if you do not wish to change your password</font></b></em></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell"><font color="#FF0000">Re-enter Password </font></td>
          <td><input type="password" name="txtPasswordVerify" size="20" maxlength="20">-- <em>for verification</em></td>
        </tr>
        <tr>
          <td colspan="2" align="CENTER">
          <p align="justify"><b><font face="Arial" size="1"><br>
          </font></b><font size="1" face="verdana, geneva, arial"><b>This form submits your request for changes to our staff. Your account will NOT be immediately 
          changed. For security purposes our staff will review your request and you will be notified via email when your account has been updated. This process 
          may take up to 48 hours. Thank You.</b></font></p>
          <center>
          <p><input type="SUBMIT" value="Submit"></p>
          <p><font face="geneva,arial" size="-2">By submitting this completed form you are agreeing to our terms of use.<br>
          Please read our <a href="/disclaimer.aspx">Terms of Service Agreement.</a></font></p>
          </center></td>
        </tr>
      </form>
    </table>
    </td>
  </tr>
</table>
<Script>placeFocus()</Script>
<%    
  End If
%>
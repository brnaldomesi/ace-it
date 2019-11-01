<% If mActiveTab = "" Then response.redirect "http://www.staging-portal.ace-it.com/" %> <% 
  If Request.QueryString("a") <> "" Then 
    AuditSpec = "http://webportal.staging-portal.ace-it.com/en2/" & EncStr("AuditPDF/ACEEstID/" & Request.QueryString("a") & "/.pdf", 1) & Request.QueryString("a") & ".pdf" '& Int(Rnd() * 10000) & ".pdf"     
    Response.Clear
    Response.Redirect AuditSpec
    
   ElseIf Request.QueryString("c") <> "" Then 
    sOffice = mUserACEInsCoOfficeName
    If (mUserSecurityL >= 5) OR (mUserTypeID) = cMemberType_CoMgr Then sOffice = ""
    AuditSpec = "http://webportal.staging-portal.ace-it.com/en2/" & EncStr("AuditPDF/ClaimNo/" & mUserInsCoName & "/" & sOffice & "/" & Request.QueryString("c") & "/", 1) & Request.QueryString("c") & ".pdf" '& Int(Rnd() * 10000) & ".pdf"
    Response.Clear
    Response.Redirect AuditSpec
  End If
%>
<script language="JavaScript"> var strCode = '';</script>
<script language="JavaScript">
  AuditNewWin = -1
</script>
<script>

function tabEmpViewAudit() {
  //sEditURL = "ProcessPDFReq.asp?EstID=" & rs("EstimateIDPK") & "&ref=" & Server.URLEncode(sURL)
  //DetailLink = " OnClick=""window.open ('" & sEditURL & "', '', 'toolbar=0, menubar=0, location=0');"" "
  var sEditURL = "ProcessPDFReq.asp?EstID=" + document.frmACENumSubmit.txtAuditNum.value;
  window.open (sEditURL, '', 'toolbar=0, menubar=0, location=0');
}

function tabEmpViewStatus() {
  if (document.frmACENumSubmit.txtAuditNum.value != ''){
    var sEditURL = "rptStatusSheet.asp?EstID=" + document.frmACENumSubmit.txtAuditNum.value;
    window.open (sEditURL, '', 'toolbar=1, menubar=1, location=0');
  }
}
</script>
<div align="center">
  <center>
  <% If mUserSecurityL >= 5 Then 
  %>
  <!--<table border="0" cellspacing="1" width="100%" cellpadding="5">-->
  <table border="0" cellspacing="1" cellpadding="5">
    <tr>
      <td><img border="0" src="../images/spacer.gif" width="119" height="1"></td>
      <td></td>
    </tr>
    <form name="frmACENumSubmit" onsubmit="tabEmpViewAudit(); return false" method="POST">
      <tr>
        <td align="right"><font face="Verdana" size="2"><b>ACE Audit #:</b></font></td>
        <td align="left"><input type="text" name="txtAuditNum" size="9">
        <!--<input type="submit" value="View..." name="cmdAudit" tabindex="0">&nbsp; --><a href="Javascript:tabEmpViewAudit()" tabindex="0">View Audit</a>&nbsp;&nbsp;
        <a href="Javascript:tabEmpViewStatus()">View Status</a> </td>
      </tr>
      <tr>
        <td align="right"><font face="Verdana" size="2"><b>&nbsp;</b></font></td>
        <td width="100%">
        <p align="left"><a href="frmRetrieveFileAttachments.asp">View Claim Submission Attachments</a>&nbsp; <br>
        <a target="_blank" href="http://webportal.staging-portal.ace-it.com/en2/<%=EncStr("bsntest/" & mUserName, 2)%>">BSN Test Site </a>&nbsp;<!--
            <a href="javascript:void(window.open('frmMembers_List.asp','Members','height=200,width=500,top='+(screen.width-500)/2+',left='+(screen.height-200)/2));">
            <font face="Verdana" size="1">View Member List</font></a>
            --> </p>
 </td>
      </tr>
      <tr>
        <!--
      <td align="right" width="125"><font face="Verdana" size="2" color="#0000FF"><b>Office:</b></font></td>
      <td align="left" width="375"><input type="text" name="T2" size="51" value="<%=mUserInsCoName & " - " & mUserACEInsCoOfficeName%>"></td>
      -->
        <td align="right"><font face="Verdana" size="2"><b>Office:</b></font></td>
        <td align="left" width="100%"><b><font size="2" face="Arial"><%=mUserInsCoName & " - " & mUserACEInsCoOfficeName%></font></b></td>
      </tr>
      <tr>
        <!--
      <td align="right" width="125"><font face="Verdana" size="2" color="#0000FF"><b>Claim Rep:</b></font></td>
      <td align="left" width="375"><input type="text" name="T1" size="33" value="<%=mRepName%>"> 
      -->
        <td align="right"><font face="Verdana" size="2"><b>Rep:</b></font></td>
        <td align="left" width="100%"><b><font size="2" face="Arial"><%=mRepName%></font></b> <font face="Verdana" size="2"><% 
              Select Case mUserTypeID
                Case cMemberType_ClaimRep
                  Response.Write " Claim Rep"
                Case cMemberType_CoOfficeMgr
                  Response.Write " Office Manager"
                Case cMemberType_CoMgr
                  Response.Write " Company Manager"
              End Select
            %> </font></td>
      </tr>
      <tr>
        <td colspan="2">
        <p align="left"><font face="Verdana" size="1">
		<!--
		You can change the rep you are mimicking by clicking here:
        <a href="javascript:void(window.open('frmSelRep.asp','SelectRep','height=350,width=500,resizable=1,scrollbars=1,left='+(screen.width-500)/2+', top='+(screen.height-200)/2));">Change Rep</a>.
		-->
		Note: ACE logins are automatically set for company manager access regardless of the rep you choose.</font>
        <!--
            <a href="javascript:void(window.open('frmRepViewLogin.asp','ViewRepLogin','height=200,width=500,top='+(screen.width-500)/2+',left='+(screen.height-200)/2));">
            <font face="Verdana" size="1">View Rep Login</font></a>
            -->
		</p>
        </td>
      </tr>
      <tr>
        <td width="100%" colspan="2">
        <table border="0" cellspacing="0" width="100%" cellpadding="0">
          <tr>
            <td width="100%" bgcolor="#DEDEDE"><img border="0" src="../images/spacer.gif" width="1" height="1"></td>
          </tr>
        </table>
        </td>
      </tr>
    </form>
  </table>
  <%Else%> You are not authorized to view this page. <% End If %> </center>
</div>
<p>
<!--<iframe name="ifrmClaimsListing" src="frmClaims_List.asp?SupressHdr=1" width="100%" height="400" marginwidth="6" marginheight="1" border="0" frameborder="0">-->
<iframe name="ifrmClaimsListing" src="frmClaims_List.asp?SupressHdr=1" width="80%" height="400" marginwidth="6" marginheight="1" border="0" frameborder="0">
Your browser does not support inline frames or is currently configured not to display inline frames.</iframe>
</p>
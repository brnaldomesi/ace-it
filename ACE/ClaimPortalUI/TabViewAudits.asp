<% If mActiveTab = "" Then response.redirect "http://www.staging-portal.ace-it.com/" %> 
<script language="JavaScript"> var strCode = '';</script>
<script language="JavaScript">


 //var jsparam = 'Rep/Recent/<%=mUserInsCoName%>/<%=mRepName%>/';</script>
 <script language="JavaScript"> var jsparam = 'en/<%=EncStr("Rep/Recent/" & mUserInsCoName & "/" & mRepName & "/", 1)%>';
</script>
<script LANGUAGE="JavaScript">
<!--
AuditNewWin = -1
// -->
</script>
<% 
  If Request.QueryString("a") <> "" Then 
    '--- Fetch Audit PDF ---
    Dim AuditSpec 
    AuditSpec = "http://webportal.staging-portal.ace-it.com/en2/" & EncStr("AuditPDF/ACEEstID/" & Request.QueryString("a") & "/.pdf", 1) & Request.QueryString("a") & ".pdf" '& Int(Rnd() * 10000) & ".pdf"     
    Response.Clear
    Response.Redirect AuditSpec
    
   ElseIf Request.QueryString("c") <> "" Then 
    '--- Fetch Audit PDF for claim # ---
    sOffice = mUserACEInsCoOfficeName
    If (mUserSecurityL >= 5) OR (mUserTypeID) = cMemberType_CoMgr Then sOffice = ""
    AuditSpec = "http://webportal.staging-portal.ace-it.com/en2/" & EncStr("AuditPDF/ClaimNo/" & mUserInsCoName & "/" & sOffice & "/" & Request.QueryString("c") & "/", 1) & Request.QueryString("c") & ".pdf" '& Int(Rnd() * 10000) & ".pdf"
    Response.Clear
    Response.Redirect AuditSpec
  End If
%> <script>
function DofrmACENumSubmit() {
  if (document.frmACENumSubmit.txtAuditNum.value != '') {
    NewWin('frmMain.asp?t=<%=mActiveTab%>&a=' + document.frmACENumSubmit.txtAuditNum.value);
    //document.frmACENumSubmit.txtAuditNum.value = '';
  }
}

function DofrmClaimNumSubmit() {
  if (document.frmClaimNumSubmit.txtClaimNum.value != '') {
    document.frmClaimNumSubmit.cmdClaimNoView.disabled=true
    NewWin('frmMain.asp?t=<%=mActiveTab%>&c=' + document.frmClaimNumSubmit.txtClaimNum.value);
    //document.frmClaimNumSubmit.txtClaimNum.value = '';
    document.frmClaimNumSubmit.cmdClaimNoView.disabled=false
  }
}

function DofrmClaimNumFind() {
  if (document.frmClaimNumFind.txtClaimNum.value != '') {
    document.frmClaimNumFind.cmdView.disabled=true
    NewWin('ProcessPDFReq.asp?txtState=2&txtClaimNum='+document.frmClaimNumFind.txtClaimNum.value);
    //document.frmClaimNumSubmit.txtClaimNum.value = '';
    document.frmClaimNumFind.cmdView.disabled=false
  }
}

</script>
  <table border="0" cellspacing="1" cellpadding="5" style="border-collapse: collapse">
    <tr>
      <td valign="top">
      <table border="0" width="100%" cellspacing="1" cellpadding="5">
        <tr>
          <td><img border="0" src="../images/spacer.gif" width="100" height="1"></td>
          <td></td>
        </tr>
        <form name="frmAudit" onsubmit="return DoNothing()" method="POST">
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
            <td align="left" width="100%"><b><font size="2" face="Arial"><%=mRepName%></font></b> 
            <font face="Verdana" size="2">
            <% 
              Select Case mUserTypeID
                Case cMemberType_ClaimRep
                  Response.Write " Claim Rep"
                Case cMemberType_CoOfficeMgr
                  Response.Write " Office Manager"
                Case cMemberType_CoMgr
                  Response.Write " Company Manager"
              End Select
            %>
            </font>            
            </td>
          </tr>
          <tr>
            <td colspan="2">
            <table border="0" cellspacing="0" width="100%" cellpadding="0">
              <tr>
                <td width="100%" ><img border="0" src="../images/spacer.gif"></td>
              </tr>
            </table>
            </td>
          </tr>
		  <% If false Then %>
          <tr>
            <td>&nbsp;</td>
            <td>
            <p align="left"><font face="Verdana" size="2" color="#808080"><b>Your most recent completed claims are shown below: </b></font></p>
            </td>
          </tr>
          <tr>
            <td align="right"><b><font face="Verdana" size="2">Click </font></b><font face="Verdana" size="2"><b>Claim to View Audit:</b></font></td>
            <td align="left" width="100%"><select size="4" name="cmbNewAudits" multiple 
            onchange="if (this.options[this.selectedIndex].value != '0') {var url=this.options[this.selectedIndex].value; this.selectedIndex=0; NewWin(url);}" 
            tabindex="3">
            <!------- Include code -------------><script language="JavaScript" src="../Inc/jsPullExtranet.js"></script><script language="JavaScript">document.write(strCode);
            </script>
            <!------- End Include code --------------->
            </select></td>
          </tr>
		  <% End If %>
        </form>
        <% 
      'If mUserSecurityL >= 5 Then 
      If 1 Then 
%>
        <tr>
          <td colspan="2">
          <table border="0" cellspacing="0" width="100%" cellpadding="0">
            <tr>
              <td width="100%"><img border="0" src="../images/spacer.gif"></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td colspan="2" style="padding-left: 15px;">
          <p><font face="Verdana" size="2"><b>
          <!--
          <a target="_blank" title="" href="ViewRecievedFiles.asp?State=1">
          <font color="#008000">Click HERE to see a list of all the claims received by ACE <br> from your Office/Company.</font>
          </a>
          -->
          <a title="" href="frmClaims_List.asp">
          <font color="#008000">
          <%
            Select Case mUserTypeID
              Case cMemberType_CoOfficeMgr
                Response.Write "Click HERE to see Pending and Completed files for your office"
              Case cMemberType_CoMgr
                Response.Write "Click HERE to see Pending and Completed files for your company"
              Case cMemberType_ClaimRep
                Response.Write "Click HERE to see your Pending and Completed files"
              Case Else
                Response.Write "Click HERE to see Pending and Completed files for your company"
            End Select
          %>
          </font>
          </a>
          </b></font></p>
          </td>
        </tr>
        <tr>
          <td colspan="2">&nbsp;
          <!--
          <table border="0" cellspacing="0" width="100%" cellpadding="0">
            <tr>
              <td width="100%" bgcolor="#DEDEDE"><img border="0" src="../images/spacer.gif"></td>
            </tr>
          </table>
          -->
          </td>
        </tr>
<!-- 1/6/2006 change       
        <tr>
          <td align="right" colspan="2">
          <p align="left"><font face="Verdana" size="1" color="#808080"><b>If you don't see the claim you wish to view in the menu above, enter all or part of 
          the claim number in the box below and click find. NOTE: 
          This function will only allow you to access claims that you have submitted. If you are an office manager you can view audits for claims submitted by your 
          office. If you are a company manager you can view all audits for your company.</b></font></p>
          </td>
        </tr>
-->        
<!-- Old code
    <tr>
      <form name="frmClaimNumSubmit" onsubmit="DofrmClaimNumSubmit(); return false" method="POST">
        <td align="right" width="119"><font face="Verdana" size="2" color="#0000FF"><b>Claim Number:</b></font></td>
        <td align="left" width="408"><input type="text" name="txtClaimNum" size="33" value> <input type="submit" value="View..." name="cmdClaimNoView">
        <br>
        <b><font size="1" color="#808080" face="Verdana">Enter the </font><font size="1" face="Verdana">complete</font><font size="1" color="#808080" face="Verdana"> 
        claim number and click View</font></b></td>
      </form>
    </tr>
--><% End If %>
        <!--
    <tr>
      <td colspan="2" width="538">
      <table border="0" cellspacing="0" width="100%" cellpadding="0">
        <tr>
          <td width="100%" bgcolor="#FFCC66"><img border="0" src="../images/spacer.gif"></td>
        </tr>
      </table>
      </td>
    </tr>
-->
<!-- 1/6/2006 change       
        <tr>
          <form name="frmClaimNumFind" onsubmit="DofrmClaimNumFind(); return false" method="POST">
            <td align="right" valign="top"><img src="../Images/fsviews.gif" border="0" align="bottom"><b><font size="2" face="Arial"> Find a Claim:</font></b></td>
            <td align="left" width="100%"><input type="text" name="txtClaimNum" size="30" tabindex="2"> <input type="submit" value="Find..." name="cmdView" 
            tabindex="3"></td>
          </form>
        </tr>
        <tr>
          <td colspan="2">
          <table border="0" cellspacing="0" width="100%" cellpadding="0">
            <tr>
              <td width="100%" bgcolor="#FFCC66"><img border="0" src="../images/spacer.gif"></td>
            </tr>
          </table>
          </td>
        </tr>
        <tr>
          <td align="center"><img border="0" src="../images/spacer.gif" width="1" height="8"></td>
          <td width="100%"></td>
        </tr>
-->        
      </table>
      <!--<table border="0" cellspacing="0" width="650px">-->
      <table border="0" cellspacing="0" width="100%">
        <tr>
          <td style="text-align:right; vertical-align:top;">
			  <a target="_blank" href="http://www.adobe.com/products/acrobat/readstep2.html"><font color="#808080"><img border="0" src="../Images/getacro.gif" width="88" height="31"></font></a>
          </td>
		  <td style="padding-left: 15px;">
			  <font face="Arial" size="2" color="#808080">Audits are in Adobe PDF format.<br>
			  You must have the </font> <a href="http://www.adobe.com/products/acrobat/readstep2.html"><font color="#808080">Adobe Acrobat plug-in </font> </a>
			  <font color="#808080">to view/print them.</font>
		  </td>
        </tr>
      </table>
	  
      </td>
      <!--
      <% If mUserSecurityL >= 5 Then %>
      <td valign="top" width="1%" height="100%">
      <table border="2" cellspacing="3" cellpadding="3" height="100%" style="border-collapse: collapse" bgcolor="#E0E0F0">
        <tr>
          <td>&nbsp;<br>
          <img border="0" src="../ACE-IT/Forms/formimages/alliance.jpg"><br>&nbsp;
          </td>
        </tr>
        <tr>
          <td align="center" bgcolor="#FFCC66">
          <p><a href="frmShopNetworkZipFind.asp">Shop Locator</a></p>
          </td>
        </tr>
        <tr height="100%">
          <td valign="top" height="100%">
          <p align="center">
          <font face="Arial" size="2">Click the link above to access the </font><i><b><font size="2" face="Times New Roman">ACE</font></b></i><font 
          face="Arial" size="2">&nbsp; Shop network.</font></p>
          <p align="center"><img border="0" src="../ACE-IT/images/logo_icon_Trans.gif"> <font face="Arial" size="2">and Dupont have teamed up to bring you a Body Shop Network with a 
          <b>difference</b>:</font></p>
          <p align="center"><b><font face="Arial" size="2">Top tier Shops<br>
          Dupont Guarantee<br>
          No Markup</font></b></p>
          <p align="center"><font face="Arial" size="2">Contact ACE here to find out more! </font> </p>
          </td>
        </tr>
      </table>
      </td>
      <% End If %>
      -->
    </tr>

<!--	
  	<tr >
		  <td width="800px">
		<p>
		<iframe name="ifrmClaimsListing" src="frmClaims_List.asp?SupressHdr=1" width="100%" height="400" marginwidth="6" marginheight="1" border="0" frameborder="0">
		Your browser does not support inline frames or is currently configured not to display inline frames.</iframe>
		</p>
	  </td>

	</tr>
-->
  </table>
    
  
<Script>placeFocus();</Script>
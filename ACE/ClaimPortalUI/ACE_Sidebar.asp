<style type="text/css">
<!--

body {
}

.ContentLink {
font-family: Arial; font-size: 10pt; color: #000066; /*#333399;*/ font-weight:bold; text-decoration:none}

.Link {
font-family: Arial; font-size: 10pt; color: #333399; font-weight:bold}
-->
</style>

<script>
function OpenTutorial(sDoc) {
  OpenChildWindow(sDoc,'ACE_Presentation','dependent,scrollbars=yes,location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=yes,width=600,screenX=0, screenY=0, height=400');  
}

  var windowOpt = 'dependent,scrollbars=yes,location=no,directories=no,status=yes,menubar=yes,toolbar=yes,resizable=yes,screenX=0,screenY=0,width=500,height=' + (screen.availHeight - 60);
</script>

<table border="0" cellspacing="0" width="150" id="AutoNumber1" cellpadding="0" style="margin:0; border-width:0; padding:0; border-style:solid; border-collapse: collapse; ">
  <!--
  <tr>
    <td width="100%"><img border="0" src="../images/spacer.gif" width="6" height="6"></td>
  </tr>
  <tr>
    <td width="100%">
    <img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="../Login/LogOut.asp">Log Out</a></td>
  </tr>
  -->
 <tr>
    <td width="100%"><font size="1">&nbsp;</font></td>
  </tr>
  <!--
  <tr>
    <td width="100%" align="center" <%=fLinkHighlite(1)%>> <a class="ContentLink" href="frmMain.asp?t=1">View Audits</a></td>
  </tr>
  <tr>
    <td width="100%" align="center" <%=fLinkHighlite(2)%>> <a class="ContentLink" href="frmMain.asp?t=2">Submit Claims</a></td>
  </tr>
  <tr>
    <td width="100%" align="center" <%=fLinkHighlite(3)%>> <a class="ContentLink" href="frmMain.asp?t=3">Reports</a></td>
  </tr>
  <tr>
    <td width="100%" align="center" <%=fLinkHighlite(4)%>> <a class="ContentLink" href="frmMain.asp?t=4">ACE Info</a></td>
  </tr>
  <tr>
    <td width="100%" align="center" <%=fLinkHighlite(5)%>> <a class="ContentLink" href="frmMain.asp?t=5">My Profile</a></td>
  </tr>
  <tr>
    <td width="100%" align="center" <%=fLinkHighlite(6)%>> <a class="ContentLink" href="frmMain.asp?t=6">Support</a></td>
  </tr>
  <tr>
    <td width="100%">&nbsp;</td>
  </tr>
  -->
  <tr>
    <td width="100%">
    <b><font face="Arial" size="4" color="#333399">Submit a Claim:</font></b><br>
    <%
		Dim addon

        addon = ""

		If Trim(LCase(mUserInsCoName)) = "westfield acs" Or Trim(LCase(mUserInsCoName)) = "westfield subro" Or Trim(LCase(mUserInsCoName)) = "westfield insurance" Then
            addon = "Westfield"
        End If
		
		' Must also change frmClaims_List.asp as well!
    %>
    <img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="JavaScript:OpenChildWindow('frmSubmitAutoReview<%=addon %>.asp','ACE',windowOpt);void(0);">Auto</a><br>
    <img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="JavaScript:OpenChildWindow('frmSubmitSubReview.asp','ACE',windowOpt);void(0);">Subrogation</a><br>
    <img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="JavaScript:OpenChildWindow('frmSubmitHeavySpecialtyReview.asp','ACE',windowOpt);void(0);">Heavy/Specialty</a><br>
    <img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="JavaScript:OpenChildWindow('frmSubmitPropReview.asp','ACE',windowOpt);void(0);">Property</a><br>
    <img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="JavaScript:OpenChildWindow('frmTotalLoss.asp','ACE',windowOpt);void(0);">Auto Total Loss</a>
    <br/><img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="JavaScript:OpenChildWindow('frmSubmitPhotoQuote.asp','ACE',windowOpt);void(0);">PhotoQuote</a>
	</td>
  </tr>
  <tr>
    <td width="100%" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%"><b><font face="Arial" size="3" color="#333399">Additional Resources:<br>
    </font></b>
    <img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" target="_blank" href="/brochures.aspx">Online Brochure</a><br>
    <!--
    <a class="Link" href="JavaScript:OpenTutorial('../ACE-IT/brochure.html');void(0);">Online Brochure</a><br>
    <img height="12" src="../Images/arrow_05.gif" width="13"> 
    <a class="Link" href="JavaScript:OpenTutorial('ACEClaimsPortalHelp.pdf');void(0);">Portal Tutorial</a>
    -->
&nbsp;</td>
  </tr>
  <!--<tr>
    <td width="100%"><font size="5">&nbsp;</font></td>
  </tr>
  <tr>
    <td width="100%" align="center"><a class="Link" target="_blank" href="http://webportal.ace-it.com/en2/<%=EncStr("bsn/" & mUserName, 2)%>">
    <img border="0" src="images/BodyShopNet_smaller.gif" width="139" height="72"></a><br><font size="1"><br>
    </font>
    <font color="#0057A6" face="Arial" size="1">
    A network of over 7,500 DuPont Performance Alliance and BodyShopNet Class A shops.</font><font color="#333399" face="Arial" size="2"><br>
    </font>
    <font color="#333399" face="Arial" size="1">&nbsp;</font></td>
  </tr>
  <tr>
    <td width="100%" align="left">
    <img height="12" src="../Images/arrow_05.gif" width="13">
    <a class="Link" target="_blank" href="http://webportal.ace-it.com/en2/<%=EncStr("bsn/" & mUserName, 2)%>">Make an Assignment</a>
    <img height="12" src="../Images/arrow_05.gif" width="13"> 
    <a class="Link" target="_blank" href="http://BodyShopNet.com/">About BodyShopNet</a>
    </td>
  </tr>-->
  <tr>
    <td width="100%" height="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" height="100%">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" height="100%">
<!-- footer -->
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
      <td width="100%">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#808080" size="1" face="Arial">Unless otherwise noted all content Copyright© American Computer Estimating, Inc.
        </font>
      <font face="Arial" size="-2"><a target="_top" href="../ACE-IT_sv/companyinfo.html">Contact Us</a></font><font color="#808080" size="1" face="Arial"><br>
      <br>
        Web Site design, programs and web code Copyright © <a href="http://www.puppozone.com/">PUPPO SOFTWARE SOLUTIONS, LLC.</a>&nbsp;&nbsp; All Rights Reserved.<br>
      <br>
        Send mail to <br>
      <a href="mailto:support@staging-portal.ace-it.com?subject=Regarding the ACE-IT Website">support@staging-portal.ace-it.com</a> with questions or comments about this site.</font></tr>
  </table>
</div>

    </td>
  </tr>
  <tr>
    <td width="100%" height="100%"><img border="0" src="../images/spacer.gif" width="1" height="5">
    </td>
  </tr>
</table>
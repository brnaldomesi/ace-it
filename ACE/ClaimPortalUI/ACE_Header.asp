<%
  Dim sTitleBarText
  sTitleBarText = "Welcome: " & mName & " "
%> <%
  Dim sTitleBarTxt
  
  On Error Resume Next
  sTitleBarTxt = sTitleBarText
  On Error Goto 0
  If sTitleBarTxt = "" then sTitleBarTxt = "W W W . A C E - I T . C O M"
%>
<!-- START header -->
<style type="text/css">
<!--
.titlebartxt { font-family: Verdana, Geneva, Arial, 'sans serif';
	           font-size: .8em;
               color: #ffffff; 
               font-weight: bold
               }
-->
</style>
<div style="margin:0; padding:0; background-image: url('images/ACEPortalRight.jpg'); background-repeat: repeat-x; width:100%">
  <div style="margin:0; padding:0; background-image: url('images/ACEPortalBannerWhole.jpg'); background-repeat: no-repeat; width:100%; height:90px">
    <table border="0" cellspacing="0" width="100%" id="AutoNumber1" cellpadding="0" style="margin:0; padding:0; border-collapse: collapse; ">
      <tr>
        <td width="100%" align="right" valign="top"><img border="0" src="../Images/spacer.gif" width="1" height="51" align="left">
			<img height="12" src="../Images/arrow_05.gif" width="13"> <a class="Link" href="../Login/LogOut.asp"><font color="#FFFFFF">Log Out</font></a><img border="0" src="../Images/spacer.gif" width="6" height="1">
        </td>
      </tr>
      <tr>
        <td class="titlebartxt" width="100%" style="padding-left: 290px;"><%=sTitleBarTxt%></td>
      </tr>
      <tr>
        <td width="100%"><img border="0" src="../Images/spacer.gif" width="1" height="3" align="left"></td>
      </tr>
      <tr>
        <td class="titlebartxt" width="100%" style="padding-left: 142px;">
        <div id="topMenu">
          <!-- #INCLUDE FILE="ACE_TopMenu.asp" -->
        </div>
        </td>
      </tr>
    </table>
  </div>
</div>
<!--
<table border="0" cellpadding="0" cellspacing="0" bgcolor="#000000">
  <tr>
    <td><img src="../ACE-IT/images/spacer.gif" width="580" height="2"></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="../ACE-IT/images/spacer.gif" width="600" height="2"></td>
  </tr>
</table>
-->
<!-- END header / START main body -->
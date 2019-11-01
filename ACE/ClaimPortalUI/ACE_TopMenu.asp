<script language="JavaScript1.2">
function changeto(highlightcolor){
  source = event.srcElement;
  if (source.tagName == "TR" || source.tagName == "TABLE")
    return;
  while(source.tagName != "TD"){
    source=source.parentElement;
    if (source.style.backgroundColor != highlightcolor && source.id != "ignore")
      source.style.backgroundColor=highlightcolor;
  }
}

function changeback(originalcolor){
if (event.fromElement.contains(event.toElement)||source.contains(event.toElement)||source.id=="ignore")
return
if (event.toElement!=source)
source.style.backgroundColor=originalcolor
}


// onMouseOver="changeto('#cccccc')" onMouseOut="changeback('white')"
</script>
<%
  Function fLinkHighlite(iTab)
    fLinkHighlite = ""
    If CStr(mActiveTab) = CStr(iTab) Then fLinkHighlite = "bgcolor='#FFFFFF'" ''#FFFFC0'
  End Function
%>

<table border="0" cellspacing="0" width="100%" id="AutoNumber1" cellpadding="2" style="margin:0; padding:0; border-collapse: collapse">
 <tr>
    <td><font size="1">&nbsp;</font><img border="0" src="../images/spacer.gif" width="25" height="1"></td>
    <td align="center" nowrap style="padding-left:8px; padding-right:8px;" <%=fLinkHighlite(1)%>> <a class="ContentLink" href="frmMain.asp?t=1">View Audits</a></td>
    <td align="center" nowrap style="padding-left:8px; padding-right:8px;" <%=fLinkHighlite(3)%>> <a class="ContentLink" href="frmMain.asp?t=3">Reports</a></td>
    <td align="center" nowrap style="padding-left:8px; padding-right:8px;" <%=fLinkHighlite(5)%>> <a class="ContentLink" href="frmMain.asp?t=5">My Profile</a></td>
    <td align="center" nowrap style="padding-left:8px; padding-right:8px;" <%=fLinkHighlite(6)%>> <a class="ContentLink" href="frmMain.asp?t=6">Support</a></td>
<% If mUserSecurityL >= 5 Then %>
    <td align="center" nowrap style="padding-left:8px; padding-right:8px;" <%=fLinkHighlite(7)%>> <a class="ContentLink" href="frmMain.asp?t=7">Employees Only</a></td>
<% End If%>    
    <td width="100%">&nbsp;</td>
  </tr>
</table>
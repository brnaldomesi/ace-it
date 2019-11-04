<% Option Explicit %>
<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->

<%
'--------------------------------------------
Sub PostActivity()
  Dim oDataActivity
  
  If mUserID <> 0 Then
    Set oDataActivity = New  clsACEDataConn
    oDataActivity.ACEMemberActivityInsert2 mUserID, 10, Request.ServerVariables("SCRIPT_NAME") & " " & Request.ServerVariables("QUERY_STRING")
    Set oDataActivity = Nothing
  End If
End Sub
'--------------------------------------------

  PostActivity
%>

<%
  Dim mState, mClaimNum

  mState=0
  On Error Resume Next
  mState=Request.QueryString("txtState")
  mClaimNum=Request.QueryString("txtClaimNum")
  On Error Goto 0
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>ACE - Audit PDF Req. Processing</title>
<meta name="Microsoft Border" content="none">
</head>

<body>

<p align="center"><img border="0" src="../Images/ace_logo.gif" width="149" height="71"><br>
<img border="0" src="../Images/LineGradBlue.jpg" width="422" height="11"></p>
<%
  If mState=0 Then
    'Response.Clear
    'Response.Redirect "http://webportal.ace-it.com/en2/" & EncStr("AuditPDF/ACEEstID/" & Request.QueryString("EstID") & ".pdf", 1) & "/" & Request.QueryString("EstID") & ".pdf"
%>
 <p align="center"><font color="#808000" size="4">... Your document request is being processed please wait ...</font></p>

 <%
	Dim url,strArr,xmlhttp,lineno, EstID
	
	EstID = Request.QueryString("EstID")
	If Len(EstID) > 10 Then
		'--- its encrypted so decrypt
		EstID = DeObfiscateEstimateID(Request.QueryString("EstID"))
	End If
	If EstID = 0 Then
		Response.Write "... ERROR: The file cannot be found or you are not authorized to view the file."
		Response.End
	End If
    
    url = "http://webportal.ace-it.com/en/" & EncStr("AuditPDF/ACEEstID/" & EstID & ".pdf", 1) & "/" & Request.QueryString("EstID") & ".pdf" 
 
    
    set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
	xmlhttp.open "GET", url, false 
    xmlhttp.send "" 
	Response.ContentType = "application/pdf"
    Response.BinaryWrite xmlhttp.ResponseBody 'xmlhttp.responseText 
	Response.End
 %>
 
 
 
<form name="frmPDF" method="POST" 
action="<%="http://webportal.ace-it.com/en/" & EncStr("AuditPDF/ACEEstID/" & Request.QueryString("EstID") & ".pdf", 1) & "/" & Request.QueryString("EstID") & ".pdf"%>" >
  <input type = "hidden" size="20">
</form>
<script language="JavaScript">
 window.title="ACE"
 //window.location.replace("<%="http://webportal.ace-it.com/en/" & EncStr("AuditPDF/ACEEstID/" & Request.QueryString("EstID") & ".pdf", 1) & "/" & Request.QueryString("EstID") & ".pdf"%>");
 document.frmPDF.submit();
</script>
<%
  Else
    Select Case mUserTypeID 
      Case cMemberType_Administrator, cMemberType_ACEEmployee, cMemberType_ACEAdmin, cMemberType_CoOfficeMgr, cMemberType_CoMgr
%>
<script language="JavaScript"> var strCode = '';</script>
<script language="JavaScript"> var jsparam = 'en/<%=EncStr("Search/" & mUserInsCoName & "/" & "" & "/" & mClaimNum, 1)%>/';</script>
<%
      Case Else
%>      
<script language="JavaScript"> var strCode = '';</script>
<script language="JavaScript"> var jsparam = 'en/<%=EncStr("Rep/AuditPDF/" & mUserInsCoName & "/" & mRepName & "/" & mClaimNum, 1)%>/';</script>
<%
    End Select
%>
    
<div id="divProcessing" >
<p align="center"><font color="#808080">... Please wait while we find all your claim numbers that contain the search string ...</font></p>
</div>
<table border="0" cellspacing="1" width="100%">
  <tr>
    <td width="100%" align="center">
      <form method="POST" action="javascript:ProcessPDFReq.asp" target="_blank" name="frmView">
      <p align="center">The box below contains <b>closed</b> files matching the results of your search. <br>
      Clicking a file will display the finished audit.</p>
        <p align="center">
        <select size="5" style="color: #000080" name="cmbClosedAudits" 
        onchange="if (this.options[this.selectedIndex].value != '0') {window.location=(this.options[this.selectedIndex].value);}" tabindex="1">
          <option selected value="0">...Select a claim below to view it...</option>
                                  <!------- Include code ------------->
<script src="../Inc/jsPullExtranet.js"></script>
<script language="JavaScript">document.write(strCode);</script>
                                  <!------- End Include code --------------->                                     
        </select>
        </p>
      </form>
    </td>
  </tr>
</table>  
<table border="0" cellspacing="1" width="100%">
  <tr>
    <td width="100%" align="center">
      <p align="center"><br>
      The box below contains files matching the results of your search<br>
      that have been <b>received</b> by <i><b>ACE</b></i>
      but are still in process.</p>
<script language="JavaScript">strCode = '';</script>
<script language="JavaScript">jsparam = 'en/<%=EncStr("SearchInProcess/" & mUserInsCoName & "/" & "" & "/" & mClaimNum, 1)%>/';</script>
        <p align="center">
        <select size="5" style="color: #000080" name="cmbOpenAudits"
        onchange="if (this.options[this.selectedIndex].value != '0') {window.location=(this.options[this.selectedIndex].value);}" tabindex="2">
          <option selected value="0">...Select a claim below to view it's 
          status...</option>
                                  <!------- Include code ------------->
<script src="../Inc/jsPullExtranet.js"></script>
<script language="JavaScript">document.write(strCode);</script>
                                  <!------- End Include code --------------->                                     
        </select>
        </p>
      <p align="center"><font color="#0000FF"><b>** Search Complete **</b> </font></p>
    </td>
  </tr>
</table>

<script language="JavaScript">
<!--
ns4 = (document.layers)? true:false
ie4 = (document.all)? true:false

  if (ns4) {
    //divProcessing.visibility = "hide"
    document.layers['divProcessing'].visibility = "hide"
  } else if (ie4) {
    //divProcessing.visibility = "hidden"
    document.all['divProcessing'].style.visibility = "hidden"
  }
-->
</Script>

<%
  End If
%>
<p>&nbsp;</p>

</body>

</html>
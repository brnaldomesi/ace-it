<% Option Explicit %>

<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->

<%
  Dim mSelectedCo, mSelectedOffice, mSelectedRep
  Dim mState
  Dim sSQL, rs

  Dim oDataConn, mrs

  On Error Resume Next
  mSelectedCo = Session("AdminSelectedCo")
  mSelectedOffice = Session("AdminSelectedOffice")
  mSelectedRep = Session("AdminSelectedRep")
  On Error Goto 0

  mState = 0
  mState = Request.Form("txtState")

  Select Case mUserTypeID
    Case 1,2,4 'Claim Rep or Agent
    	mState = 99 'not allowed
    Case 9 'Company Office manager
    	If mState <= 2 Then
    	  mState = 2
         mSelectedCo = mUserInsCoName
         Session("AdminSelectedCo") = mSelectedCo
         mSelectedOffice = mUserACEInsCoOfficeName
         Session("AdminSelectedOffice") = mSelectedOffice
    	End If
    Case 10 'Company manager
    	If mState <= 1 Then
    	  mState = 1
         mSelectedCo = mUserInsCoName
         Session("AdminSelectedCo") = mSelectedCo
    	End If
  End Select  

'  If mSelectedRep = "" Then mState = 1
'  If mSelectedCo = "" Then mState = 0

  Select Case mState
    Case 1 'Pick Office
      If Request.Form("cmbCompany") <> "" Then
        mSelectedCo=Request.Form("cmbCompany")
        Session("AdminSelectedCo")=mSelectedCo
      End If
    Case 2 'Pick Rep
      If request.Form("cmbOffice") <> "" Then
        mSelectedOffice = request.Form("cmbOffice")
        Session("AdminSelectedOffice")=  mSelectedOffice
      End If
    Case 3 'Show Submit
      mSelectedRep = request.Form("cmbRep")
      Session("AdminSelectedRep")=  mSelectedRep
    Case 4 'Submit Clicked
      '--- Write changes to user account and update session varaibles ---
      Session("RepName")=mSelectedRep         '<--Session vars used by rep menu
      Session("InsCompanyName")=mSelectedCo   '<--/
      Session("ACEInsCoOfficeName")=mSelectedOffice   '<--/

      If mSelectedRep <> "" and mSelectedCo <> "" Then
        Set oDataConn = New clsACEDataConn

        Set mrs = oDataConn.ACEMemberGet(mUserID,"","")
        mrs("ACERepName") = mSelectedRep
        
        'mrs("InsCompanyID")=oDataConn.ACEGetCoID(mSelectedCo)
        'mrs("ACEInsCoOfficeName")=mSelectedOffice

        If mSelectedOffice = "ALL" Then
          sSQL = _
            "SELECT CR.InsuranceCompanyID, Co.InsuranceCompanyOfficeLocation " & _
            "FROM   dbo.wv_tblClaimReps CR INNER JOIN " & _
                   "dbo.vw_tblInsuranceCompanies Co ON CR.InsuranceCompanyID = Co.InsuranceCompanyIDPK " & _
            "WHERE     (CR.ClaimRepName = N'" & mSelectedRep & "') AND (Co.InsuranceCompanyName = '" & mSelectedCo & "')"
          Set rs = oDataConn.RSGet(sSQL)
          mrs("ACEInsCoOfficeID") = rs(0)
          Session("ACEInsCoOfficeID") = rs(0)
          Session("ACEInsCoOfficeName") = rs(1)
          rs.Close
        Else
          sSQL = _
            "SELECT InsuranceCompanyIDPK " & _
            "FROM   vw_tblInsuranceCompanies " & _
            "WHERE  (InsuranceCompanyName = '" & mSelectedCo & "') AND (InsuranceCompanyOfficeLocation = '" & mSelectedOffice & "')"
          Set rs = oDataConn.RSGet(sSQL)
          mrs("ACEInsCoOfficeID") = rs(0)
          Session("ACEInsCoOfficeID") = rs(0)
          rs.Close
        End If
        
        oDataConn.ACEMemberUpdate mrs

        mrs.close
        Set mrs = Nothing

        Session.Contents.Remove "AdminSelectedCo"
        Session.Contents.Remove "AdminSelectedOffice"
        Session.Contents.Remove "AdminSelectedRep"

        'response.redirect "../RepMenu2.asp"
      Else
        'response.redirect "ACEAdminSelRep.asp"
      End If
  End Select
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>ACE Select Rep</title>

<!-- ~~~~~ Display Loading message ~~~~~~ -->
<script language="JavaScript">
<!--
ns4 = (document.layers)? true:false
ie4 = (document.all)? true:false

var PageLoaded = false

function ShowLoading() {
  if (ns4) block = document.layers['blockDiv'] //document.blockDiv
  if (ie4) block = blockDiv.style
  block.xpos = parseInt(block.left)
  slide()
}

function HideLoading() {
  if (ns4) {
    block.visibility = "hide"
    document.layers['load'].visibility = "hide"
  } else if (ie4) {
    block.visibility = "hidden"
    document.all['load'].style.visibility = "hidden"
  }
  PageLoaded = true
}

function slide() {
  if (PageLoaded != true) {
    if (block.xpos < 270) {
      block.xpos += 10
      block.left = block.xpos
      setTimeout("slide()",100)
  	} else {
      restart()
    }
  }
}

function restart() {
  block.xpos = 200
  block.left = block.xpos
	slide()
}
//-->
</SCRIPT>
<!-- End display loading message -->


<script language="JavaScript" >
function NextStep(ThisCtrl, Step) {

  if (ThisCtrl.options[ThisCtrl.selectedIndex].value != '0') {
    document.forms[0].txtState.value=Step;
    document.forms[0].submit();
  }
}
</script>

</head>

<body oncontextmenu="return false">

<!--- ~~~~~~~~~~ Show Loading Message ~~~~~~~~~~~~ --->
<div id="load" style="position:absolute; left:110px; top:88px;color: #000000; font:12pt Arial, Helvetica, sans-serif">
  Page Loading
</div>
<div id="blockDiv" style="position:absolute; left:200px; top:100px; width:5px; height:5px; clip:rect(0px 5px 5px 0px); background-color:#567886; layer-background-color:#567886;">
</div>
<SCRIPT Language = "JavaScript">
ShowLoading()
</script>
<!--- ~~~~~~~~~~ End Show Loading Message ~~~~~~~~~~~~ --->
<br>
<div align="center">
  <center>
  <table border="1" cellspacing="1" width="450">
    <tr>
      <td width="100%" align="center" bgcolor="#356C91"><font color="#FFFFFF"><b><i>ACE</i> Claim Rep Selection</b></font></td>
    </tr>
    <font color="#808000">
    <tr>
        <form method="POST" action="frmSelRep.asp">
      <td>
        <table border="0" cellspacing="1" width="100%">
            <%
	          Select Case mState
	              Case 0
            %>
          <tr>
            <td width="100%" align="center">Please select the company of the Rep you wish to alias: &nbsp;</font></td>
        </tr>
        <tr>
          <!-- <td width="100%" align="center"><select size="1" name="cmbCompany" onchange="NextStep(this, 1);"> -->
          <td width="100%" align="center"><select size="15" name="cmbCompany">
            <%
              'Set oDataConn = New clsACEDataConn

              'sSQL = _
              '  "SELECT DISTINCT InsuranceCompanyName " & _
              '  "FROM   vw_tblInsuranceCompanies " & _
              '  "WHERE  (NOT (InsuranceCompanyName LIKE '%Do No%')) " & _
              '  "ORDER BY InsuranceCompanyName"
              'Set mrs = oDataConn.RSGet(sSQL)
             ' 
             ' Do While Not mrs.eof
             '   Response.Write "<option value=""" & mrs(0) & """>" & mrs(0) & "</option> "
             '   mrs.MoveNext
             ' Loop
             ' mrs.Close
            %>
            
<script language="JavaScript" src="http://webportal.staging-portal.ace-it.com/Admin/CoList/">document.write("<br>Database server did not respond<br>");</script>
<script language="JavaScript">document.write(strCode);</script>
          
            </select> <input type="button" value="Next &gt;&gt;" name="cmdNext" onclick="NextStep(document.forms[0].cmbCompany, 1);"></td>
        </tr>
            <%   Case 1  %>
        <tr>
          <td width="100%" align="center"></td>
        </tr>
        <tr>
          <td width="100%" align="center">Select an office for <%=mSelectedCo%>:</td>
        </tr>
        <tr>
          <!-- <td width="100%" align="center"><select size="1" name="cmbOffice" onchange="NextStep(this, 2);"> -->
          <td width="100%" align="center"><select size="5" name="cmbOffice">
<script language="JavaScript" src="http://webportal.staging-portal.ace-it.com/Admin/CoOffList/<%=mSelectedCo%>/">document.write("<br>Database server did not respond<br>");</script>
<script language="JavaScript">document.write(strCode);</script>
            </select> <input type="button" value="Next &gt;&gt;" name="cmdNext" onclick="NextStep(document.forms[0].cmbOffice, 2);"></td>
        </tr>
            <%   Case 2  %>
        <tr>
          <td width="100%" align="center"></td>
        </tr>
        <tr>
          <td width="100%" align="center">Select a Rep for <%=mSelectedCo & "-" & mSelectedOffice%>:</td>
        </tr>
        <tr>
          <!-- <td width="100%" align="center"><select size="1" name="cmbRep" onchange="NextStep(this, 3);"> -->
          <td width="100%" align="center"><select size="5" name="cmbRep">
<script language="JavaScript" src="http://webportal.staging-portal.ace-it.com/Admin/Reps/<%=mSelectedCo & "/" & mSelectedOffice%>/">document.write("<br>Database server did not respond<br>");</script>
<script language="JavaScript">document.write(strCode);</script>
            </select> <input type="button" value="Next &gt;&gt;" name="cmdNext" onclick="NextStep(document.forms[0].cmbRep, 3);"></td>
        </tr>
            <%   Case 3  %>
        <tr>
          <td width="100%" align="center"></td>
        </tr>
        <tr>
          <td width="100%" align="center">Pressing the submit button will commit your changes:<br>
               <%=mSelectedCo & " - " & mSelectedOffice & " - " & mSelectedRep%>
          </td>
        </tr>
        <tr>
          <td width="100%" align="center"><input type="button" value="Submit" name="cmdSubmit" onclick="document.forms[0].txtState.value=4;document.forms[0].submit();"></td>
        </tr>
            <%   Case 99  %>
        <tr>
          <td width="100%" align="center">Sorry, your membership level does not allow this operation.</td>
        </tr>
            <%   Case 4  %>
<Script LANGUAGE="JavaScript">
<!--
  window.opener.location.reload(true)
  HideLoading()
  window.close()
-->
            </script>
            <%  End Select  %>

      </table>
      <input type="hidden" name="txtState" value="0">
      </form>
    </tr>
  </table>
  </center>
</div>
<p>&nbsp;</p>

<Script LANGUAGE="JavaScript">
  HideLoading()
</script>

</body>

</html>
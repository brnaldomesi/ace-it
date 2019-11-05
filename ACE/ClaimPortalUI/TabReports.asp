<% If mActiveTab = "" Then response.redirect "http://www.staging-portal.ace-it.com/" %> 
<% 
  Dim MaxEndDate, DaysToDelay
  
  Randomize

  DaysToDelay = 1
  If DateDiff("n", Date & " 12:00 pm", Now) < 0 Then DaysToDelay = 2

  If mUserSecurityL >= 5 Then DaysToDelay = 0
  
'  If mUserSecurityL < 5 Then 
  Select Case mUserTypeID 
    Case cMemberType_Administrator, cMemberType_ACEEmployee, cMemberType_ACEAdmin, cMemberType_CoOfficeMgr, cMemberType_CoMgr
    
    Case Else
%>
<table border="0" cellspacing="1" width="100%">
  <tr>
    <td width="100%" align="center">&nbsp;<p><b>Sorry, Reports are only available to Office and Company managers</b>.<p>&nbsp;</td>
  </tr>
</table>

<Script LANGUAGE="JavaScript">
  HideLoading()
</script>


<% 
    Response.End '*** this leaves the Page loading twirly running so we do the HideLoading above***
  End Select
%>

<%    
    On Error Resume Next
    sCompany = Request.Form("lstCompany")
    sOffice = Request.Form("lstOffice")
    On Error Goto 0
    
    If Len(sCompany) = 0 Then sCompany = mUserInsCoName
    If mUserTypeID = cMemberType_CoOfficeMgr Then If Len(sOffice) = 0 Then sOffice = mUserACEInsCoOfficeName
    
    If Request.QueryString("txtState") = 1 Then
      '--- Report URL: ---
      ' ReportID/CompanyName/Start Date/End Date/Office Name

      '--- Fetch Report PDF ---
      Response.Clear
      Dim ReportSpec 
      'ReportSpec = "http://webportal.ace-it.com/en2/" & EncStr("Report/" & Request.Form("lstReport") & "/" & sCompany & "/" & Replace(Request.Form("txtDateStart"),"/","-") & "/" & Replace(Request.Form("txtDateEnd"),"/","-") & "/" & sOffice, 1) & "/rpt" & Int(Rnd() * 10000) & ".pdf"
      ReportSpec = "http://webportal.ace-it.com/" & "Report/" & Request.Form("lstReport") & "/" & sCompany & "/" & Replace(Request.Form("txtDateStart"),"/","-") & "/" & Replace(Request.Form("txtDateEnd"),"/","-") & "/" & sOffice & "/rpt" & Int(Rnd() * 10000) & ".pdf"
      'ReportSpec = "http://webportal.ace-it.com/test"
      Response.Redirect ReportSpec    

    End If
%>


<style>
<!--
  td.styLabelCell        { font-family: verdana, geneva, arial; font-size: 8pt; font-weight:bold; /*background-color:#F9F9E3 #FFFFCC*/ }
-->
</style>


<!-- PopUp Calendar BEGIN -->
<script type="text/JavaScript" src="../JavaScript/PopupCal/PopupCal.js"></script>
<!-- PopUp Calendar END -->
	

<script Language="JavaScript" Type="text/javascript"><!--
function EndDateCheck(txtRptEnd, bNoAlert) {
  var DaysDelay = <%=DaysToDelay%>

  var dtNow = new Date() //; if (dtNow.getYear() < 1000) dtNow.setYear(dtNow.getYear() + 1900);
  var dtYest = new Date(); dtYest.setDate(dtNow.getDate()-1);
  var dt2Ago = new Date(); dt2Ago.setDate(dtNow.getDate()-2); 
  
 // var dtYest = new Date(dtNow.getYear(), dtNow.getMonth(), dtNow.getDate()-1); if (dtYest.getYear() < 1000) dtYest.setYear(dtYest.getYear() + 1900);
 // var dt2Ago = new Date(dtNow.getYear(), dtNow.getMonth(), dtNow.getDate()-2); if (dt2Ago.getYear() < 1000) dt2Ago.setYear(dt2Ago.getYear() + 1900);


  dtRptEnd = new Date(Date.parse(txtRptEnd.value));
  
  if (DaysDelay == 2) {
    if ((dt2Ago - dtRptEnd) < 0) {
      dtRptEnd = dt2Ago;
      if (bNoAlert != 1) alert ('The End date has been adjusted because files are delayed two days when running reports before 12:00 PM EDT.') 
    } 
  } else if (DaysDelay == 1) {
    if ((dtYest - dtRptEnd) < 0) {
      dtRptEnd = dtYest;
      if (bNoAlert != 1) alert ('The End date has been adjusted because files are delayed one day when running reports.') 
    }
  }  
  txtRptEnd.value = FormatDate(dtRptEnd);
}

function FormatDate(dtDate) {
  var dtY; var dtM; var dtD;  
  dtY = dtDate.getYear(); if (dtY < 1000) dtY += 1900;
  dtM = dtDate.getMonth() + 1; if (dtM < 9) dtM = '0' + dtM.toString();
  dtD = dtDate.getDate(); if (dtD < 9) dtD = '0' + dtD.toString();
  return (dtM + '/' + dtD + '/' + dtY);  
}
--></script>

<script language="JavaScript">var strCode = '';</script>

<script Language="JavaScript" Type="text/javascript"><!--

function FrmReports_Validator(theForm)
{

  if (theForm.lstReport.selectedIndex < 0)
  {
    alert("Please select one of the \"Report Selection\" options.");
    theForm.lstReport.focus();
    return (false);
  }

  if (theForm.lstReport.value == 0)
  {
    alert("You selected an invalid \"Report Selection\" option. Please select a report.");
    theForm.lstReport.focus();
    return (false);
  }

<%
  Select Case mUserTypeID 
    Case cMemberType_Administrator, cMemberType_ACEEmployee, cMemberType_ACEAdmin
%> 
  if (theForm.lstCompany.selectedIndex < 0)
  {
    alert("Please select one of the \"Company Selection\" options.");
    theForm.lstCompany.focus();
    return (false);
  }

  if (theForm.lstCompany.selectedIndex == 0)
  {
    alert("The first \"Company Selection\" option is not a valid selection.  Please choose one of the other options.");
    theForm.lstCompany.focus();
    return (false);
  }
<% End Select %>

  if (theForm.txtDateStart.value == "")
  {
    alert("Please enter a value for the \"Start date\" field.");
    //theForm.txtDateStart.focus();
    return (false);
  }

  if (theForm.txtDateEnd.value == "")
  {
    alert("Please enter a value for the \"End Date\" field.");
    //theForm.txtDateEnd.focus();
    return (false);
  }
  theForm.txtDateStart.disabled = 0;
  theForm.txtDateEnd.disabled = 0;
  return (true);
}
//--></script>


<script language="JavaScript">
	// *** Script for callback to dynamically populate the offices list from company selection ***
	
	//var strCode; //make page scope var
	var g_intervalID_Populate_lstOffices

	function Load_lstOfficesScript() {	
		var InsCoName = "<%=mUserInsCoName%>"
		if (document.getElementById('lstCompany')) {
			var olstCompany = document.getElementById('lstCompany');
			if (olstCompany.options[olstCompany.selectedIndex].value == '0') return;
			var InsCoName = olstCompany.options[olstCompany.selectedIndex].value 
		}
		strCode = "";
		
		var remoteServer = "http://webportal.ace-it.com/Admin/CoOffList/" + InsCoName + "/";
		if (document.getElementById('chkShowACEOfficeCodes').checked == true) 
			remoteServer = "http://webportal.ace-it.com/Admin/CoOffListWithACECode/" + InsCoName + "/";
		
		var head = document.getElementsByTagName('head').item(0);
		var old  = document.getElementById('jsLoadOffices');
		if (old) head.removeChild(old);
		var Script = document.createElement('script');
		Script.src = remoteServer;
		Script.type = 'text/javascript';
		Script.defer = true;
		Script.id = 'jsLoadOffices';
		Script.name = 'jsLoadOffices';
		//void(head.appendChild(Script));
		head.appendChild(Script);
				
		Populate_lstOffices();		
	}
	
	function Populate_lstOffices() {
		var lstOfficesOptions = document.getElementById('lstOffices').options
		lstOfficesOptions.length = 0; //clear
		lstOfficesOptions[0] = new Option("Loading Offices, Please Wait ...", "");
		//ShowLoading();
		
		clearInterval(g_intervalID_Populate_lstOffices);
		if (strCode == ""){
			g_intervalID_Populate_lstOffices = setInterval(Populate_lstOffices, 1000);
			return;
		}
		
		//document.getElementById('txt1').innerText = strCode;
		
		// parse returned HTML -- I initially created the company, office and rep callbacks for straight HTML via javascript document.write so we have to parse out the values to use via javascript object model
		var aOptions = strCode.split("<option "); 
		var i; var x; var y; var z
		for(i = 1; i < aOptions.length; i++) { //we start at i=1 instead of i=0 because the first element needs to be ignored since it is the stuff that preceeded the first <option "
			x = aOptions[i].indexOf('value="') + 7;
			y = aOptions[i].indexOf('">', x);
			z = aOptions[i].indexOf('</option>', y);
			lstOfficesOptions[i-1] = new Option(aOptions[i].substring(y+2,z), aOptions[i].substring(x,y)); // y+2,z coords = what's displayed x,y=value
		}
		if (aOptions.length == 3) lstOfficesOptions[0] = null //If only one office remove ALL option, length ==3 because we ignore element [0] since it is the text that preceeds the <option " for the split
		lstOfficesOptions[0].selected = true;
		//HideLoading();
	}
	
	// *** Script for callback to dynamically populate the offices list from company selection END ***
</script>


<!--<table width="100%" border="2" cellspacing="3" style="border-collapse: collapse"> -->
<!--<table width="100%" border="0" cellspacing="3" style="border-collapse: collapse"> -->
<table border="0" cellspacing="3" style="border-collapse: collapse">
  <tr>
    <td align="left" width="100%">
    <!--<form method="POST" target="_blank" name="frmReports" action="frmMain.asp?t=<%=mActiveTab%>&txtState=1" onsubmit="return FrmReports_Validator(this)">-->
    <form method="POST" target="_blank" name="frmReports" action="http://webportal.ace-it.com/ReportEx" onsubmit="return FrmReports_Validator(this)">
      <input type="hidden" name="txtUserCompanyName" value="<%=mUserInsCoName%>"><input type="hidden" name="txtUserOfficeName" value="<%=mUserACEInsCoOfficeName%>">
      <table border="0" cellspacing="3" width="100%" cellpadding="3">
<!--        <tr>
          <td align="right" width="33%">&nbsp;</td>
          <td width="66%">&nbsp;</td>
        </tr>
-->        
        <tr>
          <td align="right" class="styLabelCell" width="33%">Select a Report:</td>
          <td width="66%"><select size="4" name="lstReport" style="font-family: Courier New, Courier New (monospace), Lucida Console, Courier, monospace;">
          <option value="1" selected>Auto Results - Summary</option>
          <option value="2">Auto Results - Detail</option>
          <option value="3">Parts Analysis - Total</option>
          <option value="4">Claim Type Analysis by Office &amp; State - Summary</option>
          <option value="0">--- Net Results Reports ---</option>
          <option value="5">Auto Net Results - Summary</option>
          <option value="6">Auto Net Results - Detail</option>
          </select> </td>
        </tr>
<%
  Select Case mUserTypeID 
    Case cMemberType_Administrator, cMemberType_ACEEmployee, cMemberType_ACEAdmin
%>        
<!--
        <tr>
          <td align="right" class="styLabelCell" width="33%">Select Report Format:</td>
          <td bgcolor="#EEEEEE" width="66%">
            <input type="radio" value="1" checked id="optOutputFormat" name="optOutputFormat"><img height="32" alt="Adobe PDF icon" src="../Images/Apps/pdficon_large.gif" width="32" border="0" onclick="Javascript:document.getElementById('optOutputFormat').value='1';">&nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="optOutputFormat" value="2"><img height="32" alt="Excel icon" src="../Images/Apps/Excel-32.gif" width="32">
          </td>
        </tr>
-->
        <tr>
          <td align="right" class="styLabelCell" width="33%">Select Company:</td>
          <td width="66%"><select size="5" id="lstCompany" name="lstCompany" onchange="Javascript: void(Load_lstOfficesScript());" style="font-family: Courier New, Courier New (monospace), Lucida Console, Courier, monospace;">            
				<script language="JavaScript" src="http://webportal.ace-it.com/Admin/CoList/">document.write("<br>Database server did not respond<br>");</script>
				<script language="JavaScript">
					//build control and and set default to current company
					document.write(strCode.replace(/<option selected value="0">/i, '<option value="0">')
	                      .replace('<option value="<%=mUserInsCoName%>">', '<option selected value="<%=mUserInsCoName%>">'));
				</script>       
            </select></td>
        </tr>
		<%=fOfficeSelection()%>
        
		<!-- !!! When changes are made to the row below copy them to the co mgr case also -->
	<%Private Function fOfficeSelection
		If false Then
	%>
        <tr>
          <td align="right" class="styLabelCell" width="33%">Include Office(s):<br>
          <span style="font-weight: 400"><font face="Arial" color="#808080">To select multiple offices<br>
          use Shift+Click <br>
          or CTRL+Click<br>
&nbsp;</font></span> <br>
          <input type="checkbox" id="chkShowACEOfficeCodes" name="chkShowACEOfficeCodes" value="ON" onclick="Javascript: void(Load_lstOfficesScript());"><font face="Arial" size="1" color="#808080">Show ACE office codes</font>
			&nbsp;</td>
          <td width="66%"><select size="5" id="lstOffices" name="lstOffices" multiple style="font-family: Courier New, Courier New (monospace), Lucida Console, Courier, monospace;">
				<script language="JavaScript" src="http://webportal.ace-it.com/Admin/CoOffList/<%=mUserInsCoName%>/">document.write("<br>Database server did not respond<br>");</script>
				<script language="JavaScript">document.write(strCode);</script>
            </select></td>
        </tr>
 <%
    Else
	%>
        <tr>
          <td align="right" class="styLabelCell" width="33%">Include Office(s):<br>
          <span style="font-weight: 400"><font face="Arial" color="#808080">To select multiple offices<br>
          use Shift+Click <br>
          or CTRL+Click<br>
          &nbsp;</font>
          </span> <br>
        </td>
          <td width="66%"><select size="5" id="lstOffices" name="lstOffices" multiple style="font-family: Courier New, Courier New (monospace), Lucida Console, Courier, monospace;">
<%
        Dim oData, oSQL
        Dim rs
        Dim RecCount
        
        Set oData = New clsACEDataConn

        '--- Populate recordset ---
        'Set rs = oData.rsCallStoredProc("Portal_CoOfficeList", "InsCoListHTML", mUserInsCoName)
        Dim cmd
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.CursorLocation = 3 'adUseClient
        'Set cn = Server.CreateObject("ADODB.Connection")
        Set cmd = Server.CreateObject("ADODB.Command")
        'cn.Open mConnStr
        'Set cmd.ActiveConnection = cn
        cmd.ActiveConnection = oData.mConnStr
        cmd.CommandText = "Portal_CoOfficeList"
        cmd.CommandType = 4
        ' Ask the server about the parameters for the stored proc
        cmd.Parameters.Refresh
        ' Assign values to parameters. Index of 0 represents first parameter.
        cmd.Parameters(1) = "InsCoListHTML"
        cmd.Parameters(2) = mUserInsCoName
        rs.Open cmd

        rs.MoveFirst
%>
   <%=rs(0)%>
      <% 
      rs.Close
      %>


            </select></td>
        </tr>
 <%
    
		End If
	  End Function%>
<% Case cMemberType_CoOfficeMgr %>
        <tr>
          <td align="right" class="styLabelCell" width="33%">Company:</td>
          <td width="66%"><%=mUserInsCoName%></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell" width="33%">Office:</td>
          <td width="66%"><%=mUserACEInsCoOfficeName%></td>
        </tr>
<% Case cMemberType_CoMgr %>
        <tr>
          <td align="right" class="styLabelCell" width="33%">Company:</td>
          <td width="66%"><%=mUserInsCoName%></td>
        </tr>
<!--
        <tr>
          <td align="right" class="styLabelCell" width="33%">Office:</td>
          <td width="66%">ALL</td>
        </tr>
-->
		<!-- This is a copy of the table row in the ACE employee case -->
		<%=fOfficeSelection()%>
<% End Select %>        
<!--
        <tr>
          <td align="right" class="styLabelCell" width="33%">Include Office(s):</td>
          <td bgcolor="#EEEEEE" width="66%"><select size="4" name="lstOffices" multiple>
          <option value="B01A" selected>Office1</option>
          <option value="B01B" selected>Office2</option>
          <option value="B01C" selected>Office3</option>
          </select>&nbsp; </td>
        </tr>
-->        
        <tr>
          <a name="bkmarkBot"></a>
          <td colspan="2"><b><font face="Verdana" size="1" color="#808080">Use the dropdown arrows below to select a date range for your report. Please note: 
          When running reports after 12:00 pm EDT, files available for reporting are delayed one day (the files closed today will not appear). When running 
          reports before noon EDT, files are delayed two days.</font></b></td>
        </tr>
        <tr>
          <td align="right" class="styLabelCell" width="33%">Start Date:</td>
          <td width="66%"> <input type="text" name="txtDateStart" size="13" disabled="1"><img id="btnStartDt" border="0" src="../Images/calendar.gif" width="16px" height="19px" onclick="javascript:scwShow(document.frmReports.txtDateStart, event);">
          <!--<a href="#" onClick="getCalendarFor(document.frmReports.txtDateStart);return false"><img border="0" src="../JavaScript/SpiffyPopupCal/btn_date2_up.gif"></a>-->
          </td> 
        </tr>
        <tr>
          <td align="right" class="styLabelCell" width="33%">End Date:</td>
          <td width="66%">
          <input type="text" id="txtDateEnd" name="txtDateEnd" size="13" disabled="1" value="<%=Date()-2%>"><img id="btnEndDt" border="0" src="../Images/calendar.gif" width="16" height="19" onclick="javascript:scwShow(document.frmReports.txtDateEnd, event);">
          <!--<a href="#" onClick="getCalendarFor(document.frmReports.txtDateEnd);txtDateEnd.value=document.body.clientHeight; return false;"></a>-->
          </td>
        </tr>
        <tr>
          <td align="right" width="33%"><img border="0" src="../images/spacer.gif" width="1" height="3"></td>
          <td width="66%"><img border="0" src="../images/spacer.gif" width="1" height="3"></td>
        </tr>
        <tr>
          <td align="right" width="33%">&nbsp;</td>
          <td width="66%"><input type="submit" value="View Report" name="cmdSubmit"></td>
        </tr>
      </table>
    </form>
    <!--
    <table border="0" cellspacing="0" width="100%" cellpadding="0">
      <tr>
        <td width="100%" bgcolor="#DEDEDE"><img border="0" src="../images/spacer.gif"></td>
      </tr>
    </table>
    -->
    <table border="0" cellspacing="3" width="100%" cellpadding="3" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td align="center" width="100%">
        <font face="Arial" size="2" color="#808080"><p><br>
        <a target="_blank" href="http://www.adobe.com/products/acrobat/readstep2.html"><img border="0" src="../Images/getacro.gif" width="88" height="31"></a><br>
		Reports are provided in Adobe PDF format.<br>
        You must have the</font>
        <a href="http://www.adobe.com/products/acrobat/readstep2.html"><font color="#808080">Adobe Acrobat plug-in </font> </a>
        <font color="#808080">to view/print them.</font></font>
        <br>
        <font color="#808080" face="Arial" size="2">
        &nbsp;</font></td>
      </tr>
      <tr>
        <td align="center" width="100%">
        <img border="0" src="../images/spacer.gif" width="1" height="30"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<script Language="JavaScript" Type="text/javascript"><!--
  m_dtNow = new Date(); //m_dtNow.setYear(m_dtNow.getYear());
  m_dtFOM = new Date(m_dtNow.getFullYear(), m_dtNow.getMonth() - 1, 1);
  m_dtLOM = new Date(m_dtFOM.getFullYear(), m_dtFOM.getMonth() + 1, 0);

  document.frmReports.txtDateStart.value = FormatDate(m_dtFOM);
  document.frmReports.txtDateEnd.value = FormatDate(m_dtLOM);
  EndDateCheck(document.frmReports.txtDateEnd, 1);
--></script>

<Script>placeFocus();</Script>
<% Option Explicit %>

<!-- #INCLUDE Virtual="/ACE/Login/Security.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsACEDataConn.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsSQLStmt.asp" -->
<%
  Dim sURL, URLTemp 
  
  sURL = "http://" & _
	    Request.ServerVariables("SERVER_NAME") & _
        Request.ServerVariables("SCRIPT_NAME")

	'URLTemp = Request.ServerVariables("QUERY_STRING")
	If URLTemp & "" <> "" Then sURL = sURL & "?" & URLTemp
  
  Dim sPage
  
  sPage = Request.Querystring("Page")
  If sPage = "" Then sPage = Session("ClaimsListPage")
  If sPage = "" Then sPage = "A"
  Session("ClaimsListPage") = sPage

  Dim oData, oSQL
  Dim rsClaims, rsCompanies, rsOffices
  Dim RecCount
  
  Set oData = New clsACEDataConn
  'Set oSQL = New clsSQLStmt
    
  Dim sOwner, sStat, sType, sFind, sRcvDt, sSort, sFilterCompany, sFilterOffice
  
  
  sFind = Request.Querystring("Find")
  sOwner = Request.Querystring("Owner")
  sType = Request.Querystring("Type")
  If sType & "" = "" Then sType = "All"
  sStat = Request.Querystring("Stat")
  sRcvDt = Request.Querystring("RcvDt")
  If sRcvDt & "" = "" or sRcvDt = "0" Then sRcvDt = 30
  sSort = Request.Querystring("Sort")
  If sSort & "" = "" Then sSort = "1"
  sFilterCompany = Request.Querystring("FilterCompany")
  sFilterOffice = Request.Querystring("FilterOffice")
  
  If Len(sFilterCompany) = 0 Then sFilterCompany = mUserInsCoName
  If Len(sFilterOffice) = 0 Then sFilterOffice = "ALL"
  
  '--- If either claim # or owner name set then make search max period ---
  If len(sOwner) > 0 Or Len(sFind) > 0 Then sRcvDt = "150"
    
  '--- Populate recordset ---
  'Set rsClaims = oData.rsClaimList(mUserInsCoName, mUserACEInsCoOfficeID, mUserACEInsCoOfficeName, sOwner, sStat, sClaimNum)
  Set rsClaims = oData.rsClaimList(mUserACEInsCoOfficeID, Trim(sOwner), Trim(sFind), sStat, sRcvDt, sSort, sType, sFilterCompany, sFilterOffice)
  RecCount = rsClaims.RecordCount
  
  '--- If ACE employee allow to switch company
  If RecCount > 0 And sFilterCompany <> mUserInsCoName And mUserSecurityL >= 5 Then 
    '--- update the session variables defined in SECURITY.asp
	Session("InsCompanyName") = sFilterCompany
	mUserInsCoName = Session("InsCompanyName")
	Session("ACEInsCoOfficeID") = rsClaims("InsuranceCompanyIDPK")
	mUserACEInsCoOfficeID = Session("ACEInsCoOfficeID")
	Session("ACEInsCoOfficeName") = rsClaims("InsuranceCompanyOfficeLocation")
	mUserACEInsCoOfficeName = Session("ACEInsCoOfficeName")
  End If


  Dim OffsetTop
  If mUserSecurityL >= 5 Then 
	  '--- Get company HTML 
	  Dim sCompanyHTMLOptionList
	  sCompanyHTMLOptionList = oData.sCompanyHTMLOptionList(sFilterCompany)
  End If
  
%>
<!DOCTYPE html>
<html>

<head>
<meta http-equiv="expires" content="Tue, 20 Aug 1996 01:00:00 GMT">
<meta http-equiv="Pragma" content="no-cache">
<meta name="Author" content="Bob Puppo">
<meta name="Copyright" content="(C) Bob Puppo, 1999-2019, All rights reserved">
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ACE - Claims List</title>
<link rel="stylesheet" id="PEStyles" type="text/css" href="PE_Styles.css">
<style>
	.loader {
	  border: 3px solid #f3f3f3;
	  border-top: 3px solid blue;
	  border-right: 3px solid blue;
	  #border-bottom: 3px solid blue;
	  border-radius: 50%;
	  width: 15px;
	  height: 15px;
	  -webkit-animation: spin 2s linear infinite;
	  animation: spin 2s linear infinite;
	}

	@-webkit-keyframes spin {
	  0% { -webkit-transform: rotate(0deg); }
	  100% { -webkit-transform: rotate(360deg); }
	}

	@keyframes spin {
	  0% { transform: rotate(0deg); }
	  100% { transform: rotate(360deg); }
	}
</style>


<meta name="Microsoft Border" content="b">

<META HTTP-EQUIV="refresh" CONTENT="900">
<SCRIPT LANGUAGE="JavaScript">
	function ShowFileList(estID) {
		document.frmFileAttachmentsList.txtEstID.value = estID;
		document.frmFileAttachmentsList.submit();
	}
</Script>
</head>

<body>
<form id="frmFileAttachmentsList" name="frmFileAttachmentsList" method="POST" action="frmAdditionalFilesList.asp" target="_blank">
	<!--- form to launch an audits file list - including original assiment files. --> 
	<INPUT type="hidden" name="txtState" value="FileList">
	<INPUT type="hidden" name="txtEstID" value="">
</form>

<!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">

<script id="sKeepAlive"></script>

<%
	Dim addon

    addon = ""

	If Trim(LCase(mUserInsCoName)) = "westfield acs" Or Trim(LCase(mUserInsCoName)) = "westfield subro" Or Trim(LCase(mUserInsCoName)) = "westfield insurance" Then
        addon = "Westfield"
    End If
%>

<% If Request.QueryString("SupressHdr") & "" = "" Then %>
<table width="100%" style="border-collapse: collapse" bordercolor="#111111" cellpadding="1" cellspacing="1" border="0">
  <tr><td width="50%"><b><font size="2" face="Arial">&nbsp;
    <a href="frmMain.asp"><font color="#333399">&#9668; Back to Portal</font></a></font></b><a href="frmMain.asp"><font color="#333399"></font></a>
	<img id="keepAliveIMG" img border="0" width="1" height="1" src="../Images/spacer.gif" /><br>
    <font size="2" face="Arial"> 
    <b>&nbsp;&nbsp;&nbsp;&nbsp; <a href="JavaScript:OpenChildWindow('frmSubmitAutoReview<%=addon %>.asp','ACE',windowOpt);void(0);"><font color="#333399">Submit an Auto Claim</font></a></font></b>
    <br>
    <b><font color="#0000FF" size="2" face="Arial">
    <font color="#333399">&nbsp;&nbsp;&nbsp;&nbsp; </font>
    <a href="JavaScript:OpenChildWindow('frmSubmitPropReview.asp','ACE',windowOpt);void(0);"><font color="#333399">Submit a Property Claim</font></a></font></b>   
    <br><b>&nbsp;&nbsp;&nbsp;&nbsp; <a href="JavaScript:OpenChildWindow('frmSubmitPhotoQuote.asp','ACE',windowOpt);void(0);"><font color="#333399">Submit a PhotoQuote</font></a></b>
    </td>
    <td><img border="0" src="../Images/ace_logo.gif"></td>
    <td width="50%">&nbsp;</td>
    </tr>
<!--<tr><td colspan="3" align="center"><img border="0" src="../Images/spacer.gif" width="1" height="2"></td></tr> -->
</table>
<%End If%>

<script LANGUAGE="JavaScript">
  var normDPI = 96;
  if ((screen.deviceXDPI == screen.logicalXDPI) && (screen.deviceXDPI > normDPI)) {

  }  

  function fnDPIScaleFactorX() {
    return screen.deviceXDPI / screen.logicalXDPI;

  }
  
</script>

<script LANGUAGE="JavaScript">
function ApplyFilter(sFilter) {
  /* reload page with filter parameters set in query string */
  document.getElementById("loader").style.display = "block";

  //lnkFilter.style.backgroundColor = "";
  
  //var frmFilter = document.forms[0];
  var frmFitler = document.frmFilter;
  
  var sPage = "";
  var sFind = "";
  
  switch (sFilter) {
    case "Find":
      sFind = frmFilter.txtFind.value;
      break;
    default: 
      sPage = sFilter;
  }
/*

  sFind = frmFilter.txtFind.value;	//-- this is the claim #
*/
  //window.location = "frmClaims_List.asp?Stat=" + frmFilter.cmbStatusID.value + "&Owner=" + frmFilter.txtOwner.value + "&Find=" + sFind + "&RcvDt=" + frmFilter.cmbRcvDt.value + "&Sort=" + frmFilter.cmbSortID.value + "&Page=" + sPage + "&SupressHdr=<%=Request.QueryString("SupressHdr") & ""%>";
  
  window.location = "frmClaims_List.asp?Stat=" + frmFilter.cmbStatusID.value 
					+ "&Owner=" + frmFilter.txtOwner.value 
					+ "&Find=" + frmFilter.txtFind.value	//-- this is the claim # 
					+ "&RcvDt=" + frmFilter.cmbRcvDt.value 
					+ "&Sort=" + frmFilter.cmbSortID.value 
					+ "&Type=" + frmFilter.cmbType.value
					+ "&Page=" + sPage 
					+ "&SupressHdr=<%=Request.QueryString("SupressHdr") & ""%>"
					+ (!frmFilter.cmbFilterCompany ? "" : "&FilterCompany=" + frmFilter.cmbFilterCompany.value) 
					+ (!frmFilter.cmbFilterOffice ? "" : "&FilterOffice=" + frmFilter.cmbFilterOffice.value) 
					;
					
}
</script>

<script>
function OpenChildWindow(sPath,sName,sOptions) {
  //self.name = "cover";
  var x = window.open(sPath,sName,sOptions);
  if(parseInt(navigator.appVersion)>3) {x.focus();}
  return x;
}

  var windowOpt = 'dependent,scrollbars=yes,location=no,directories=no,status=yes,menubar=yes,toolbar=yes,resizable=yes,screenX=0,screenY=0,width=500,height=' + (screen.availHeight - 60);
</script>

<table border="0" cellspacing="0" style="border-collapse: collapse;" bordercolor="#111111" width="100%" cellpadding="0">
  <tr>
    <td width=10% bgcolor="#1A74A1" align="Left" nowrap>&nbsp;
    <!-- <b><font color="#0000FF" size="2" face="WingDings">�</font></b> -->
    </td>
    <td width="80%" bgcolor="#1A74A1" align="center" nowrap>
    <b><font color="#FFFFFF" size="3" face="Arial">Claims List - Click the file you wish to view.</font></b></td>
    <td width="10%" bgcolor="#1A74A1" align="right" nowrap>
      <!-- <b><a accesskey="R" href="JavaScript:window.location.reload();"><font color="#FFFF99">Refresh</font></a></b> -->
      <font color="#FFFF99" size="2" face="Arial">Count: <%=RecCount%> &nbsp;&nbsp;&nbsp;</font>
    </td>
  </tr>
  <!--
  <tr>
    <td width="100%" colspan="3"><img border="0" src="../Images/spacer.gif" width="1" height="2"></td>
  </tr>
  -->
  <tr>
	<td colspan="3" bgcolor="#BBC7CE" style="padding-bottom: 5px; padding-left: 5px;">
		<form id="frmFilter" name="frmFilter" method="POST" action="JavaScript:ApplyFilter('Find');">
			<!--<button name="B3" type="submit" style="width: 0; height: 0">1</button>-->

			<style>
			table#tblFilters {
				width: 100%;
				border-collapse: collapse;
				border: 0px
				margin-top: 10px;
				/*
				POSITION: absolute;
				Z-INDEX: 1;
				LEFT: 23;
				TOP: 130;
				*/
			}

			table#tblFilters th, table#tblFilters td {
				padding-top: 0px;
				padding-bottom: 0px;
				padding-right: 10px;
			}

			/*
			table#tblFilters tr:nth-child(even) {
				background-color: #eee;
			}
			table#tblFilters tr:nth-child(odd) {
				background-color: #fff;
			}
			*/

			</style>

			<table id = "tblFilters">
				<% If mUserSecurityL >= 5 And Len(sCompanyHTMLOptionList) > 0 Then 
				%>
				<tr>
					<td colspan=2 nowrap style="white-space: nowrap">
						<span class="PE_Label">Company:</span>
					</td>
					<% If 0 Then %>
					<td nowrap style="white-space: nowrap">
						<span class="PE_Label">Office:</span>
					</td>
					<%Else%>
					<td>
					</td>					
					<%End If%>
					<td colspan=5>
					</td>
				</tr>
				<tr>
					<td colspan=2>
						<select class="PE_ComboBox" size="1" name="cmbFilterCompany" id="cmbFilterCompany" 
							Style="WIDTH: 220px;">
						<%=sCompanyHTMLOptionList%>
						</select> 
					</td>
					<% If 0 Then %>
					<td>
						<select class="PE_ComboBox" size="1" name="cmbFilterOffice" id="cmbFilterOffice" 
							Style="WIDTH: 150px;">
						<option <%IF sType = "ALL" Then Response.Write("selected")%> value="ALL">ALL</option>
						</select> 
					</td>
					<%Else%>
					<td>
					</td>					
					<%End If%>
					<td colspan=5>
					</td>
				</tr>
				<% End If %>
			  

				<tr>
					<td nowrap style="white-space: nowrap">
						<span class="PE_Label">Claim # Starts With:</span>
					</td>
					<td nowrap style="white-space: nowrap">
						<span class="PE_Label">Owner Last Name Starts With:</span>
					</td>
					<td nowrap style="white-space: nowrap">
						<span class="PE_Label">Type:</span>
					</td>
					<td nowrap style="white-space: nowrap">
						<span class="PE_Label">Status:</span>
					</td>
					<td nowrap style="white-space: nowrap">
						<span class="PE_Label">Received Within:</span>
					</td>
					<td nowrap style="white-space: nowrap">
						<span class="PE_Label">Sort By:</span>
					</td>
					<td nowrap style="white-space: nowrap">
					</td>
					<td nowrap style="white-space: nowrap" Width="100%">
					</td>
				</tr>
				<tr>
					<td>
						<input type="text" class="PE_Textbox" name="txtFind" id="txtFind" size="20" value="<%=sFind%>" 
						Style="WIDTH: 111px;" tabIndex="1">
					</td>
					<td>
						<input type="text" class="PE_Textbox" name="txtOwner" id="txtOwner" size="20" value="<%=sOwner%>" 
						Style="WIDTH: 150px;">
					</td>
					<td>
						<select class="PE_ComboBox" size="1" name="cmbType" id="cmbType" 
							Style="WIDTH: .75in;">
						<option <%IF sType = "ALL" Then Response.Write("selected")%> value="ALL">ALL</option>
						<option <%IF sType = "Auto" Then Response.Write("selected")%> value="Auto">Auto</option>
						<option <%IF sType = "Subro" Then Response.Write("selected")%> value="Subro">Subro</option>
						<option <%IF sType = "Heavy" Then Response.Write("selected")%> value="Heavy">Heavy</option>
						<option <%IF sType = "Over5K" Then Response.Write("selected")%> value="Over5K">Over5K</option>
						<option <%IF sType = "PhotoQuote" Then Response.Write("selected")%> value="PhotoQuote">PhotoQuote</option>
						</select> 
					</td>
					<td>
						<select class="PE_ComboBox" size="1" name="cmbStatusID" id="cmbStatusID" 
							Style="WIDTH: .75in;">
						<option <%IF sStat = 0 Then Response.Write("selected")%> value="0">ALL</option>
						<option <%IF sStat = 1 Then Response.Write("selected")%> value="1">Closed</option>
						<option <%IF sStat = 2 Then Response.Write("selected")%> value="2">Pending</option>
						<option <%IF sStat = 3 Then Response.Write("selected")%> value="3">Cancelled</option>
						</select> 
					</td>
					<td>
						<select class="PE_ComboBox" size="1" name="cmbRcvDt" id="cmbRcvDt" 
						Style="WIDTH: 91px;">

						<option <%IF sRcvDt = 1 Then Response.Write("selected")%> value="1">1 days</option>
						<option <%IF sRcvDt = 5 Then Response.Write("selected")%> value="5">5 days</option>
						<option <%IF sRcvDt = 10 Then Response.Write("selected")%> value="10">10 days</option>
						<option <%IF sRcvDt = 15 Then Response.Write("selected")%> value="15">15 days</option>
						<option <%IF sRcvDt = 30 Then Response.Write("selected")%> value="30">30 days</option>
						<option <%IF sRcvDt = 60 Then Response.Write("selected")%> value="60">60 days</option>
						<option <%IF sRcvDt = 90 Then Response.Write("selected")%> value="90">90 days</option>
						<option <%IF sRcvDt = 120 Then Response.Write("selected")%> value="120">120 days</option>
						<option <%IF sRcvDt = 150 Then Response.Write("selected")%> value="150">150 days</option>
						<option <%IF sRcvDt = 365 Then Response.Write("selected")%> value="365">365 days</option>

						</select> 
					</td>
					<td>
						<select class="PE_ComboBox" size="1" name="cmbSortID" id="cmbSortID" 
						Style="WIDTH: 1in;">
						<option <%IF sSort = 0 Then Response.Write("selected")%> value="0">Loss Date</option>
						<option <%IF sSort = 3 Then Response.Write("selected")%> value="3">Closed Date</option>
						<option <%IF sSort = 1 Then Response.Write("selected")%> value="1">Received Date</option>
						<option <%IF sSort = 2 Then Response.Write("selected")%> value="2">Claim #</option>
						</select> 
					</td>
					<td nowrap style="white-space: nowrap">
						<a class="PE_Label" id="lnkFilter" Style="FONT-SIZE: 10pt;" href="JavaScript:ApplyFilter('Find');">Apply Filter</a>
					</td>
					<td Width="100%">
						<div id="loader" class="loader"></div>
					</td>
				</tr>


			</table>


		</form>
		<script>
			var frm = document.getElementById("frmFilter");
			frm.txtFind.addEventListener("keypress", highlightApplyFilter);
			frm.txtOwner.addEventListener("keypress", highlightApplyFilter);
			frm.cmbType.addEventListener("change", highlightApplyFilter);
			frm.cmbStatusID.addEventListener("change", highlightApplyFilter);
			frm.cmbRcvDt.addEventListener("change", highlightApplyFilter);
			frm.cmbSortID.addEventListener("change", highlightApplyFilter);
			if(!(!frm.cmbFilterCompany)) frm.cmbFilterCompany.addEventListener("change", highlightApplyFilter);

			function highlightApplyFilter() {
				lnkFilter.style.backgroundColor = "yellow";
			}
		</script>

	</td>
  </tr>
</table>

<style type="text/css" media="print,screen" >  
  th {  
     font-family:Arial;  
     color:white;  
     background-color: #1A74A1;  
     background-image: url('../Images/Gradients/Aqua.gif');
  }  
  thead {  
     display:table-header-group;  
  }  
  tbody {  
     display:table-row-group;  
  }
  td {  
     font-family: Arial;  
     font-size: 10pt;
  }  
    
</style>  


<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <thead><tr>
    <th width="15%">Claim #</th>
    <th width="0%"> </th>
    <th width="0%"> </th>
    <th width="30%">Vehicle Owner</th>
    <th width="30%">Insured</th>
    <th width="5%">Rcv.<br>Date</th>
    <th width="5%">Closed<br>Date</th>
    <th width="5%">Shop Estimate</th>
    <th width="5%">ACE Audit</th>
    <th width="5%">Status</th>
  </tr></thead>
  <tbody>
  
  
<%
  
  
  
  
  Dim sBColor, sFont, EvenRow
  Dim sCoURL, sPersonURL, sEditURL, sClaimActivityURL
  
  sFont = " <font face=""Arial"" size=""2"">"
  EvenRow = False
  
  
  If rsClaims.EOF Then
    Response.Write _
      "<tr><td colspan=""7"">... No Files meet your filter criteria ...</td><tr>"
  End If
  
  Do While Not rsClaims.EOF
    If EvenRow Then sBColor = " bgcolor=""#E0E0E0"" " Else sBColor = " bgcolor=""#FFFFFF"" "

    sEditURL = rsClaims("EstimateIDPK") & "_ACE_Audit.pdf"

    sClaimActivityURL = "frmViewClaimStatusActivity.asp?EstID=" & rsClaims("EstimateIDPK") & "&ref=" & Server.URLEncode(sURL)

	Dim DetailLink
	Dim PhotosLink
	Dim StatusLink
	Dim FilesLink	'--- 2017-08-14
	
	DetailLink = " OnClick=""ShowFile('" & sEditURL & "', 'view', " & rsClaims("EstimateIDPK") & ");"" "
	If rsClaims("Status") <> "Closed" Then
		DetailLink = " OnClick=""alert('Audit not available. This file is not yet closed.');"" "
	End If

	StatusLink = "<a target=""_blank"" href=""rptStatusSheet.asp?EstID=" & rsClaims("EstimateIDPK") & """>" & rsClaims("Status") & "</a>"
		
	PhotosLink = ""
	If rsClaims("HasPhotos")=1 Then
	  PhotosLink = " <img border=""0"" src=""../Images/CAMERA.gif"" width=""16"" height=""16"" alt=""This file has photos attached"">"
	  PhotosLink = "<td width=""0%"" title=""Click to view file photos"" OnClick=""window.open ('frmViewPhotos.asp?EstID=" & rsClaims("EstimateIDPK") & "', '', 'resizable=1, scrollbars=1, toolbar=0, menubar=0, location=0');"">" & sFont &  PhotosLink & "</font></td>"
	Else
	  PhotosLink = "<td " & DetailLink & " width=""0%"">" & sFont & " </font></td>"
	End If

	'--- 2017-08-14
	FilesLink = ""

	If rsClaims("IsPhotoQuote")=1 Then
	  FilesLink = " <img border=""0"" src=""../Images/folderAndCamera.png"" width=""16"" height=""16"" alt=""This is a PhotoQuote file and has photos and additional files attached"">"
	  FilesLink = "<td width=""0%"" title=""PhotoQuote: Click to view photos and additional file attachements to this file"" OnClick=""ShowFileList(" & rsClaims("EstimateIDPK") & ");"">" & sFont &  FilesLink & "</font></td>"
	ElseIf rsClaims("HasAdditionalFiles")=1 Then
	  FilesLink = " <img border=""0"" src=""../Images/folder.png"" width=""16"" height=""16"" alt=""This file has additional files attached"">"
	  FilesLink = "<td width=""0%"" title=""Click to view additional file attachements to this file"" OnClick=""ShowFileList(" & rsClaims("EstimateIDPK") & ");"">" & sFont &  FilesLink & "</font></td>"
	Else
	  FilesLink = "<td " & DetailLink & " width=""0%"">" & sFont & " </font></td>"
	End If
	
    Response.Write _
                   "<tr onMouseOver=""this.bgColor = '#FFFF00';this.style.cursor='pointer';"" onMouseOut =""this.bgColor = '" & Mid(sBColor, 11, 7) & "'"" " & sBColor & ">" & _
                     "<td " & DetailLink & "width=""15%"">" & Server.HTMLEncode(rsClaims("ClaimNo") & "") & "</td>" & _
                     FilesLink & _
                     PhotosLink & _
                     "<td " & DetailLink & "width=""30%"">" & Server.HTMLEncode(rsClaims("Owner") & "") & "</td>" & _
                     "<td " & DetailLink & "width=""30%"">" & Server.HTMLEncode(rsClaims("Insured") & "") & "</td>" & _
                     "<td " & DetailLink & "width=""5%"" align=""center"">" & Server.HTMLEncode(rsClaims("ReceivedDate") & "") & "</td>" & _
                     "<td " & DetailLink & "width=""5%"" align=""center"">" & Server.HTMLEncode(rsClaims("ClosedDate") & "") & "</td>" & _
                     "<td " & DetailLink & "width=""5%"" align=""right"">" & Server.HTMLEncode(rsClaims("BodyShopEstAmt") & "") & "</td>" & _
                     "<td " & DetailLink & "width=""5%"" align=""right"">" & Server.HTMLEncode(rsClaims("AgreedPrice") & "") & "</td>" & _
                     "<td width=""5%"" align=""right"">" & StatusLink & "</td>" & _
                   "</tr>"


    EvenRow = Not EvenRow
    rsClaims.MoveNext


  Loop
  rsClaims.close
%>
</tbody>
</table>


<form id="frmViewFile" name="frmViewFile" method="POST" action="frmAdditionalFilesList.asp" target="_blank">
	<!-- Form to show file content -->
	<INPUT type="hidden" name="txtState" value="Show" >
	<INPUT type="hidden" name="txtEstID" value="">
	<INPUT type="hidden" name="txtFileName" value="">
	<INPUT type="hidden" name="txtAction" value="download">
</form>
<SCRIPT LANGUAGE="JavaScript">
	function ShowFile(fileName, action, EstID) {
		document.frmViewFile.txtFileName.value = fileName;
		document.frmViewFile.txtAction.value = action;
		document.frmViewFile.txtEstID.value = EstID;
		document.frmViewFile.submit();
	}
</Script>



<!--<span Style="POSITION: absolute; LEFT: 0; TOP: <%=OffsetTop%>">-->
<!--
<span Style="POSITION: absolute; LEFT: 10; TOP: 125">

</span>
-->

<!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<hr>
<p><font size="1" color="#666633">Unless otherwise noted all content �
<a href="http://www.staging-portal.ace-it.com">AMERICAN COMPUTER ESTIMATING, INC.</a><br>
Website design, programs and web code � <a href="http://www.puppozone.com/">PUPPO 
SOFTWARE SOLUTIONS, LLC.</a><br>
All Rights Reserved. <br>
Last modified:
January 09, 2008
</font> </p>



</td></tr><!--msnavigation--></table>

<script src="https://www.google-analytics.com/urchin.js" type="text/javascript"></script>
<script type="text/javascript">_uacct = "UA-1644909-1"; urchinTracker(); </script>
<script type="text/javascript">
	document.getElementById("loader").style.display = "none";
</script>

</body>
</html>
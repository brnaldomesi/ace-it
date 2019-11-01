<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<title>ACE PropReview</title>
<link rel="stylesheet" media="screen" href="css/forms.css">
<script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.5.5/angular.min.js"></script>
<script Language="JavaScript">
/*function SubmitForm(oButton) {
  var bDoSubmit = true;
  oButton.disabled = 1;
  if (ValidateFrm() == 0) {
    //alert ("You have not attached any files. Please attach your files and resubmit.");
    //bDoSubmit = fasle;
    bDoSubmit = window.confirm("You have not attached any files.\n\nTypically you need to attach the estimate when submitting claims. Click the 'Cancel' button to return to the form and attach a file. \n\nIf you are sure you don't need to attach a file click 'OK' and your claim will be submitted.");
  }
  if (bDoSubmit != false) {
    document.frmPropReview.submit();
  } else {
    oButton.disabled = 0;
  }
}*/

function SubmitForm(oButton) {
	var bDoSubmit = true;
	oButton.disabled = 1;
	var r = ValidateFrm();
	if (r == 2) {
		bDoSubmit = window.alert("Please complete all required information.\n\nRequired fields are denoted by the astrix (*).\n\n");
		bDoSubmit = false;
	} else if (r == 0) {
		//alert ("You have not attached any files. Please attach your files and resubmit.");
		//bDoSubmit = fasle;
		bDoSubmit = window.confirm("You have not attached any files.\n\nTypically you need to attach the estimate when submitting claims. Click the 'Cancel' button to return to the form and attach a file. \n\nIf you are sure you don't need to attach a file click 'OK' and your claim will be submitted.");
	}
	if (bDoSubmit != false) {
		document.frmPropReview.submit();
	} else {
		oButton.disabled = 0;
	}
}

function ValidateFrm() {
  var bAttachedFile = 0;
  if (document.forms[0].attachedfile1.value != "") bAttachedFile = 1;
  if (document.forms[0].attachedfile2.value != "") bAttachedFile = 1;
  if (document.forms[0].attachedfile3.value != "") bAttachedFile = 1;
  if (document.forms[0].attachedfile4.value != "") bAttachedFile = 1;
  
	writeError("txtCompany", document.forms[0].txtCompany.value != "");
	writeError("txtSubmittedBy", document.forms[0].txtSubmittedBy.value != "");
	writeError("policytype", validateRequiredRadios("policytype"));
	writeError("Description_Of_Loss", document.forms[0].Description_Of_Loss.value != "");
	writeError("claimnumber", document.forms[0].claimnumber.value != "");
	writeDateError("dol", document.forms[0].dol.value == "" || isDate(document.forms[0].dol.value));
	writeError("Insured", document.forms[0].Insured_Business.value != "" || document.forms[0].Insured_Person_First_Name.value != "" || document.forms[0].Insured_Person_Last_Name.value != "");
	writeError("Property_Owner", document.forms[0].Property_Owner_Business.value != "" || document.forms[0].Property_Owner_Person_First_Name.value != "" || document.forms[0].Property_Owner_Person_Last_Name.value != "");
	writeError("Comments", document.forms[0].Comments.value != "");
	writeError("contactname", document.forms[0].contactname.value != "");
	writeError("contactphone", document.forms[0].contactphone.value != "");
	writeError("Reach_Agreement", validateRequiredRadios("Reach_Agreement"));
	
	
	
	
	if (document.forms[0].txtCompany.value == "") bAttachedFile = 2;
	if (document.forms[0].txtSubmittedBy.value == "") bAttachedFile = 2;
	if (!validateRequiredRadios("policytype")) bAttachedFile = 2;
	if (document.forms[0].Description_Of_Loss.value == "") bAttachedFile = 2;
	if (document.forms[0].claimnumber.value == "") bAttachedFile = 2;
	if (document.forms[0].contactname.value == "") bAttachedFile = 2;
	if (document.forms[0].contactphone.value == "") bAttachedFile = 2;
	if (document.forms[0].Insured_Business.value == "" && document.forms[0].Insured_Person_First_Name.value == "" && document.forms[0].Insured_Person_Last_Name.value == "") bAttachedFile = 2;
	if (document.forms[0].Property_Owner_Business.value == "" && document.forms[0].Property_Owner_Person_First_Name.value == "" && document.forms[0].Property_Owner_Person_Last_Name.value == "") bAttachedFile = 2;
	if (!validateRequiredRadios("Reach_Agreement")) bAttachedFile = 2;
	
  return bAttachedFile;
}

if (navigator.appName.indexOf("Microsoft") > -1) {
	var canSee = 'block'
} else {
	var canSee = 'table-row';
}

function writeError(sID, bClearErr) {
	if (bClearErr == false) {
		bAttachedFile = 2;
		/*
				document.getElementById(sID + "_error").innerHTML = "
				<tr>
				  <td bgcolor='#FF0000' align='right' valign='top'>&nbsp;</td>
				  <td>The above field is required</td>
				</tr>
			";
		*/
		document.getElementById(sID + "_error_span").innerHTML = "<img border=\"0\" src=\"images/upArrowErr.jpg\" width=\"25\" height=\"24\">  The above field is required";
		document.getElementById(sID + "_error_tr").style.display = canSee;

	} else {
		document.getElementById(sID + "_error_span").innerHTML = "";
		document.getElementById(sID + "_error_tr").style.display = "none";
	}
}

function isDate(date) {
	var parsedDate = Date.parse(date);

	// You want to check again for !isNaN(parsedDate) here because Dates can be converted
	// to numbers, but a failed Date parse will not.
	if (isNaN(date) && !isNaN(parsedDate)) {
		return true;
	}
	
	return false;
}

function writeDateError(sID, bClearErr) {
	if (bClearErr == false) {
		bAttachedFile = 2;
		/*
				document.getElementById(sID + "_error").innerHTML = "
				<tr>
				  <td bgcolor='#FF0000' align='right' valign='top'>&nbsp;</td>
				  <td>The above field is required</td>
				</tr>
			";
		*/
		document.getElementById(sID + "_error_span").innerHTML = "<img border=\"0\" src=\"images/upArrowErr.jpg\" width=\"25\" height=\"24\">  Please enter a valid date";
		document.getElementById(sID + "_error_tr").style.display = canSee;

	} else {
		document.getElementById(sID + "_error_span").innerHTML = "";
		document.getElementById(sID + "_error_tr").style.display = "none";
	}
}

function validateRequiredRadios(RadioGroupName) {
	var radios = document.getElementsByName(RadioGroupName);
	var isChecked = false;

	var i = 0;
	while (!isChecked && i < radios.length) {
		if (radios[i].checked) isChecked = true;
		i++;
	}

	//if (!formValid) alert("Must check some option!");
	return isChecked;
}

function hideRequired(sID, bHide) {
	if (bHide == true) {
		document.getElementById(sID).innerHTML = "";
		//document.getElementById(sID).style.display = "none";	
	} else {
		document.getElementById(sID).innerHTML = "*";
		//document.getElementById(sID).style.display = canSee;
	};
}
</script>
<meta name="Microsoft Border" content="t">
</head>

<body ng-app><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<p align="center"><a href="../ACE-IT/index.htm"><img border="0" src="../Images/ace_logo.gif" width="149" height="71"></a><br>
<img border="0" src="../Images/LineGradBlue.jpg" width="422" height="11"><br>
</p>

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">

<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->

<%
'--------------------------------------------
Sub PostActivity(lActivityType, sInfo)
  Dim oDataActivity
  
  If mUserID <> 0 Then
    Set oDataActivity = New clsACEDataConn
    oDataActivity.ACEMemberActivityInsert2 mUserID, lActivityType, Request.ServerVariables("SCRIPT_NAME") & " " & Request.ServerVariables("QUERY_STRING") & " " & sInfo
    Set oDataActivity = Nothing
  End If
End Sub
'--------------------------------------------

  Dim CoName, SubmittedBy, OfficeID, OfficeName
  Dim PhoneAc, PhoneExchg, PhoneL4, PhoneEx
  Dim FaxAc, FaxExchg, FaxL4
  Dim EmailAddr
  Dim oData, mrs
  Dim sClaimNum
  'Dim oFormMailer
  Dim sAttachments
  Dim sBody
  Dim sFileName

  Dim oUpload, oFile, oItem
  
  Set oData = New clsACEDataConn

  'If Request.Form("txtState") <> "1" Then
  If Request.Querystring("txtState") <> "1" Then
  
    PostActivity Activity_ViewPage, ""

    Set mrs = oData.ACEMemberGet(mUserID, "","")

    If Not mrs.eof Then
      EmailAddr = mrs("MemberEmail") & ""

      PhoneEx = mrs("MemberPhoneEx") & ""
          
      If Len(mrs("MemberPhone") & "") > 0 Then Phone = Left(mrs("MemberPhone") & "", 3) & " " & Mid(mrs("MemberPhone") & "", 4, 3) & "-" & Mid(mrs("MemberPhone") & "",7)
      If Len(mrs("MemberFax") & "") > 0 Then Fax = Left(mrs("MemberFax") & "", 3) & " " & Mid(mrs("MemberFax") & "", 4, 3) & "-" & Mid(mrs("MemberFax") & "",7)

    End If
    mrs.Close
  
    CoName = mUserInsCoName
    SubmittedBy = mRepName
    OfficeID = mUserInsCoOfficeID_PropertyForm 'mUserACEInsCoOfficeID
    OfficeName = mUserACEInsCoOfficeName
  Else
%>

<p><b><font size="4">Click <a href="frmSubmitPropReview.asp">Here</a> to Submit 
another claim.</font></b></p>
<p>The information you submitted is listed below:</p>
<%
    PostActivity Activity_SubmitForm, ""

    Set oUpload = Server.CreateObject ("Persits.Upload.1")
    With oUpload
      .OverwriteFiles = False ' Generate unique names
	  .SetMaxSize 83886080 ' Truncate files above 80MB
     .SaveVirtual "/data/ACE" ' Save to data directory
      For Each oFile in .Files
	    If Instr(1, oFile.ExtractFileName, ",") Then
	      oFile.MoveVirtual("/data/ACE/" & Replace(oFile.ExtractFileName, ",", "_")) 'Move method has a side effect: once this method is successfully called, the Path property of this file object will point to the new location/name. 
	    End If
	    If Instr(1, oFile.ExtractFileName, "..") Then
	      oFile.MoveVirtual("/data/ACE/" & Replace(oFile.ExtractFileName, "..", ".")) 'Move method has a side effect: once this method is successfully called, the Path property of this file object will point to the new location/name. 
	    End If
      Next
	  

      sAttachments = "File attachments to this Claim:  "
      For Each oFile in .Files
        Response.Write "File saved as: " & oFile.ExtractFileName & "<br>"
       sAttachments = sAttachments & oFile.ExtractFileName & ", "
      Next
      sAttachments = Left(sAttachments, Len(sAttachments) - 2)
     
	  sBody = "<<< PropReview Submission >>>" & VBCrLf & Now & VBCrLf & VBCrLf

      For Each oItem in .Form
        Response.Write oItem.Name & " = " & oItem.Value & "<BR>"
        'oFormMailer.sBody = oFormMailer.sBody &  oItem.Name & " = " & oItem.Value & VBCrLf
        sBody = sBody &  oItem.Name & " = " & oItem.Value & VBCrLf
        If oItem.Name = "claimnumber" Then sClaimNum = oItem.Value
      Next  
    End With 

    PostActivity Activity_SubmitForm, "Claim#: " & sClaimNum
    
    Dim lErr
    SaveClaimSubmitToDb
    
    lErr = Err.Number
    On Error Goto 0
          
    If lErr > 0 Then
      PostActivity Activity_SubmitForm, "ERROR #" & err.Number & " Submitting Claim#: " & sClaimNum
%>

    <p align="center"><font size="4" color="#FF0000">There has been a problem trying to submit your request. The Server may temporarily be experiencing problems, please click the back button and resubmit your request.</font></p>

    <p align="center"><font size="4" color="#FF0000">PLEASE Contact ACE if this problem persists.</font></p>
    <p></p>
<%    
    End If
    
    On Error Goto 0
%>    
    <BR> <b><font color="#008000">... Your Claim has been Submitted ...</font></b> <BR>
<%    
    Response.End
  End If  
%>

<%
  Sub SaveClaimSubmitToDb()    
    Dim rsClaim 
    Dim sSQL
    Dim i
    Dim sAttach(5)
    
    sSQL = "SELECT * FROM tblClaimsSubmitted WHERE ClaimIDPK = 0"
    
    Set rsClaim = Server.CreateObject("ADODB.Recordset")
    rsClaim.Open sSQL, oData.ConnStr, 3, adLockOptimistic     'adOpenDynamic 3, adLockOptimistic


    rsClaim.AddNew
      rsClaim("ClaimBody") = sBody & VBCrLf & sAttachments
      rsClaim("ClaimNum") = sClaimNum
      rsClaim("SubmittedBy") =  oUpload.Form("txtSubmittedBy").value 
	  rsClaim("EmailTo") = "property@staging-portal.ace-it.com"
    rsClaim.Update  '--- Do update here to make sure we get record in case there's an issue with the attachments
    rsClaim.MoveFirst
      
      sAttach(1) = "" : sAttach(2) = "" : sAttach(3) = "" : sAttach(4) = "" : sAttach(5) = ""
      i = 1
      For Each oFile in oUpload.Files
        sAttach(i) = Server.Mappath("/Data/ACE/" & oFile.ExtractFileName)
        '--- This No Longer works: rsClaim("Attach" & i) = oFile.Binary
        rsClaim("Attach" & i).AppendChunk(oFile.Binary)
        i = i + 1
      Next      
      
    rsClaim.Update
    rsClaim.Close
    
  End Sub


  Sub SaveClaimSubmitToDbz()
  
    Dim rsClaim 
    Dim sSQL
    Dim i
    Dim sAttach(5)
    
    sSQL = "SELECT * FROM tblClaimsSubmitted WHERE ClaimIDPK = 0"
    
    Set rsClaim = Server.CreateObject("ADODB.Recordset")
    rsClaim.Open sSQL, oData.ConnStr, 3, adLockOptimistic     'adOpenDynamic 3, adLockOptimistic


    rsClaim.AddNew
      rsClaim("ClaimBody") = sBody & VBCrLf & sAttachments
      rsClaim("ClaimNum") = sClaimNum
      rsClaim("SubmittedBy") =  oUpload.Form("txtSubmittedBy").value 
      rsClaim("NotificationSent") = 1
    rsClaim.Update  '--- Do update here to make sure we get record in case there's an issue with the attachments
    rsClaim.Close


  
    Dim oFormMailer  
    Set oFormMailer = New clsFormMailer
    With oFormMailer
      .sHeader = "AutoReview Claim Submission Form."
      .sFooter = ""

      .sFrom = "ACEPortal@staging-portal.ace-it.com" '<--- no spaces
      .sTo = "aceitweb@comcast.net"
      .sSubject = "ACE-IT Portal, Claim Submission Form"
	End With
	
    With oUpload	
      For Each oFile in .Files
       oFormMailer.AttachFile = Server.Mappath("/Data/ACE/" & oFile.ExtractFileName)    
      Next
     
      For Each oItem in .Form
        oFormMailer.sBody = oFormMailer.sBody &  oItem.Name & " = " & oItem.Value & VBCrLf
      Next  
    End With 
    
    oFormMailer.Send    
    
  End Sub
  
%>


<div align="center">
  <center>

      <style>
          input[type=text]:disabled {
              background: #dddddd;
          }
      </style>

  <table border="1" cellpadding="6" cellspacing="0" bgcolor="#FFFFFF" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td>
      <table border="0" cellpadding="1" cellspacing="6">
        <tr>
          <td colspan="2" align="center"><b><i><font face="Times New Roman" size="4">PropReview 
          Submission Form</font></i></b></td>
        </tr>
		<tr>
		  <td colspan="2"><small><center>* Denotes a required field.</center></small></td>
		</tr>
        <form method="POST" enctype="multipart/form-data" name="frmPropReview" action="FrmSubmitPropReview.asp?txtState=1">
          <tr>
            <td colspan="2" align="center">
            <img SRC="../ACE-IT_sv/Forms/formimages/coinfo.jpg" border="0" width="405" height="19">
            </td>
          </tr>
          <!--
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Company&nbsp;&nbsp;
            </font></td>
            <td><input type="text" name="company" size="40" value=" ANYTOWN MUTUAL"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Submitted by&nbsp;&nbsp;
            </font></td>
            <td><input type="text" name="submittedby" size="40" value=" Betty Smith"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> ACE #&nbsp;&nbsp; </font>
            </td>
            <td><input type="text" name="officeid" size="40" value=" Z99"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Office&nbsp;&nbsp; </font>
            </td>
            <td><input type="text" name="office" size="40" value=" central region"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font>&nbsp; Phone Number&nbsp;&nbsp;
            </font></td>
            <td><b>(</b><input type="text" name="phoneac" size="3" maxlength="3" value="999"><b>)</b> <input type="text" name="phonep1" size="3" maxlength="3" 
            value="555"> <b>-</b> <input type="text" name="phonep2" size="4" maxlength="4" value="9999"> <font size="1">&nbsp;Ext</font> <input type="text" 
            name="phoneextn" size="5" maxlength="5" value="555"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Fax Number&nbsp;&nbsp;
            </font></td>
            <td><b>(</b><input type="text" name="faxac" size="3" maxlength="3" value="999"><b>)</b> <input type="text" name="faxp1" size="3" maxlength="3" 
            value="555"> <b>-</b> <input type="text" name="faxp2" size="4" maxlength="4" value="5555"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Email Address&nbsp;&nbsp;
            </font></td>
            <td><input type="text" name="email" size="40" value=" myname@anytown.com"></td>
          </tr>
-->
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            Company</td>
            <td>
            <input type="text" name="txtCompany" style="width: 88%" value="<%=CoName%>"></td>
          </tr>
		  <tr ID="txtCompany_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="txtCompany_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            Submitted by</td>
            <td>
            <input type="text" name="txtSubmittedBy" style="width: 88%" value="<%=SubmittedBy%>"></td>
          </tr>
		  <tr ID="txtSubmittedBy_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="txtSubmittedBy_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            ACE # </font></td>
            <td>
            <input type="text" name="txtOfficeID" style="width: 88%" value="<%=OfficeID%>"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            Office </font></td>
            <td>
            <input type="text" name="txtOfficeName" style="width: 88%" value="<%=OfficeName%>"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            Phone Number </font></td>
            <td><b>(</b><input type="text" name="txtPhoneAc" size="3" maxlength="3" value="<%=PhoneAc%>"><b>)</b>
            <input type="text" name="txtPhoneExchg" size="3" maxlength="3" value="<%=PhoneExchg%>">
            <b>-</b>
            <input type="text" name="txtPhoneL4" size="4" maxlength="4" value="<%=PhoneL4%>">
            <font size="1">&nbsp;Ext</font>
            <input type="text" name="txtPhoneEx" size="10" maxlength="10" value="<%=PhoneEx%>"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            Fax Number </font></td>
            <td><b>(</b><input type="text" name="txtFaxAc" size="3" maxlength="3" value="<%=FaxAc%>"><b>)</b>
            <input type="text" name="txtFaxExchg" size="3" maxlength="3" value="<%=FaxExchg%>">
            <b>-</b>
            <input type="text" name="txtFaxL4" size="4" maxlength="4" value="<%=FaxL4%>"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            Email Address </font></td>
            <td>
            <input type="text" name="txtEmail" style="width: 88%" value="<%=EmailAddr%>"></td>
          </tr>
          <tr>
            <td colspan="2" align="center">
            <img SRC="../ACE-IT_sv/Forms/formimages/claiminfo.jpg" border="0" width="405" height="19">
            </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            *Policy Type </font></td>
			
			<td>
				<fieldset name="Policy_Types">
					<input id="optReplacementCost" type="radio" value="Replacement Cost" name="policytype" required>
					<label for="optReplacementCost"> Replacement Cost</label>
					<input id="optACV" type="radio" value="ACV" name="policytype" required>
					<label for="optACV"> ACV</label>
				</fieldset>
			</td>
			
			
			
            <!--<td><select name="policytype">
            <option>Select</option>
            <option>Replacement Cost</option>
            <option>ACV</option>
            </select></td>-->
          </tr>
		  <tr ID="policytype_error_tr" style="display:none;">
			  <td>&nbsp;</td>
			  <td><SPAN STYLE=color:red ID="policytype_error_span"></SPAN></td>
		  </tr>
		  <tr>
		    <td bgcolor="#FFFFCC" align="right">*Description of Loss</td>
		    <td><textarea rows="3" name="Description_Of_Loss" style="width: 88%"></textarea></td>
		  </tr>
		  <tr ID="Description_Of_Loss_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="Description_Of_Loss_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            *Claim # </font></td>
            <td><input type="text" name="claimnumber" style="width: 88%"></td>
          </tr>
		  <tr ID="claimnumber_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="claimnumber_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            Policy Number </font></td>
            <td><input type="text" name="policynumber" style="width: 88%"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            *DOL </font></td>
            <td><input type="text" name="dol" style="width: 88%"></td>
          </tr>
		  <tr ID="dol_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="dol_error_span"></SPAN></td>
		  </tr>
          <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" ><font size="1" face="verdana, geneva, arial">Insured (*if business)</font></td>
              <td><input type="text" name="Insured_Business" style="width: 88%" placeholder="Business Name" ng-model="input.insuredBusiness" ng-disabled="!!input.insuredPersonFirstName || !!input.insuredPersonLastName"></td>
          </tr>
          <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" style="width: 88%"><font size="1" face="verdana, geneva, arial">Insured (*if person)</font></td>
              <td>
                  <input type="text" name="Insured_Person_First_Name" style="width: 43%;float: left;" placeholder="First Name" ng-model="input.insuredPersonFirstName" ng-disabled="!!input.insuredBusiness">
                  <input type="text" name="Insured_Person_Last_Name" style="width: 43%;margin-left: 3px" placeholder="Last Name" ng-model="input.insuredPersonLastName" ng-disabled="!!input.insuredBusiness">
              </td>
          </tr>
		  <tr ID="Insured_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="Insured_error_span"></SPAN></td>
		  </tr>
          <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial">Property Owner (*if business)</font></td>
              <td><input type="text" name="Property_Owner_Business" style="width: 88%" placeholder="Business Name" ng-model="input.propertyOwnerBusiness" ng-disabled="!!input.propertyOwnerPersonFirstName || !!input.propertyOwnerPersonLastName"></td>
          </tr>
          <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial">Property Owner (*if person)</font></td>
              <td>
                  <input type="text" name="Property_Owner_Person_First_Name" style="width: 43%;float: left;" placeholder="First Name" ng-model="input.propertyOwnerPersonFirstName" ng-disabled="!!input.propertyOwnerBusiness">
                  <input type="text" name="Property_Owner_Person_Last_Name" style="width: 43%;margin-left: 3px" placeholder="Last Name" ng-model="input.propertyOwnerPersonLastName" ng-disabled="!!input.propertyOwnerBusiness">
              </td>
          </tr>
		  <tr ID="Property_Owner_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="Property_Owner_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right">
            <font face="verdana, geneva, arial" size="1">*Comments</font></td>
            <td><textarea rows="3" name="Comments" cols="34"></textarea></td>
          </tr>
		  <tr ID="Comments_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="Comments_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td colspan="2" align="center">
            <img SRC="../ACE-IT_sv/Forms/formimages/crcontactinfo.jpg" border="0" width="405" height="19">
            </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            *Contact Name </font></td>
            <td><input type="text" name="contactname" style="width: 88%" maxlength="40">
            </td>
          </tr>
		  <tr ID="contactname_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="contactname_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            *Phone Number </font></td>
            <td><input type="text" name="contactphone" style="width: 88%" maxlength="12">
            </td>
          </tr>
		  <tr ID="contactphone_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="contactphone_error_span"></SPAN></td>
		  </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
             
            *Reach Agreement </font></td>
            <td>
				<fieldset name="Agreement_Types">
				<input id="optYes" type="radio" value="Yes" name="Reach_Agreement" required>
				<label for="optYes"> Yes</label>
				<input id="optNo" type="radio" value="No" name="Reach_Agreement" required>
				<label for="optNo"> No</label>
			</fieldset>
			
			
                <!--<font size="1" face="verdana, geneva, arial"><label><input type="checkbox" name="auditonly" value="1">Check here if you do not want us to settle.</label></font>-->
            </td>
          </tr>
		  <tr ID="Reach_Agreement_error_tr" style="display:none;">
			<td>&nbsp;</td>
			<td><SPAN STYLE=color:red ID="Reach_Agreement_error_span"></SPAN></td>
		  </tr>
          <!--          
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Attachments&nbsp;&nbsp;
            </font></td>
            <td><font size="1" face="verdana, geneva, arial">To attach a file related to this claim use the browse<br>
            button and navigate your computer to locate and<br>
            attach the necessary file.</font><br>
            <input type="file" name="attachedfile" size="27"> </td>
          </tr>
-->
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="bottom">
             
            Attachment 1 </font></td>
            <td>
                <font size="1" face="verdana, geneva, arial">To attach a file related to this claim use the<br />browse button and navigate your computer to<br />locate and attach the necessary file.
                <br>
                <b><font color="#008080">Please note file size cannot exceed 80MB.</font></b>
                <br>

            </font><input type="file" name="attachedfile1" size="27"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
            <font size="1" face="verdana, geneva, arial">Attachment 2
            </font></td>
            <td><input type="file" name="attachedfile2" size="27"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
            <font size="1" face="verdana, geneva, arial">Attachment 3
            </font></td>
            <td><input type="file" name="attachedfile3" size="27"> </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
            <font size="1" face="verdana, geneva, arial">Attachment 4
            </font></td>
            <td><input type="file" name="attachedfile4" size="27"> </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFCC" align="right" valign="top">
            <font size="1" face="verdana, geneva, arial">Attachment 5
            </font></td>
            <td><input type="file" name="attachedfile5" size="27"> </td>
          </tr>
          <!--          
    <tr>
      <td bgcolor="#FFFFCC" colspan="2" align="center">You must click submit for your request to be processed <br>
      <input type="submit" value="Submit" name="cmdSubmit" onClick="this.disabled = 1; document.frmPropReview.submit();"></td>
    </tr>
-->
          <tr>
            <td bgcolor="#FFFFCC" colspan="2" align="center"><b>
            <font face="Arial" size="2" color="#800000">You must click submit for 
            your request to be processed</font><font color="#000080"> </font>
            </b><br>
            <input type="button" value="Submit" name="cmdSubmit" onClick="SubmitForm(this);" style="font-weight: bold"></td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
  </table>
  </center>
</div>
<!-- END main body / START footer -->

<!--msnavigation--></td></tr><!--msnavigation--></table></body>

</html>
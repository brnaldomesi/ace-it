<!DOCTYPE html>
<html>
<head>
	<meta http-equiv="expires" content="Tue, 20 Aug 1996 01:00:00 GMT">
	<meta http-equiv="Pragma" content="no-cache">
	<meta name="Author" content="Bob Puppo">
	<meta name="Copyright" content="(C) Bob Puppo, 1999-2014, All rights reserved">
	<meta http-equiv="Content-Language" content="en-us">
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>ACE AutoReview</title>
	<link rel="stylesheet" media="screen" href="css/forms.css">
    <script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.5.5/angular.min.js"></script>
	<script language="JavaScript">
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
				document.frmAutoReview.submit();
			} else {
				oButton.disabled = 0;
			}
		}

		var bAttachedFile;
		function ValidateFrm() {
			//var bAttachedFile = 0;
			bAttachedFile = 0;
			if (document.forms[0].attachedfile1.value != "") bAttachedFile = 1;
			if (document.forms[0].attachedfile2.value != "") bAttachedFile = 1;
			if (document.forms[0].attachedfile3.value != "") bAttachedFile = 1;
			if (document.forms[0].attachedfile4.value != "") bAttachedFile = 1;
			if (document.forms[0].attachedfile5.value != "") bAttachedFile = 1;

			writeError("Company_Name", document.forms[0].Company_Name.value != "");
			writeError("Submitted_By", document.forms[0].Submitted_By.value != "");
			writeError("Description_Of_Loss", document.forms[0].Description_Of_Loss.value != "");
			writeError("Claim_Number", document.forms[0].Claim_Number.value != "");
			writeDateError("Date_Of_Loss", document.forms[0].Date_Of_Loss.value == "" || isDate(document.forms[0].Date_Of_Loss.value));
			writeError("Insured", document.forms[0].Insured_Business.value != "" || document.forms[0].Insured_Person_First_Name.value != "" || document.forms[0].Insured_Person_Last_Name.value != "");
			writeError("Vehicle_Owner", document.forms[0].Vehicle_Owner_Business.value != "" || document.forms[0].Vehicle_Owner_Person_First_Name.value != "" || document.forms[0].Vehicle_Owner_Person_Last_Name.value != "");
			writeError("Vehicle_Year", document.forms[0].Vehicle_Year.value != "");
			writeError("Vehicle_Make", document.forms[0].Vehicle_Make.value != "");
			writeError("Vehicle_VIN_Number", document.forms[0].Vehicle_VIN_Number.value != "");
			writeError("Vehicle_Body_Style", document.forms[0].Vehicle_Body_Style.value != "");
			writeError("Reach_Agreement", validateRequiredRadios("Reach_Agreement"));
			/*if (document.forms[0].AuditOnly.checked == true) {	//--- if Review Only
				//--- Don't skip cleanup since they may not have initially selected Review Only
				writeError("Shop_Name", true);
				writeError("Contact_Phone", true);
			} else {
				writeError("Shop_Name", document.forms[0].Shop_Name.value != "");
				writeError("Contact_Phone", document.forms[0].Contact_Phone.value != "");
			};*/

			//writeError("Percent_Negotiate", !(document.forms[0].optLiabilityNotAccepted.checked == true && document.forms[0].Percent_Negotiate.value == "0"));
			//writeError("Liability_Accepted", !(document.forms[0].optLiabilityNotAccepted.checked != true && document.forms[0].optLiabilityAccepted.checked != true));
				
				
			//writeError("Audit_Type", validateRequiredRadios("Audit_Type"));

			/*
              if (document.forms[0].Submitted_By.value == "") bAttachedFile = 2;
              if (document.forms[0].Claim_Number.value == "") bAttachedFile = 2;
              if (document.forms[0].Insured.value == "") bAttachedFile = 2;
              if (document.forms[0].Vehicle_Owner.value == "") bAttachedFile = 2;
              if (document.forms[0].Vehicle_Year.value == "") bAttachedFile = 2;
              if (document.forms[0].Vehicle_Make.value == "") bAttachedFile = 2;
              if (document.forms[0].Vehicle_VIN_Number.value == "") bAttachedFile = 2;
              if (document.forms[0].Vehicle_Body_Style.value == "") bAttachedFile = 2;
              if (document.forms[0].Shop_Name.value == "") bAttachedFile = 2;
              if (document.forms[0].Contact_Phone.value == "") bAttachedFile = 2;
              
            //  if (document.forms[0].optAuditTypeSettle.value == "" && document.forms[0].optAuditTypeReviewOnly.value == "") bAttachedFile = 2;
              if (validateRequiredRadios("Audit_Type") == false) bAttachedFile = 2;
            */

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
	<meta name="Microsoft Border" content="none">
</head>

<body ng-app>

	<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
	<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->

	<%
'--------------------------------------------------------------
' Send Email  --- NO LONGER USED ---
'--------------------------------------------------------------
	%>
	<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows Library" -->
	<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
	<%
SUB sendmailz(sFrom, sTo, sCC, sSubject, sTextBody, sHTMLBody, sURL, sAttach1, sAttach2, sAttach3, sAttach4, sAttach5)
  Dim objCDO
  Dim iConf
  Dim Flds

  Const cdoSendUsingPort = 2

  Set objCDO = Server.CreateObject("CDO.Message")
  Set iConf = Server.CreateObject("CDO.Configuration")

  Set Flds = iConf.Fields
  With Flds
    .Item(cdoSendUsingMethod) = cdoSendUsingPort
    .Item(cdoSMTPServer) = "mail-fwd"
    .Item(cdoSMTPServerPort) = 25
    .Item(cdoSMTPconnectiontimeout) = 10
    .Update
  End With

  Set objCDO.Configuration = iConf
  
  With objCDO
  
    .From = sFrom
    .To = sTo
    .CC = sCC
    .Subject = sSubject
    .TextBody = sTextBody
    If Len(sHTMLBody) <> 0 Then .HTMLBody = sHTMLBody
    
    If Len(sURL) <> 0 Then 
      'iMsg.CreateMHTMLBody "http://server.example.com", cdoSuppressAll, "domain\username", "password"
      .CreateMHTMLBody sURL, cdoSuppressNone, "", ""
    End If
  
    '.AddAttachment "http://server.example.com/picture.gif"
    '.AddAttachment "file://d:/temp/test.doc"
    '.AddAttachment "C:\files\another.doc"

    If Len(sAttach1) > 0 then .AddAttachment sAttach1
    If Len(sAttach2) > 0 then .AddAttachment sAttach2
    If Len(sAttach3) > 0 then .AddAttachment sAttach3
    If Len(sAttach4) > 0 then .AddAttachment sAttach4
    If Len(sAttach5) > 0 then .AddAttachment sAttach5
     
    .Send
    
  End With

  'Cleanup
  Set objCDO = Nothing
  Set iConf = Nothing
  Set Flds = Nothing
END SUB
	%>

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
    OfficeID = mUserInsCoOfficeID_AutoForm 'mUserACEInsCoOfficeID
    OfficeName = mUserACEInsCoOfficeName
  Else
	%>

	<p align="center"><b><font color="#008080" size="4"><img border="0" src="../Images/ace_logo_sm.gif"></font></b></p>
	<p align="center"><b><font size="4"><font color="#008080">Click <a href="frmSubmitAutoReview.asp">Here</a> <font color="#008080">to Submit another claim.</font></font></b></p>
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
        'Response.Write oFile.OriginalFileName & " saved as " & oFile.FileName
        Response.Write "File saved as: " & oFile.ExtractFileName & "<br>"
       sAttachments = sAttachments & oFile.ExtractFileName & ", "
      Next
      sAttachments = Left(sAttachments, Len(sAttachments) - 2)
      
	  sBody = "<<< AutoReview Submission >>>" & VBCrLf & Now & VBCrLf & VBCrLf

      For Each oItem in .Form
        Response.Write oItem.Name & " = " & oItem.Value & "<BR>"
        'oFormMailer.sBody = oFormMailer.sBody &  oItem.Name & " = " & oItem.Value & VBCrLf
        sBody = sBody &  oItem.Name & " = " & oItem.Value & VBCrLf
        If oItem.Name = "Claim_Number" Then sClaimNum = oItem.Value
      Next  
    End With 
    
    '--- Post Activity ---
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
	<br>
	<b><font color="#008000">... Your Claim has been Submitted ...</font></b>
	<br>
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
      rsClaim("SubmittedBy") =  oUpload.Form("Submitted_By").value 
	  rsClaim("EmailTo") = "eclaims@ace-it.com"
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
    
    '--- EMAIL ---
'    sendmail "support@ace-it.com", "support@ace-it.com", "", "ACE-IT Portal, Claim Submission Form", sBody, "", "", sAttach(1), sAttach(2), sAttach(3), sAttach(4), sAttach(5)
    
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
      rsClaim("SubmittedBy") =  oUpload.Form("Submitted_By").value 
      rsClaim("NotificationSent") = 1
    rsClaim.Update  '--- Do update here to make sure we get record in case there's an issue with the attachments
    rsClaim.Close


  
    Dim oFormMailer  
    Set oFormMailer = New clsFormMailer
    With oFormMailer
      .sHeader = "AutoReview Claim Submission Form."
      .sFooter = ""

      .sFrom = "ACEPortal@ace-it.com" '<--- no spaces
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

  <form  method="POST" enctype="multipart/form-data" name="frmAutoReview" action="FrmSubmitAutoReview.asp?txtState=1">
<table border="1" cellspacing="3"  cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td align="center" width="100%">
<table border="0" cellpadding="1" cellspacing="6">
  <tr>
    <td colspan="2" align="center">
    <b><i><font face="Times New Roman" size="4"><img border="0" src="../Images/ACELogo_icon.gif">&nbsp; AutoReview Submission Form</font></i></b></td>
  </tr>
    <tr>
      <td colspan="2"><small><center>* Denotes a required field.</center></small></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><img SRC="../ACE-IT_sv/Forms/formimages/coinfo.jpg" border="0"> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp; Company
      </td>
      <td><input type="text" name="Company_Name" style="width: 88%" value="<%=CoName%>"></td>
    </tr>
    <tr ID="Company_Name_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Company_Name_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp; Submitted by
      </td>
      <td><input type="text" name="Submitted_By" style="width: 88%" value="<%=SubmittedBy%>"></td>
    </tr>
    <tr ID="Submitted_By_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Submitted_By_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;ACE #</td>
      <td><input type="text" name="Office_ID" style="width: 88%" value="<%=OfficeID%>"></td>
    </tr>
    <tr ID="Office_ID_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Office_ID_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Office</td>
      <td><input type="text" name="Office_Name" style="width: 88%" value="<%=OfficeName%>"></td>
    </tr>
    <tr ID="Office_Name_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Office_Name_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Phone Number
      </td>
      <td><input type="text" name="Submitted_By_Phone" size="20" value="<%=Phone%>">&nbsp;&nbsp;Ext<input type="text" 
      name="Submitted_By_Extension" size="10" maxlength="10" value="<%=PhoneEx%>"></td>
    </tr>
    <tr ID="Submitted_By_Phone_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Submitted_By_Phone_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Fax Number
      </td>
      <td><input type="text" name="Submitted_By_Fax" size="20" value="<%=Fax%>"></td>
    </tr>
    <tr ID="Submitted_By_Fax_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Submitted_By_Fax_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Email Address
     </td>
      <td><input type="text" name="Submitted_By_Email" style="width: 88%" value="<%=EmailAddr%>"></td>
    </tr>
    <tr ID="Submitted_By_Email_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Submitted_By_Email_error_span"></SPAN></td>
	</tr>
    <tr>
      <td colspan="2" align="center"><img SRC="../ACE-IT_sv/Forms/formimages/claiminfo.jpg" border="0"> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Loss Type
      </td>
      <td><select name="Loss_Type">
      <option>--- Select ---</option>
      <option>Collision</option>
      <option>Comprehensive</option>
      <option>Property Damage</option>
      <option>Uninsured Motorist</option>
      </select></td>
    </tr>
    <tr ID="Loss_Type_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Loss_Type_error_span"></SPAN></td>
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
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Claim #
      </td>
      <td><input type="text" name="Claim_Number" style="width: 88%"></td>
    </tr>
    <tr ID="Claim_Number_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Claim_Number_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Policy Number
     </td>
      <td><input type="text" name="Policy_Number" style="width: 88%"></td>
    </tr>
    <tr ID="Policy_Number_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Policy_Number_error_span"></SPAN></td>
	</tr>
	<tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Deductible
      </td>
      <td><input type="text" name="Deductible" style="width: 88%"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;DOL</td>
      <td><input type="text" name="Date_Of_Loss" style="width: 88%"></td>
    </tr>
    <tr ID="Date_Of_Loss_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Date_Of_Loss_error_span"></SPAN></td>
	</tr>
    <tr>
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Insured (*if business)</td>
        <td><input type="text" name="Insured_Business" style="width: 88%" placeholder="Business Name" ng-model="input.insuredBusiness" ng-disabled="!!input.insuredPersonFirstName || !!input.insuredPersonLastName"></td>
    </tr>
    <tr>
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Insured (*if person)</td>
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
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Vehicle Owner (*if business)</td>
        <td><input type="text" name="Vehicle_Owner_Business" style="width: 88%" placeholder="Business Name" ng-model="input.vehicleOwnerBusiness" ng-disabled="!!input.vehicleOwnerPersonFirstName || !!input.vehicleOwnerPersonLastName"></td>
    </tr>
    <tr>
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Vehicle Owner (*if person)</td>
        <td>
            <input type="text" name="Vehicle_Owner_Person_First_Name" style="width: 43%;float: left;" placeholder="First Name" ng-model="input.vehicleOwnerPersonFirstName" ng-disabled="!!input.vehicleOwnerBusiness">
            <input type="text" name="Vehicle_Owner_Person_Last_Name" style="width: 43%;margin-left: 3px" placeholder="Last Name" ng-model="input.vehicleOwnerPersonLastName" ng-disabled="!!input.vehicleOwnerBusiness">
        </td>
    </tr>
    <tr ID="Vehicle_Owner_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Owner_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right">Comments</td>
      <td><textarea rows="3" name="Comments" style="width: 88%"></textarea></td>
    </tr>
    <tr ID="Comments_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Comments_error_span"></SPAN></td>
	</tr>
    <tr>
      <td COLSPAN="2" ALIGN="CENTER"><img SRC="../ACE-IT_sv/Forms/formimages/vehicleinfo.jpg" border="0"> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Year</td>
      <td><input type="text" name="Vehicle_Year" style="width: 88%"></td>
    </tr>
    <tr ID="Vehicle_Year_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Year_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Make</td>
      <td><input type="text" name="Vehicle_Make" style="width: 88%"></td>
    </tr>
    <tr ID="Vehicle_Make_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Make_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Model</td>
      <td><input type="text" name="Vehicle_Body_Style" style="width: 88%"></td>
    </tr>
    <tr ID="Vehicle_Body_Style_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Body_Style_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Vin #</td>
      <td><input type="text" name="Vehicle_VIN_Number" style="width: 88%"></td>
    </tr>
    <tr ID="Vehicle_VIN_Number_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_VIN_Number_error_span"></SPAN></td>
	</tr>
	<tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Odometer</td>
      <td><input type="text" name="Odometer" style="width: 88%"></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><img SRC="../ACE-IT_sv/Forms/formimages/contactinfo.jpg" border="0"> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top"><span id="Shop_req_span">*</span>Shop
     </td>
      <td><input type="text" name="Shop_Name" style="width: 88%" maxlength="40"> </td>
    </tr>
    <tr ID="Shop_Name_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Shop_Name_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top"><span id="Contact_Phone_req_span">*</span>Phone<br>Number&nbsp;
     </td>
      <td><input type="text" name="Contact_Phone" style="width: 88%" maxlength="12"> </td>
    </tr>
    <tr ID="Contact_Phone_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Contact_Phone_error_span"></SPAN></td>
	</tr>
<!--
    <tr bgcolor="#ffe4b5">
      <td colspan="2"><font size="2" face="verdana,geneva,arial" color="#FF0000"><input type="radio" name="mmchoice" value="0" checked 
      onClick="enablepercopt()"><font size="2" face="verdana,geneva,arial" color="#CC0000">Liability accepted</font></td>
    </tr>
    <tr bgcolor="#ffe4b5">
      <td colspan="2"><font size="2" face="verdana,geneva,arial" color="#FF0000"><input type="radio" name="mmchoice" value="0" onClick="disablepercopt()">
     <font size="2" face="verdana,geneva,arial" color="#CC0000">Negotiate our liability at</font></td>
    </tr>
-->    
    <!--<tr bgcolor="#ffe4b5">
      <td colspan="2"><font size="2" face="verdana,geneva,arial" color="#FF0000"><input id="optLiabilityAccepted" type="radio" name="Liability_Accepted" value="0" >
	  <label for="optLiabilityAccepted">
      <font size="2" face="verdana,geneva,arial" color="#CC0000">Liability accepted 100%</font></label></td>
    </tr>

    <tr ID="Liability_Accepted_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Liability_Accepted_error_span"></SPAN></td>
	</tr>

    <tr bgcolor="#ffe4b5">
      <td colspan="2"><font size="2" face="verdana,geneva,arial" color="#FF0000"><input id="optLiabilityNotAccepted" type="radio" name="Liability_Accepted" value="0" >
	  <label for="optLiabilityNotAccepted">
     <font size="2" face="verdana,geneva,arial" color="#CC0000">Negotiate our liability at</font></label></td>
    </tr>
    <tr bgcolor="#ffe4b5">
      <td align="right"><font size="2" face="verdana,geneva,arial&quot;" color="000000">Percentage&nbsp;</td>
      <!-- <td><select name="Percent_Nagerate" id="percentagerate" disabled> --><!--
      <td><select name="Percent_Negotiate" id="percentNegotiate">
		  <option value="0">n/a</option>
		  <option value="5">5%</option>
		  <option value="10">10%</option>
		  <option value="15">15%</option>
		  <option value="20">20%</option>
		  <option value="25">25%</option>
		  <option value="30">30%</option>
		  <option value="35">35%</option>
		  <option value="40">40%</option>
		  <option value="45">45%</option>
		  <option value="50">50%</option>
		  <option value="55">55%</option>
		  <option value="60">60%</option>
		  <option value="65">65%</option>
		  <option value="70">70%</option>
		  <option value="75">75%</option>
		  <option value="80">80%</option>
		  <option value="85">85%</option>
		  <option value="90">90%</option>
		  <option value="95">95%</option>
		  <option value="100">100%</option>
      </select></td>
    </tr>
    <tr ID="Percent_Negotiate_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Percent_Negotiate_error_span"></SPAN></td>
	</tr>-->
<!--
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Audit Only&nbsp;&nbsp;
      </td>
      <td><font size="1" face="verdana, geneva, arial"><input type="checkbox" name="Audit_Only" value="1">Check here if you do not want us to settle.</font></td>
    </tr>
-->
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Reach Agreement
      </td>
      <td>
          <!--- 2015-08-08 removed
		<fieldset name="Audit_Types">
		<input id="optAuditTypeSettle" type="radio" value="Review and Settle" name="Audit_Type" required onClick="hideRequired('Contact_Phone_req_span', false);hideRequired('Shop_req_span', false);"">
		<label for="optAuditTypeSettle"> Review and Settle</label>
		<input id="optAuditTypeReviewOnly" type="radio" value="Review Only" name="Audit_Type" required onClick="hideRequired('Contact_Phone_req_span', true);hideRequired('Shop_req_span', true);">
		<label for="optAuditTypeReviewOnly"> Review Only</label>
		</fieldset>
              -->
			  
			  
			  
			  <fieldset name="Agreement_Types">
				<input id="optYes" type="radio" value="Yes" name="Reach_Agreement" required>
				<label for="optYes"> Yes</label>
				<input id="optNo" type="radio" value="No" name="Reach_Agreement" required>
				<label for="optNo"> No</label>
			</fieldset>
			
			

            <!--<font size="1" face="verdana, geneva, arial"><label><input type="checkbox" name="AuditOnly" value="1">Check here if you do not want us to settle.</label></font>-->
      </td>
    </tr>
    <tr ID="Reach_Agreement_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Reach_Agreement_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="bottom">&nbsp;Attachment 1
      </td>
      <td>
          <font size="1" face="verdana, geneva, arial">To attach a file related to this claim use the<br />browse button and navigate your computer to<br />locate and attach the necessary file.
          <br>
          <b><font color="#008080">Please note file size cannot exceed 80MB.</font></b>
          <br>
     
      <input type="file" name="attachedfile1" size="27"/></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">Attachment&nbsp;2
      </td>
      <td>
      <input type="file" name="attachedfile2" size="27"/></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">Attachment&nbsp;3
      </td>
      <td>
      <input type="file" name="attachedfile3" size="27"/> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">Attachment&nbsp;4
      </td>
      <td>
      <input type="file" name="attachedfile4" size="27"/> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">Attachment&nbsp;5
      </td>
      <td>
      <input type="file" name="attachedfile5" size="27"/> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" colspan="2" align="center"><b><font face="Arial" size="2" color="#800000">You must click submit for your request to be processed</font>
      </b> <br>
      <input type="button" value="Submit" name="cmdSubmit" onClick="SubmitForm(this);" style="font-weight: bold"></td>
    </tr>
</table>

    </td>
  </tr>
</table>
  </form>

  </center>
	</div>

	<script type="text/javascript" language="JavaScript">
		document.forms[0].elements["Loss_Type"].focus();
	</script>

	<script>
  window.moveTo(0, 25);
  window.resizeTo(500, screen.availHeight-25);
	</script>

	<p>&nbsp;</p>

	</font>

</body>
</html>

<%
  Set oUpload = Nothing
%>
<%
  Function IIf(i,j,k)
    If i Then IIf = j Else IIf = k
  End Function
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
	<title>ACE PhotoQuote</title>

	<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
	<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->
	

	<link rel="stylesheet" media="screen" href="css/forms.css">
    <!--<script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.5.5/angular.min.js"></script>-->
    <script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.6.4/angular.min.js"></script>
	<style>
		input:enabled.ng-invalid {
		 border-width:1px;
		 border-color:red;
		}
		input.ng-valid {
		}
		textarea.ng-invalid {
		 outline-color:red;*/
		 border-width:1px;
		 border-color:red;
		}
	</style>
	<script language="JavaScript">
		function loadYearMakeModel() {
			var xhttp = new XMLHttpRequest();
			xhttp.onreadystatechange = function() {
				if (this.readyState == 4 && this.status == 200) {
					//document.getElementById("Vehicle_Year").innerHTML = this.responseText;
					var jObj = JSON.parse(this.responseText);
					var err = jObj.Results[0].ErrorCode; 
					if (!err.startsWith("0")) {
						alert("VIN decoding error:\n " + err);
					}
					
					document.forms[0].Vehicle_Year.value = jObj.Results[0].ModelYear;
					document.forms[0].Vehicle_Make.value = jObj.Results[0].Make;
					document.forms[0].Vehicle_Body_Style.value = jObj.Results[0].Model;
				}
			};
			var vin = document.forms[0].Vehicle_VIN_Number.value
			xhttp.open("GET", "VIN_Request.asp?vin=" + vin, true);
			xhttp.send();
		}
	</script>
	<script language="JavaScript">
	
		//--- For compatibility with <IE9 (doesn't have trim() so add it):
		if(typeof String.prototype.trim !== 'function') {
		 String.prototype.trim = function() {
			return this.replace(/^\s+|\s+$/g, ''); 
		 }
		}

		function SubmitForm(oButton) {
			var bDoSubmit = true;
			oButton.disabled = true; //1;
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
				document.frmPhotoQuote.submit();
			} else {
				oButton.disabled = false; //0;
			}
		}
		
		
		function inputTypeIsSupported(field) {
			return (field.type == "text");
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
			<% 
				If mUserACEInsCoOfficeID = "A16S" Or mUserACEInsCoOfficeID = "A16P" Then 
			%> 
				writeError("Feature_Number", document.forms[0].Feature_Number.value != "");
			<%End If%>
			
			fld = document.forms[0].Insured_Business;
			fld.value = fld.value.trim();
			fld = document.forms[0].Insured_Person_First_Name;
			fld.value = fld.value.trim();
			fld = document.forms[0].Insured_Person_Last_Name;
			fld.value = fld.value.trim();
			fld = document.forms[0].Vehicle_Owner_Business;
			fld.value = fld.value.trim();
			fld = document.forms[0].Vehicle_Owner_Person_First_Name;
			fld.value = fld.value.trim();
			fld = document.forms[0].Vehicle_Owner_Person_Last_Name;
			fld.value = fld.value.trim();
			
			writeError("Deductible", document.forms[0].Deductible.value != "");
			writeError("Insured", document.forms[0].Insured_Business.value != "" || document.forms[0].Insured_Person_Last_Name.value != "");
			writeError("Vehicle_Owner", document.forms[0].Vehicle_Owner_Business.value != "" || document.forms[0].Vehicle_Owner_Person_Last_Name.value != "");
			writeError("Zip_Code", document.forms[0].Zip_Code.value != "");

			writeError("Permission_to_SMS", document.forms[0].Permission_to_SMS.value != "");
			
			writeError("Vehicle_Year", document.forms[0].Vehicle_Year.value.trim() != "");
			writeError("Vehicle_Make", document.forms[0].Vehicle_Make.value.trim() != "");
			writeError("Vehicle_VIN_Number", document.forms[0].Vehicle_VIN_Number.value != "");
			writeError("Vehicle_Body_Style", document.forms[0].Vehicle_Body_Style.value != "");

			writeError("Cell_Num", validatePhonenumber(document.forms[0].Cell_Num.value));
			 
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
				document.getElementById(sID + "_error_span").innerHTML = "This field is required <img border=\"0\" src=\"images/upArrowErr.jpg\" width=\"14\" height=\"14\">";
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

			return isChecked;
		}

		function hideRequired(sID, bHide) {
			if (bHide == true) {
				document.getElementById(sID).innerHTML = "";
			} else {
				document.getElementById(sID).innerHTML = "*";
			};
		}
		
		function validatePhonenumber(sPhoneNum) {
			
			var formats = "(999)999-9999|(999) 999-9999|999 999-9999|999-999-9999|9999999999";
			var r = RegExp("^(" +
				formats
				.replace(/([\(\)])/g, "\\$1")
				.replace(/9/g,"\\d") +
				")$");
			if (sPhoneNum.match(r)) {
				return true;
			} else {
				return false;
			}
		}  

	</script>
	<script language="JavaScript">
		if (!String.prototype.startsWith) {
		  String.prototype.startsWith = function(searchString, position) {
			position = position || 0;
			return this.indexOf(searchString, position) === position;
		  };
		}	
	</script>
	<meta name="Microsoft Border" content="none">
</head>

<body ng-app>


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
    OfficeID = mUserInsCoOfficeID_PhotoQuoteForm 'mUserACEInsCoOfficeID
    OfficeName = mUserACEInsCoOfficeName
  Else '<--- txtState = 1
	%>

	<p align="center"><b><font color="#008080" size="4"><img border="0" src="../Images/ace_logo_sm.gif"></font></b></p>
	<p align="center"><b><font size="4"><font color="#008080">Click <a href="frmSubmitPhotoQuote.asp">Here</a> <font color="#008080">to Submit another claim.</font></font></b></p>
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
	  sBody = "<<< PhotoQuote Submission >>>" & VBCrLf & Now & VBCrLf & VBCrLf

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

  <form  method="POST" enctype="multipart/form-data" name="frmPhotoQuote" action="FrmSubmitPhotoQuote.asp?txtState=1">
<table border="1" cellspacing="3"  cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td align="center" width="100%">
<table border="0" cellpadding="1" cellspacing="6">
  <tr>
    <td colspan="2" align="center">
    <b><i><font face="Times New Roman" size="4"><img border="0" src="../Images/ACELogo_icon.gif">&nbsp; PhotoQuote Submission Form</font></i></b></td>
  </tr>
    <tr>
      <td colspan="2"><small><center>* Denotes a required field.</center></small></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><img SRC="../ACE-IT_sv/Forms/formimages/coinfo.jpg" width="100%" border="0"> </td>
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
      <td><input type="text" name="Submitted_By_Phone" size="17" value="<%=Phone%>">&nbsp;&nbsp;Ext<input type="text" 
      name="Submitted_By_Extension" size="8" maxlength="10" value="<%=PhoneEx%>"></td>
    </tr>
    <tr ID="Submitted_By_Phone_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Submitted_By_Phone_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Fax Number
      </td>
      <td><input type="text" name="Submitted_By_Fax" size="17" value="<%=Fax%>"></td>
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
      <td colspan="2" align="center"><img SRC="../ACE-IT_sv/Forms/formimages/claiminfo.jpg" width="100%" border="0"> </td>
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
      <td><textarea rows="3" name="Description_Of_Loss" style="width: 88%" ng-model="input.Description_Of_Loss" required></textarea></td>
    </tr>
    <tr ID="Description_Of_Loss_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Description_Of_Loss_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Claim #
      </td>
      <td><input type="text" name="Claim_Number" style="width: 88%" ng-model="input.Claim_Number" required></td>
    </tr>
    <tr ID="Claim_Number_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Claim_Number_error_span"></SPAN></td>
	</tr>
	<% 
		If mUserACEInsCoOfficeID = "A16S" Or mUserACEInsCoOfficeID = "A16P" Then %> 
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Feature #
      </td>
      <td><input type="text" name="Feature_Number" maxlength="3" style="width: 88%" ng-model="input.Feature_Number" required></td>
    </tr>
    <tr ID="Feature_Number_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Feature_Number_error_span"></SPAN></td>
	</tr>
	<% End If%>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Deductible
      </td>
      <td><input type="text" name="Deductible" style="width: 88%" ng-model="input.Deductible" required></td>
    </tr>
    <tr ID="Deductible_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Deductible_error_span"></SPAN></td>
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
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;DOL</td>
      <td><input type="text" name="Date_Of_Loss" style="width: 88%"></td>
    </tr>
    <tr ID="Date_Of_Loss_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Date_Of_Loss_error_span"></SPAN></td>
	</tr>
    <tr>
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Insured (*if business)</td>
        <td><input type="text" name="Insured_Business" style="width: 88%" placeholder="Business Name" ng-model="input.insuredBusiness" ng-disabled="!!input.insuredPersonFirstName || !!input.insuredPersonLastName" required></td>
    </tr>
    <tr>
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Insured (*if person)</td>
        <td>
            <input type="text" name="Insured_Person_First_Name" style="width: 43%;float: left;" placeholder="First Name" ng-model="input.insuredPersonFirstName" ng-disabled="!!input.insuredBusiness">
            <input type="text" name="Insured_Person_Last_Name" style="width: 43%;margin-left: 3px" placeholder="Last Name" ng-model="input.insuredPersonLastName" ng-disabled="!!input.insuredBusiness" required>
        </td>
    </tr>
    <tr ID="Insured_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Insured_error_span"></SPAN></td>
	</tr>
    <tr>
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Vehicle Owner (*if business)</td>
        <td><input type="text" name="Vehicle_Owner_Business" style="width: 88%" placeholder="Business Name" ng-model="input.vehicleOwnerBusiness" ng-disabled="!!input.vehicleOwnerPersonFirstName || !!input.vehicleOwnerPersonLastName" required></td>
    </tr>
    <tr>
        <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;Vehicle Owner (*if person)</td>
        <td>
            <input type="text" name="Vehicle_Owner_Person_First_Name" style="width: 43%;float: left;" placeholder="First Name" ng-model="input.vehicleOwnerPersonFirstName" ng-disabled="!!input.vehicleOwnerBusiness">
            <input type="text" name="Vehicle_Owner_Person_Last_Name" style="width: 43%;margin-left: 3px" placeholder="Last Name" ng-model="input.vehicleOwnerPersonLastName" ng-disabled="!!input.vehicleOwnerBusiness" required>
        </td>
    </tr>
    <tr ID="Vehicle_Owner_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Owner_error_span"></SPAN></td>
	</tr>
	

	
	<tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Zip Code</td>
      <td><input type="text" name="Zip_Code" style="width: 88%" ng-model="input.Zip_Code" required></td>
    </tr>
    <tr ID="Zip_Code_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Zip_Code_error_span"></SPAN></td>
	</tr>
	
	<tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Cell Phone #</td>
      <td><input type="tel" placeholder="123 123-1234" 
			name="Cell_Num" style="width: 88%"
			 ng-model="input.Cell_Num" required>
	  </td>
    </tr>
    <tr ID="Cell_Num_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Cell_Num_error_span"></SPAN></td>
	</tr>
	
	<tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">
		  <div style="max-width:200px;">
		  *Permission from owner to receive SMS message
		  </div>
      </td>
      <td><select name="Permission_to_SMS" ng-model="input.Permission_to_SMS" required>
      <option value="">--- Select ---</option>
      <option>Yes</option>
      <option>No</option>
      </select></td>
    </tr>
    <tr ID="Permission_to_SMS_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Permission_to_SMS_error_span"></SPAN></td>
	</tr>

    <tr>
    	<td colspan="2" style="padding-left:10px;padding-right:10px;">
		<div style="max-width:450px;text-align:center;">
		<font color="#800000">If the owner does not agree then a PhotoQuote is not possible and you should not submit this form.</font>
		<br>A Text (SMS) message containing a link to the PhotoQuote App in the App Store will be sent within 1 hour of successfully submitting this form.
		</div>
		</td>
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
      <td COLSPAN="2" ALIGN="CENTER"><img SRC="../ACE-IT_sv/Forms/formimages/vehicleinfo.jpg" width="100%" border="0"> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Vin #</td>
      <td><input type="text" name="Vehicle_VIN_Number" onblur="loadYearMakeModel();" style="width: 88%" ng-model="input.Vehicle_VIN_Number" required></td>
    </tr>
    <tr ID="Vehicle_VIN_Number_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_VIN_Number_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Year</td>
      <td><input type="text" name="Vehicle_Year" id="Vehicle_Year" style="width: 88%" ng-model="input.Vehicle_Year" required></td>
    </tr>
    <tr ID="Vehicle_Year_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Year_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Make</td>
      <td><input type="text" name="Vehicle_Make" name="Vehicle_Make" style="width: 88%" ng-model="input.Vehicle_Make" required></td>
    </tr>
    <tr ID="Vehicle_Make_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Make_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;*Model</td>
      <td><input type="text" name="Vehicle_Body_Style" name="Vehicle_Body_Style" style="width: 88%" ng-model="input.Vehicle_Body_Style" required></td>
    </tr>
    <tr ID="Vehicle_Body_Style_error_tr" style="display:none;">
    	<td>&nbsp;</td>
	    <td><SPAN STYLE=color:red ID="Vehicle_Body_Style_error_span"></SPAN></td>
	</tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="bottom">&nbsp;Attachment 1
      </td>
      <td>
          <font size="1" face="verdana, geneva, arial">To attach a file related to this claim use the<br />browse button and navigate your computer to<br />locate and attach the necessary file.</font>
          <br>
          <b><font color="#008080">Please note file size cannot exceed 80MB.</font></b>
          <br>
     
      <input type="file" name="attachedfile1" size="27"/>
	  </td>
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


</body>
</html>

<%
  Set oUpload = Nothing
%>
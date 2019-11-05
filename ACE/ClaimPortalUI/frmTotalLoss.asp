<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->

<!DOCTYPE html>
<html>
<head>
<meta http-equiv="expires" content="Tue, 20 Aug 1996 01:00:00 GMT">
<meta http-equiv="Pragma" content="no-cache">
<meta name="Author" content="Bob Puppo">
<meta name="Copyright" content="(C) Bob Puppo, 1999-2008, All rights reserved">
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ACE Total Loss Apraisal</title>
<link rel="stylesheet" media="screen" href="css/forms.css">
<script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.5.5/angular.min.js"></script>
<script Language = "JavaScript">
function SubmitForm(oButton) {
  var bDoSubmit = true;
  oButton.disabled = 1;
  //if (ValidateFrm() == 0) {
    //alert ("You have not attached any files. Please attach your files and resubmit.");
    //bDoSubmit = fasle;
    //bDoSubmit = window.confirm("You have not attached any files.\n\nTypically you need to attach the estimate when submitting claims. Click the 'Cancel' button to return to the form and attach a file. \n\nIf you are sure you don't need to attach a file click 'OK' and your claim will be submitted.");
  //}
  if (bDoSubmit != false) {
    document.frmTotalLoss.submit();
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
  return bAttachedFile;
}
</script>

</head>

<body ng-app>

<%
'--------------------------------------------------------------
' Send Email
'--------------------------------------------------------------
%>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows Library" -->
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
<%
SUB sendmail(sFrom, sTo, sCC, sSubject, sTextBody, sHTMLBody, sURL, sAttach1, sAttach2, sAttach3, sAttach4, sAttach5)
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
    OfficeID = mUserInsCoOfficeID_AutoTotalLossForm 'mUserACEInsCoOfficeID
    OfficeName = mUserACEInsCoOfficeName
  Else
%>
<p align="center"><b><font color="#008080" size="4"><img border="0" src="../Images/ace_logo_sm.gif"></font></b></p>
<p align="center"><b><font size="4"><font color="#008080">Click </font><a href="frmTotalLoss.asp">Here</a> <font color="#008080">to Submit another Total Loss claim.</font></font></b></p>
<p>The information you submitted is listed below:</p>
<%
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

      sAttachments = "File attachments to this Claim:  " '<---- note that we search for this strig when parsing out the attachment names
      For Each oFile in .Files
        'Response.Write oFile.OriginalFileName & " saved as " & oFile.FileName
        Response.Write "File saved as: " & oFile.ExtractFileName & "<br>"
       sAttachments = sAttachments & oFile.ExtractFileName & ", "
      Next
      sAttachments = Left(sAttachments, Len(sAttachments) - 2)
      
	  sBody = "<<< Total Loss Submission >>>" & VBCrLf & Now & VBCrLf & VBCrLf

      For Each oItem in .Form
        Response.Write oItem.Name & " = " & oItem.Value & "<BR>"
        'oFormMailer.sBody = oFormMailer.sBody &  oItem.Name & " = " & oItem.Value & VBCrLf
        sBody = sBody &  oItem.Name & " = " & oItem.Value & VBCrLf
        If oItem.Name = "Claim_Number" Then sClaimNum = oItem.Value
      Next  
    End With 
    
    PostActivity Activity_SubmitForm, "Claim#: " & sClaimNum
        
    Dim lErr
    On Error Resume Next

    SaveClaimSubmitToDb
    lErr = err.Number
          
    On Error Goto 0

    If lErr > 0 Then
      PostActivity Activity_SubmitForm, "ERROR #" & err.Number & " Submitting Claim#: " & sClaimNum

%>
<p align="center"><font size="4" color="#FF0000">There has been a problem trying to submit your request. The Server may temporarily be experiencing problems, please 
click the back button and resubmit your request.</font></p>
<p align="center"><font size="4" color="#FF0000">PLEASE Contact ACE if this problem persists.</font></p>
<p></p>
<%    
    End If
    
    On Error Goto 0
%> <br>
<b><font color="#008000">... Your Total Loss Valuation Request has been Submitted ...</font></b> <br>
<%    
    Response.End
  End If  


 Sub SaveClaimSubmitToDb()    
    Dim rsClaim 
    Dim sSQL
    Dim i
    Dim sAttach(5)
    
    sSQL = "SELECT * FROM tblClaimsSubmitted WHERE ClaimIDPK = 0"
    
    Set rsClaim = Server.CreateObject("ADODB.Recordset")
    rsClaim.Open sSQL, oData.ConnStr, 3, adLockOptimistic     'adOpenDynamic 3, adLockOptimistic
    'rsClaim.Open sSQL, oData.ConnStr, adOpenDynamic, adLockOptimistic     ' 3, adLockOptimistic

    rsClaim.AddNew
      rsClaim("ClaimBody") = sBody & VBCrLf & sAttachments
      rsClaim("ClaimNum") = sClaimNum
      rsClaim("SubmittedBy") =  oUpload.Form("Submitted_By").value 
	  rsClaim("EmailTo") = "subro@ace-it.com"
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

%>

            <style>
                input[type=text]:disabled {
                    background: #dddddd;
                }
            </style>

<table border="1" cellspacing="3"  cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td align="center" width="100%">
        <form  method="POST" enctype="multipart/form-data" name="frmTotalLoss" action="FrmTotalLoss.asp?txtState=1">
          <table border="0" cellspacing="6" width="433">
            <tr>
              <td colspan="3"><b><i><font face="Times New Roman" size="4"><img border="0" src="../Images/ACELogo_icon.gif">&nbsp; Total Loss Valuation Submission 
              Form</font></i></b></td>
            </tr>
            <tr>
              <td colspan="3" align="center" width="419"><img SRC="../ACE-IT_sv/Forms/formimages/coinfo.jpg" border="0" width="405" height="19"> </td>
            </tr>
            <tr>
              <td colspan="3" width="419"><small><center>&nbsp;&nbsp;&nbsp; All fields required.</center></small></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Company</font></td>
              <td width="294" colspan="2"><input type="text" name="Company_Name" style="width: 88%" value="<%=CoName%>"></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Submitted by</font></td>
              <td width="294" colspan="2"><input type="text" name="Submitted_By" style="width: 88%" value="<%=SubmittedBy%>"></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> ACE #</font></td>
              <td width="294" colspan="2"><input type="text" name="Office_ID" style="width: 88%" value="<%=OfficeID%>"></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Office</font></td>
              <td width="294" colspan="2"><input type="text" name="Office_Name" style="width: 88%" value="<%=OfficeName%>"></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font>&nbsp; Phone 
              Number</font></td>
              <td width="294" colspan="2"><input type="text" name="Submitted_By_Phone" size="20" value="<%=Phone%>"><font size="1">&nbsp; 
              Ext</font> <input type="text" 
      name="Submitted_By_Extension" size="5" maxlength="5" value="<%=PhoneEx%>"></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Fax Number</font></td>
              <td width="294" colspan="2"><input type="text" name="Submitted_By_Fax" size="20" value="<%=Fax%>"></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Email Address</font></td>
              <td width="294" colspan="2"><input type="text" name="Submitted_By_Email" style="width: 88%" value="<%=EmailAddr%>"></td>
            </tr>
            <tr>
              <td colspan="3" align="center" width="419"><img SRC="../ACE-IT_sv/Forms/formimages/claiminfo.jpg" border="0" width="405" height="19"> </td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Claim #</font></td>
              <td width="294" colspan="2"><input type="text" name="Claim_Number" style="width: 88%"></td>
            </tr>
			<tr>
			<tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Deductible</font></td>
              <td width="294" colspan="2"><input type="text" name="Deductible" style="width: 88%"></td>
            </tr>
            <tr>
                <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;<font size="1" face="verdana, geneva, arial">Insured (*if business)</font></td>
                <td width="294" colspan="2"><input type="text" name="Insured_Business" style="width: 88%" placeholder="Business Name" ng-model="input.insuredBusiness" ng-disabled="!!input.insuredPersonFirstName || !!input.insuredPersonLastName"></td>
            </tr>
            <tr>
                <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;<font size="1" face="verdana, geneva, arial">Insured (*if person)</font></td>
                <td width="294" colspan="2">
                    <input type="text" name="Insured_Person_First_Name" style="width: 43%;float: left;" placeholder="First Name" ng-model="input.insuredPersonFirstName" ng-disabled="!!input.insuredBusiness">
                    <input type="text" name="Insured_Person_Last_Name" style="width: 43%;margin-left: 3px" placeholder="Last Name" ng-model="input.insuredPersonLastName" ng-disabled="!!input.insuredBusiness">
                </td>
            </tr>
            <tr>
                <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;<font size="1" face="verdana, geneva, arial">Claimant (*if business)</font></td>
                <td width="294" colspan="2"><input type="text" name="Claimant_Business" style="width: 88%" placeholder="Business Name" ng-model="input.claimantBusiness" ng-disabled="!!input.claimantPersonFirstName || !!input.claimantPersonLastName"></td>
            </tr>
            <tr>
                <td bgcolor="#FFFFCC" align="right" valign="top">&nbsp;<font size="1" face="verdana, geneva, arial">Claimant (*if person)</font></td>
                <td width="294" colspan="2">
                    <input type="text" name="Claimant_Person_First_Name" style="width: 43%;float: left;" placeholder="First Name" ng-model="input.claimantPersonFirstName" ng-disabled="!!input.claimantBusiness">
                    <input type="text" name="Claimant_Person_Last_Name" style="width: 43%;margin-left: 3px" placeholder="Last Name" ng-model="input.claimantPersonLastName" ng-disabled="!!input.claimantBusiness">
                </td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> DOL</font></td>
              <td width="294" colspan="2"><input type="text" name="DOL" style="width: 88%"></td>
            </tr>
            <tr>
              <td colspan="3" align="center" width="419"><img SRC="../ACE-IT_sv/Forms/formimages/vehicleinfo.jpg" border="0" width="405" height="19"> </td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Year</font></td>
              <td width="294" colspan="2"><input type="text" name="Vehicle_Year" style="width: 88%" maxlength="40"> </td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Make</font></td>
              <td width="294" colspan="2"><input type="text" name="Vehicle_Make" style="width: 88%" maxlength="40"> </td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Model</font></td>
              <td width="294" colspan="2"><input type="text" name="Vehicle_Model" style="width: 88%" maxlength="40"> </td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Vin #</font></td>
              <td width="294" colspan="2"><input type="text" name="Vehicle_VIN" style="width: 88%"></td>
            </tr>
			<tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Odometer</font></td>
              <td width="294" colspan="2"><input type="text" name="Odometer" style="width: 88%"></td>
            </tr>
            <tr>
              <td bgcolor="#FFFFCC" align="right" valign="top" width="117"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> State</font></td>
              <td colspan="2" width="294"><select name="Vehicle_State" style="width: 264">
              <option>Select</option>
              <option value="AL">AL - Alabama</option>
              <option value="AK">AK - Alaska</option>
              <option value="AS">AS - American Samoa</option>
              <option value="AZ">AZ - Arizona</option>
              <option value="AR">AR - Arkansas</option>
              <option value="CA">CA - California</option>
              <option value="CO">CO - Colorado</option>
              <option value="CT">CT - Connecticut</option>
              <option value="DE">DE - Delaware</option>
              <option value="DC">DC - District of Columbia</option>
              <option value="FM">FM - Federated States of Micronesia</option>
              <option value="FL">FL - Florida</option>
              <option value="GA">GA - Georgia</option>
              <option value="GU">GU - Guam</option>
              <option value="HI">HI - Hawaii</option>
              <option value="ID">ID - Idaho</option>
              <option value="IL">IL - Illinois</option>
              <option value="IN">IN - Indiana</option>
              <option value="IA">IA - Iowa</option>
              <option value="KS">KS - Kansas</option>
              <option value="KY">KY - Kentucky</option>
              <option value="LA">LA - Louisiana</option>
              <option value="ME">ME - Maine</option>
              <option value="MH">MH - Marchall Islands</option>
              <option value="MD">MD - Maryland</option>
              <option value="MA">MA - Massachusetts</option>
              <option value="MI">MI - Michigan</option>
              <option value="MN">MN - Minnesota</option>
              <option value="MS">MS - Mississippi</option>
              <option value="MO">MO - Missouri</option>
              <option value="MT">MT - Montana</option>
              <option value="NE">NE - Nebraska</option>
              <option value="NV">NV - Nevada</option>
              <option value="NH">NH - New Hampshire</option>
              <option value="NJ">NJ - New Jersey</option>
              <option value="NM">NM - New Mexico</option>
              <option value="NY">NY - New York</option>
              <option value="NC">NC - North Carolina</option>
              <option value="ND">ND - North Dakota</option>
              <option value="MP">MP - Northern Mariana Islands</option>
              <option value="OH">OH - Ohio</option>
              <option value="OK">OK - Oklahoma</option>
              <option value="OR">OR - Oregon</option>
              <option value="PW">PW - Palau</option>
              <option value="PA">PA - Pennsylvania</option>
              <option value="PR">PR - Puerto Rico</option>
              <option value="RI">RI - Rhode Island</option>
              <option value="SC">SC - South Carolina</option>
              <option value="SD">SD - South Dakota</option>
              <option value="TN">TN - Tennessee</option>
              <option value="TX">TX - Texas</option>
              <option value="UT">UT - Utah</option>
              <option value="VT">VT - Vermont</option>
              <option value="VI">VI - Virgin Islands</option>
              <option value="VA">VA - Virginia</option>
              <option value="WA">WA - Washington</option>
              <option value="WV">WV - West Virginia</option>
              <option value="WI">WI - Wisconsin</option>
              <option value="WY">WY - Wyoming</option>
              </select></td>
            </tr>
            <tr>
              <td colspan="3" align="center" width="419"><img SRC="../ACE-IT_sv/Forms/formimages/enginedtls.jpg" border="0" width="405" height="19"> </td>
            </tr>
            <tr>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Power Options</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Power_Steering">Steering<br>
                  <input type="checkbox" value="1" name="Options_Power_Brakes">Brakes<br>
                  <input type="checkbox" value="1" name="Options_Power_Windows">Windows<br>
                  <input type="checkbox" value="1" name="Options_Power_Locks">Locks<br>
                  <input type="checkbox" value="1" name="Options_Power_Driver_Seat">Driver Seat<br>
                  <input type="checkbox" value="1" name="Options_Power_Passenger_Seat">Pass Seat<br>
                  <input type="checkbox" value="1" name="Options_Power_Antenna">Antenna<br>
                  <input type="checkbox" value="1" name="Options_Power_Mirrors">Mirrors<br>
                  <input type="checkbox" value="1" name="Options_Power_Trunk">Trunk</font></td>
                </tr>
              </table>
              </td>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Radio</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Radio_AM">AM<br>
                  <input type="checkbox" value="1" name="Options_Radio_FM">FM<br>
                  <input type="checkbox" value="1" name="Options_Radio_Stereo">Stereo<br>
                  <input type="checkbox" value="1" name="Options_Radio_Cassette">Cassette<br>
                  <input type="checkbox" value="1" name="Options_Radio_SeekScan">Seek/Scan<br>
                  <input type="checkbox" value="1" name="Options_Radio_Equalizer">Equalizer<br>
                  <input type="checkbox" value="1" name="Options_Radio_CDPlayer">CD Player<br>
                  <input type="checkbox" value="1" name="Options_Radio_CDRadio">CB Radio<br>
                  </font></td>
                </tr>
              </table>
              </td>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Seats</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Seats_Cloth">Cloth<br>
                  <input type="checkbox" value="1" name="Options_Seats_Leather">Leather<br>
                  <input type="checkbox" value="1" name="Options_Seats_Third">Third<br>
                  <input type="checkbox" value="1" name="Options_Seats_Buckets">Buckets<br>
                  <input type="checkbox" value="1" name="Options_Seats_ReclineLounge">Recline/Lounge<br>
                  <input type="checkbox" value="1" name="Options_Seats_Hiback_Buckets">Hiback Buckets<br>
                  <input type="checkbox" value="1" name="Options_Seats_Rear_Buckets">Rear Buckets<br>
                  <input type="checkbox" value="1" name="Options_Seats_Recaro">Recaro<br>
                  <input type="checkbox" value="1" name="Options_Seats_Split_Bench">Split Bench </font></td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Decor</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Decor_Tinted_Glass">Tinted Glass<br>
                  <input type="checkbox" value="1" name="Options_Decor_Bodyside_Moulding">Bodyside<br>
&nbsp;&nbsp;&nbsp;&nbsp; Moulding<br>
                  <input type="checkbox" value="1" name="Options_Decor_Woodgrain">Woodgrain<br>
                  <input type="checkbox" value="1" name="Options_Decor_Bumper_Cushings">Bumper Cushings<br>
                  <input type="checkbox" value="1" name="Options_Decor_Bumper_Guards">Bumper Guards<br>
                  <input type="checkbox" value="1" name="Options_Decor_Dual_Mirrors">Dual Mirrors<br>
                  <input type="checkbox" value="1" name="Options_Decor_Custom_Int">Custom Int.<br>
                  <input type="checkbox" value="1" name="Options_Decor_Luxury_Interior">Luxury Interior<br>
                  <input type="checkbox" value="1" name="Options_Decor_Privacy_Glass">Privacy Glass<br>
                  <input type="checkbox" value="1" name="Options_Decor_Roof_Console">Roof Console </font></td>
                </tr>
              </table>
              </td>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Safety</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Safety_Antilock_Bks_4">Antilock 
                  Brakes (4)<br>
                  <input type="checkbox" value="1" name="Options_Safety_Antilock_Bks_2">Antilock Brakes (2)<br>
                  <input type="checkbox" value="1" name="Options_Safety_Auto_Restraint">Auto Restraint<br>
                  <input type="checkbox" value="1" name="Options_Safety_Driver_Airbag">Driver Airbag<br>
                  <input type="checkbox" value="1" name="Options_Safety_Pass_Airbag">Pass Airbag<br>
                  <input type="checkbox" value="1" name="Options_Safety_4_WL_Disc_Bks">4 Wl Disc Brakes<br>
                  <input type="checkbox" value="1" name="Options_Safety_Rollbar">Roll Bar<br>
                  <input type="checkbox" value="1" name="Options_Safety_Positraction">Positraction </font></td>
                </tr>
              </table>
              </td>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Wheels</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Wheels_Aluminum">Aluminum<br>
                  <input type="checkbox" value="1" name="Options_Wheels_Wire">Wire<br>
                  <input type="checkbox" value="1" name="Options_Wheels_Wire_Covers">Wire Covers<br>
                  <input type="checkbox" value="1" name="Options_Wheels_Styled_Steel">Styled Steel<br>
                  <input type="checkbox" value="1" name="Options_Wheels_Alloy">Alloy<br>
                  <input type="checkbox" value="1" name="Options_Wheels_Spoke_Aluminum">Spoke<br>
&nbsp;&nbsp;&nbsp;&nbsp; Aluminum<br>
                  <input type="checkbox" value="1" name="Options_Wheels_Locking_Covers">Locking Covers<br>
                  <input type="checkbox" value="1" name="Options_Wheels_Locking_Wheels">Locking<br>
&nbsp;&nbsp;&nbsp; Wheels</font></td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Convenience</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Convenience_Airconditioning">Air 
                  Conditioning<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Rear_Defogger">Rear Defogger<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Tilewheel">Tilt Wheel<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Cruisecontrol">Cruise Control<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Telescopic_Wheel">Telescopic Wheel<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Intermittent_wipers">Intermittent <br>
&nbsp;&nbsp;&nbsp;&nbsp; Wipers<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Aux_Fuel_Tank">Aux Fuel Tank<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Autolevel">Auto Level<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Climate_control">Climate Control<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Electronic_inst_panel">Electronic <br>
&nbsp;&nbsp;&nbsp;&nbsp; Inst. Panel<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Rear_window_wiper">R Window Wiper<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Theft_alarm">Theft Alarm<br>
                  <input type="checkbox" value="1" name="Options_Convenience_Fog_lights">Fog Lights</font></td>
                </tr>
              </table>
              </td>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Roof</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial">
                  <input type="checkbox" value="1" name="Options_Roof_Vinyl">Vinyl<br>
                  <input type="checkbox" value="1" name="Options_Roof_Landau">Landau<br>
                  <input type="checkbox" value="1" name="Options_Roof_Luggage_rf_rack">Luggage Rf Rack<br>
                  <input type="checkbox" value="1" name="Options_Roof_Padded_vinyl">Padded Vinyl<br>
                  <input type="checkbox" value="1" name="Options_Roof_Padded_landau">Padded Landau<br>
                  <input type="checkbox" value="1" name="Options_Roof_Cabriolet">Cabriolet<br>
                  <input type="checkbox" value="1" name="Options_Roof_Elec_steel_sunrf">Elec Steel Sunrf<br>
                  <input type="checkbox" value="1" name="Options_Roof_Elec_glass_sunrf">Elec Glass Sunrf<br>
                  <input type="checkbox" value="1" name="Options_Roof_Man_steel_sunrf">Man Steel Sunrf<br>
                  <input type="checkbox" value="1" name="Options_Roof_Man_glass_sunrf">Man Glass Sunrf<br>
                  <input type="checkbox" value="1" name="Options_Roof_Glass_t_tops">Glass T-Tops<br>
                  <input type="checkbox" value="1" name="Options_Roof_T-top">T-Top<br>
                  <input type="checkbox" value="1" name="Options_Roof_FLip_roof">Flip Roof</font></td>
                </tr>
              </table>
              </td>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Paint</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial"><input type="checkbox" value="1" name="paintclearcoatpaint">Clear Coat<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  Paint<br>
                  <input type="checkbox" value="1" name="painttwotonepaint">Two Tone Paint<br>
                  <input type="checkbox" value="1" name="paintthreestagepaint">Three Stage<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Paint<br>
                  <input type="checkbox" value="1" name="paintmetallic">Metallic<br>
                  <input type="checkbox" value="1" name="paintdeluxetwotone">Deluxe Two<br>
&nbsp;&nbsp;&nbsp;&nbsp; Tone<br>
                  <input type="checkbox" value="1" name="paintcustom">Custom</font></td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Truck</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial"><input type="checkbox" value="1" name="truckrearstepbumper">Rear Step<br>
&nbsp;&nbsp;&nbsp;&nbsp; 
                  Bumper<br>
                  <input type="checkbox" value="1" name="truckslidingrearwindow">Sliding R Window<br>
                  <input type="checkbox" value="1" name="truckrunningboards">Running Boards<br>
                  <input type="checkbox" value="1" name="truckbedlinerduraliner">Bedliner/Duraliner<br>
                  <input type="checkbox" value="1" name="truckchromebed">Chrome Bed<br>
                  <input type="checkbox" value="1" name="trucktoolbox">Tool Box<br>
                  <input type="checkbox" value="1" name="truckgrilleguards">Grille Guards<br>
                  <input type="checkbox" value="1" name="truckdualrearwheels">Dual Rear Wheels<br>
                  <input type="checkbox" value="1" name="trucktraileringpackage">Trailering<br>
&nbsp;&nbsp;&nbsp;&nbsp; Package<br>
                  <input type="checkbox" value="1" name="truckduelac">Dual A/C </font></td>
                </tr>
              </table>
              </td>
              <td valign="top">
              <table>
                <tr>
                  <td valign="top" width="130" bgcolor="#FFFFCC"><font size="1" face="verdana, geneva, arial"><center>Motorcycle</center></font></td>
                </tr>
                <tr>
                  <td valign="top" width="130"><font size="1" face="verdana, geneva, arial"><input type="checkbox" value="1" name="motorcyclesheader">Header<br>
                  <input type="checkbox" value="1" name="motorcyclesfullflairing">Full Fairing<br>
                  <input type="checkbox" value="1" name="motorcyclesplexiglass">Plexiglass<br>
                  <input type="checkbox" value="1" name="motorcyclescustomseats">Custom Seats<br>
                  <input type="checkbox" value="1" name="motorcyclessaddlebags">Saddle Bags<br>
                  <input type="checkbox" value="1" name="motorcyclestraveltrunk">Travel Trunk<br>
                  <input type="checkbox" value="1" name="motorcyclesengineguard">Engine Guard<br>
                  <input type="checkbox" value="1" name="motorcyclesbackrest">Back Rest<br>
                  <input type="checkbox" value="1" name="motorcyclescruisecontrol">Cruise Control<br>
                  <input type="checkbox" value="1" name="motorcyclesluggagerack">Luggage Rack </font></td>
                </tr>
              </table>
              </td>
            </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" width="117"><font face="verdana, geneva, arial" size="1">Other Options or Comments</font></td>
      <td colspan="2"><textarea rows="3" name="Comments" cols="34"></textarea></td>
    </tr>


<!-- -------------------------------------------------------------------------------------- -->
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="bottom"><font size="1" face="verdana, geneva, arial"><font size="3">&nbsp;</font> Attachment 1&nbsp;&nbsp; </font>
      </td>
      <td colspan="2">
          <font size="1" face="verdana, geneva, arial">To attach a file related to this claim use the<br />browse button and navigate your computer to<br />locate and attach the necessary file.
          <br>
          <b><font color="#008080">Please note file size cannot exceed 80MB.</font></b>
          <br>
      </font>
      <input type="file" name="attachedfile1" size="27"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial">Attachment&nbsp;2&nbsp;&nbsp; </font>
      </td>
      <td colspan="2">
      <input type="file" name="attachedfile2" size="27"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial">Attachment&nbsp;3 &nbsp; </font>
      </td>
      <td colspan="2">
      <input type="file" name="attachedfile3" size="27"> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial">Attachment&nbsp;4 &nbsp; </font>
      </td>
      <td colspan="2">
      <input type="file" name="attachedfile4" size="27"> </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFCC" align="right" valign="top"><font size="1" face="verdana, geneva, arial">Attachment&nbsp;5 &nbsp; </font>
      </td>
      <td colspan="2">
      <input type="file" name="attachedfile5" size="27"> </td>
    </tr>

<!-- -------------------------------------------------------------------------------------- -->
            
    <tr>    
      <td bgcolor="#FFFFCC" colspan="3" align="center"><b><font face="Arial" size="2" color="#800000">You must click submit for your request to be processed</font><font 
      color="#000080"> </font></b> <br>
      <input type="button" value="Submit" name="cmdSubmit" onClick="SubmitForm(this);" style="font-weight: bold"></td>
    </tr>
            
          </table>
        </form>
        </td>
        </tr>
        </table>
<!-- END main body / START footer -->
<script>
  window.moveTo(0, 25);
  window.resizeTo(500, screen.availHeight-25);
</script>


</body>

</html>
<%
  Set oUpload = Nothing
%>
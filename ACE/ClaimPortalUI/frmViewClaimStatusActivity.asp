<% Option Explicit %>

<!-- #INCLUDE Virtual="/ACE/Login/Security.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsACEDataConn.asp" -->
<!-- #INCLUDE Virtual="/ACE/inc/clsSQLStmt.asp" -->
<%
  Dim oData, oSQL
  Dim rs
  Dim RecCount
  
  Set oData = New clsACEDataConn
  'Set oSQL = New clsSQLStmt
    
  Dim sEstID
        
  sEstID = Request.Querystring("EstID")
  
  '--- Populate recordset ---
  Set rs = oData.rsStatusActivityList(sEstID)
  RecCount = rs.RecordCount
  
  
  Set rAssignmentData = dbSQLSvr_ACE_Auto_Audits.OpenRecordset( _
     "SELECT InsuranceCompanyIDPK, InsuranceCompanyName, InsuranceCompanyAddr1, InsuranceCompanyAddr2, InsuranceCompanyCity, InsuranceCompanyState, InsuranceCompanyZip, Advertisement, ACEInfoOnly, SupressInvoice, InsLastName, InsFirstName, ClmLastName, ClmFirstName, FileNumber, PolicyNo, ClaimNo, LossDate, BillingDate, Notes, ClaimRepName, ClaimRepFax, ClaimRepPhone, ClaimRepPhoneExt, Hours, AceFeeUsed, ClosedDate, Free, EmailClaim, Email, BodyShopEstAmt, AgreedPrice " & _
     "FROM ((tblAssignments LEFT JOIN tblInsuranceCompanies  ON tblAssignments.InsuranceCompanyID = tblInsuranceCompanies.InsuranceCompanyIDPK) LEFT JOIN tblClaimReps ON tblAssignments.ClaimRepID = tblClaimReps.ClaimRepIDPK) " & _
     "WHERE (" & ASSIGNMENTS_TABLE_NAME & ".EstimateIDPK = " & EstimateID & ");" _
     , dbOpenSnapshot, dbSQLPassThrough)


    "SELECT sh.IDPK, AssignmentID, StatusDTM, StatusDesc, Notes " & _
    "FROM   tblAssignmentsStatusHistory sh " & _
              "INNER JOIN tblAssignmentsStatusTypes " & _
                "ON sh.StatusID = tblAssignmentsStatusTypes.IDPK " & _
    "WHERE  (AssignmentID = " & AuditID & ") " & _
    "ORDER BY StatusDTM Desc, sh.IDPK Desc " _
    , dbOpenForwardOnly, dbSQLPassThrough)
  
%>


<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ACE - Claim Status Activity</title>
</head>

<body>

</body>

</html>
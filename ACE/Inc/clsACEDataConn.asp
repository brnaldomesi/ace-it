<%
Const Activity_Login = 1
Const Activity_WelcomeEmail = 2
Const Activity_AccountCreated = 3
Const Activity_ViewAuditPDF = 4
Const Activity_SearchForAudit = 5
Const Activity_ValidationEmailRcv = 6
Const Activity_AccountEnabled = 7
Const Activity_AccountDisabled = 8
Const Activity_SeachBodyShopNet = 9
Const Activity_ViewPage = 10
Const Activity_SubmitForm = 11


Const MemberType_ClaimRep = 1
Const MemberType_Agent = 2
Const MemberType_Administrator = 3
Const MemberType_OfficeAdmin = 4
Const MemberType_ACEEmployee = 5
Const MemberType_ACEAuditor = 6
Const MemberType_ACEClerical = 7
Const MemberType_ACEAdmin = 8

Private Const PSSWebServIP = "webportal.staging-portal.ace-it.com" '"webportal.staging-portal.ace-it.com" ' "64.221.61.162"

Private Const SQLSvrIP = "sql.ace-it.com" 
'Private Const SQLSvrPort = "56102"
Private Const SQLSvrPort = "1433"
Private Const SQLSvrUserName = "WebPortal"
Private Const SQLSvrPassword = "Portal4ACE"
Private Const SQLSvrInitCatalog = "ACE_WEB_Portal"


'Private Const SQLSvrOLEDBConnStr = _
'                "Provider=SQLOLEDB.1;" & _
'                "Network Library=DBMSSOCN;" & _
'                "Data Source=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
'                "User ID=" & SQLSvrUserName & ";" & _
'                "Password=" & SQLSvrPassword & ";" & _
'                "Initial Catalog=" & SQLSvrInitCatalog

'Private Const mMemberDBName = "Data/Members.gif"
Private Const mMemberDBName = "Data/ACEMembers.gif"

'=========================================================================================
' Start Class definition
'=========================================================================================

Class clsACEDataConn

  Public mConnStr

  Public sAdminConn 
  
  public sAdminConn2

'-----------------------------------------------------------------------------------------
  Private Sub Class_Initialize

    'mConnStr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & Server.MapPath("\" & mMemberDBName) & ";"

    mConnStr = _
            "Provider=SQLNCLI11;" & _
            "Network Library=DBMSSOCN;" & _
            "Data Source=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
            "User ID=" & SQLSvrUserName & ";" & _
            "Password=" & SQLSvrPassword & ";" & _
            "Initial Catalog=" & SQLSvrInitCatalog & ";" & _
			"Use Encryption for Data=True;Trust Server Certificate=True;"
			'"Encrypt=Yes;TrustServerCertificate=yes;"


'    mConnStr = _
'			"Driver={SQL Server Native Client 11.0};" & _
'			"Server=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
'			"Database=" & SQLSvrInitCatalog & ";" & _
'			"user id=" & SQLSvrUserName & ";" & _
'			"password=" & SQLSvrPassword & ";" & _
'			"Encrypt=yes;TrustServerCertificate=yes"

    sAdminConn = _
            "Provider=SQLOLEDB;" & _
            "Network Library=DBMSSOCN;" & _
            "Data Source=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
            "Persist Security Info=True;" & _
            "User ID=SA;" & _
            "Password=i8@enzos;" & _
            "Initial Catalog=ACE_Auto_Audits" & ";" & _
			"Encrypt=true;TrustServerCertificate=true;"
			
			
    sAdminConn = _
			"Driver={SQL Server Native Client 11.0};" & _
			"Server=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
			"Database=ACE_Auto_Audits" & ";" & _
			"user id=SA;" & _
			"password=i8@enzos;" & _
			"Encrypt=yes;TrustServerCertificate=yes"			


    sAdminConn2 = _
            "Provider=SQLOLEDB;" & _
            "Network Library=DBMSSOCN;" & _
            "Data Source=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
            "Persist Security Info=True;" & _
            "User ID=SA;" & _
            "Password=i8@enzos;" & _
            "Initial Catalog=" & SQLSvrInitCatalog & ";" & _
			"Encrypt=true;TrustServerCertificate=true;"
            
    sAdminConn2 = _
			"Driver={SQL Server Native Client 11.0};" & _
			"Server=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
			"Database=" & SQLSvrInitCatalog & ";" & _
			"user id=SA;" & _
			"password=i8@enzos;" & _
			"Encrypt=yes;TrustServerCertificate=yes"
			
			
			
            
'    mConnStr = _
'            "Provider=SQLOLEDB;" & _
'			"Data Source=" & SQLSvrIP & "," & SQLSvrPort & ";" & _
'			"Network Library=DBMSSOCN;" & _
'			"Initial Catalog=" & SQLSvrInitCatalog & ";" & _
'			"User ID=" & SQLSvrUserName & ";" & _
'			"Password=" & SQLSvrPassword & ";" 

'    mConnStr = _
'			"Driver={SQL Server};Server=" & SQLSvrIP & ";Address=" & SQLSvrIP & "," & SQLSvrPort & ";" & ";Network=DBMSSOCN;Database=" & SQLSvrInitCatalog & ";UID=" & SQLSvrUserName & ";PWD=" & SQLSvrPassword 

            
  End Sub

'-----------------------------------------------------------------------------------------
  Private Sub Class_Terminate
     On Error Resume Next
'    mConn.Close
'    Set mConn = Nothing
  End Sub

'-----------------------------------------------------------------------------------------
  Public Function ConnStr()
    ConnStr = mConnStr
  End Function
  
'-----------------------------------------------------------------------------------------
  Public Function ConnMemberDB()
    Dim Conn

    Set Conn = Server.CreateObject("adodb.connection")
    Conn.Open mConnStr
    Set ConnMemberDB = Conn
  End Function


'-----------------------------------------------------------------------------------------
' Call stored procedure methods
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
  Public Function rsCallStoredProc (sProcName, sParam1, sParam2, sParam3)
    'This method queries the data source about the parameters
    'of the stored procedure. This is the least efficient method of calling
    'a stored procedure.
    Dim cmd
    Set rsCallStoredProc = Server.CreateObject("ADODB.Recordset")
    rsCallStoredProc.CursorLocation = 3 'adUseClient
    'Set cn = Server.CreateObject("ADODB.Connection")
    Set cmd = Server.CreateObject("ADODB.Command")
    'cn.Open mConnStr
    'Set cmd.ActiveConnection = cn
    cmd.ActiveConnection = mConnStr
    cmd.CommandText = sProcName
    cmd.CommandType = 4
    ' Ask the server about the parameters for the stored proc
    cmd.Parameters.Refresh
    ' Assign values to parameters. Index of 0 represents first parameter.
    If Len(CStr(sParam1)) > 0 Then cmd.Parameters(1) = sParam1
    If Len(CStr(sParam2)) > 0 Then cmd.Parameters(2) = sParam2
    If Len(CStr(sParam3)) > 0 Then cmd.Parameters(3) = sParam3
    rsCallStoredProc.Open cmd
  End Function
  
  Public Function rsCallStoredProc2 (sProcName, sParam1, sParam2, sParam3)
   '** Not finished
   'This method declares the stored procedure, and then explicitly declares the parameters.
   Set cn = Server.CreateObject("ADODB.Connection")
   cn.Open mConnStr
   Set cmd = Server.CreateObject("ADODB.Command")
   Set cmd.ActiveConnection = cn
   cmd.CommandText = sProcName
   cmd.CommandType = 4
   cmd.Parameters.Append cmd.CreateParameter("RetVal", adInteger, adParamReturnValue)
   cmd.Parameters.Append cmd.CreateParameter("Param1", adInteger, adParamInput)
   ' Set value of Param1 of the default collection to 22
   cmd("Param1") = 22
   cmd.Execute
   'ReturnValue = &lt;% Response.Write cmd(0) %&gt;&lt;P&gt;
  End Function

  Public Function rsCallStoredProc3 (sProcName, sParam1, sParam2, sParam3)
   '** Not finished
   'This method is probably the most formal way of calling a stored procedure.
   'It uses the canocial name
   Set cn = Server.CreateObject("ADODB.Connection")
   cn.Open mConnStr
   Set cmd = Server.CreateObject("ADODB.Command")
   Set cmd.ActiveConnection = cn
   ' Define the stored procedure's inputs and outputs
   ' Question marks act as placeholders for each parameter for the
   ' stored procedure
   'cmd.CommandText = "{?=call sp_test(?)}"
   cmd.CommandText = "{?=call " & sProcName & "sp_test(?)}"
   ' specify parameter info 1 by 1 in the order of the question marks
   ' specified when we defined the stored procedure
   cmd.Parameters.Append cmd.CreateParameter("RetVal", adInteger, adParamReturnValue)
   cmd.Parameters.Append cmd.CreateParameter("Param1", adInteger, adParamInput)
   cmd.Parameters("Param1") = 33
   cmd.Execute
   'ReturnValue = &lt;% Response.Write cmd("RetVal") %&gt;&lt;P&gt;
  End Function




'-----------------------------------------------------------------------------------------
  Public Function ACEMemberGet(MemberID, strMemberUserName, strOrderBy)
    Dim Conn, rs, SQL

'    SQL = "SELECT tblMembers.*, tblInsCompanies.InsCoName " & _
'          "FROM tblInsCompanies RIGHT JOIN tblMembers ON tblInsCompanies.InsCompanyIDPK = tblMembers.InsCompanyID "

'    SQL = "SELECT tblMembers.*, Ins.InsuranceCompanyName AS InsCoName, Ins.InsuranceCompanyOfficeLocation AS ACEInsCoOfficeName " & _
'          "FROM vw_tblInsuranceCompanies Ins RIGHT JOIN tblMembers ON Ins.InsuranceCompanyIDPK = tblMembers.ACEInsCoOfficeID "


     SQL = "SELECT MemberIDPK, MemberTypeID, MemberLogonName, MemberPassword, MemberSecurityLevel, MemberForumAdmin, MemberForumModerator, " & _
                  "MemberEmailBody, MemberConfEmailRespRecvd, MemberAccountEnabled, MemberIP, MemberFirstName, MemberLastName, MemberEmail, " & _
                  "MemberGender, MemberPhone, MemberPhoneEx, MemberFax, ACERepName, ACEInsCoOfficeID, MemberRejected, " & _
                  "Ins.InsuranceCompanyName AS InsCoName, Ins.InsuranceCompanyOfficeLocation AS ACEInsCoOfficeName, ShopNetworkAccess, ShopDemoAccess " & _
                  ", pq.InsCoOfficeID as PQ_OfficeID, pq.PortalSubmissionFormName as PQ_FormName " & _
                  ", a.InsCoOfficeID as Auto_OfficeID, a.PortalSubmissionFormName as Auto_FormName " & _
                  ", s.InsCoOfficeID as Subro_OfficeID, s.PortalSubmissionFormName as Subro_FormName " & _
                  ", hs.InsCoOfficeID as HeavySpecialty_OfficeID, hs.PortalSubmissionFormName as HeavySpecialty_FormName " & _
                  ", p.InsCoOfficeID as Property_OfficeID, p.PortalSubmissionFormName as Property_FormName " & _
                  ", tl.InsCoOfficeID as AutoTotalLoss_OfficeID, tl.PortalSubmissionFormName as AutoTotalLoss_FormName " & _
           "FROM vw_tblInsuranceCompanies Ins " & _
			"RIGHT JOIN tblMembers ON Ins.InsuranceCompanyIDPK = tblMembers.ACEInsCoOfficeID " & _
			"Left Outer Join tblMembers_InsOffices_Forms_XRef pq On pq.MemberID = tblMembers.MemberIDPK And pq.PortalSubmissionFormTypeCode = 'PhotoQuote' " & _
			"Left Outer Join tblMembers_InsOffices_Forms_XRef a On a.MemberID = tblMembers.MemberIDPK And a.PortalSubmissionFormTypeCode = 'Auto' " & _
			"Left Outer Join tblMembers_InsOffices_Forms_XRef s On s.MemberID = tblMembers.MemberIDPK And s.PortalSubmissionFormTypeCode = 'Subrogation' " & _
			"Left Outer Join tblMembers_InsOffices_Forms_XRef hs On hs.MemberID = tblMembers.MemberIDPK And hs.PortalSubmissionFormTypeCode = 'Heavy/Specailty' " & _
			"Left Outer Join tblMembers_InsOffices_Forms_XRef p On p.MemberID = tblMembers.MemberIDPK And p.PortalSubmissionFormTypeCode = 'Property' " & _
			"Left Outer Join tblMembers_InsOffices_Forms_XRef tl On tl.MemberID = tblMembers.MemberIDPK And tl.PortalSubmissionFormTypeCode = 'AutoTotalLoss' "


'    If MemberID = 0 Then
'      SQL = SQL & "WHERE (((tblMembers.MemberLogonName)=""" &  strMemberUserName & """)) "
'    ElseIf Len(strMemberUserName) <> 0 Then
'      SQL = SQL & "WHERE (((tblMembers.MemberIDPK)=" &  MemberID & ")) "
'    End If

    If Len(strMemberUserName) <> 0 Then
      SQL = SQL & "WHERE (((tblMembers.MemberLogonName)='" &  strMemberUserName & "')) "
    ElseIf MemberID <> 0 Then
      SQL = SQL & "WHERE (((tblMembers.MemberIDPK)=" &  MemberID & ")) "
    End If

    IF len(strOrderBy) <> 0 Then SQL = SQL & "ORDER BY tblMembers." & strOrderBY & " "

    SQL = SQL & ";"

    'Set Conn = ConnMemberDB

    Set rs = Server.CreateObject("ADODB.Recordset")

    ' Setting the cursor location to client side is important to get a disconnected recordset.
    rs.CursorLocation = 3 'adUseClient
    'rs.Open SQL, Conn, 0, 4
    'rs.Open SQL, Conn, adOpenDynamic, 4
    rs.Open SQL, mConnStr, 3, 4
    
    ' Disconnect the recordset.
    Set Rs.ActiveConnection = Nothing

'    Conn.Close
'    Set Conn = Nothing

    Set ACEMemberGet = rs

    'rs.close
    'Set rs=Nothing
  End Function

'-----------------------------------------------------------------------------------------
'  Public Function ACEGetCoID(CoName)
'    Dim Conn, rs, SQL
'
'    ACEGetCoID=0
'
'    SQL = "SELECT tblInsCompanies.InsCompanyIDPK " & _
'          "FROM tblInsCompanies " & _
'          "WHERE tblInsCompanies.InsCoName = '" & CoName & "';"'
'
'    Set rs = Server.CreateObject("ADODB.Recordset")
'    rs.Open SQL, mConnStr, 3, 4
'
'    If Not rs.eof Then ACEGetCoID=rs(0)
'
'    rs.close
'    Set rs=Nothing
'  End Function

'-----------------------------------------------------------------------------------------
  Public Sub ACEMemberUpdate(ByRef rs)
    Dim Conn, UserName

    ' Now edit the value and save it.
    'Rs.Fields("au_lname").Value = "NewValue"


    ' reopen the connection and attach it to the recordset.
    'Set Conn = ConnMemberDB
    'Set rs.ActiveConnection = Conn

    rs.ActiveConnection = mConnStr  '<--- We don't need to create a connection object just set to connection string
    rs.UpdateBatch
    'rs.resync

    '--- Disconnect the recordset. ---
    Set rs.ActiveConnection = Nothing

    'Conn.Close
    'Set Conn = Nothing
  End Sub

'-----------------------------------------------------------------------------------------
  Public Sub ACEMemberActivityInsert(MemberID, ActivityID)
    Dim Conn, SQL, rs

'    SQL = "INSERT INTO tblMemberActivity ( MemberID, MActivityDateTime, ActivityTypeID ) " & _
'          "VALUES (" & MemberID & ", #" & Now() & "#, " & ActivityID & ");"

'    '--- we don't need to explicitly set time timestamp ---
'    SQL = "INSERT INTO tblMemberActivity ( MemberID, ActivityTypeID ) " & _
'          "VALUES (" & MemberID & ", " & ActivityID & ");"

'Exit Sub

    '--- we don't need to explicitly set time timestamp ---
    SQL = "INSERT INTO tblMemberActivity ( MemberID, ActivityTypeID, MActivityMemberIP ) " & _
          "VALUES (" & MemberID & ", " & ActivityID & ", '" & Request.ServerVariables("REMOTE_HOST") & "');" 

    Set Conn = ConnMemberDB
    Conn.execute(SQL)
    conn.close
    Set conn=Nothing

'--- use recordset to add new rec ---
'    SQL = "SELECT * FROM tblMemberActivity WHERE MActivityIDPK=0;"
'    Set rs = Server.CreateObject("adodb.Recordset")
'    rs.Open SQL, Conn, 3, 4
'    rs.AddNew
'    rs("MemberID")=MemberID
'    rs("MActivityDateTime")=Now()
'    rs("ActivityTypeID")=ActivityID
'    rs.Update
'    rs.close
'    Set rs = Nothing

  End Sub

'-----------------------------------------------------------------------------------------
  Public Sub ACEMemberActivityInsert2(MemberID, ActivityID, sInfo)
    Dim Conn, SQL

'Exit Sub

    '--- we don't need to explicitly set time timestamp ---
    SQL = "INSERT INTO tblMemberActivity ( MemberID, ActivityTypeID, MActivityInfo1 ) " & _
          "VALUES (" & MemberID & ", " & ActivityID & ", '" & Replace(sInfo,"'","''") & "');"

    Set Conn = ConnMemberDB
    Conn.execute(SQL)
    conn.close
    Set conn=Nothing

  End Sub

'-----------------------------------------------------------------------------------------
  Public Sub ACEMemberInsert(rs)
    ' We should change this to a function and return the ID of the newly inserted record
    ' with a SELECT @@IDENTITY query - See MSDN favorites SQL-
    Dim Conn, SQL

    '--- First See If Company already exists ---

'    SQL = "INSERT INTO tblMembers( MemberTypeID, MemberLogonName, MemberPassword, MemberSecurityLevel, MemberForumAdmin, MemberForumModerator, MemberConfEmailRespRecvd, MemberAccountEnabled, MemberIP, MemberFirstName, MemberLastName, MemberEmail, MemberGender, MemberPhone, MemberPhoneEx, MemberFax, ACERepName, MInsCompanyName, InsCompanyID ) " & _
'          "VALUES (" & MemberID & ", #" & Now() & "#, " & ActivityID & ");"
'    Set Conn = ConnMemberDB
'    Conn.execute(SQL)
'    conn.close
'    Set conn=Nothing

    rs.AddNew
    rs("MemberTypeID") = 1
    rs("MemberLogonName") = Request.Form("txtUserName")
    rs("MemberPassword") = Request.Form("txtPassword")
    rs("MemberSecurityLevel") = 1
'    rs("MemberForumAdmin")
'    rs("MemberForumModerator")
'    rs("MemberConfEmailRespRecvd")
'    rs("MemberAccountEnabled")
'    rs("MemberIP")
    rs("MemberFirstName") = Request.Form("txtFirstName")
    rs("MemberLastName")  = Request.Form("txtLastName")
    rs("MemberEmail") = Request.Form("txtEmail")
    rs("MemberGender") = Request.Form("grpGender")
    rs("MemberPhone") = zn(Request.Form("txtPhoneAc") & Request.Form("txtPhoneExch") & Request.Form("txtPhoneL4"))
    rs("MemberPhoneEx") = zn(Request.Form("txtPhoneExtn"))
    rs("MemberFax") = zn(Request.Form("txtCoFaxAc") & Request.Form("txtCoFaxExch") & Request.Form("txtCoFaxL4"))
    rs("ACERepName") = zn(Trim(UCase(Request.Form("txtFirstName")))) & " " & zn(Trim(UCase(Request.Form("txtLastName"))))
'    rs("MInsCompanyName") = zn(Request.Form("txtCoName"))
'    rs("InsCompanyID")

    ACEMemberUpdate(rs)

  End Sub
  
'  Public Function CoList() 
'    Dim rs
'    Set rs = RsGet("SELECT InsuranceCopmanyName FROM vw_tblInsuranceCompanies Order By InsuranceCompanyName")
'  End Function




'=========================================================================================
' Company List
'=========================================================================================
  Public Function sCompanyHTMLOptionList(sFilterCompany)
	'Dim sConn, rs, sSQL
    Dim oCn, oCmd, oPrm, rs
	
    Set oCn = Server.CreateObject("ADODB.Connection")
    oCn.Open mConnStr 'sAdminConn2 '
    Set oCmd = Server.CreateObject("ADODB.Command")
    oCmd.CommandText = "Portal_CompanyList"
    oCmd.CommandType = 4 '4
    Set oCmd.ActiveConnection = oCn

    ' Automatically fill in parameter info from stored procedure.  
    oCmd.Parameters.Refresh  

    ' Set the param values.  
    oCmd.Parameters(1).Value = mUserID
    oCmd.Parameters(2).Value = "HTML_OptionList"	'--- Function
	oCmd.Parameters(3).Value = sFilterCompany

	Set rs = Server.CreateObject("ADODB.Recordset")
    ' Setting the cursor location to client side is important to get a disconnected recordset.
    rs.CursorLocation = 3 'adUseClient
	rs.CursorType = 3 'adOpenStatic
	rs.LockType = 1 'adLockReadOnly
	Set rs.Source = oCmd    'this is the key
	
	
    ' Execute...  
    'rs.Open sSQL, sAdminConn, 3, 4
	rs.Open
	
	' Disconnect the recordset.
	Set rs.ActiveConnection = Nothing	

	
	If not rs.EOF Then 
		sCompanyHTMLOptionList = rs(0)	'first field of first row
	End If
	
	rs.Close
	
	exit function
End Function


'=========================================================================================
' Claim List
'=========================================================================================

  Public Function rsClaimList(sOfficeID, sOwner, sClaimNum, sStatus, sRcvDt, sSort, sType, sFilterCompany, sFilterOffice)
	Dim sConn, rs, sSQL, sWhere, sOrderBy
    Dim TopN 

  	TopN = "TOP 300"
	
	'=========
    Dim oCn, oCmd, oPrm
	
    Set oCn = Server.CreateObject("ADODB.Connection")
    oCn.Open mConnStr 'sAdminConn2 '
    Set oCmd = Server.CreateObject("ADODB.Command")
    oCmd.CommandText = "Portal_ClaimsGridList"
    oCmd.CommandType = 4 '4
    Set oCmd.ActiveConnection = oCn

    ' Automatically fill in parameter info from stored procedure.  
    oCmd.Parameters.Refresh  

    ' Set the param values.  
    oCmd.Parameters(1).Value = mUserID
    oCmd.Parameters(2).Value = sOwner
    oCmd.Parameters(3).Value = sClaimNum
    oCmd.Parameters(4).Value = sStatus
    oCmd.Parameters(5).Value = sRcvDt
    oCmd.Parameters(6).Value = sSort
	oCmd.Parameters(7).Value = sType
	oCmd.Parameters(8).Value = sFilterCompany
	oCmd.Parameters(9).Value = sFilterOffice

	Set rs = Server.CreateObject("ADODB.Recordset")
    ' Setting the cursor location to client side is important to get a disconnected recordset.
    rs.CursorLocation = 3 'adUseClient
	rs.CursorType = 3 'adOpenStatic
	rs.LockType = 1 'adLockReadOnly
	Set rs.Source = oCmd    'this is the key
	
	
    ' Execute...  
    'rs.Open sSQL, sAdminConn, 3, 4
	rs.Open
	
	
	' Disconnect the recordset.
	Set rs.ActiveConnection = Nothing	

    Set rsClaimList = rs

	
	exit function
	
	'------------------------------------------------------------------------------------------
	' BELOW is NO LONGER USED

	Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select 1 Where 1=0", sAdminConn, 3, 4
    Set rsClaimList = rs
	exit function

    
	'=========

	'    Dim oCn, oCmd, oPrm
'    
'    Set oCn = Server.CreateObject("ADODB.Connection")
'    oCn.Open  mConnStr
'    Set oCmd = Server.CreateObject("ADODB.Command")
'    oCmd.CommandText = "PE_Portal_Claim_List"
'    Set oCmd.ActiveConnection = oCn
'    oCmd.CommandType = 4 '4
'    'Set oPrm = oCmd.CreateParameter("@OfficeID", 130, 1, 5)
'    'cmdQuery.Parameters.Append prm
'    oCmd.Parameters.Append oCmd.CreateParameter("@OfficeID", adVarChar, adParamInput, 10, sOfficeID)
'    oCmd.Parameters.Append oCmd.CreateParameter("@Owner", adVarChar, adParamInput, 30, sOwner)
'    oCmd.Parameters.Append oCmd.CreateParameter("@ClaimNum", adVarChar, adParamInput, 30, sClaimNum)
'    oCmd.Parameters.Append oCmd.CreateParameter("@Status", adVarChar, adParamInput, 30, sStatus)
'    
'    Set rs = oCmd.Execute
    
    
		      

'    "SELECT " & TopN & " A.ClaimNo, A.EstimateIDPK, A.ClmLastName, " & _
'            "CASE WHEN (ISNULL(E.Approved,'1/1/90') > ISNULL(E.LastProcessed,'1/1/90')) THEN 'C' ELSE 'P' END AS RStatus, " & _
'            "A.ReceivedDate, A.InsuranceCompanyID " & _
'    "FROM   tblAssignments A INNER JOIN " & _
'           "tblInsuranceCompanies I ON A.InsuranceCompanyID = I.InsuranceCompanyIDPK LEFT OUTER JOIN " & _
'           "tblAssignmentsEx E ON A.EstimateIDPK = E.EstimateIDPK " & _
'    "WHERE  (A.InsuranceCompanyID = '" & sOfficeID & "') "

    '"WHERE (I.InsuranceCompanyName = '" & sCo & "') " 
    
	sSQL = _	  
      "SELECT " & TopN & " " & _
       "A.EstimateIDPK, A.ClaimNo, A.PolicyNo, " & _
       "ReceivedDate = Convert(varchar, A.ReceivedDate, 101), " & _
       "LossDate = Convert(varchar, A.LossDate, 101), ClosedDate = Convert(varchar, A.ClosedDate, 101),  " & _
       "Status = CASE WHEN A.ClmLastName Like 'Cancel%' THEN 'Cancelled' WHEN (ISNULL(E.Approved,'1/1/90') > ISNULL(E.LastProcessed,'1/1/90')) THEN 'Closed' ELSE 'Pending' END, " & _
       "A.ClmLastName, A.ClmFirstName, ISNULL(A.ClmLastName, N'') + N', ' + ISNULL(A.ClmFirstName, N'') AS Owner,  " & _
       "ISNULL(A.InsLastName, N'') + N', ' + ISNULL(A.InsFirstName, N'') AS Insured,  " & _
       "BodyShopEstAmt = Case When A.BodyShopEstAmt = 0 Then NULL Else Convert(varchar, A.BodyShopEstAmt, 1) End,  " & _
       "AgreedPrice = Case When E.FileStatus IN(0, 2) Then NULL Else Convert(varchar, A.AgreedPrice, 1) End, " & _
       "E.FileStatus, E.LastProcessed, " & _
       "HasPhotos = Case IsNull(p.EstimateID,0) When 0 Then 0 Else 1 End " & _
      "FROM   tblAssignments A " & _
       "INNER JOIN tblAssignmentsEx E ON A.EstimateIDPK = E.EstimateIDPK " & _
       "INNER JOIN tblInsuranceCompanies I ON A.InsuranceCompanyID = I.InsuranceCompanyIDPK " & _
       "INNER JOIN tblClaimReps R ON R.ClaimRepIDPK = A.ClaimRepID " & _
       "Left Outer JOIN (Select EstimateID From tblPhotos Group By EstimateID) p ON p.EstimateID = A.EstimateIDPK " & _
      "WHERE  (A.ReceivedDate > (Getdate() - " & sRcvDt & ")) "
      
      '--- If co manager (or an ACE sign in) then filter just for the Co name otherwise filter for the office ---
      If mUserTypeID = cMemberType_CoMgr OR (mUserSecurityL >= 5) Then 
        sSQL = sSQL & _
          "AND (I.InsuranceCompanyName = (SELECT InsuranceCompanyName FROM tblInsuranceCompanies WHERE InsuranceCompanyIDPK = N'" & sOfficeID & "')) "  
      Else
        sSQL = sSQL & _
          "AND (A.InsuranceCompanyID = N'" & sOfficeID & "') "        
      End If
      
'      IF mUserTypeID = cMemberType_CoOfficeMgr Then
'      End If
      
      If mUserTypeID = cMemberType_ClaimRep Then
        sSQL = sSQL & _
          "AND (R.ClaimRepName = N'" & mRepName & "') "  
      End IF
      
'      "WHERE  ((A.LossDate > (Getdate() - " & sLossDt & ")) OR A.LossDate IS NULL) AND (A.InsuranceCompanyID = N'" & sOfficeID & "') "   'CONVERT(DATETIME, '2005-06-01 00:00:00', 102)


    If instr(1, sClaimNum, "--") = 0 _
       And instr(1, sClaimNum, "*") = 0 _
       AND instr(1, sClaimNum, "drop") = 0 _
       AND instr(1, sClaimNum, ";") = 0 Then
      If Len(sClaimNum & "") > 0 Then
        sSQL = sSQL & _
          "AND (A.ClaimNo LIKE N'" & sClaimNum & "%') "
      End If
    End If

    If instr(1, sOwner, "--") = 0 _
       And instr(1, sOwner, "*") = 0 _
       AND instr(1, sOwner, "drop") = 0 _
       AND instr(1, sOwner, ";") = 0 Then
      If Len(sOwner & "") > 0 Then
        sSQL = sSQL & _
          "AND (A.ClmLastName LIKE N'" & sOwner & "%') "
      End If
    End If
	  
	If (sStatus & "" <> "0") And (sStatus & "" <> "") Then 
	  Select Case sStatus & ""
	    Case "1" 'Closed
          sSQL = sSQL & _
            "AND (ISNULL(E.Approved,'1/1/90') > ISNULL(E.LastProcessed,'1/1/90')) AND (A.ClmLastName NOT LIKE N'Cancel%') "
            '"AND (A.ClosedDate IS NOT NULL) "
	    Case "2" 'Pending
          sSQL = sSQL & _
            "AND (ISNULL(E.Approved,'1/1/90') < ISNULL(E.LastProcessed,'1/1/2050')) "          
            '"AND (A.ClosedDate IS NULL) "
	    Case "3" 'Cancelled
          sSQL = sSQL & _
            "AND (A.ClmLastName LIKE N'Cancel%') "
	  End Select
	End If
	
	'If Len(sWhere) > 0 Then sWhere = " Where " & Mid(sWhere, 4)


    sOrderBy = "ORDER BY Cast(A.LossDate AS datetime) DESC" '0 case
    
	If (sSort & "" <> "0") And (sSort & "" <> "") Then 
	  Select Case sSort & ""
	    Case "1" 'Received date
          sOrderBy = "ORDER BY Cast(A.ReceivedDate AS datetime) DESC "  '--- The CAST is needed WHY ?? I dunno.
	    Case "2" 'Claim #
          sOrderBy = "ORDER BY A.ClaimNo " 
	    Case "3" 
          sOrderBy = "ORDER BY Cast(A.ClosedDate AS datetime) DESC " 
	  End Select
	End If
    

	sSQL = sSQL & _
	  " " & sWhere & " " & sOrderBy
	  
	
'      "ORDER BY A.ReceivedDate DESC"            

    Set rs = Server.CreateObject("ADODB.Recordset")

    ' Setting the cursor location to client side is important to get a disconnected recordset.
    rs.CursorLocation = 3 'adUseClient

    rs.Open sSQL, sAdminConn, 3, 4

    ' Disconnect the recordset.
    Set rs.ActiveConnection = Nothing

    Set rsClaimList = rs	                
  End Function
'=========================================================================================
' End Claim List
'=========================================================================================



'=========================================================================================
' Remote data Admin function area
'=========================================================================================

'-----------------------------------------------------------------------------------------
  Public Sub StreamRecordsetToResponse(rs)
    rs.Save Response, adPersistXML   ' saves recordset to Response object in XML format
  End Sub

'-----------------------------------------------------------------------------------------
  Public Function RSGet(sSQL)
    Dim rs

    Set rs = Server.CreateObject("ADODB.Recordset")

    ' Setting the cursor location to client side is important to get a disconnected recordset.
    rs.CursorLocation = 3 'adUseClient

    rs.Open sSQL, mConnStr, 3, 4

    ' Disconnect the recordset.
    Set rs.ActiveConnection = Nothing

    Set RSGet = rs
  End Function
'-----------------------------------------------------------------------------------------
  Public Function RSBatchUpdate(rs)
    rs.ActiveConnection = mConnStr  '<--- We don't need to create a connection object just set to connection string
    rs.UpdateBatch
    'rs.resync

    '--- Disconnect the recordset. ---
    Set rs.ActiveConnection = Nothing
  End Function

'-----------------------------------------------------------------------------------------
  Public Function SQLExecute(sSQL)
    Dim Conn

    Set Conn = ConnMemberDB
    Conn.execute(sSQL)
    conn.close
    Set conn=Nothing
  End Function
'=========================================================================================
' End Remote data Admin function area
'=========================================================================================


'=========================================================================================
' Admin function area
'=========================================================================================
'-----------------------------------------------------------------------------------------
'  Public Function ACEMemberList(EmailPending, Inactive)
'    Dim Conn, rs, SQL
'
'    SQL = "SELECT tblMembers.MemberIDPK AS ID, [MemberLastName] & "", "" & [MemberFirstName] AS Name, tblInsCompanies.InsCoName AS CoName, tblMembers.MemberLogonName AS LogonName " & _
'          "FROM tblInsCompanies RIGHT JOIN tblMembers ON tblInsCompanies.InsCompanyIDPK = tblMembers.InsCompanyID "
'
'    If EmailPending <> 0 Then SQL = SQL & "WHERE (((tblMembers.MemberConfEmailRespRecvd)=0)) "
'    If Inactive <> 0 Then SQL = SQL & "WHERE (((tblMembers.MemberAccountEnabled)=0)) "
'
'    SQL = SQL & "ORDER BY [MemberLastName] & "", "" & [MemberFirstName]"
'
'    SQL = SQL & ";"
'
'    Set rs = Server.CreateObject("ADODB.Recordset")
'
'    ' Setting the cursor location to client side is important to get a disconnected recordset.
'    rs.CursorLocation = 3 'adUseClient
'
'    'Set Conn = ConnMemberDB
'    rs.Open SQL, mConnStr, 3, 4
'
'    ' Disconnect the recordset.
'    Set Rs.ActiveConnection = Nothing
'
'    'Conn.Close
'    'Set Conn = Nothing
'
'    Set ACEMemberList = rs
'
'  End Function


'-----------------------------------------------------------------------------------------
  Public Function ACEGetMemberTables()
    Dim Conn, objSchema, strData, strTblName

    Set Conn = ConnMemberDB

    Set objSchema = Conn.OpenSchema(adSchemaPrimaryKeys)
    Do Until objSchema.EOF
      strTblName = objSchema("TABLE_NAME")
      If Not Instr(1,strTblName,"MSys") Then
        strData = strData & "<a href=""ShowTable.asp?db=" & Server.URLEncode(mMemberDBName) & "&table=" & Server.URLEncode(strTableName) & """>" & strTblname & "-" & objSchema("Column_Name") & "</a><br>"
      End If
      objSchema.MoveNext
    Loop

    objSchema.close
    Set objSchema = Nothing
    conn.close
    Set conn = Nothing

    ACEGetMemberTables = strData 
  End Function


'-----------------------------------------------------------------------------------------
  Public Function ACEGetMemberTableData(TableName)
    Dim Conn, objSchema, strData, strTblName
  End Function

'-----------------------------------------------------------------------------------------
  Private Function ZN(param)
    If param="" Then
      ZN = NULL
    Else
      ZN = param
    End If
  End Function

End Class

%>
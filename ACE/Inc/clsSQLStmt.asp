<%

  Const List_Contact_Address_Types = 1
  Const List_Contact_Phone_Types = 2
  Const List_Contact_Record_Types = 3 'Company, Office, Contact
  Const List_Contact_Categories = 4 'Client, Prospect, Supplier

'=========================================================================================
' Start Class definition
'=========================================================================================

Class clsSQLStmt

'-----------------------------------------------------------------------------------------
  Public Function SQL_Lists(ListID, UnionAll)
    SQL_Lists = _
      "SELECT LookupIDPK, LookupDesc " & _
      "FROM tblLookups " & _
      "WHERE (LookupCategoryID = " & ListID & ") "
      
    If UnionAll Then
      SQL_Lists = SQL_Lists & _
        "UNION ALL " & _
        "SELECT 0, '" & Chr(1) & " ALL " & Chr(1) & "' "
    End If
    
    SQL_Lists = SQL_Lists & _
      "ORDER BY LookupDesc "
      
  End Function
  

'-----------------------------------------------------------------------------------------
  Public Function SQL_ClaimList(sDateStart, sDateEnd, StatusID, sOwner, sFind, sPage)
    '--- Create SQL to pull all claims meeting criteria
    '--- Filter by date range if not empty "" 
    '--- Filter by status if they are not 0
    '--- If sFind or sOwner is not empty perform the find for the given status selections
    
    Dim sSQL 
    
    sSQL = _
        ""
            
    If Len(sFind) > 0 Then      
      sSQL = sSQL & _
        "AND (" & _
        "(AccountNum LIKE '%" & sFind & "%') OR " & _
        "(Name LIKE '%" & sFind & "%') OR " & _
        "(City LIKE '%" & sFind & "%') OR " & _
        "(Addr1 LIKE '%" & sFind & "%') OR " & _
        "(State LIKE '%" & sFind & "%') OR " & _
        "(Zip LIKE '%" & sFind & "%') OR " & _
        "(ReferredBy LIKE '%" & sFind & "%') OR " & _
        "(Email LIKE '%" & sFind & "%') OR " & _
        "(Phone LIKE '%" & sFind & "%') OR " & _
        "(PhoneFax LIKE '%" & sFind & "%') OR " & _
        "(PhoneMobile LIKE '%" & sFind & "%') OR " & _
        "(PhoneHome LIKE '%" & sFind & "%') OR " & _
        "(PrimaryContact LIKE '%" & sFind & "%') " & _
        ") "
        
    ElseIf (Len(sPage) > 0) AND (sPage <> "ALL") Then
       sSQL = sSQL & _
        "AND (" & _
        "(Name LIKE '" & sPage & "%')" & _
        ") "
    End If
    
    sSQL = sSQL & _
      "ORDER BY Name "
        
    SQL_ContactParentFindList = _
      sSQL
   ' Response.Write sSQL
  End Function
      
  
End Class
'=========================================================================================
' END Class definition
'=========================================================================================

%>
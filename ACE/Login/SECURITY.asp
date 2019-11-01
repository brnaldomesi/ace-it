
<%

  '--- Security Include File



%>
<%
  

  Dim mUserID, mUserName, mUserTypeID, mUserSecurityL
  Dim mName, mRepName, mUserInsCoName, mUserACEInsCoOfficeID, mUserACEInsCoOfficeName
  Dim mUserGender
  Dim mEncPass
  Dim mAllowShopAccess
  Dim mShopDemoAccess
  
  Dim mUserInsCoOfficeID_AutoForm
  Dim mUserInsCoOfficeID_SubrogationForm
  Dim mUserInsCoOfficeID_HeavySpecialtyForm
  Dim mUserInsCoOfficeID_PropertyForm
  Dim mUserInsCoOfficeID_AutoTotalLossForm
  Dim mUserInsCoOfficeID_PhotoQuoteForm
  
  'On Error Resume Next

  mUserInsCoOfficeID_AutoForm = Session("UserInsCoOfficeID_AutoForm") & ""
  mUserInsCoOfficeID_SubrogationForm = Session("UserInsCoOfficeID_SubrogationForm") & ""
  mUserInsCoOfficeID_HeavySpecialtyForm = Session("UserInsCoOfficeID_HeavySpecialtyForm") & ""
  mUserInsCoOfficeID_PropertyForm = Session("UserInsCoOfficeID_PropertyForm") & ""
  mUserInsCoOfficeID_AutoTotalLossForm = Session("UserInsCoOfficeID_AutoTotalLossForm") & ""
  mUserInsCoOfficeID_PhotoQuoteForm = Session("UserInsCoOfficeID_PhotoQuoteForm") & ""
  
  If mUserInsCoOfficeID_AutoForm = "" Then mUserInsCoOfficeID_AutoForm = Session("ACEInsCoOfficeID")
  If mUserInsCoOfficeID_SubrogationForm = "" Then mUserInsCoOfficeID_SubrogationForm = Session("ACEInsCoOfficeID")
  If mUserInsCoOfficeID_HeavySpecialtyForm = "" Then mUserInsCoOfficeID_HeavySpecialtyForm = Session("ACEInsCoOfficeID")
  If mUserInsCoOfficeID_PropertyForm = "" Then mUserInsCoOfficeID_PropertyForm = Session("ACEInsCoOfficeID")
  If mUserInsCoOfficeID_AutoTotalLossForm = "" Then mUserInsCoOfficeID_AutoTotalLossForm = Session("ACEInsCoOfficeID")
  If mUserInsCoOfficeID_PhotoQuoteForm = "" Then mUserInsCoOfficeID_PhotoQuoteForm = Session("ACEInsCoOfficeID")

  'On Error Goto 0

  Dim Temp
  
  Const cMemberType_ClaimRep = 1
  Const cMemberType_Agent = 2
  Const cMemberType_Administrator = 3
  Const cMemberType_Temp = 4
  Const cMemberType_ACEEmployee = 5
  Const cMemberType_ACEAuditor = 6
  Const cMemberType_ACEClerical = 7
  Const cMemberType_ACEAdmin = 8
  Const cMemberType_CoOfficeMgr = 9
  Const cMemberType_CoMgr = 10
  	
  mUserID = 0
  On Error Resume Next
  mUserID=Session("MemberID")  '-- this is not the login name but unique id
  mUserName=Session("MemberUserName")
  mUserTypeID=Session("MemberTypeID")
  mUserSecurityL=Session("MemberSecurityLevel")
  mName=Session("MemberName")
  mRepName=Session("RepName")
  mUserInsCoName=Session("InsCompanyName")
  mUserACEInsCoOfficeID=Session("ACEInsCoOfficeID")
  mUserACEInsCoOfficeName=Session("ACEInsCoOfficeName")
  mUserGender=Session("MemberGender")
  mAllowShopAccess = Session("ShopNetworkAccess")
  mShopDemoAccess = Session("ShopDemoAccess")
  
  'mSecToken = 

  On Error Goto 0


  If mUserID = 0 Then
    'Session("RequestedURL") = Request.ServerVariables("HTTP_REFERER")
	Session("RequestedURL") = "http://" & _
	    Request.ServerVariables("SERVER_NAME") & _
        Request.ServerVariables("SCRIPT_NAME")

	Temp = Request.ServerVariables("QUERY_STRING")
	If Temp & "" <> "" Then
      Session("RequestedURL") = Session("RequestedURL") & "?" & Temp
	End If
    response.redirect "https://www.staging-portal.ace-it.com/ace/login/ACELogin.asp?txtState=0" 'Server.MapPath("\Login\ACELogin.asp")
    response.end
  Else
    'PostActivity
  End If



'---------------------------------
Function ObfiscateEstimateID(ByVal sEstimateID)
	Dim strHex, i, c, sEnc
	
	sEnc = Right("0" & Len(sEstimateID), 2) & sEstimateID & "iul454benfb" & sEstimateID

	sEnc = EnDeCrypt(sEnc, "eikdn744fdcnv@#")
	
	For i= 1 To Len(sEnc)
		strHex = strHex & Right("0" & Hex(Asc(Mid(sEnc,i,1))), 2)
	Next

	ObfiscateEstimateID = strHex
End Function

Function DeObfiscateEstimateID(ByVal sEstimateID)
	Dim str, i 

	For i= 1 To Len(sEstimateID) Step 2
		str = str & Chr("&H" & Mid(sEstimateID,i,2))
	Next

	str = EnDeCrypt(str, "eikdn744fdcnv@#") ' get decrypted string
	i = mid(str, 3, cInt(Left(str, 2))) ' parse off EstID
	str = Mid(str, 3)	'get everything but len of EstID
	If str <> i & "iul454benfb" & i Then i = "0" '--- see if string has been tampered
	DeObfiscateEstimateID = i 'mid(str, 6, len(str) - 18)
End Function

'---------------------------------
Private Function EncryptText(ByVal strEncryptionKey, ByVal strTextToEncrypt)
 Dim outer, inner, Key, strTemp

 For outer = 1 To Len(strEncryptionKey)

  key = Asc(Mid(strEncryptionKey, outer, 1))

  For inner = 1 To Len(strTextToEncrypt)
   strTemp = strTemp & Chr(Asc(Mid(strTextToEncrypt, inner, 1)) Xor key)
   key = (key + Len(strEncryptionKey)) Mod 256
  Next

  strTextToEncrypt = strTemp
  strTemp = ""
 
 Next

 EncryptText = strTextToEncrypt
End Function



'---------------------------------

Function EncStr(ByVal sIn, bURLEnc)
  Dim sURLParts
  Dim sOut
  Dim sTemp
  Dim i 

  If Right(sIn, 1) <> "/" Then sIn = sIn & "/"
  sURLParts = Split(sIn, "/")
  For i = 0 To Ubound(sURLParts)
    sTemp = EncString(sURLParts(i), 0)
    If bURLEnc = 1 Then sTemp = sURLEncode(sTemp)
    If bURLEnc = 2 Then sTemp = sURLEncode2(sTemp)
    sOut = sOut & "/" & sTemp
  Next
  EncStr = mid(sOut, 2) 
  
End Function

'---------------------------------
Function EncString(ByVal sIn, Decode)
  ' IE7 can't handle %00 in the URL so we have to convert all 0's to two byte sequence FFFE and FF to FFFF
  Dim i
  Dim sPassword
  Dim lPassWordIndex
  Dim sOut
  Dim sChar
  
  If Len(sIn) > 0 Then
    sPassword = "ZACE-OIU#YRHjK"    ' reset the password to the array"
    
    lPassWordIndex = 1
    For i = 1 To Len(sIn)
      If Decode <> 0 Then
        If Mid(sIn, i, 1) = Chr(255) Then
          i = i + 1
          If Mid(sIn, i, 1) = Chr(254) Then 
          	Mid(sIn, i, 1) = "?"
          ElseIf Mid(sIn, i, 1) <> Chr(255) Then 
            Mid(sIn, i, 1) = Chr(Asc(Mid(sIn, i, 1)) - 65)
          End If
        End If
      End If
      
      sChar = Chr(Asc(Mid(sIn, i, 1)) Xor Asc(Mid(sPassword, lPassWordIndex, 1)))
      
      If Decode = 0 Then
        If Asc(sChar) < 35 Then
          sChar = Chr(255) & Chr(Asc(sChar) + 65)
        ElseIf sChar = "?" Then
          sChar = Chr(255) & Chr(254)
        ElseIf sChar = Chr(255) Then
          sChar = Chr(255) & Chr(255)
        End If
      End If
      
      sOut = sOut & sChar
      
      If lPassWordIndex < Len(sPassword) Then
        sPassword = Left(sPassword, lPassWordIndex - 1) & _
              Chr((Asc(Mid(sPassword, lPassWordIndex, 1)) + Asc(Mid(sPassword, lPassWordIndex + 1, 1))) And &HFF) & _
              Right(sPassWord, Len(sPassWord) - lPassWordIndex)
        lPassWordIndex = lPassWordIndex + 1
      Else
        sPassword = Left(sPassword, lPassWordIndex - 1) & _
              Chr((Asc(Mid(sPassword, lPassWordIndex, 1)) + Asc(Mid(sPassword, 1, 1))) And &HFF) & _
              Right(sPassWord, Len(sPassWord) - lPassWordIndex)
        lPassWordIndex = 1
      End If      
      
    Next
    
    EncString = sOut
  End If
  
End Function

'---------------------------------
Public Function sURLEncode(sRawURL)
  Dim i
  Dim sRtn 
  Dim sTmp 
  Dim sValidChars 
  
  sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:.?=_-$(){}~&"

  If Len(sRawURL) > 0 Then
      For i = 1 To Len(sRawURL)
          sTmp = Mid(sRawURL, i, 1)

          If InStr(1, sValidChars, sTmp, 0) = 0 Then 'vbBinaryCompare
              sTmp = Hex(Asc(sTmp)): If Len(sTmp) < 2 Then sTmp = "0" & sTmp
              sTmp = "%" & sTmp
              sTmp = Replace(sTmp, "%20", "+")
          End If
          sRtn = sRtn & sTmp
      Next 'iLoop
      sURLEncode = sRtn
  End If
End Function

'---------------------------------
Public Function sURLEncode2(sRawURL)
  Dim i
  Dim sRtn 
  Dim sTmp 
  Dim sValidChars 
  
  sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:.?=_-$(){}~&"

  If Len(sRawURL) > 0 Then
      ' Loop through each char
      For i = 1 To Len(sRawURL)
          sTmp = Mid(sRawURL, i, 1)

          If InStr(1, sValidChars, sTmp, 0) = 0 Then 'vbBinaryCompare
              ' If not ValidChar, convert to HEX and prefix with %
              sTmp = Hex(Asc(sTmp)): If Len(sTmp) < 2 Then sTmp = "0" & sTmp
              sTmp = "%25" & sTmp       '----------------- HERE WE PREFIX %25 (hex 25 = %), needed for <a tag href but not other places 
              ' If sTmp is %20 then replace with +
              sTmp = Replace(sTmp, "%20", "+")
          End If
          sRtn = sRtn & sTmp
      Next 'iLoop
      sURLEncode2 = sRtn
  End If
End Function


Public Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If
	
    sOutput = REPLACE(sConvert, "+", " ")
    aSplit = Split(sOutput, "%")
	
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & Chr("&H" & Left(aSplit(i + 1), 2)) & _
                     Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
	
    URLDecode = sOutput
End Function



   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::                                                             :::
   ':::  This script performs 'RC4' Stream Encryption               :::
   ':::  (Based on what is widely thought to be RSA's RC4           :::
   ':::  algorithm. It produces output streams that are identical   :::
   ':::  to the commercial products)                                :::
   ':::                                                             :::
   ':::  This script is Copyright © 1999 by Mike Shaffer            :::
   ':::  ALL RIGHTS RESERVED WORLDWIDE                              :::
   ':::                                                             :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::                                                             :::
   ':::  This routine passes the standard test vectors for the      :::
   ':::  RC4 algorithm. The test vectors are included below for     :::
   ':::  easy cut-and-paste verification. It is recommended that    :::
   ':::  you remove these comments for actual production to         :::
   ':::  reduce initial script parsing/processing time.             :::
   ':::                                                             :::
   ':::                                                             :::
   '::: Test harness vectors:                                       :::
   ':::                                                             :::
   ':::   Test vector 0                                             :::
   ':::   Key: &h01 &h23 &h45 &h67 &h89 &hab &hcd &hef              :::
   ':::   Input: &h01 &h23 &h45 &h67 &h89 &hab &hcd &hef            :::
   ':::   Output: &h75 &hb7 &h87 &h80 &h99 &he0 &hc5 &h96           :::
   ':::                                                             :::
   ':::   Test vector 1                                             :::
   ':::   Key: &h01 &h23 &h45 &h67 &h89 &hab &hcd &hef              :::
   ':::   Input: &h00 &h00 &h00 &h00 &h00 &h00 &h00 &h00            :::
   ':::   Output: &h74 &h94 &hc2 &he7 &h10 &h4b &h08 &h79           :::
   ':::                                                             :::
   ':::   Test vector 2                                             :::
   ':::   Key: &h00 &h00 &h00 &h00 &h00 &h00 &h00 &h00              :::
   ':::   Input: &h00 &h00 &h00 &h00 &h00 &h00 &h00 &h00            :::
   ':::   Output: &hde &h18 &h89 &h41 &ha3 &h37 &h5d &h3a           :::
   ':::                                                             :::
   ':::   Test vector 3                                             :::
   ':::   Key: &hef &h01 &h23 &h45                                  ::: 
   ':::   Input: &h00 &h00 &h00 &h00 &h00 &h00 &h00 &h00 &h00 &h00  ::: 
   ':::   Output: &hd6 &ha1 &h41 &ha7 &hec &h3c &h38 &hdf &hbd &h61 :::
   ':::                                                             :::
   ':::   Test vector 4                                             :::
   ':::   Key: &h01 &h23 &h45 &h67 &h89 &hab &hcd &hef              :::
   ':::   Input:                                                    :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01    :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01 &h01         :::
   ':::   &h01                                                      :::
   ':::   Output:                                                   :::
   ':::   &h75 &h95 &hc3 &he6 &h11 &h4a &h09 &h78 &h0c &h4a &hd4    :::
   ':::   &h52 &h33 &h8e &h1f &hfd &h9a &h1b &he9 &h49 &h8f         :::
   ':::   &h81 &h3d &h76 &h53 &h34 &h49 &hb6 &h77 &h8d &hca         :::
   ':::   &hd8 &hc7 &h8a &h8d &h2b &ha9 &hac &h66 &h08 &h5d         :::
   ':::   &h0e &h53 &hd5 &h9c &h26 &hc2 &hd1 &hc4 &h90 &hc1         :::
   ':::   &heb &hbe &h0c &he6 &h6d &h1b &h6b &h1b &h13 &hb6         :::
   ':::   &hb9 &h19 &hb8 &h47 &hc2 &h5a &h91 &h44 &h7a &h95         :::
   ':::   &he7 &h5e &h4e &hf1 &h67 &h79 &hcd &he8 &hbf &h0a         :::
   ':::   &h95 &h85 &h0e &h32 &haf &h96 &h89 &h44 &h4f &hd3         :::
   ':::   &h77 &h10 &h8f &h98 &hfd &hcb &hd4 &he7 &h26 &h56         :::
   ':::   &h75 &h00 &h99 &h0b &hcc &h7e &h0c &ha3 &hc4 &haa         :::
   ':::   &ha3 &h04 &ha3 &h87 &hd2 &h0f &h3b &h8f &hbb &hcd         :::
   ':::   &h42 &ha1 &hbd &h31 &h1d &h7a &h43 &h03 &hdd &ha5         :::
   ':::   &hab &h07 &h88 &h96 &hae &h80 &hc1 &h8b &h0a &hf6         :::
   ':::   &h6d &hff &h31 &h96 &h16 &heb &h78 &h4e &h49 &h5a         :::
   ':::   &hd2 &hce &h90 &hd7 &hf7 &h72 &ha8 &h17 &h47 &hb6         :::
   ':::   &h5f &h62 &h09 &h3b &h1e &h0d &hb9 &he5 &hba &h53         :::
   ':::   &h2f &haf &hec &h47 &h50 &h83 &h23 &he6 &h71 &h32         :::
   ':::   &h7d &hf9 &h44 &h44 &h32 &hcb &h73 &h67 &hce &hc8         :::
   ':::   &h2f &h5d &h44 &hc0 &hd0 &h0b &h67 &hd6 &h50 &ha0         :::
   ':::   &h75 &hcd &h4b &h70 &hde &hdd &h77 &heb &h9b &h10         :::
   ':::   &h23 &h1b &h6b &h5b &h74 &h13 &h47 &h39 &h6d &h62         :::
   ':::   &h89 &h74 &h21 &hd4 &h3d &hf9 &hb4 &h2e &h44 &h6e         :::
   ':::   &h35 &h8e &h9c &h11 &ha9 &hb2 &h18 &h4e &hcb &hef         :::
   ':::   &h0c &hd8 &he7 &ha8 &h77 &hef &h96 &h8f &h13 &h90         :::
   ':::   &hec &h9b &h3d &h35 &ha5 &h58 &h5c &hb0 &h09 &h29         :::
   ':::   &h0e &h2f &hcd &he7 &hb5 &hec &h66 &hd9 &h08 &h4b         :::
   ':::   &he4 &h40 &h55 &ha6 &h19 &hd9 &hdd &h7f &hc3 &h16         :::
   ':::   &h6f &h94 &h87 &hf7 &hcb &h27 &h29 &h12 &h42 &h64         :::
   ':::   &h45 &h99 &h85 &h14 &hc1 &h5d &h53 &ha1 &h8c &h86         :::
   ':::   &h4c &he3 &ha2 &hb7 &h55 &h57 &h93 &h98 &h81 &h26         :::
   ':::   &h52 &h0e &hac &hf2 &he3 &h06 &h6e &h23 &h0c &h91         :::
   ':::   &hbe &he4 &hdd &h53 &h04 &hf5 &hfd &h04 &h05 &hb3         :::
   ':::   &h5b &hd9 &h9c &h73 &h13 &h5d &h3d &h9b &hc3 &h35         :::
   ':::   &hee &h04 &h9e &hf6 &h9b &h38 &h67 &hbf &h2d &h7b         :::
   ':::   &hd1 &hea &ha5 &h95 &hd8 &hbf &hc0 &h06 &h6f &hf8         :::
   ':::   &hd3 &h15 &h09 &heb &h0c &h6c &haa &h00 &h6c &h80         :::
   ':::   &h7a &h62 &h3e &hf8 &h4c &h3d &h33 &hc1 &h95 &hd2         :::
   ':::   &h3e &he3 &h20 &hc4 &h0d &he0 &h55 &h81 &h57 &hc8         :::
   ':::   &h22 &hd4 &hb8 &hc5 &h69 &hd8 &h49 &hae &hd5 &h9d         :::
   ':::   &h4e &h0f &hd7 &hf3 &h79 &h58 &h6b &h4b &h7f &hf6         :::
   ':::   &h84 &hed &h6a &h18 &h9f &h74 &h86 &hd4 &h9b &h9c         :::
   ':::   &h4b &had &h9b &ha2 &h4b &h96 &hab &hf9 &h24 &h37         :::
   ':::   &h2c &h8a &h8f &hff &hb1 &h0d &h55 &h35 &h49 &h00         :::
   ':::   &ha7 &h7a &h3d &hb5 &hf2 &h05 &he1 &hb9 &h9f &hcd         :::
   ':::   &h86 &h60 &h86 &h3a &h15 &h9a &hd4 &hab &he4 &h0f         :::
   ':::   &ha4 &h89 &h34 &h16 &h3d &hdd &he5 &h42 &ha6 &h58         :::
   ':::   &h55 &h40 &hfd &h68 &h3c &hbf &hd8 &hc0 &h0f &h12         :::
   ':::   &h12 &h9a &h28 &h4d &hea &hcc &h4c &hde &hfe &h58         :::
   ':::   &hbe &h71 &h37 &h54 &h1c &h04 &h71 &h26 &hc8 &hd4         :::
   ':::   &h9e &h27 &h55 &hab &h18 &h1a &hb7 &he9 &h40 &hb0         :::
   ':::   &hc0                                                      :::
   ':::                                                             :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


   Dim sbox(255)
   Dim key(255)


   Sub RC4Initialize(strPwd)
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::  This routine called by EnDeCrypt function. Initializes the :::
   ':::  sbox and the key array)                                    :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

      dim tempSwap
      dim a
      dim b
	  dim intLength

      intLength = len(strPwd)
      For a = 0 To 255
         key(a) = asc(mid(strpwd, (a mod intLength)+1, 1))
         sbox(a) = a
      next

      b = 0
      For a = 0 To 255
         b = (b + sbox(a) + key(a)) Mod 256
         tempSwap = sbox(a)
         sbox(a) = sbox(b)
         sbox(b) = tempSwap
      Next
   
   End Sub
   
   Function EnDeCrypt(plaintxt, psw)
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   ':::  This routine does all the work. Call it both to ENcrypt    :::
   ':::  and to DEcrypt your data.                                  :::
   ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

      dim temp
      dim a
      dim i
      dim j
      dim k
      dim cipherby
      dim cipher

      i = 0
      j = 0

      RC4Initialize psw

      For a = 1 To Len(plaintxt)
         i = (i + 1) Mod 256
         j = (j + sbox(i)) Mod 256
         temp = sbox(i)
         sbox(i) = sbox(j)
         sbox(j) = temp
   
         k = sbox((sbox(i) + sbox(j)) Mod 256)

         cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
         cipher = cipher & Chr(cipherby)
      Next

      EnDeCrypt = cipher

   End Function



%>
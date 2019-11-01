<%

  Class clsFormMailer

  %>
  <!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows Library" -->
  <!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->
  <%

  Private oMailer

  Public sHeader
  Public sFooter
  Public sTo, sFrom, sCC, sBCC, sSubject, sBody, sImportance, sAttachFile

  Private Sub Class_Initialize
    InitCDO
    sImportance = 3
  End Sub

  Private Sub Class_Terminate
    Set oMailer = Nothing
  End Sub


  Public Property Let AttachFile(sFileSpec)
    oMailer.AddAttachment sFileSpec
  End Property
  
Private Sub InitCDO
  Dim iConf
  Dim Flds

  Const cdoSendUsingPort = 2

  Set oMailer = server.createobject("CDO.Message")
  Set iConf = Server.CreateObject("CDO.Configuration")

  Set Flds = iConf.Fields
  With Flds
    .Item(cdoSendUsingMethod) = cdoSendUsingPort
    .Item(cdoSMTPServer) = "mail-fwd"
    .Item(cdoSMTPServerPort) = 25
    .Item(cdoSMTPconnectiontimeout) = 10
    .Update
  End With

  Set oMailer.Configuration = iConf

  Set iConf = Nothing
  Set Flds = Nothing
End Sub

SUB sendCDOMail()
  Dim objCDO
  
  With oMailer ' objCDO
  
    .From = sFrom
    .To = sTo
    .CC = sCC
    .Subject = sSubject
    .TextBody = sBody
    'If Len(sHTMLBody) <> 0 Then .HTMLBody = sHTMLBody
    

    .Send
    
  End With

  'Cleanup
  Set objCDO = Nothing
  Set oMailer = Nothing
END SUB


  Public Sub BuildBody()
    if Request.Form.Count<>0 then

      Dim strBody
      strBody = strBody & "FORM Data: " & vbCrLf & _
            "FORM submitted at " & Now() & vbCrLf & vbCrLf

      Dim myElement, ix
      For ix = 1 to Request.Form.Count
        myElement = Request.Form.Key(ix)   'name of control
        Select Case Left(myElement,3)
          Case "txt","cmb","grp":
            strBody = strBody & Replace(Mid(myElement,4,len(myElement)),"."," ") &  ": " 
            if Len(Request.Form(myElement)) = 0 then
              strBody = strBody & "--UNANSWERED--"
            else
              strBody = strBody & Request.Form(myElement)
            end if

            strBody = strBody & vbCrLf

          Case "chk":
            strBody = strBody & Replace(Mid(myElement,4,len(myElement)),"."," ") & _
              ": " & Request.Form(myElement) & vbCrLf
        End Select
      Next

      sBody = sBody & strBody
    End If
  End Sub

  '2017-02-12
  Public Function fBuildBody()
	fBuildBody = ""
    if Request.Form.Count<>0 then
      Dim strBody
      strBody = strBody & "FORM Data: " & vbCrLf & _
            "FORM submitted at " & Now() & vbCrLf & vbCrLf

      Dim myElement, ix
      For ix = 1 to Request.Form.Count
        myElement = Request.Form.Key(ix)   'name of control
        Select Case Left(myElement,3)
          Case "txt","cmb","grp":
            strBody = strBody & Replace(Mid(myElement,4,len(myElement)),"."," ") &  ": " 'ditch control prefix "txt" and change periods to spaces
            if Len(Request.Form(myElement)) = 0 then
              strBody = strBody & "--UNANSWERED--"
            else
              strBody = strBody & Request.Form(myElement)
            end if

            strBody = strBody & vbCrLf

          Case "chk":
            strBody = strBody & Replace(Mid(myElement,4,len(myElement)),"."," ") & _
              ": " & Request.Form(myElement) & vbCrLf
        End Select
      Next
      'sBody = sBody & strBody
	  fBuildBody = strBody
    End If
  End Function
  
  
  
  
  Public Sub Send()
      'Time to send the email

      sBody = sHeader & vbCrLf & vbCrLf & sBody & vbCrLf & vbCrLf & sFooter
    SendCDOMail
  End Sub

  Private Sub SendCDONTS()
     Dim objCDO
     Set objCDO = Server.CreateObject("CDONTS.NewMail")

     objCDO.To = sTo
     objCDO.From = sFrom
     objCDO.CC = sCC
     objCDO.BCC = sBCC
     objCDO.Subject = sSubject
     objCDO.Body = sBody
     objCDO.Importance = sImportance

      objCDO.Send
      Set objCDO = Nothing
  End Sub

   Public Function MailerObj()
     Set MailerObj = oMailer
   End Function

  Private Sub SendASPMail()
     oMailer.WordWrap = True
     oMailer.WordWrapLen = 120

     oMailer.AddRecipient "", sTo
     oMailer.FromName = sFrom
     oMailer.FromAddress = sFrom
     'oMailer.RemoteHost = "mail-fwd"
     oMailer.RemoteHost = "mail-fwd.boca15-verio.com" '"mail-fwd.verio-hosting.net" "mail.verio-hosting.net"
     oMailer.Subject = sSubject
     If len(sCC) <> 0 Then oMailer.AddCC "", sCC
     If len(sBCC) <> 0 Then oMailer.AddBCC "", sBCC
     oMailer.BodyText = sBody
     oMailer.Priority = sImportance

     If Len(sAttachFile) >0 Then
       oMailer.AddAttachment sAttachFile
       sAttachedFile = ""
     End If

  if oMailer.SendMail then
  else
     Response.Write ("An error has occurred.<BR>")
     Response.Write ("The error was " & oMailer.Response)
  end if

    Set oMailer = Nothing
  End Sub

  Private Sub SendJMail()
      Dim JMail
      Set JMail = Server.CreateObject("JMail.SMTPMail")

      JMail.ServerAddress = "64.225.184.100" '"208.233.88.4" ' "mail3.puppozone.com"  
     JMail.Sender = sFrom
     JMail.AddRecipient sTo
     JMail.Subject = sSubject
     'JMail.CC = sCC
     'JMail.BCC = sBCC
     JMail.Body = sBody
     JMail.Priority = sImportance  ' 1 - highest priority (Urgent) 3 - normal 5 - lowest

     JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")

      JMail.Execute
      Set JMail = Nothing
  End Sub

  End Class
%>
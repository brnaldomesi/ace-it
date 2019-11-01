<!-- #INCLUDE VIRTUAL="/ACE/Login/Security.asp" -->
<!-- #INCLUDE VIRTUAL="/ACE/inc/clsACEDataConn.asp" -->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>


<%
  If mUserSecurityL < 5 Then
%>
<p align="center"><b>... Sorry you do not have permission to view this page ...</b></p>
<%
    Response.End
  End If
%>



<% 


	Dim strThisPage
	strThisPage = Request.ServerVariables("SCRIPT_NAME")
	strThisPage = Right(strThisPage, Len(strThisPage) - 1)
		
	FILE_FOLDER = Server.Mappath("\Data\ACE")	
%>

<HEAD>
	<TITLE>File Download List For <%= Date() %></TITLE>
	<STYLE TYPE="TEXT/CSS">
	.TabHeader { Font-Family: Arial; Font-Weight: Bold; Font-Size: 12px; Background: Silver }
	.DataCol { Font-Family: Verdana; Font-Size: 12px }
	</STYLE>
	<SCRIPT>
		function msg() {
			self.status = 'File Downloads For <%= Date() %>';
			return true
		}
	</SCRIPT>
</HEAD>


<BODY onLoad="msg()">
<TABLE BORDER=1 ID=tblFileData BACKGROUND="">
	<TR>
		<TD CLASS=TabHeader><A HREF="frmRetrieveFileAttachments.asp?sort=Name">File Name</A></TD>
		<TD CLASS=TabHeader><A HREF="frmRetrieveFileAttachments.asp?sort=Type">File Type</A></TD>
		<TD CLASS=TabHeader align="right"><A HREF="frmRetrieveFileAttachments.asp?sort=Size">File Size</A></TD>
<!--		<TD CLASS=TabHeader><A HREF="<%=strThisPage%>?sort=Path">File Path</A></TD> -->
		<TD CLASS=TabHeader><A HREF="frmRetrieveFileAttachments.asp?sort=Date%20Desc">Last Modified</A></TD>
	</TR>
<%  
	strSortHeader = Request.QueryString("sort")
	
	IF strSortHeader = "" Then
		Call GetAllFiles("Date Desc")
	Else
		Call GetAllFiles(strSortHeader)
	End IF
%>


</TABLE>
</BODY>
</HTML>
<%  

Sub GetAllFiles(strSortBy)
	Dim oFS, oFolder, oFile
	Set oFS = Server.CreateObject("Scripting.FileSystemObject")
		
	'Set Folder Object To Proper File Directory
	Set oFolder = oFS.getFolder(FILE_FOLDER)
	
	Dim intCounter
	
	intCounter = 0
	
	IF strSortBy = "" Then 'UnSorted (default)
		Dim FileArray()
		ReDim Preserve FileArray(oFolder.Files.Count, 5)
	
		For Each oFile in oFolder.Files
			strFileName = oFile.Name
			strFileType = oFile.Type
			strFileSize = oFile.Size
			strFilePath = oFile.Path
			strFileDtMod = oFile.DateLastModified
				
			FileArray(intCounter, 0) = strFileName
			FileArray(intCounter, 1) = "<A HREF=" & Chr(34) & "startDownload.asp?File=" _
				& Server.urlEncode(strFilePath) & "&Name=" & Server.urlEncode(strFileName) & "&Size=" & strFileSize & Chr(34) _
				& " onMouseOver=" & Chr(34) & "self.status='" & strFileName & "'; return true;" & Chr(34) _
				& " onMouseOut=" & Chr(34) & "self.status=''; return true;" & Chr(34) & ">" & strFileName & "</A>"
			FileArray(intCounter, 2) = strFileType
			FileArray(intCounter, 3) = strFileSize
			FileArray(intCounter, 4) = strFilePath
			FileArray(intCounter, 5) = strFileDtMod
		
			intCounter = (intCounter + 1)
		Next
		
		intRows = uBound(FileArray, 1)
		intCols = uBound(FileArray, 2)
	
		For x = 0 To intRows -1
			Echo("<TR>")
			For z = 0 To intCols
				If z > 0  Then
					BuildTableCol FileArray(x, z),""
				End IF
			Next
			Echo("</TR>")
		Next
		
	Else
	
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Fields.Append "Name", adVarChar, 500
		oRS.Fields.Append "Type", adVarChar, 500
		oRS.Fields.Append "Size", adInteger
		oRS.Fields.Append "Path", adVarChar, 500
		oRS.Fields.Append "Date", adFileTime
		oRS.Open
		
		For Each oFile in oFolder.Files
			strFileName = oFile.Name
			strFileType = oFile.Type
			strFileSize = oFile.Size
			strFilePath = oFile.Path
			strFileDtMod = oFile.DateLastModified
			
			oRS.AddNew
			oRS.Fields("Name").Value = "<A HREF=" & Chr(34) & "startDownload.asp?File=" _
				& Server.urlEncode(strFilePath) & "&Name=" & Server.urlEncode(strFileName) & "&Size=" & strFileSize & Chr(34) _
				& " onMouseOver=" & Chr(34) & "self.status='" & strFileName & "'; return true;" & Chr(34) _
				& " onMouseOut=" & Chr(34) & "self.status=''; return true;" & Chr(34) & ">" & strFileName & "</A>"
			oRS.Fields("Type").Value = strFileType
			oRS.Fields("Size").Value = strFileSize
			oRS.Fields("Path").Value = strFilePath
			oRS.Fields("Date").Value = strFileDtMod
		Next
		
		oRS.Sort = strSortBy 
		
		Do While Not oRS.EOF
			Echo("<TR>")
				BuildTableCol oRS("Name"),""
				BuildTableCol oRS("Type"),""
				BuildTableCol FormatNumber(oRS("Size"),0,-1,0,-1),"right"
				'BuildTableCol oRS("Path"),""
				BuildTableCol oRS("Date"),""
			Echo("</TR>")
		oRS.MoveNext
		Loop			
		
		oRS.Close
		Set oRS = Nothing
	End IF
	
	EchoB("<B>" & oFolder.Files.Count & " Files Available</B>")
		
	Cleanup oFile
	Cleanup oFolder
	Cleanup oFS
End Sub

Function Echo(str)
	Echo = Response.Write(str & vbCrLf)
End Function

Function EchoB(str)
	EchoB = Response.Write(str & "<BR>" & vbCrLf)
End Function

Sub Cleanup(obj)
	IF isObject(obj) Then
		Set obj = Nothing
	End IF
End Sub

Function StripFileName(strFile)
	StripFileName = Left(strFile, inStrRev(strFile, "\"))
End Function

Sub BuildTableCol(strData, align)
  If align = "" Then
	Echo("<TD CLASS=DataCol> " & strData & " </TD>")
  Else
	Echo("<TD CLASS=DataCol Align=""" & align & """> " & strData & " </TD>")
  End If
End Sub

'Not implemented
Sub BuildTableRow(arrData)
	Dim intCols
	intCols = uBound(arrData)
	For y = 0 To intCols
		Echo("<TD CLASS=DataCol>" & arrData(y) & "</TD>")
	Next
End Sub

%>
<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function mkArray(n) ' Small utility function
	Dim s
	s = Space(n)
	mkArray = Split(s," ")
End Function

Sub writeBinary(bstr, path)
	Dim fso
	Dim ts
	Set fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
	Set ts = fso.createTextFile(path)
	If Err.number <> 0 Then
    	'MsgBox(Err.message)
    	Exit Sub
	End If
On Error GoTo 0
	ts.Write(bstr)
	ts.Close
End Sub

lngID = Request("iid")
If lngID < 1 Then
	Session("MSG") = "survey response index is missing"
	Response.Redirect "survey.list.asp"
End If

'Response.Write Request("URL") & "<br />" & vbCrLf
'Response.Write Request("REQUEST_METHOD") & "<br />" & vbCrLf
'Response.Write Request("HTTPS") & "<br />" & vbCrLf
'Response.Write Request("SERVER_NAME") & "<br />" & vbCrLf
'Response.Write Request("LOCAL_ADDR") & "<br />" & vbCrLf
'Response.Write Request("PATH_INFO") & "<br />" & vbCrLf
'Response.Write Request("PATH_TRANSLATED") & "<br />" & vbCrLf
'Response.Write Request("SERVER_PROTOCOL") & "<br />" & vbCrLf

strSrc = Request("URL")
arrUrl = Split(strSrc, "/")

strServerName = Request("SERVER_NAME")
strPDF = pdfStr & "Survey" & lngID & ".18.pdf"
If Request("HTTPS") = "on" Then
	strUrl = "https://"
Else
	strUrl = "http://"
End If
strUrl = strUrl & strServerName & "/"
For lngI = 1 To UBound(arrUrl) - 1
	strUrl = strUrl & arrUrl(lngI) & "/"
Next
strUrl = strUrl & "survey.report.asp?ix=" & lngID

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(strPDF) Then
	Call fso.DeleteFile(strPDF, TRUE)
End If

'Response.Write strPDF & "<br />" & vbCrLf
'Response.Write strUrl & "<br />" & vbCrLf

Set http = CreateObject("Microsoft.XmlHttp")
On Error Resume Next
http.open "GET", strUrl, False
http.send ""
strRep = ""
if err.Number = 0 Then
'	Response.Write "got it" 'http.responseText
	strRep = http.ResponseText
Else
'    Response.Write "error " & Err.Number & ": " & Err.Description
End If
Set http = Nothing

On Error Goto 0
Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT * FROM [surveyreports] WHERE [iid]=" & lngID
strGUID = Z_GenerateGUID()
' Response.Write "GUID: " & strGUID & "<br />" & vbCrLf
rsSurv.Open strSQL, g_strCONN, 1, 3
If rsSurv.EOF Then
	rsSurv.AddNew
	rsSurv("iid")		= lngID
	rsSurv("guid") 		= strGUID
	'rsSurv("release") 	= 0
End If
rsSurv("report") = strRep
rsSurv("release") = 1
rsSurv.Update
rsSurv.Close
Set rsSurv = Nothing

On Error Resume Next
'Set rsSurv = Server.CreateObject("ADODB.RecordSet")
Set theDoc = Server.CreateObject("ABCpdf6.Doc") 'converts html to pdf
If Err.Number <> 0 Then
	theDoc.HtmlOptions.PageCacheClear
	theDoc.HtmlOptions.RetryCount = 3
	theDoc.HtmlOptions.Timeout = 120000
	theDoc.Pos.X = 10
	theDoc.Pos.Y = 10
	theID = theDoc.AddImageUrl(strUrl)
	Do
		If Not theDoc.Chainable(theID) Then Exit Do
		theDoc.Page = theDoc.AddPage()
		theID = theDoc.AddImageToChain(theID)
	Loop
	
	For i = 1 To theDoc.PageCount
		theDoc.PageNumber = i
		theDoc.Flatten
	Next

	theDoc.Save strPDF	
''	Set pdf = fso.getFile(strPDF)
''	Set ots = pdf.OpenAsTextStream()
''	a = mkArray(pdf.Size)
''	i = 0
''	While Not ots.AtEndOfStream
''		a(i) = ots.Read(1)
''		i = i + 1
''	Wend
''	ots.Close
	
End If

'survey.release.asp
Response.Redirect "survey.list.asp"
%>

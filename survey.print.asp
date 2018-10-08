<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
lngID = Request("ix")
If lngID < 1 Then
	lngID = Session("UIntr")
	If lngID < 1 Then
		Response.Status = "404 File Not Found"
		Response.End
		' Session("MSG") = "survey response index is missing"
		' Response.Redirect "survey.list.asp"
	End If
'	Session("MSG") = "survey response index is missing"
'	Response.Redirect "survey.list.asp"
End If


strSrc = Request("URL")
arrUrl = Split(strSrc, "/")

strServerName = Request("SERVER_NAME")
strNm = "Survey" & lngID & ".18.pdf"
strPDF = pdfStr & strNm 
If Request("HTTPS") = "on" Then
	strUrl = "https://"
Else
	strUrl = "http://"
End If
strUrl = strUrl & strServerName & "/"
For lngI = 1 To UBound(arrUrl) - 1
	strUrl = strUrl & arrUrl(lngI) & "/"
Next
strUrl = strUrl & "survey.reppdf.asp?ix=" & lngID

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(strPDF) Then
	Call fso.DeleteFile(strPDF, TRUE)
End If

Set rsSurv = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT [release], [signature], [viewed], [medical_viewed] FROM [surveyreports] WHERE [iid]=" & lngID
rsSurv.Open strSQL, g_strCONN, 3, 1
blnRelease = FALSE
dtSig = ""

If Not rsSurv.EOF Then
	blnRelease = CBool( rsSurv("release") )
	dtSig = Z_CDate(rsSurv("signature"))
	dtSVw = Z_CDate(rsSurv("viewed"))
	dtMVw = Z_CDate(rsSurv("medical_viewed"))
End If
rsSurv.Close


Randomize
strUrl = strUrl & "&rnd=" & CLng(Rnd * 10000000)
'blnRelease = True
'If Not blnRelease Then
''	Response.Redirect "survey.list.asp"
'End If
'Set rsSurv = Nothing

'On Error Resume Next
Set theDoc = Server.CreateObject("ABCpdf6.Doc")

thedoc.HtmlOptions.PageCacheClear
theDoc.HtmlOptions.RetryCount = 3
theDoc.HtmlOptions.Timeout = 120000
theDoc.Pos.X = 10
theDoc.Pos.Y = 10
theDoc.Rect.Inset 50, 50
theDoc.Page = theDoc.AddPage()

theID = theDoc.AddImageUrl(strUrl, True, 950, True)

Do
  theDoc.Framerect
  If Not theDoc.Chainable(theID) Then Exit Do
  theDoc.Page = theDoc.AddPage()
  theID = theDoc.AddImageToChain(theID)
Loop

For i = 1 to theDoc.PageCount
     theDoc.PageNumber = i
     theDoc.Flatten
Next

Set rsSurv = Server.CreateObject("ADODB.Recordset")
'If Err.Number <> 0 Then
	Err.Clear
	strSQL = "SELECT [release], [viewed] FROM [surveyreports] WHERE [iid]=" & lngID
	rsSurv.Open strSQL, g_strCONN, 1, 3
	If Not rsSurv.EOF Then
		rsSurv("viewed") = Now
		rsSurv.Update
	End If
	rsSurv.Close
Set rsSurv = Nothing
'theDoc.Save "C:\work\apr_pdf\zz.pdf"
theData = theDoc.GetData() 

theDoc.Save strPDF
theDoc.Clear
Response.Addheader "Content-Disposition", "attachment; filename=""InSurvey." & lngID & ".pdf"""
Response.AddHeader "Content-Length", UBound(theData) - LBound(theData) + 1 
Response.Addheader "Content-Transfer-Encoding", "binary"
Response.ExpiresAbsolute = #January 1, 2001 01:00:00#
Response.CacheControl = "Private"
Response.ContentType = "application/pdf"

Response.BinaryWrite theData
Response.Flush
Response.End
%>
<!-- #inc  lude file="_Files.asp" -->
<%
DIM z_SMTP_From
z_SMTP_From = "language.services@thelanguagebank.org"

DIM z_SMTPServer(1), z_SMTP_Port(1), z_SMTP_User(1), z_SMTP_Pass(1)
z_SMTPServer(0) = "smtp.socketlabs.com"
z_SMTP_Port(0) = 2525
z_SMTP_User(0) = "server3874"
z_SMTP_Pass(0) = "c8W4Tmt5R3BaHn2"
z_SMTPServer(1) = "smtp.mailgun.org"
z_SMTP_Port(1) = 587
z_SMTP_User(1) = "postmaster@alt.thelanguagebank.org"
z_SMTP_Pass(1) = "d53256ad805ddbcf269221d16db0f6d1"


Function zSendMessage(strTo, strBCC, strSubject, strMSG)
	'SEND EMAIL
	lngIdx = 0
	blnOK = False
	Set mlMail = CreateObject("CDO.Message")
	mlMail.To = strTo
	'mlMail.To = "hagee@zubuk.com"
	mlMail.Bcc = strBCC
	mlMail.From = z_SMTP_From
	mlMail.Subject= strSubject
	If (InStr(strMSG, "<html>")>0) Then
		strMSG = "<!doctype html><html lang=""en""><head><meta charset=""utf-8"">" & _
				"<title>" & strSubject & "</title><meta name=""description"" content=""Notification"">" & _
				"<meta name=""author"" content=""Language Services""></head><body>" & vbCrLf & _
				strMSG & vbCrLf & "</body></html>"
	End If
	mlMail.HTMLBody = strMSG
	lngRet = 0
On Error Resume next
	Do
		With mlMail.Configuration.Fields
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")			= 2
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")			= z_SMTPServer(lngIdx)
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")		= z_SMTP_Port(lngIdx)
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")	= 1 'basic (clear-text) authentication
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")		= z_SMTP_User(lngIdx)
			.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")		= z_SMTP_Pass(lngIdx)
			.Update
		End With

		mlMail.Send
		lngRet = Err.Number

		If Err.Number = 0 Then
			blnOK = zLogMailMessage(Err.Number, mlMail.To, mlMail.Subject, z_SMTPServer(lngIdx), mlMail.HTMLBody, mlMail.Bcc)
			blnOK = True
		Else
			lngIdx = lngIdx + 1
			If lngIdx > 0 Then
				mlMail.Bcc = ""
			End If
			If lngIdx <= UBound(z_SMTPServer) Then
				lngMo = zMsgsLastMonth("smtp.mailgun.org")
				lngHr = zMsgsLastHour("smtp.mailgun.org")
				If lngHr >= 100 Or lngMo >= 10000 Then
					' we're unable to send a message
					blnOK = zLogMailMessageRem(Err.Number, mlMail.To, mlMail.Subject, z_SMTPServer(0) _
							, mlMail.HTMLBody, mlMail.Bcc _
							, "OVERLIMIT: " & lngHr & "|" & lngMo )
					blnOK = True
				End If
			Else
				blnOK = zLogMailMessageRem(Err.Number, mlMail.To, mlMail.Subject, z_SMTPServer(0) _
						, mlMail.HTMLBody, mlMail.Bcc _
						, "TOTAL FAILURE ON " & lngIdx & " SERVERS")
				blnOK = True
			End If
		End If
	Loop Until blnOK
On Error Goto 0
	zSendMessage = lngRet
	Set mlMail = Nothing
End Function

Function zLogMailMessage(lngerr, strto, subject, smtp, body, cc)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	rsLog.Open "[log_email]", g_strCONN, 1, 3
	rsLog.AddNew
	rsLog("err") = lngerr
	rsLog("to") = strto
	rsLog("subject") = subject
	rsLog("smtp") = smtp
	rsLog("body") = body
	rsLog("cc") = cc
	rsLog.Update
	rsLog.Close
	Set rsLog = Nothing
	zLogMailMessage = True
End Function

Function zLogMailMessageRem(lngerr, strto, subject, smtp, body, cc, remk)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	rsLog.Open "[log_email]", g_strCONN, 1, 3
	rsLog.AddNew
	rsLog("err") = lngerr
	rsLog("to") = strto
	rsLog("subject") = subject
	rsLog("smtp") = smtp
	rsLog("body") = body
	rsLog("cc") = cc
	rsLog("rem") = remk
	rsLog.Update
	rsLog.Close
	Set rsLog = Nothing
	zLogMailMessage = True
End Function

Function zMsgsLastHour(smtp)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	dtLsHour   = DateAdd("h", -1, Now)
	strLsHour  = DatePart("yyyy", dtLsHour) & "-" & DatePart("m", dtLsHour) & "-" & _
			DatePart("d", dtLsHour) & " " & FormatDateTime(dtLsHour, 4)

	strSQL = "EXEC [dbo].[CountMessages] '" & strLsHour & "', '" & smtp & "'"

	rsLog.Open strSQL, g_strCONN, 3, 1
	If rsLog.EOF Then
		zMsgsLastHour = 0
	Else
		zMsgsLastHour = Z_CLng(rsLog("msgs"))
	End If
	rsLog.Close
	Set rsLog = Nothing
End Function

Function zMsgsLastMonth(smtp)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	dtLsMonth  = DateAdd("m", -1, Date)
	strLsMonth = DatePart("yyyy", dtLsMonth) & "-" & DatePart("m", dtLsMonth) & "-15"

	strSQL = "EXEC [dbo].[CountMessages] '" & strLsMonth & "', '" & smtp & "'"
	rsLog.Open strSQL, g_strCONN, 3, 1
	If rsLog.EOF Then
		zMsgsLastMonth = 0
	Else
		zMsgsLastMonth = Z_CLng(rsLog("msgs"))
	End If
	rsLog.Close
	Set rsLog = Nothing
End Function

Function zGetInterpreterEmailByID(xxx)
	zGetInterpreterEmailByID = ""
	Set rsEm = Server.CreateObject("ADODB.RecordSet")
	sqlEm = "SELECT [e-mail] FROM interpreter_T WHERE [index] = " & xxx
	rsEm.Open sqlEm, g_strCONN, 1, 3
	If Not rsEm.EOF Then
		zGetInterpreterEmailByID = rsEm("e-mail")
	End If
	rsEm.Close
	Set rsEm = Nothing
End Function
%>
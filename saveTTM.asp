<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
'save entry?
	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	sqlMain = "SELECT actMil, actTT, InstActMil,InstActTT, RealTT, RealM, sentIntr FROM request_T WHERE [index] = " & Request("appID")
	rsMain.Open sqlMain, g_strCONN, 1, 3
	If Not rsMain.EOF Then
		rsMain("actMil") = Request("txtMile")
		rsMain("actTT") = Request("txtTravel")
		rsMain("InstActMil") = Z_CZero(Request("txtMileInst"))
		rsMain("InstActTT") = Z_CZero(Request("txtTravelInst"))
		rsMain("RealTT") = Z_CZero(Request("txtRTravel"))
		rsMain("RealM") = Z_CZero(Request("txtRMile"))
		rsMain("sentIntr") = Now
		rsMain.Update
	End If
	rsMain.Close
	Set rsMain = Nothing
	Session("MSG") = "Interpreter Assigned"
	Call SaveHist(Request("appID"), "openappts.asp")
	appdate = Z_GetInfoFROMAppID(Request("appID"), "appdate")
	timeframe = Z_FormatTime(Z_GetInfoFROMAppID(Request("appID"), "appTimeFrom"), 4) & " - " & Z_FormatTime(Z_GetInfoFROMAppID(Request("appID"), "appTimeTo"), 4)
	inst = GetInst(Z_GetInfoFROMAppID(Request("appID"), "InstID"))
	tmpcity = GetCity(Z_GetInfoFROMAppID(Request("appID"), "deptid"))
	If Z_GetInfoFROMAppID(Request("appID"), "cliAdd") Then tmpcity =  Z_GetInfoFROMAppID(Request("appID"), "ccity")
	'send confirmation email to institution
	Pcon = GetPrime(Z_GetInfoFROMAppID(Request("appID"), "reqID"))
	Call Z_EmailInst(pcon, Request("appID"))
	'send email to intr
	tmpEmail = GetPrime2(Z_GetInfoFROMAppID(Request("appID"), "IntrID"))
	If Z_FixNull(tmpEmail) <> "" Then
		Set mlMail = CreateObject("CDO.Message")
		'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
		'mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 26
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.socketlabs.com"
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "server3874"
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "UO2CUSxat9ZmzYD7jkTB"
		mlMail.Configuration.Fields.Update
		mlMail.To = Trim(tmpEmail)
		mlMail.Cc = "language.services@thelanguagebank.org"
		mlMail.Bcc = "sysdump1@ascentria.org"
		mlMail.From = "DO-NOT-REPLY@thelanguagebank.org"
		mlMail.Subject = "[LBIS] Appointment assigned to you"
		strBody = "<p>Language Bank has assigned you to an appointment:<br><br>" & _
			"ID: " & Request("appID") & "<br>" & _
			"Institution: " & inst & "<br>" & _
			"Date: " &appdate & "<br>" & _
			"Timeframe: " & timeframe & "<br>" & _
			"Location: " & tmpcity & "<br><br>" & _
			"Please log into the <a href='https://interpreter.thelanguagebank.org/'>LB database</a> and download the verification form for this appointment.</p>" & _
			"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email.</font>"
	'response.write strBody & "<br>"
	
		mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"	
		mlMail.Send
		set mlMail = Nothing
	End If
	'send email to unassigned "yes" interpreters
	Response.Redirect "openappts.asp?reload=1&frmdte=" & Request("frmdte") & "&todte=" & Request("todte")& "&selLang=" & Request("SelLang")
%>
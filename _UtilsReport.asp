<%
Function Z_GetBillhrsCourt(timefrom, timeto) 'finish this
	Z_GetBillhrsCourt = 1.5
	If DateDiff("n", timefrom, timeto) > 90 Then
		tmpBillMin = DateDiff("n", timefrom, timeto)
		timebefore75 = tmpBillMin / 60
		tmpBillHrs = timebefore75 * 0.75
		tmpBillMHrs = Int(tmpBillHrs)
		tmpLen = Len(tmpBillHrs)
		tmpPosDec = Instr(tmpBillHrs, ".")
		tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
		If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) < 0.22 Then
			Z_GetBillhrsCourt = tmpBillMHrs
		ElseIf Cdbl(tmpBillMMin) => 0.22 And  Cdbl(tmpBillMMin) < 0.38 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 0.25
		ElseIf Cdbl(tmpBillMMin) => 0.38 And  Cdbl(tmpBillMMin) < 0.63 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 0.50
		ElseIf Cdbl(tmpBillMMin) => 0.63 And  Cdbl(tmpBillMMin) < 0.88 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 0.75
		ElseIf Cdbl(tmpBillMMin) => 0.88 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 1
		Else
			Z_GetBillhrsCourt = tmpBillMHrs
		End If
	End If
End Function
Function Z_GetBillhrs(timefrom, timeto)
	Z_GetBillhrs = 2
	If DateDiff("n", timefrom, timeto) > 120 Then
		tmpBillMin = DateDiff("n", timefrom, timeto)
		tmpBillHrs = tmpBillMin / 60
		tmpBillMHrs = Int(tmpBillHrs)
		tmpLen = Len(tmpBillHrs)
		tmpPosDec = Instr(tmpBillHrs, ".")
		tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
		If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.5 Then
			Z_GetBillhrs = tmpBillMHrs + 0.5
		ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.99 Then
			Z_GetBillhrs = tmpBillMHrs + 1
		Else
			Z_GetBillhrs = tmpBillMHrs
		End If
	End If
End Function
Function GetReqCSV(zzz)
	GetReqCSV = "N/A"
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT Lname, Fname FROM requester_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		GetReqCSV = rsRP("Lname") & """,""" & rsRP("Fname")
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
Function Z_GetReqEmail(zzz)
	Z_GetReqEmail = "N/A"
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT Email FROM requester_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		Z_GetReqEmail = rsRP("Email")
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
Function Z_GetApptState(rid)
	Z_GetApptState = "N/A"
	Set rsAdr = Server.CreateObject("ADODB.RecordSet")
	rsAdr.Open "SELECT cliadd, deptid, cstate FROM request_T WHERE [index] = " & rid, g_strCONN, 3, 1
	If Not rsAdr.EOF Then
		If rsAdr("CliAdd") Then
			Z_GetApptState = rsAdr("CState")
		Else
			Z_GetApptState = GetDeptState(rsAdr("deptid"))
		End If
	End If
	rsAdr.Close
End Function
Function GetDeptState(xxx)
	GetDeptState = "N/A"
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = "SELECT State FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetDeptState = rsDept("State")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
Function Z_GetApptAddr(rid)
	Z_GetApptAddr = "N/A"
	Set rsAdr = Server.CreateObject("ADODB.RecordSet")
	rsAdr.Open "SELECT cliadd, deptid, caddress, CliAdrI, ccity, cstate, czip FROM request_T WHERE [index] = " & rid, g_strCONN, 3, 1
	If Not rsAdr.EOF Then
		If rsAdr("CliAdd") Then
			Z_GetApptAddr = rsAdr("CAddress") & ", " & rsAdr("CliAdrI") & ", " & rsAdr("CCity") & ", " & rsAdr("CState") & ", " & rsAdr("CZip")
		Else
			Z_GetApptAddr = GetDeptAdr(rsAdr("deptid"))
		End If
	End If
	rsAdr.Close
	Set rsAdr = Nothing
End Function
Function Z_GetHigherPay(defrate, intrid)
	Z_GetHigherPay = defrate
	Set rsElig = Server.CreateObject("ADODB.RecordSet")
	hpay = "HigherPay"
	If IsIntrerpreterI(intrid) Then hpay = "higherPay2"
	rsElig.Open "SELECT " & hpay & " FROM [EmergencyFee_T]", g_strCONN, 3, 1
	If Not rsElig.EOF Then
		Z_GetHigherPay = rsElig(hpay)
	End If
	rsElig.Close
	Set rsElig = Nothing
End Function
Function Z_EligibleHigherPay(langid)
	Z_EligibleHigherPay  = False
	if langid <= 0 Then Exit Function
	Set rsElig = Server.CreateObject("ADODB.RecordSet")
	rsElig.Open "SELECT HPay FROM Language_T WHERE [index] = " & langid, g_strCONN, 3, 1
	If Not rsElig.EOF Then
		If rsElig("Hpay") Then Z_EligibleHigherPay = True
	End If
	rsElig.Close
	Set rsElig = Nothing
End Function
Function SearchArraysIntrZip(xzip, xintr, tmpZip, tmpintr)
	DIM	lngMax, lngI
	SearchArraysIntrZip = -1
	On Error Resume Next	
	lngMax = UBound(tmpintr)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpZip(lngI) = xzip And tmpintr(lngI) = xintr Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysIntrZip = lngI
End Function
Function Z_GetCustID(deptid)
	Z_GetCustID = ""
	If Z_CZero(deptid) = 0 Then Exit Function
	Set rsCust = Server.CreateObject("ADODB.RecordSet")
	rsCust.Open "SELECT CustID FROM dept_T WHERE [index] = " & deptid, g_strCONN, 3, 1
	If Not rsCust.EOF Then Z_GetCustID = rsCust("CustID")
	rsCust.Close
	Set rsCust = Nothing
End Function
Function Z_CleanDate(dte)
	If dte = "" Then Exit Function
	myMonth = Right(0 & Month(dte), 2)
	myDay = Right(0 & Day(dte), 2)
	myYear = Right(Year(dte), 2)
	Z_CleanDate = myMonth & myDay & myYear
End Function
Function Z_GetUIDHP(hpid)
	Z_GetUIDHP = 0
	If Z_CZero(hpid) = 0 Then Exit Function
	Set rsHp = Server.CreateObject("ADODB.RecordSet")
	rsHp.Open "SELECT [UID] FROM appointment_T WHERE [index] = " & hpid, g_strCONNHP, 3, 1
	If Not rsHp.EOF Then Z_GetUIDHP = rsHp("UID")
	rsHp.Close
	Set rsHp = Nothing
End Function
Function Z_GetLoginHP(uid)
	Z_GetLoginHP = "N/A"
	If Z_CZero(uid) = 0 Then Exit Function
	Set rsLogin = Server.CreateObject("ADODB.RecordSet")
	rsLogin.Open "SELECT [user] FROM user_T WHERE [index] = " & uid, g_strCONNHP, 3, 1
	If Not rsLogin.EOF Then Z_GetLoginHP = rsLogin("user")
	rsLogin.Close
	Set rsLogin = Nothing
End Function
Function Z_GetDefRate(IntrID)
	Z_GetDefRate = 0
	If Z_CZero(IntrID) <= 0 Then Exit Function
	Set rsIntrRate = Server.CreateObject("ADODB.RecordSet")
	rsIntrRate.Open "SELECT [Rate] FROM Interpreter_T WHERE [index] = " & IntrID, g_strCONN, 3, 1
	If Not rsIntrRate.EOF Then
		Z_GetDefRate = Z_CZero(rsIntrRate("Rate"))
	End If
	rsIntrRate.Close
	Set rsIntrRate = Nothing
End Function
Function AppAssigned(appId)
	AppAssigned = False
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	rsReq.Open "SELECT IntrID FROM request_T WHERE [index] = " & appID, g_strCONN, 3, 1
	If Z_CZero(rsReq("IntrID")) > 0 Then AppAssigned = True
	rsReq.Close
	Set rsReq = Nothing
End Function
Function SkedCheckLenient(intrID, UID, appdate, timefrom, timeto)
	SkedCheckLenient = 0
	Meron = 0
	If Not intrID > 0 Then Exit Function
	'check if same start time
	Set rsSked = Server.CreateObject("ADODB.RecordSet")
	sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
		"AND Request_T.[index] <> " & UID
	rsSked.Open sqlSked, g_strCONN, 3, 1
	If Not rsSked.EOF Then
		Meron = 1
	End If	 
	rsSked.Close
	Set rsSked = Nothing
	If Meron = 0 Then
		'check if same end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if same time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND apptimeto = '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt start time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timefrom & "' " & _
			"AND apptimeto > '" & timefrom & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timeto & "' " & _
			"AND apptimeto >= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if overlap app
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom >= '" & timefrom & "' " & _
			"AND apptimeto <= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	SkedCheckLenient = Meron
End Function
Function Z_PrevDRG(deptID)
	Z_PrevDRG = False
	Set rsDrg = Server.CreateObject("ADODB.RecordSet")
	rsDrg.Open "SELECT * FROM prevMCO_T WHERE DeptID = " & deptID, g_strCONN, 3, 1
	If Not rsDrg.EOF Then
		Z_PrevDRG = True
	End If
	rsDrg.Close
	Set rsDrg = Nothing
End Function
Function GetLangSurvey(lngID)
	If lngID = 3 Then 
		GetLangSurvey = "clientsurveyArabic.pdf"
	ElseIf lngID = 10 Then
		GetLangSurvey = "clientsurveyFarsi.pdf"
	ElseIf lngID = 17 Then
		GetLangSurvey = "clientsurveyKorean.pdf"
	ElseIf lngID = 49 Then
		GetLangSurvey = "clientsurveyNepali.pdf"
	ElseIf lngID = 21 Then
		GetLangSurvey = "clientsurveyPortuguese.pdf"
	ElseIf lngID = 22 Then
		GetLangSurvey = "clientsurveyRussian.pdf"
	ElseIf lngID = 24 Then
		GetLangSurvey = "clientsurveySomali.pdf"
	ElseIf lngID = 25 Then
		GetLangSurvey = "clientsurveySpanish.pdf"
	ElseIf lngID = 29 Then
		GetLangSurvey = "clientsurveyVietnamese.pdf"
	Else
		GetLangSurvey = "clientsurveyEnglish.pdf"
	End If
End Function

Function Z_EmailInst(pcon, appid)
	'GET REQUEST INFO
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT r.[CLname], r.[CFname], r.[IntrID], r.[LangID], r.[appDate], r.[appTimeFrom], r.[appTimeTo] " & _
				", r.[InstID], r.[DeptID], r.[MRRec], r.[ReqID], r.[timestamp], r.[DocNum], r.[CrtRumNum] " & _
				", r.[SentReq], r.[CliAdd], r.[CAddress], r.[CliAdrI], r.[CCity], r.[CState], r.[CZip] " & _
				", r.[intrcomment], r.[HPID], r.[actTT], r.[actMil], r.[Claimant], r.[Judge] " & _
				", l.[Language], r.[cc_addr] " & _
				", COALESCE(i.[First Name], '') + ' ' + COALESCE(i.[Last Name], '') AS [intr_nm] " & _
				", COALESCE(q.[lname], '') + ', ' + COALESCE(q.[fname], '') AS [requester_nm] " & _
				", d.[Dept] AS [department_nm] " & _
				", n.[Facility] AS [inst_nm] " & _
				", d.[Address] As [DepAddr], d.[InstAdrI] AS [InstAdrI], d.[City] As [DepCity]" & _
				", d.[State] AS [DepState], d.[ZIP] AS [DepZIP]" & _
				", d.[BAddress], d.[BCity], d.[BState], d.[BZip], d.[Blname], d.[Bfname]" & _
				", COALESCE(a.[ReqName], '') AS [AppReqName] " & _
			"FROM [request_T] AS r " & _
				"INNER JOIN [language_T] AS l ON r.[langID]=l.[index] " & _
				"INNER JOIN [dept_T] AS d ON r.[DeptID]=d.[index] " & _
				"INNER JOIN [institution_T] AS n ON r.[InstID]=n.[index] " & _
				"LEFT JOIN [interpreter_T] AS i ON r.[IntrID]=i.[index] " & _
				"LEFT JOIN [requester_T] AS q ON r.[ReqID]=q.[index] " & _
				"LEFT JOIN [interpretersql].[dbo].[appointment_T] AS a ON r.[HPID]=a.[index] " & _
			"WHERE r.[index] = " & appid
	rsReq.Open sqlReq, g_strCONN, 3, 1
	If Not rsReq.EOF Then
		CliName = rsReq("Clname") & ", " & rsReq("Cfname")
		IntrName = rsReq("Intr_nm")
		tmpIntrName = IntrName
		tmpIntr = IntrName
		LangName = rsReq("Language")
		LangID = rsReq("LangID")
		AppFrame = rsReq("appDate") & " (" & rsReq("appTimeFrom") & " - " & rsReq("appTimeTo") & ")" 
		AppDate = rsReq("appDate")
		InstID = rsReq("InstID")
		DeptID = rsReq("DeptID")
		'tmpDOB = rsReq("DOB")
		'tmpComment = rsReq("Comment")
		mrrec = rsReq("mrrec")
		ReqName = rsReq("requester_nm")
		timestamp = rsReq("timestamp")
		tmpOther = rsReq("DocNum") & ",  " & rsReq("CrtRumNum")
		
		tmpCCAddr = Z_FixNull(rsReq("cc_addr"))
		If (Not Z_Blank(tmpCCAddr)) Then
			If Len(tmpCCAddr) > 5 Then
				If  (InStr(tmpCCAddr, "@")<2) Then
					' it's a fax!
					tmpCCAddr = tmpCCAddr & "@emailfaxservice.com"
				End If
				tmpCCAddr = tmpCCAddr & ";"
			Else
				tmpCCAddr = ""
			End If
		End If

		tmpdept =  rsReq("department_nm")
		tmpCon = rsReq("SentReq")
		If rsReq("CliAdd") = True Then InstAddr =  rsReq("CAddress") & ", " & rsReq("CliAdrI") & ", " & rsReq("CCity") & ", " & rsReq("CState") & ", " & rsReq("CZip")
		If rsReq("CliAdd") = True Then SubCity = rsReq("CCity")
		tmpcomintr = rsReq("intrcomment")
		tmpHPID = rsReq("HPID")
		tmpDecTT = z_fixNull(rsReq("actTT"))
		tmpDecMile = z_fixNull(rsReq("actMil"))
		tmpclaim = rsReq("claimant")
		tmpjudge = rsReq("judge")
		InstName = rsReq("inst_nm")
		subInst = rsReq("inst_nm")

		InstName = InstName & " - " & tmpdept
		If InstAddr = "" Then _
				InstAddr = rsReq("DepAddr") & ", " & rsReq("InstAdrI") & ", " & _
						rsReq("DepCity") & ", " & rsReq("DepState") & ", " & rsReq("DepZIP")
		If SubCity = "" Then SubCity = rsReq("DepCity")
		BillAddr =  rsReq("BAddress") &", " & rsReq("BCity") & ", " & rsReq("BState") & ", " & rsReq("BZip")
		tmpBContact = rsReq("Blname") & ", " & rsReq("Bfname")
		If Z_CZero(tmpHPID) <> 0 Then
			ReqName = rsReq("AppReqName")
		End If
	End If
	rsReq.Close
	Set rsReq = Nothing

	strTo = FixEmail(pcon)
	strBCC = tmpCCAddr & "language.services@thelanguagebank.org"
	strSubject= "Interpreter Confirmation - The Language Bank"
	strBody = "<table cellpadding='0' cellspacing='0' border='0' align='center'>" & vbCrLf & _
			"<tr><td align='center'>" & vbCrLf & _
				"<img src='http://languagebank.lssne.org/lsslbis/images/LBISLOGOBandW.jpg'>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
			"<tr><td align='center'>" & vbCrLf & _
				"<font size='2' face='trebuchet MS'><b>Appointment Confirmation:</b></font><br>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td>" & vbCrLf & _
				"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Appointment Requested by:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & ReqName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Date of Request:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & timestamp & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Date and Time of Confirmation:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpCon & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Interpreter:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpIntrName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
				"</table>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
			"<tr><td>" & vbCrLf & _
				"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right' width='225px'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Project ID Number:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & appid & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Client Name:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & CliName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>MR #:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & mrrec & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Date of Appointment:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & AppFrame & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Address of Appointment:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & InstAddr & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Language:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & LangName & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Department:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpdept & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Doc Num / Court / Delivery Ticket:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpOther & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr>" & vbCrLf & _
						"<td align='right'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>Comment:</font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
						"<td align='left'>" & vbCrLf & _
							"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpComment & "</b></font><br>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf
				If InstID = 757 Or InstID = 777 Then 'SSA
					strBody = strBody & "<tr>" & vbCrLf & _
							"<td align='right'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>Claimant:</font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
							"<td align='left'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpclaim & "</b></font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
						"</tr>" & vbCrLf & _
						"<tr>" & vbCrLf & _
							"<td align='right'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>Judge:</font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
							"<td align='left'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpjudge & "</b></font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
						"</tr>" & vbCrLf & _
						"<tr>" & vbCrLf & _
							"<td align='right'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>Interpreter:</font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
							"<td align='left'>" & vbCrLf & _
								"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpIntr & "</b></font><br>" & vbCrLf & _
							"</td>" & vbCrLf & _
						"</tr>" & vbCrLf
				End If
				strBody = strBody & "</table>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
			"<tr><td align='left'>" & vbCrLf & _
				"<font size='2' face='trebuchet MS'>The request for the above appointment has been received and confirmed.  A Language Bank Interpreter will be<br>" & vbCrlf & _
				"present for this appointment.  If any of the above information is not correct, changes or you have additional questions,<br>" & vbCrLf & _
				"please contact the LanguageBank office immediately at 410-6183 or email us at <a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>.<br>"  & vbCrlf & _
				"Please refer to the Project ID Number when calling or emailing the office. If there are any difficulties in completing<br>" & vbCrLf & _
				"this assignment you will be notified.<br><br>" & vbCrLf & _
				"Language Bank Cancellation Policy:<br>" & vbCrLf & _
				"Under the following conditions you, or your agency, will still be responsible for full payment to The Language Bank:<br>" & vbCrLf & _
				"&nbsp;&nbsp;&nbsp;&nbsp;• If you or your agency cancels a request with less than 24-hour notice prior to the scheduled service.<br>" & vbCrLf & _
		    "&nbsp;&nbsp;&nbsp;&nbsp;• If a patient cancels or reschedules an appointment with less than 24-hour notice prior to the scheduled service.<br>" & vbCrLf & _
		    "&nbsp;&nbsp;&nbsp;&nbsp;• If a patient does not show up for scheduled appointment.<br>" & vbCrLf & _
				"By making this appointment you agree to our cancellation policy.<br><br>" & vbCrLf & _
				"Emergency Request (less than 24 hours notice):<br>" & vbCrLf & _
				"&nbsp;&nbsp;&nbsp;&nbsp;• Requests received less than 24 hours notice are subject to additional $20.00 fee<br><br>" & vbCrLf & _
				"DISCLAIMER: This cancellation policy does not apply to American Sign Language (ASL) appointments as Language Bank must<br>" & vbCrlf & _
				"follow NH regulations/rules related to ASL interpretation. For questions, please contact Language Bank.<br><br>" & vbCrLf & _
				"Thank you for using LanguageBank for your interpretation needs.</font><br><br>" & vbCrLf & _
				"<font size='2' face='Script MT Bold'><i>Alen Omerbegovic</i></font><br><br>" & vbCrLf & _
				"<font size='1' face='trebuchet MS'><i><b>Alen Omerbegovic</b><br>" & vbCrLf & _
				"Program Manager<br>" & vbCrLf & _
				"LanguageBank<br>" & vbCrLf & _
				"Ascentria Care Alliance<br>" & vbCrLf & _ 
				"340 Granite Street, 3rd Floor <br>" & vbCrLf & _ 
				"Manchester, NH 03102 <br>" & vbCrLf & _ 
				"603-410-6183  603-410-6186 fax<br><br>" & vbCrLf & _ 
				"PLEASE NOTE: This email/fax is intended only for the use of the individual or entity to which it is addressed, and may contain<br>" & vbCrLf & _
				"information that is privileged, confidential and exempt from disclosure under applicable law.  If you are not the intended recipient,<br>" & vbCrLf & _
				"then dissemination, distribution or copying of this communication is strictly prohibited.  If you have received this communication in<br>" & vbCrLf & _
				"error, please notify LSS immediately at 1-800-244-8119 and return the original email to us at the above address." & vbCrLf & _
				"</i></font><br><br>" & vbCrLf & _
				"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email. Use the information above for questions.</font>" & vbCrLf & _
			"</td></tr>" & vbCrLf & _
		"</table>"
	retErr = zSendMessage(strTo, strBCC, strSubject, strBody)
	Call SaveHist(appID, "email.asp") 

	sqlSent = "UPDATE request_T SET SentReq = '" & Now & "' WHERE [index] = " & appid
	Set rsSent = Server.CreateObject("ADODB.RecordSet")
	rsSent.Open sqlSent, g_strCONN, 1, 3
	Set rsSent = Nothing

	'CREATE LOG
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(EmailLog, 8, True)
	strLog = Now & vbCrLf & "----- EMAIL SENT -----" & vbCrLf & _
		"TO: " & pcon & vbCrLf & _
		"REQUEST ID: " & appID
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
End Function 	' Z_EmailInst

Function FixEmail(stremail)
	If stremail = "" Then Exit Function
	notfixemail = trim(stremail)
	If left(notfixemail, 1) = "'" Then
		notfixemail = Mid(notfixemail, 2, Len(notfixemail) - 1)
	End If
	If right(notfixemail, 1) = "'" Then
		notfixemail = Mid(notfixemail, 1, Len(notfixemail) - 1)
	End If
	notfixemail = Replace(notfixemail,"'", "")
	notfixemail = Replace(notfixemail," ", "")
	notfixemail = Replace(notfixemail,"(", "")
	notfixemail = Replace(notfixemail,")", "")
	'notfixemail = Replace(notfixemail,"-", "")
	FixEmail = notfixemail
End Function

Function GetPrime2(xxx)
	GetPrime2 = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime2 = rsRP("E-mail")
		ElseIf rsRP("prime") = 1 Or rsRP("prime") = 2 Then
			GetPrime2 = ""
		ElseIf rsRP("prime") = 3 Then
			GetPrime2 = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
Function GetPrime(xxx)
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime = rsRP("Email")
		ElseIf rsRP("prime") = 1 Then
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
Function Z_GetPayStart(appdate)
	If Z_FixNull(appdate) = "" Then
		dte = date
	Else
		dte = Z_DateNull(appdate)
	End If
	wk1 = "12/16/2012"
	difwk = DateDiff("ww", wk1, dte)
	If Not Z_IsOdd2(difwk) Then
		myDate = dte
	Else
		myDate = DateAdd("d", -7, dte)
	End If
	If Weekday(myDate) = 1 Then
		sunDate = myDate
	Else
		difDate = DatePart("w", myDate)
		sunDate = DateAdd("d", -Cint(difDate - 1), myDate)
	End If	
	Z_GetPayStart = sunDate
End Function
Function Z_GetPayEnd(appdate)
	If Z_FixNull(appdate) = "" Then
		dte = date
	Else
		dte = Z_DateNull(appdate)
	End If
	wk1 = "12/16/2012"
	difwk = DateDiff("ww", wk1, dte)
	If Not Z_IsOdd2(difwk) Then
		myDate = dte
	Else
		myDate = DateAdd("d", -7, dte)
	End If
	If Weekday(myDate) = 1 Then
		sunDate = myDate
		satDate = DateAdd("d", 6, sunDate)
	Else
		difDate = DatePart("w", myDate)
		tmpDates = DateAdd("d", -Cint(difDate - 1), myDate)
		sunDate = DateAdd("d", 7, tmpDates)
		satDate = DateAdd("d", 6, sunDate)
	End If
	Z_GetPayEnd = satDate
End Function
Function Z_GetAppDate(appID)
	Set rsDte = Server.CreateObject("ADODB.RecordSet")
	rsDte.Open "SELECT appDate FROM request_T WHERE [index] = " & appID, g_strCONN, 3, 1
	If Not rsDte.EOF Then
		Z_GetAppDate = rsDte("appDate")
	End If
	rsDte.Close
	Set rsDte = Nothing
End Function
Function Z_GetPastSked(intrID, appdate)
	tmpPast = ""
	startPay = Z_GetPayStart(appdate)
	endPay = Z_GetPayEnd(appdate)
	Set rsPast = Server.CreateObject("ADODB.RecordSet")
	rsPast.Open "SELECT appDate, appTimeFrom, appTimeTo, InstID FROM request_T WHERE intrID = " & intrID & " AND AppDate = '" & appdate & _
		"' AND status <> 2 AND status <> 3 ORDER BY appTimeFrom, appTimeTo", g_strCONN, 3, 1
	If Not rsPast.EOF Then
		Do Until rsPast.EOF
			tmpPast = tmpPast & "• " & rsPast("appDate") & " " & Ctime(rsPast("appTimeFrom")) & " - " & Ctime(rsPast("appTimeTo")) & " - " & GetInst(rsPast("InstID")) & vbCrLf
			rsPast.MoveNext
		Loop
	Else
		pastctr = 0
		tmpPast = ""
	End If
	rsPast.Close
	Set rsPast = Nothing
	Set rsCount = Server.CreateObject("ADODB.RecordSet")
	rsCount.Open "SELECT Count([index]) AS MyCount FROM request_T WHERE intrID = " & intrID & " AND AppDate >= '" & startPay & "' AND appdate <= '" & endPay & _
		"' AND status <> 2 AND status <> 3", g_strCONN, 3, 1
	pastctr = 0
	If Not rsCount.EOF Then pastctr = rsCount("myCount")
	rsCount.Close
	Set rsCount = Nothing
	IntrName = GetIntr(intrID) 
	Z_GetPastSked =  pastctr & " appointment/s found for " & intrName & " for the pay period " & startPay & " - " & endPay & "." & vbCrLf & tmpPast
	'Z_GetPastSked =  pastctr & " appointment/s found for " & intrName & " on " & appdate & "." & vbCrLf & tmpPast
End Function
Function Z_GetIntrCity(intrID)
	Z_GetIntrCity = ""
	Set rsPID = Server.CreateObject("ADODB.RecordSet")
	sqlPID = "SELECT city FROM interpreter_T WHERE [index] = " & intrID
	rsPID.Open sqlPID, g_strCONN, 3, 1
	If Not rsPID.EOF Then
		Z_GetIntrCity = Z_FixNull(rsPID("city"))
	End If
	rsPID.Close
	Set rsPID = Nothing
End Function
Function Z_IntrYesNo(appID)
	Z_IntrYesNo = False
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	rsRes.Open "SELECT accept FROM appt_T WHERE appID = " & appID, g_strCONN, 3, 1
	Do Until rsRes.EOF
		If rsRes("accept") = 1 Or rsRes("accept") = 2 Or rsRes("accept") = 0 Then
			Z_IntrYesNo = True
			Exit Do
		End If
		rsRes.MoveNext
	Loop
	rsRes.Close
	Set rsRes = Nothing
End Function
Function Z_ResetIntr(appID)
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	'rsRes.Open "UPDATE appt_T Set accept = 0 WHERE appID = " & appID, g_strCONN, 1, 3
	rsRes.Open "SELECT * FROM appt_T WHERE appID = " & appID, g_strCONN, 1, 3
	If Not rsRes.EOF Then
		Do Until rsRes.EOF
			rsRes("accept") = 0
			rsRes.Update
			rsRes.MoveNext
		Loop
	Else
		'for old appt
		Call Z_EmailJob(appID)
	End If
	rsRes.Close
	Set rsRes = Nothing
End Function
Function Z_ResetIntr2(appID)
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	rsRes.Open "DELETE FROM appt_T  WHERE appID = " & appID, g_strCONN, 1, 3
	Set rsRes = Nothing
	Call Z_EmailJob(appID)
End Function
Function Z_EmailJob(AppID)
	ts = now
	DeptID = Z_GetInfoFROMAppID(appID, "DeptID")
	DeptClass = ClassInt(DeptID)
	LangID = Z_GetInfoFROMAppID(appID, "LangID")
	IDtoLang = UCase(GetLang(LangID))
	LangSQL = " (Upper(Language1) = '" & IDtoLang & "' OR Upper(Language2) = '" & IDtoLang & "' OR Upper(Language3) = '" & IDtoLang & _
		"' OR Upper(Language4) = '" & IDtoLang & "' OR Upper(Language5) = '" & IDtoLang & "' OR Upper(Language6) = '" & IDtoLang & "') AND Active = 1"
	If DeptClass = 1 Then classSql = " Social = 1"
	If DeptClass = 2 Then classSql = " Private = 1"
	If DeptClass = 3 Then classSql = " Court = 1"
	If DeptClass = 4 Then classSql = " Medical = 1"
	If DeptClass = 5 Then classSql = " Legal = 1"
	If DeptClass = 6 Then classSql = " Mental = 1"
	AppDate = Z_GetInfoFROMAppID(appID, "AppDate")
	'If AppDate > DateAdd("m", 2, Date) Then Exit Function ' do not send if more than 2 months
	InstID = Z_GetInfoFROMAppID(appID, "InstID")
	appTimeFrom = Z_GetInfoFROMAppID(appID, "appTimeFrom")
	tmpAvail = Weekday(AppDate) & "," & Hour(appTimeFrom)
	appTimeTo = Z_GetInfoFROMAppID(appID, "appTimeTo")
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [index] as myIntrID, [e-mail], Phone1, sendonce FROM interpreter_T WHERE" & classSql & " AND" & LangSQL
	rsIntr.Open sqlIntr, g_strCONN, 1, 3
	Do Until rsIntr.EOF
		If Not OnVacation(rsIntr("myIntrID"), AppDate) Then
			If Avail(rsIntr("myIntrID"), tmpAvail) And NotRestrict(rsIntr("myIntrID"), InstID, DeptID) Then
				If SkedCheckLenient(rsIntr("myIntrID"), appID, AppDate, appTimeFrom, appTimeTo) = 0 Then
					'send email here
					If Z_FixNull(rsIntr("e-mail")) <> "" Then
						If AppDate < DateAdd("m", 2, Date) Then ' do not send if more than 2 months
							Urgent = ""
							If DateDiff("n", Now, appTimeFrom) >= 0 And DateDiff("n", Now, appTimeFrom) < 1440 Then Urgent = "URGENT"
							If Not rsIntr("sendonce") Or Urgent = "URGENT" Then
								rsIntr("sendonce") = True
								rsIntr.Update

								strBody = "<p>Language Bank has received new request for your language(s) and skills.<br>" & _
									"Please log into the <a href='https://interpreter.thelanguagebank.org/interpreter/'>LB database</a> and let us know if you are available.</p>" & _
									"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email." & appID & "</font>"

							End If
						End If
					End If
					'save to db
					Set rsApp = Server.CreateObject("ADODB.RecordSet")
					rsApp.Open "SELECT * FROM appt_T WHERE timestamp = '" & ts & "'", g_strCONN, 1, 3
					rsApp.AddNew
					rsApp("timestamp") = ts
					rsApp("appID") = appID
					rsApp("IntrID") = rsIntr("myIntrID")
					rsApp.Update
					rsApp.Close
					Set rsApp = Nothing
				End If
			End If
		End If
		rsIntr.MoveNext
	Loop
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function Z_GetInfoFROMAppID(AppID, infoneeded)
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT " & infoneeded & " FROM request_T WHERE [index] = " & AppID, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		Z_GetInfoFROMAppID = rsIntr(infoneeded)
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function ActiveSage(deptID)
	If Z_CZero(deptID) = 0 Then Exit Function
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	rsDept.Open "UPDATE Dept_T SET SageActive = 1 WHERE [index] = " & deptID, g_strCONN, 1, 3
	Set rsDept = Nothing
End Function
Function Z_IntrAdr(intIntr)
	Z_IntrAdr = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT Address1 FROM interpreter_T WHERE [index] = " & intIntr , g_strCONN, 3, 1
	If Not rsIntr.EOF Then Z_IntrAdr = rsIntr("Address1")
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function Z_IntrCity(intIntr)
	Z_IntrCity = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT city FROM interpreter_T WHERE [index] = " & intIntr , g_strCONN, 3, 1
	If Not rsIntr.EOF Then Z_IntrCity = rsIntr("city")
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function Z_IntrZip(intIntr)
	Z_IntrZip = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT [zip code] as myZip FROM interpreter_T WHERE [index] = " & intIntr , g_strCONN, 3, 1
	If Not rsIntr.EOF Then Z_IntrZip = rsIntr("myZip")
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function IsMedApp(strmed)
	IsMedApp = False
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT medicaid FROM medapprove_T WHERE [medicaid] = '" & strmed & "'", g_strCONN, 3, 1
	If Not rsIntr.EOF Then IsMedApp = True
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function IsIntrerpreterI(intrID)
	IsIntrerpreterI = False
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT InterpreterI FROM Interpreter_T WHERE [index] = " & intrID, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		If rsIntr("interpreterI") Then IsIntrerpreterI = True
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function GetReqlname(zzz)
	GetReqlname = "N/A"
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT Lname FROM requester_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		GetReqlname = rsRP("Lname")
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
Function GetReqfname(zzz)
	GetReqfname = "N/A"
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT Fname FROM requester_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		GetReqfname = rsRP("fname")
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
Function ClassInt(deptid)
	ClassInt = 0
	Set rsClass = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT class FROM Dept_T WHERE [index] = " & deptid
	rsClass.Open sqlReq, g_strCONN, 3, 1
	If not rsClass.EOF Then
		ClassInt = Z_Czero(rsClass("class"))
	End If
	rsClass.Close
	Set rsClass = Nothing
End Function
Function Progcode(hmoused)
	If hmoused = 0 Then Progcode = "LB"
	If hmoused = 1 Then Progcode = "MHP"
	If hmoused = 2 Then Progcode = "NHHF"
	If hmoused = 3 Then Progcode = "WSHP"
	'If Z_Czero(wid) = 0 Then Exit Function 
	'Progcode = "LB" & Right("00" & wid, 3)
End Function
Function LastApp(xxx)
	LastApp = "N/A"
	Set rsPID = Server.CreateObject("ADODB.RecordSet")
	sqlPID = "SELECT TOP 1 intrID, appdate FROM Request_T WHERE intrID = " & xxx & " ORDER BY Appdate DESC"
	rsPID.Open sqlPID, g_strCONN, 3, 1
	If Not rsPID.EOF Then
		LastApp = rsPID("appdate")
	End If
	rsPID.Close
	Set rsPID = Nothing
End Function
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
Function OnVacation(IntrID, appDate)
	OnVacation = False
	Set rsVac = Server.CreateObject("ADODB.RecordSet")
	sqlVac = "SELECT vacto, vacfrom, vacto2, vacfrom2 FROM interpreter_T WHERE [index] = " & intrID
	rsVac.Open sqlVac, g_strCONN, 3, 1
	If Not rsVac.EOF Then
		If Not IsNull(rsVac("vacfrom")) Then
			If appDate >= rsVac("vacfrom") And appDate <= rsVac("vacto") Then 
				OnVacation = True
			End If
		End If
		If onVacation = False Then
			If Not IsNull(rsVac("vacfrom2")) Then
				If appDate >= rsVac("vacfrom2") And appDate <= rsVac("vacto2") Then 
					OnVacation = True
				End If
			End If
		End If
	End If
	rsVac.Close
	Set rsVac = Nothing
End Function
Function NotRestrict(IntrID, InstID, DeptID)
	NotRestrict = True
	tmpNotRestrict = 1
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	sqlRes = "SELECT * FROM Restrict_T WHERE IntrID = " & IntrID & " AND InstID = " & InstID
	'response.write sqlRes
	rsRes.Open sqlRes, g_strCONN, 3, 1
	If Not rsRes.EOF Then
		tmpNotRestrict = 0
	End If
	rsRes.Close
	Set rsRes = Nothing
	If tmpNotRestrict = 1 Then
		Set rsRes = Server.CreateObject("ADODB.RecordSet")
		sqlRes = "SELECT * FROM Restrict2_T WHERE IntrID = " & IntrID & " AND DeptID = " & DeptID
		rsRes.Open sqlRes, g_strCONN, 3, 1
		If Not rsRes.EOF Then
			tmpNotRestrict = 0
		End If
		rsRes.Close
		Set rsRes = Nothing
	End If
	If tmpNotRestrict = 0 Then NotRestrict = False
End Function
Function Avail(myID, myTime)
	Avail = False
	Set rsAvail = Server.CreateObject("ADODB.RecordSet")
	sqlAvail = "SELECT * FROM Avail_T WHERE intrID = " & myID & " AND Avail = '" & myTime & "'"
	rsAvail.Open sqlAvail, g_strCONN, 3, 1
	If Not rsAvail.EOF Then Avail = True
	rsAvail.Close
	set rsAvail = Nothing
	If Avail Then Exit Function
	Set rsAvail2 = Server.CreateObject("ADODB.RecordSet")
	sqlAvail2 = "SELECT * FROM Avail_T WHERE IntrID = " & myID
	rsAvail2.Open sqlAvail2, g_strCONN, 3, 1
	If rsAvail2.EOF Then Avail = True
	rsAvail2.Close
	set rsAvail2 = Nothing
End Function
Function FixDateFormat(xxx)
		FixDateFormat = Right("0" & DatePart("m", xxx), 2) & "/" & Right("0" & DatePart("d", xxx), 2) & "/" & Year(xxx)
	End Function	
Function GetLBCode(hrs, medicaid, wid, appdate, hmoused)
	tmplbcode = ""
	tmplbcode = """" & medicaid & """,""" & wid & """,""" & hmoused & """,""" & "LBHR" & """,""" & "" & """,""" & _
		FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	If hrs > 2 Then
		newhrs = hrs - 2
		If Z_Mod(newhrs, 0.25) = 0 Then
			myUnits = newhrs / 0.25
			tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & hmoused & """,""" & "LB3QH" & """,""" & "" & """,""" & _
				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & myUnits & """,""" & """" & vbCrLf
		Else
			myUnits = Int(newhrs / 0.25) + 1
			tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & hmoused & """,""" & "LB3QH" & """,""" & "" & """,""" & _
				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & myUnits & """,""" & """" & vbCrLf
		End If
	End If
	'If hrs > 0 Then tmplbcode = """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB1QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	'If hrs > 0.26 Then tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB2QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	'If hrs > 0.51 Then tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB3QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	'If hrs > 0.76 Then tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB4QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	'If hrs > 1.1 Then tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB5QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	'If hrs > 1.26 Then tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB6QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	'If hrs > 1.51 Then tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB7QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	'If hrs > 1.76 Then tmplbcode = tmplbcode & """" & medicaid & """,""" & wid & """,""" & Progcode(wid) & """,""" & "LB8QH" & """,""" & "" & """,""" & _
	'				FixDateFormat(appdate) & """,""" & FixDateFormat(appdate) & """,""" & "1" & """,""" & """" & vbCrLf 
	GetLBCode = tmplbcode
End Function
Function PaytoMedicaid(outpatient, hasmed, vermed, autoacc, wcomp, drg, intrID, medicaid, meridian, nhhealth, wellsense)
	PaytoMedicaid = False
	If Z_FixNull(medicaid) <> "" Or Z_FixNull(meridian) <> "" Or Z_FixNull(nhhealth) <> "" Or Z_FixNull(wellsense) <> "" Then
		If outpatient And hasmed And (vermed = 1 Or vermed = 0) AND autoacc = false AND wcomp = False AND drg Then PaytoMedicaid = True
		If outpatient And hasmed And vermed = 2 AND autoacc = false AND wcomp = False AND drg Then PaytoMedicaid = False
	End If
End Function

Function GetPID(intrID)
	GetPID = ""
	If IntrID < 0 Then Exit Function
	Set rsPID = Server.CreateObject("ADODB.RecordSet")
	sqlPID = "SELECT PID FROM interpreter_T WHERE [index] = " & intrID
	rsPID.Open sqlPID, g_strCONN, 3, 1
	If Not rsPID.EOF Then
		GetPID = Z_FixNull(rsPID("PID"))
	End If
	rsPID.Close
	Set rsPID = Nothing
End Function
Function GetReasonTardy(xxx)
	If Z_CZero(xxx) = 0 Then
		GetReasonTardy = "N/A"
		Exit Function
	End If
	Set rsPID = Server.CreateObject("ADODB.RecordSet")
	sqlPID = "SELECT lateres FROM tardy_t WHERE [uid] = " & xxx
	rsPID.Open sqlPID, g_strCONN, 3, 1
	If Not rsPID.EOF Then
		GetReasonTardy = rsPID("lateres")
	End If
	rsPID.Close
	Set rsPID = Nothing
End Function
Function GetStatNum(xxx)
	GetStatNum = 0
	Set rsStat = CreateObject("ADODB.RecordSet")
	sqlStat = "SELECT Status FROM Request_T WHERE [index] = " & xxx
	rsStat.Open sqlStat, g_strCONN, 3, 1
	If Not rsStat.EOF Then
		GetStatNum = rsStat("status")
	End If
	rsStat.Close
	Set rsStat = Nothing
End Function
Function SkedCheck(intrID, UID, appdate, timefrom, timeto)
	SkedCheck = 0
	Meron = 0
	If Not intrID > 0 Then Exit Function
	'check if same start time
	Set rsSked = Server.CreateObject("ADODB.RecordSet")
	sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
		"AND Request_T.[index] <> " & UID
	rsSked.Open sqlSked, g_strCONN, 3, 1
	If Not rsSked.EOF Then
		Meron = 1
	End If	 
	rsSked.Close
	Set rsSked = Nothing
	If Meron = 0 Then
		'check if same end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if same time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND apptimeto = '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt start time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timefrom & "' " & _
			"AND apptimeto > '" & timefrom & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timeto & "' " & _
			"AND apptimeto >= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if no gap between appt
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND apptimeto = '" & timefrom & "' " & _
			"AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if overlap app
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom >= '" & timefrom & "' " & _
			"AND apptimeto <= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if next appointment is less than 2 hr
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND apptimefrom >= '" & dateadd("n", -120, timefrom) & "' " & _
			"AND apptimefrom < '" & timefrom & "' " & _
			"AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONN, 3, 1
		If Not rsSked.EOF Then
			Do Until rsSked.EOF
				If datediff("n", rsSked("apptimefrom"), rsSked("apptimeto")) < 121 Then
					If dateadd("n", 120, rsSked("apptimefrom")) > timefrom Then 
						Meron = 1
						Exit Do
					End If
				End If			
				rsSked.MoveNext
			Loop
		End If
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if previous appointment is less than 2 hr
		If datediff("n", timefrom, timeto) < 121 Then
			Set rsSked = Server.CreateObject("ADODB.RecordSet")
			sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
				"AND apptimefrom >= '" & timeto & "' " & _
				"AND apptimefrom <= '" & dateadd("n", 30, timeto) & "' " & _
				"AND Request_T.[index] <> " & UID
			'If intrID = 431 Then response.write sqlsked & "<br>"
			rsSked.Open sqlSked, g_strCONN, 3, 1
			If Not rsSked.EOF Then
				'Do Until rsSked.EOF
				'	If dateadd("n", 30, rsSked("apptimeto")) > timefrom Then 
						Meron = 1
				'		Exit Do
				'	End If
				'	rsSked.MoveNext
				'Loop	
			End If
			rsSked.Close
			Set rsSked = Nothing
		End If
		
		
		
		
	End If
	SkedCheck = Meron
End Function
Function ICSFormatTime(strTime)
	tmpYear = Year(strTime)
	tmpMonth = Month(strTime)
	If Len(Month(strTime)) = 1 Then tmpMonth = "0" & Right(Month(strTime), 1)
	tmpDay = Day(strTime)
	If Len(Day(strTime)) = 1 Then tmpDay = "0" & Right(Day(strTime), 1)
	tmpHr = Replace(FormatDateTime(TimeValue(strTime), 4), ":", "")
	ICSFormatTime = tmpYear & tmpMonth & tmpDay & "T" & tmpHr & "00"
End Function
Function CreateICS(UID, strSubj)
	CreateICS = False
	If UID = "" Then 
		Exit Function
	End If
	Set fso = CreateObject("Scripting.FileSystemObject")
	oFileName = CalPath & UID & ".ICS"
	Set ICSFile = fso.OpenTextFile(oFileName, 2, True)
	Set rsICS = Server.CreateObject("ADODB.RecordSet")
	sqlICS = "SELECT appDate, appTimeFrom, appTimeTo, InstID, DeptID FROM Request_T WHERE [index] = " & UID
	rsICS.Open sqlICS, g_strCONN, 3, 1
	If Not rsICS.EOF Then
		strICS = "BEGIN:VCALENDAR" & vbCrLf & _
			"VERSION:2.0" & vbCrLf & _
			"PRODID:-//LanguageBank//LBICS v1.0//EN" & vbCrLf & _
			"BEGIN:VEVENT" & vbCrLf & _
			"UID:" & UID & vbCrLf & _
			"DTSTAMP:" & ICSFormatTime(Now) & vbCrLf & _
			"ORGANIZER;CN=LanguageBank:MAILTO:info@thelanguagebank.org" & vbCrLf & _
			"DTSTART:" & ICSFormatTime(rsICS("appTimeFrom")) & vbCrLf & _
			"DTEND:" & ICSFormatTime(rsICS("appTimeTo")) & vbCrLf & _
			"SUMMARY:" & strSubj & vbCrLf & _
			"END:VEVENT" & vbCrLf & _
			"END:VCALENDAR"
		ICSFile.WriteLine strICS
		CreateICS = True
	End If
	rsICS.Close
	Set rsICS = Nothing
	ICSFile.CLose
	Set ICSFile = Nothing
	Set fso = Nothing
End Function
Function GetUsername(xxx)
	GetUsername = "N/A"
	If Z_FixNull(xxx) = "" Then Exit Function
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT * FROM user_T WHERE [index] = " & xxx
	rsReq.Open sqlReq, g_strCONN, 3, 1
	If not rsReq.EOF Then
		GetUsername = rsReq("fname") & " " & rsReq("lname")
	End If
	rsReq.Close
	Set rsReq = Nothing
End Function
Function MileageAmt(xxx)
	MileageAmt = 0
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT * FROM MileageRate_T"
	rsReq.Open sqlReq, g_strCONN, 3, 1
	If not rsReq.EOF Then
		MileRate = rsReq("mileageRate")
	End If
	rsReq.Close
	Set rsReq = Nothing
	MileageAmt = xxx * MileRate
End Function
Function MileageRate()
	MileageRate = 0
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT * FROM MileageRate_T"
	rsReq.Open sqlReq, g_strCONN, 3, 1
	If not rsReq.EOF Then
		MileageRate = rsReq("mileageRate")
	End If
	rsReq.Close
	Set rsReq = Nothing
End Function
Function IsHoliday(xxx)
	IsHoldiday = False
	Set rsHol = Server.CreateObject("ADODB.RecordSet")
	sqlHol = "SELECT * FROM holiday_T WHERE holdate = '" & xxx & "'"
	rsHol.Open sqlHol, g_strCONN, 3, 1
	If Not rsHol.EOF Then
		IsHoliday = True
	End If
	rsHol.Close
	Set rsHol = Nothing
End Function
Function IsCourt(xxx)
	IsCourt = False
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT class FROM Dept_T WHERE [index] = " & xxx
	rsReq.Open sqlReq, g_strCONN, 3, 1
	If not rsReq.EOF Then
		If rsReq("class") = 3 Then IsCourt = True
	End If
	rsReq.Close
	Set rsReq = Nothing
End Function
Function MyMileages(IntrID, LangID, InstID, DeptID, m, strM, distcode)
	ccode = "LB Mileage .535"
	mdes = " - Mileage Rate"
	tmpM = m / 0.535 ' - changed 1/18/17 m / 0.54 '0.56 - changed 11/15 'm / 0.40 - changed on 9/23  -- 0.
	If InstID = 777 Then ' SSA ODAR MANCHESTER
		ccode = "Mileage Rate-.42"
		tmpM = m / 0.42
	End If
	If IsCourt(DeptID) Then 
		ccode = "Mileage Rate CT .535"
		tmpM = m / 0.535 
	Else
		tmpM = m / 0.535
	End If
	If IsASL(LangID) Then
		tmpM = m
		ccode = "Travel ASL"
		mdes = " - Mileage Rate (pass through)"
	End If
	MyMileages = "DOTC" & """,""" & _
		"0" & """,""" & ccode & """,""" & strM & mdes & """,""" & date & """,""" & _
		distcode & """,""" & Z_FormatNumber(tmpM, 2)
End Function
Function MyTravelTime(IntrID, LangID, InstID, DeptID, TT, strTT, distcode)
	ccode = "Travel Time"
	ttdes = " - Travel Time"
	If IsCert(IntrID) Then
		tmpTT = TT / 38
		ccode = "Travel Time Cer"
	Else
		tmpTT = TT / 28
	End If
	If IsCourt(DeptID) Then  ' - added 9/23
		ccode = "Travel Time CT"
		tmpTT = TT / 33
	End If
	If InstID = 273 Or InstID = 374 Then 'dhmc and ply pedia  
		ccode = "DHMC Travel"
		tmpTT = 1
	End If
	If IsASL2(LangID) Then 'fix this
		tmpTT = TT
		ccode = "Travel ASL"
		ttdes = " - Travel Time (pass through)"
	End If
	'If InStr(tmpTT, ".") > 0 Then
	'	If Right(tmpTT, 3) <> "." Then 
	'		tmpTT = TT
	'		ccode = "Travel ASL"
	'	End If
	'End If
	MyTravelTime = "DOTC" & """,""" & _
		"0" & """,""" & ccode & """,""" & strTT & ttdes & """,""" & date & """,""" & _
		distcode & """,""" & Z_FormatNumber(tmpTT, 2)
End Function
Function MyTravelTimeDarthCourt(IntrID, LangID, InstID, TT, strTT, distcode)
	ccode = "Travel Time"
	If IsCert(IntrID) Then
		tmpTT = TT / 38
		ccode = "Travel Time Cer"
	Else
		tmpTT = TT / 28
	End If
	If InstID = 273 Then 
		ccode = "DHMC Travel"
		tmpTT = 1
	End If
	If IsASL(LangID) Then
		tmpTT = TT
		ccode = "Travel ASL"
	End If
	If InStr(tmpTT, ".") > 0 Then
		If Right(tmpTT, 3) <> "." Then 
			tmpTT = TT
			ccode = "Travel ASL"
		End If
	End If
	MyTravelTimeDarthCourt = "DOTC" & """,""" & _
				"0" & """,""" & ccode & """,""" & strTT & " - Travel Time" & """,""" & date & """,""" & _
				distcode & """,""" & Z_FormatNumber(tmpTT, 2)
End Function
Function IsCert(xxx)
	IsCert = False
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT CertIntr FROM Interpreter_T WHERE [index] = " & xxx
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		If rsIntr("certIntr") Then IsCert = True
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function IsASL(xxx)
	IsASL = False
	if xxx = 52 then IsASL = True
End Function
Function IsASL2(xxx)
	IsASL2 = False
	if xxx = 52 Or xxx = 90 Or xxx = 81 Or xxx = 85 Or xxx = 78 then IsASL2 = True
End Function
Function SaveHist(xxx, mypage)
	'SAVE HIST SQL
	server.scripttimeout = 360000
	tmpHist = ""
	Set rsHist = Server.CreateObject("ADODB.RecordSet")
	Set rsLB = Server.CreateObject("ADODB.RecordSet")
	sqlHist = "SELECT * FROM hist_T WHERE Timestamp = '" & Now & "'"
	sqlLB = "SELECT * FROM request_T WHERE [index] = " & xxx
	rsLB.Open sqlLB, g_strCONN, 1, 3
On error resume next
	rsHist.Open sqlHist, g_strCONNHist2, 1,3 
	If not rsLB.EOF Then
		rsHist.AddNew
		rsHist("LBID") = xxx
		rsHist("Timestamp") = Now
		rsHist("Author") = Request.Cookies("LBUsrName")
		rsHist("pageused") = mypage
		x = 1
		Do Until x = rsLB.Fields.Count
			If x = 7 Then 
				tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetLang(rsLB.Fields(x).Value) & ""","
			ElseIf x = 19 Then
				tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetInst(rsLB.Fields(x).Value) & ""","
			ElseIf x = 20 Then
				tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetDept(rsLB.Fields(x).Value) & ""","
			ElseIf x = 23 Then
				tmpHist = tmpHist & """" & rsLB.Fields(x).Value & "|" & GetIntr(rsLB.Fields(x).Value) & ""","
			Else
				tmpHist = tmpHist & """" & rsLB.Fields(x).Value & ""","
			End If
        	x = x + 1
    	Loop
    	rsHist("Hist") = trim(tmpHist)
		rsHist.Update
	End If
	rsLB.CLose
	set rsLB = Nothing
	rsHist.Close
	Set rsHist = Nothing
	SaveHist = True
End Function
Function IsBlock(xxx)
	IsBlock = False
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	sqlUser = "SELECT * FROM appointment_t WHERE [index] = " & xxx
	rsUser.Open sqlUser, g_StrCONNHP, 3, 1
	If Not rsUser.EOF Then
		'If InStr(Ucase(rsUser("lbcom")), "BLOCK SCHEDULE") > 0 Then IsBlock = True
		'If InStr(Ucase(rsUser("comment")), "BLOCK SCHEDULE") > 0 Then IsBlock = True
		If rsUser("block") Then isBlock = True 
	End If
	rsUser.Close
	Set rsUser = Nothing
End Function
Function GetUserID(xxx)
	GetUserID = "N/A"
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	sqlUser = "SELECT * FROM User_T WHERE DeptLB = " & xxx
	rsUser.Open sqlUser, g_StrCONNHP, 3, 1
	If Not rsUser.EOF Then
		GetUserID = rsUser("User")
	End If
	rsUser.Close
	Set rsUser = Nothing
End Function
Function GetAppDate(xxx)
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT CONVERT(varchar(10), appdate, 101) as myAppDate FROM request_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONN, 3, 1
	If Not rsLang.EOF Then
		GetAppDate = rsLang("myAppDate")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
Function SearchArraysHours(myIntr, tmpIntr)
	DIM	lngMax, lngI
	SearchArraysHours = -1
On Error Resume Next	
	lngMax = UBound(tmpIntr)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpIntr(lngI) = myIntr Then Exit For
	Next
	SearchArraysHours = lngI
	If lngI > lngMax Then SearchArraysHours = -1 'Exit Function
End Function
Function GetFileNum(xxx)
	GetFileNum = ""
	Set rsCity = Server.CreateObject("ADODB.RecordSet")
	sqlCity = "SELECT FileNum FROM interpreter_T WHERE [index] = " & xxx
	rsCity.Open sqlCity, g_strCONN,1, 3
	If Not rsCity.EOF Then
		GetFileNum = rsCity("FileNum")
	End If
	rsCity.Close
	Set rsCity = Nothing
End Function
Function GetCity(xxx)
	GetCity = ""
	Set rsCity = Server.CreateObject("ADODB.RecordSet")
	sqlCity = "SELECT City FROM dept_T WHERE [index] = " & xxx
	rsCity.Open sqlCity, g_strCONN,1, 3
	If Not rsCity.EOF Then
		GetCity = rsCity("City")
	End If
	rsCity.Close
	Set rsCity = Nothing
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function IsActive(xxx)
	'check if interpreter is active
	IsActive = True
	Set rsAct = Server.CreateObject("ADODB.RecordSet")
	sqlAct = "SELECT Active FROM interpreter_T WHERE [index] = " & xxx
	rsAct.Open sqlAct, g_strCONN, 3, 1
	If Not rsAct.EOF Then
		If rsAct("Active") = False Then IsActive = False
	End If
	rsAct.Close
	Set rsAct = Nothing	
End Function
Function GetReas(xxx)
	If xxx = "" THen
		GetReas = ""
		exit function
	End IF
	GetReas = ""
	tmpReas = Split(xxx, "|")
	CtrReas = Ubound(tmpReas)
	x = 0
	Do Until x = CtrReas + 1
		Set rsReas = Server.CreateObject("ADODB.RecordSet")
		sqlReas = "SELECT reason FROM Reason_T WHERE [index] = " & tmpReas(x)
		rsReas.Open sqlReas, g_strCONNHP, 3, 1
		If Not rsReas.EOF Then
			GetReas = GetReas & rsReas("reason") & "<br>"
		End If
		rsReas.Close
		Set rsReas = Nothing
		x = x + 1
	Loop
End Function
'GET STATUS
Function GetStat(zzz)
	Select Case zzz
		Case 0 GetStat = "Pending"
		Case 1 GetStat = "Completed"
		Case 2 GetStat = "Missed"
		Case 3 GetStat = "Canceled"
		Case 4 GetStat = "Canceled-Billable"
	End Select
End Function
'GET LANGUAGE
Function GetLang(zzz)
	GetLang = "N/A"
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT [Language] FROM language_T WHERE [index] = " & zzz
	rsLang.Open sqlLang, g_strCONN, 3, 1
	If Not rsLang.EOF Then
		GetLang = rsLang("Language")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
'GET INSTITUTION w/ CLASSIFICATION
Function GetInst(zzz)
	GetInst = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInst = rsInst("Facility")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET CLASS
Function GetClass(zzz)
	Select Case zzz
		Case 1 GetClass = "Social Services"
		Case 2 GetClass = "Private"
		Case 3 GetClass = "Court"
		Case 4 GetClass = "Medical"
		Case 5 GetClass = "Legal"
		Case 6 GetClass = "Mental Health"
	End Select
End Function
'GET INTERPRETER
Function GetIntr(zzz)
	GetIntr = "N/A, N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [Last Name], [First Name] FROM interpreter_T WHERE [index] = " & zzz
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		GetIntr = rsIntr("Last Name") & ", " & rsIntr("First Name")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function GetIntrFN(zzz)
	GetIntrFN = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [Last Name], [First Name] FROM interpreter_T WHERE [index] = " & zzz
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		GetIntrFN = Left(rsIntr("First Name"), 1) & ". " & rsIntr("Last Name")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
'GET INTERPRETER
Function GetIntr2(zzz)
	GetIntr2 = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [Last Name], [First Name] FROM interpreter_T WHERE [index] = " & zzz
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		GetIntr2 = rsIntr("First Name") & " " & rsIntr("Last Name")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
'SEARCH FOR TOWNS
Function SearchArraysTown(xtown, xname, xlang, xclass, tmpTown, tmpName, tmpLang, tmpClass)
	DIM	lngMax, lngI
	SearchArraysTown = -1
	On Error Resume Next	
	lngMax = UBound(tmpTown)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpTown(lngI) = xtown And tmpName(lngI) = xname And tmpLang(lngI) = xlang And tmpClass(lngI) = xclass Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysTown = lngI
End Function
'SEARCH FOR INSTITUTION
Function SearchArraysInst(xinst, xdept, xname, xlang, xclass, tmpInst, tmpDept, tmpName, tmpLang, tmpClass)
	DIM	lngMax, lngI
	SearchArraysInst = -1
	On Error Resume Next	
	lngMax = UBound(tmpInst)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpInst(lngI) = xinst And tmpDept(lngI) = xdept And tmpName(lngI) = xname And tmpLang(lngI) = xlang And tmpClass(lngI) = xclass Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysInst = lngI
End Function
'SEARCH FOR INSTITUTION 2
Function SearchArraysInst2(xinst, xdept, xlang, tmpInst, tmpDept, tmpLang)
	DIM	lngMax, lngI
	SearchArraysInst2 = -1
	On Error Resume Next	
	lngMax = UBound(tmpInst)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpInst(lngI) = xinst And tmpDept(lngI) = xdept And tmpLang(lngI) = xlang Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysInst2 = lngI
End Function
'SEARCH FOR INTERPRETER
Function SearchArraysIntr(xname, xinst, xlang, xclass, tmpIntrName, tmpInst, tmpLang, tmpClass)
	DIM	lngMax, lngI
	SearchArraysIntr = -1
	On Error Resume Next	
	lngMax = UBound(tmpIntrName)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpIntrName(lngI) = xname And tmpInst(lngI) = xinst And tmpLang(lngI) = xlang And tmpClass(lngI) = xclass Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchArraysIntr = lngI
End Function
Function CheckApp(tmpdate)
	CheckApp = "#FFFFFF"
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		'sqlReq = "SELECT appDate FROM request_T WHERE appDate = #" & tmpDate & "# ORDER BY appTimeFrom"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT appDate FROM request_T WHERE appDate = '" & tmpDate & "' AND IntrID = " & Session("UIntr") & " " & _
			"AND showintr = 1 AND NOT(STATUS = 2 OR STATUS = 3) ORDER BY appTimeFrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			CheckApp = "#FFFFCE"
		End If
		rsReq.Close
		Set rsReq = Nothing
	End If
	
End Function
Function GetReq(zzz)
	GetReq = "N/A"
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT Lname, Fname FROM requester_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		GetReq = rsRP("Lname") & ", " & rsRP("Fname")
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
Function GetReqHPID(zzz)
	GetReqHPID = "N/A"
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT ReqName FROM Appointment_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONNHP, 3, 1
	If Not rsRP.EOF Then
		GetReqHPID = Z_FixNull(rsRP("ReqName"))
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
Function GetReqHPID2(zzz)
	GetReqHPID2 = 0
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT HPID FROM Request_T WHERE [index] = " & zzz
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		GetReqHPID2 = Z_CZero(rsRP("HPID"))
	End If
	rsRP.Close
	Set rsRP = Nothing
End Function
'GET INSTITUTION's NAME
Function GetInst2(zzz)
	GetInst2 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		tmpIname = rsInst("Facility") 
		GetInst2 = tmpIname
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET INSTITUTION's ADDRESS
Function GetInst3(zzz)
	GetInst3 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Address, City, State, Zip FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInst3 = rsInst("Address") & ", "& rsInst("City") & ", " & rsInst("State") & ", " & rsInst("Zip")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET INSTITUTION's NAME
Function GetInst4(zzz)
	GetInst4 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Facility, Department FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		tmpIname = """" & rsInst("Facility") & """"
		If rsInst("Department") <> "" Then tmpIname = """" & rsInst("Facility") & " - " & rsInst("Department") & """"
		GetInst4 = tmpIname
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
'GET INSTITUTION's ADDRESS for CSV
Function GetInst5(zzz)
	GetInst5 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Address, City, State, Zip FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInst5 = """" & rsInst("Address") & ""","& rsInst("City") & "," & rsInst("State") & "," & rsInst("Zip")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetDept(xxx)
	GetDept = ""
	If xxx = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Dept FROM dept_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetDept = rsInst("Dept")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetMyDept(xxx)
	GetMyDept = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT Dept FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetMyDept = " - " & rsDept("Dept")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
Function GetDeptAdr(xxx)
	GetDeptAdr = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT Address, City, State, Zip FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetDeptAdr = rsDept("Address") & ", " & rsDept("City") & ", " & rsDept("State") & ", " & rsDept("Zip")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
Function GetInstDept(xxx)
	GetInstDept = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT InstID FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONN, 3, 1
	If Not rsDept.EOF Then
		GetInstDept = GetInst2(rsDept("InstID"))
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
'SEARCH STATS
Function SearchStats(xFac, xMonthYr, tmpFac, tmpMonthYr)
	DIM lngMax, lngI
	SearchStats = -1
	On Error Resume Next
	lngMax = UBound(tmpFac)
	If Err.Number <> 0 Then Exit Function
	For lngI = 0 to lngMax
		If tmpFac(lngI) = xFac And tmpMonthYr(lngI) = xMonthYr Then Exit For
	Next
	If lngI > lngMax Then Exit Function
	SearchStats = lngI
End Function
Function GetMisReason(xxx)
	GetMisReason = "N/A"
	Set rsMis = Server.CreateObject("ADODB.RecordSet")
	sqlMis = "SELECT reason FROM missed_T WHERE [index] = " & xxx
	rsMis.Open sqlMis, g_strCONN, 3, 1
	If Not rsMis.EOF Then
		GetMisReason = rsMis("reason")
	End If
	rsMis.Close
	Set rsMis = Nothing
End Function
Function GetCanReason(xxx)
	GetCanReason = "N/A"
	Set rsMis = Server.CreateObject("ADODB.RecordSet")
	sqlMis = "SELECT reason FROM cancel_T WHERE [index] = " & xxx
	rsMis.Open sqlMis, g_strCONN, 3, 1
	If Not rsMis.EOF Then
		GetCanReason = rsMis("reason")
	End If
	rsMis.Close
	Set rsMis = Nothing
End Function
Function Z_FormatTime(xxx, zzz)
	Z_FormatTime = ""
	If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, zzz)
End Function
%>
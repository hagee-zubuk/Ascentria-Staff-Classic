<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_UtilsMedicaid.asp" -->
<%
server.scripttimeout = 360000
Dim ArrSun(), ArrMon(), ArrTue(), ArrWed(), ArrThu(), ArrFri(), ArrSat()
Function CleanMeSingleQuote(xxx)
	' clean string
	CleanMeSingleQuote = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMeSingleQuote = Replace(trim(xxx), "'", "''")
End Function
Function GetSunday(xxx)
	'get sunday of given date
	If WeekDay(xxx) = 1 Then
		GetSunday = xxx
	ElseIf WeekDay(xxx) = 2 Then
		GetSunday = DateAdd("d", -1, xxx)
	ElseIf WeekDay(xxx) = 3 Then
		GetSunday = DateAdd("d", -2, xxx)
	ElseIf WeekDay(xxx) = 4 Then
		GetSunday = DateAdd("d", -3, xxx)
	ElseIf WeekDay(xxx) = 5 Then
		GetSunday = DateAdd("d", -4, xxx)
	ElseIf WeekDay(xxx) = 6 Then
		GetSunday = DateAdd("d", -5, xxx)
	ElseIf WeekDay(xxx) = 7 Then
		GetSunday = DateAdd("d", -6, xxx)
	End If	
End Function
Function GetDate(BasisDate, MyWeekDayName, tmpLastDay, appDate)
	tmpSunday = GetSunday(BasisDate)
	If UCase(MyWeekDayName) = "SUN" Then
		Do Until WeekDay(BasisDate) = 1
			BasisDate = DateAdd("d", -1, BasisDate)
		Loop
	End If
	If UCase(MyWeekDayName) = "MON" Then
		Do Until WeekDay(BasisDate) = 2
			BasisDate = DateAdd("d", -1, BasisDate)
		Loop
	End If
	If UCase(MyWeekDayName) = "TUE" Then
		Do Until WeekDay(BasisDate) = 3
			BasisDate = DateAdd("d", -1, BasisDate)
		Loop
	End If
	If UCase(MyWeekDayName) = "WED" Then
		Do Until WeekDay(BasisDate) = 4
			BasisDate = DateAdd("d", -1, BasisDate)
		Loop
	End If
	If UCase(MyWeekDayName) = "THU" Then
		Do Until WeekDay(BasisDate) = 5
			BasisDate = DateAdd("d", -1, BasisDate)
		Loop
	End If
	If UCase(MyWeekDayName) = "FRI" Then
		Do Until WeekDay(BasisDate) = 6
			BasisDate = DateAdd("d", -1, BasisDate)
		Loop
	End If
	If UCase(MyWeekDayName) = "SAT" Then
		Do Until WeekDay(BasisDate) = 7
			BasisDate = DateAdd("d", -1, BasisDate)
		Loop
	End If
	If CDate(appDate) <= CDate(BasisDate) Then
		If tmpLastDay <> "" Then 
			If CDate(appDate) <= CDate(BasisDate) And CDate(tmpLastDay) >= CDate(BasisDate) Then
				GetDate = BasisDate
			Else
				GetDate = ""
			End If
		Else
			GetDate = BasisDate
		End If
	Else
		GetDate = ""
	End If
End Function
Function CleanMe(xxx)
	' clean string
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, ",", "''")
End Function
Function LangName(xxx)
	'get dialect from langID
	If xxx = "" Then Exit Function
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT * FROM language_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONN, 1, 3
	If Not rsLang.EOF Then
		LangName = rsLang("language")
	End If
	rsLang.Close
	Set rsLAng = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function SalitaKo(strLang, IntrID)
	'check if interpreter can speak given dialect
	SalitaKo = False
	Set rsSalita = Server.CreateObject("ADODB.RecordSet")
	sqlSalita = "SELECT * FROM interpreter_T WHERE [index] = " & IntrID 
	rsSalita.Open sqlSalita, g_strCONN, 1, 3
	If Not rsSalita.EOF Then
		If UCase(Trim(rsSalita("Language1"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language2"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language3"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language4"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language5"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
	End If
	rsSalita.Close
	Set rsSalita = Nothing
End Function
Function GetHPLang(xxx)
	'get language from HP
	GetHPLang = -1
	Set rsHPLang  =Server.CreateObject("ADODB.RecordSet")
	sqlHPLang = "SELECT * FROM Lang_T WHERE LBID = " & xxx
	rsHPLang.Open sqlHPLang, g_strCONNHP, 3, 1
	If Not rsHPLang.EOF Then
		GetHPLang = rsHPLang("index")
	End If
	rsHPLang.Close
	Set rsHPLang = Nothing
End Function
Function GetIntr(xxx)
	'gets interpreter ID
	GetIntr = 0 
	Set rsHPI = Server.CreateObject("ADODB.RecordSet")
	sqlHPI = "SELECT * FROM Request_T WHERE [index] = " & xxx
	rsHPI.Open sqlHPI, g_strCONN, 3, 1
	If Not rsHPI.EOF Then
		GetIntr = rsHPI("intrID")
	End If
	rsHPI.CLose
	Set rsHPI = Nothing
End Function
Function GetEmailIntr(xxx)
	GetEmailIntr = ""
	Set rsIntrMail = Server.CreateObject("ADODB.RecordSet")
	sqlIntrMail = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsIntrMail.Open sqlIntrMail, g_strCONN, 3, 1
	If Not rsIntrMail.EOF Then
		GetEmailIntr = rsIntrMail("E-mail")
	End If
	rsIntrMail.Close
	Set rsIntrMail = Nothing
End Function
Function GetStatus(xxx)
	GetStatus = 0 
	Set rsHPI = Server.CreateObject("ADODB.RecordSet")
	sqlHPI = "SELECT * FROM Request_T WHERE [index] = " & xxx
	rsHPI.Open sqlHPI, g_strCONN, 3, 1
	If Not rsHPI.EOF Then
		GetStatus = rsHPI("status")
	End If
	rsHPI.CLose
	Set rsHPI = Nothing
End Function

If Request("ctrl") = 1 Then
	
ElseIf Request("ctrl") = 2 Then
	
ElseIf Request("ctrl") = 3 Then
	strNow = Now
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT * FROM request_T"
	If Request("txtFromd8") <> "" Then
		If IsDate(Request("txtFromd8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' "
			'tmpFromd8 = Request("txtFromd8") 
		Else
			Session("MSG") = "ERROR: Invalid Appointment Date Range (From)."
			'Response.Redirect "reqtable.asp"
		End If
	End If
	If Request("txtTod8") <> "" Then
		If IsDate(Request("txtTod8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate <= '" & Request("txtTod8") & "' "
			'tmpTod8 = Request("txtTod8")
		Else
			Session("MSG") = "ERROR: Invalid Appointment Date Range (To)."
			'Response.Redirect "reqtable.asp"
		End If
	End If
	If Request("txtFromd8") <> "" AND Request("txtTod8") <> "" Then
		If IsDate(Request("txtFromd8")) And IsDate(Request("txtTod8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' AND appDate <= '" & Request("txtTod8") & "' "
		Else
			Session("MSG") = "ERROR: Invalid Appointment Date Range)."
			'Response.Redirect "reqtable.asp"
		End If
	End If
	rsTBL.Open sqlTBL, g_strCONN, 1, 3 
	If Not rsTBL.EOF Then 
		y = Request("Hctr")
		For ctr = 1 To y - 1
			tmpID = Request("ID" & ctr)
			tmpIndex = "Index= " & tmpID
			rsTBL.MoveFirst
			rsTBL.Find(tmpIndex)
			If Not rsTBL.EOF Then
				If Request("txtstime" & ctr) <> "" Then
					If Not IsDate(Request("txtstime" & ctr)) Then
						Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Actual Start Time in Request ID " & tmpID & "."
					End If
				End If
				If Request("txtetime" & ctr) <> "" Then
					If Not IsDate(Request("txtetime" & ctr)) Then
						Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Actual End Time in Request ID " & tmpID & "."
					End If
				End If
				If Request("txtBilHrs" & ctr) <> "" Then
					If Not IsNumeric(Request("txtBilHrs" & ctr)) Then
						Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Billable Hours in Request ID " & tmpID & "."
					End If
				End If
				If Request("txtRate" & ctr) <> "" Then
					If Not IsNumeric(Request("txtRate" & ctr)) Then
						Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Institution Rate in Request ID " & tmpID & "."
					End If
				End If
				If Request("txtIntrRate" & ctr) <> "" Then
					If Not IsNumeric(Request("txtIntrRate" & ctr)) Then
						Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Interpreter Rate in Request ID " & tmpID & "."
					End If
				End If
				If Session("MSG") = "" Then
					rsTBL("AStarttime") = empty
					rsTBL("AEndtime") = empty
					if Request("txtstime" & ctr) <> "" Then rsTBL("AStarttime") = Z_DateNull(rsTBL("appDate") & " " & Request("txtstime" & ctr))
					if Request("txtetime" & ctr) <> "" Then rsTBL("AEndtime") = Z_DateNull(rsTBL("appDate") & " " & Request("txtetime" & ctr))
					If Request("txtetime" & ctr) <> "" Then
						'rsTBL("status") = 1
					End If
					tmpBill =  Z_FixNull(Request("txtBilHrs" & ctr))
					If tmpBill = "" THen
						rsTBL("Billable") = Null
					Else
						rsTBL("Billable") = Request("txtBilHrs" & ctr)
					End If
					
					rsTBL("InstRate") = Z_Czero(Request("selInstRate" & ctr))
					rsTBL("IntrRate") = Z_Czero(Request("txtIntrRate" & ctr))
					rsTBL("Comment") = Request("txtcom" & ctr)
					rsTBL("Verified") = False
					If Request("chkVer" & ctr) <> "" Then rsTBL("Verified") = True
					rsTBL("Paid") = False
					If Request("chkbil" & ctr) <> "" Then rsTBL("Paid") = True
					If Request("txtstime" & ctr) <> "" And Request("txtstime" & ctr) <> "" And Z_CZero(Request("txtBilHrs" & ctr)) <> 0 And Request("selInstRate" & ctr) <> 0 And Z_CZero(Request("txtIntrRate" & ctr)) <> 0 Then
						If rsTBL("Status") <> 1 And rsTBL("Status") <> 4 Then 
							rsTBL("Status") = 1
							rsTBL("completed") = strNow
						End If
					End If
					rsTBL.Update
					If SaveHist(tmpID, "reqtable.asp") Then
	
					End If
				End If
			End If
		Next
	End If
	rsTBL.Close
	Set rsTBL = Nothing
	Response.Redirect "reqtable.asp?radioAss=" & Request("radioAss") & "&radioStat=" & Request("radioStat") & "&txtFromd8=" & Request("txtFromd8") & "&txtTod8=" & Request("txtTod8") & _
		"&txtFromID=" & Request("txtFromID") & "&txtToID=" & Request("txtToID") & "&selInst=" & Request("selInst") & "&seldept=" & Request("seldept") & "&selLang=" & Request("selLang") & "&tmpclilname=" & Request("txtclilname") & "&tmpclifname=" & Request("txtclifname") & _
		"&selIntr=" & Request("selIntr") & "&selClass=" & Request("selClass") & "&selAdmin=" & Request("selAdmin") & "&action=3"
ElseIf Request("ctrl") = 4 Then
	tmpMonthYear = Split(Request("Hmonth"), " - ")
	tmpMonth = tmpMonthYear(0) & "/01/" & tmpMonthYear(1)
	If IsNumeric(tmpMonthYear(1)) Then
		If Request("dir") = 0 Then
			tmpMonth = DateAdd("m", -1, tmpMonth)
		Else
			tmpMonth = DateAdd("m", 1, tmpMonth)
		End If
	End If
	'Response.Redirect "calendarview.asp?selMonth=" & Month(tmpMonth) & "&txtyear=" & Year(tmpMonth)
	If Request("type") <> 1 Then
		Response.Redirect "calendarview2.asp?selMonth=" & Month(tmpMonth) & "&txtday=1&txtyear=" & Year(tmpMonth)
	Else
		Response.Redirect "calSNHMC.asp?selMonth=" & Month(tmpMonth) & "&txtday=1&txtyear=" & Year(tmpMonth)
	End If
ElseIf Request("ctrl") = 5 Then
	'STORE ENTRIES ON COOKIE FOR PRINTING
	Response.Cookies("LBREPORT") = Z_DoEncrypt(Request("selRep")	& "|" & Request("txtRepFrom") & "|" & _
		Request("txtRepTo")	& "|" & Request("selInst")	& "|" & Request("selIntr")	& "|" & Request("selTown") & _
		"|" & Request("selLang")	& "|" & Request("selCli")	& "|" & Request("selClass") & "|" & Request("chkAddnl") & "|" & Request("selIntrStat") & "|" & Request("txtZip") & "|" & Request("txtState"))
	If Request("txtRepFrom") <> "" Then
		If Not IsDate(Request("txtRepFrom")) Then
			Session("MSG") = "ERROR: Invalid timeframe (From:)."
		End If
	End If
	If Request("txtRepTo") <> "" Then
		If Not IsDate(Request("txtRepTo")) Then
			Session("MSG") = Session("MSG") & "<br>ERROR: Invalid timeframe (To:)."
		End If
	End If
	If Request("selRep") = 28 Or Request("selRep") = 67 Then
		If DateDiff("d", Request("txtRepFrom"), Request("txtRepTo")) <> 13 Then
			Session("MSG") = Session("MSG") & "<br>ERROR: Only 2 weeks (1 pay period) is allowed for timeframe for this report."
		End If
	End If
	If Request("selRep") = 71 Then
		If DateDiff("d", Request("txtRepFrom"), Request("txtRepTo")) <> 6 Then
			Session("MSG") = Session("MSG") & "<br>ERROR: Only 1 week is allowed for timeframe for this report."
		End If
	End If
	If Session("MSG") <> "" Then
		Response.Redirect "reports.asp?rep=0&sel=0"
	Else
		Response.redirect "reports.asp?rep=1&sel=0"
	End If
ElseIf Request("ctrl") = 6 Then 'ADMIN TOOLS
	tmpLang = CInt(Request("selLang"))
	tmpUser = CInt(Request("selUser"))
	'EDIT REQUESTING PERSON
	If Request("selReq") <> "-1" Then
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		tmpReq = Request("selReq")
		sqlReq = "SELECT * FROM requester_T WHERE [index] = " & Request("selReq")
		rsReq.Open sqlReq, g_strCONN, 1, 3
		If Not rsReq.EOF Then
			If (Request("txtReqLname") <> "" Or Request("txtReqFname") <> "") And (Request("txtphone") <> "" Or Request("txtfax") <> "" Or Request("txtemail") <> "") Then
				rsReq("Lname") = CleanMe(Request("txtReqLname"))
				rsReq("Fname") = CleanMe(Request("txtReqFname"))
				rsReq("Phone") = Request("txtphone")
				rsReq("pExt") = Request("txtReqExt")
				rsReq("Fax") = Request("txtfax")
				rsReq("Email") = Request("txtemail")
				rsReq.Update
				'CREATE LOG
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Requesting person (ID: " & Request("selReq") & ") was edited by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
			Else
				If Request("txtReqLname") = "" And Request("txtReqFname") = "" Then
					Session("MSG") = Session("MSG") & "Error: Requester's name cannot be blank."
				End If
				If Request("txtphone") = "" And Request("txtfax") = "" And Request("txtemail") = "" Then
					Session("MSG") = Session("MSG") & "<br>Error: Requesting person should at least have 1 contact information."
				End If
			End If
		End If
		rsReq.Close
		Set rsReq = Nothing
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM requester_T"
		rsReq.Open sqlReq, g_strCONN, 1, 3
		If (Request("txtReqLname") <> "" Or Request("txtReqFname") <> "") And (Request("txtphone") <> "" Or Request("txtfax") <> "" Or Request("txtemail") <> "") Then
				rsReq.AddNew
				rsReq("Lname") = CleanMe(Request("txtReqLname"))
				rsReq("Fname") = CleanMe(Request("txtReqFname"))
				rsReq("Phone") = Request("txtphone")
				rsReq("pExt") = Request("txtReqExt")
				rsReq("Fax") = Request("txtfax")
				rsReq("Email") = Request("txtemail")
				rsReq.Update
				tmpReq = rsReq("index")
				'CREATE LOG
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Requesting person (ID: " & tmpReq & ") was created  by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing	
				
		Else
			If Not(Request("txtReqLname") = "" And Request("txtReqFname") = "") Then
				'Session("MSG") = "Error: Requester's name cannot be blank."
					If Request("txtphone") = "" And Request("txtfax") = "" And Request("txtemail") = "" Then
					Session("MSG") = Session("MSG") & "<br>Error: Requesting person should at least have 1 contact information."
				End If
			End If
		End If
	End If
	'EDIT LANGUAGE
	If tmpLang <> -1 Then 
		'SAVE LANGUAGE
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang = "SELECT * FROM language_T WHERE [index] = " & Request("selLang")
		rsLang.Open sqlLang, g_strCONN, 1, 3
		If Not rsLang.EOF Then
			If Request("txtLang") <> "" Then
				rsLang("Language") = Request("txtLang")
				rsLang.Update
				'CREATE LOG
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Language (ID: " & Request("selReq") & ") was edited by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
				'ADD IN HP
				Set rsLangHP = Server.CreateObject("ADODB.RecordSet")
				sqlLangHP = "SELECT * FROM Lang_T WHERE LBID = " & Request("selLang")
				rsLangHP.Open sqlLangHP, g_strCONNHP, 1, 3
				If Not rsLangHP.EOF Then
					rsLangHP("Language") = Request("txtLang")
					rsLangHP.Update
				End If
				rsLangHP.Close
				Set rsLangHP = Nothing
			Else
				Session("MSG") = Session("MSG") & "<br>Error: Language cannot be blank."
			End If
		End If
		rsLang.Close
		Set rsLang = Nothing
	Else 
		If Request("txtLang") <> "" Then
			Set rsLang = Server.CreateObject("ADODB.RecordSet")
			sqlNewLang = "SELECT * FROM language_T WHERE UCase([Language]) = '" & UCase(Request("txtLang")) & "' "
			rsLang.Open sqlNewLang, g_strCONN, 1, 3
			If rsLang.EOF Then
				rsLang.AddNew
				rsLang("Language") = Request("txtLang")
				tmpLang = rsLang("Index")
				rsLang.Update
				'CREATE LOG
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "New language (ID: " & Request("tmpLang") & ") was created by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
				'ADD IN HP
				Set rsLangHP = Server.CreateObject("ADODB.RecordSet")
				sqlLangHP = "SELECT * FROM Lang_T"
				rsLangHP.Open sqlLangHP, g_strCONNHP, 1, 3
				rsLangHP.AddNew
				rsLangHP("Language") = Request("txtLang")
				rsLangHP("LBID") = tmpLang
			Else
				Session("MSG") = Session("MSG") & "<br>Error: Language already exists."
			End If
			rsLang.Close
			Set rsLang = Nothing
		End If
	End If
	'EDIT INSTITUITION
	If Request("selInst") <> -1 Then 
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "SELECT * FROM institution_T WHERE [index] = " & Request("selInst")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		If Not rsInst.EOF Then
			If Request("txtNewInst") <> "" Then
				rsInst("Facility") = Request("txtNewInst")
				'rsInst("Department") = Request("txtInstDept")
				'rsInst("Class") = Request("selClass")
				'rsInst("Address") = CleanMe(Request("txtInstAddr"))
				'rsInst("City") = Request("txtInstCity")
				'rsInst("State") = Request("txtInstState")
				'rsInst("Zip") = Request("txtInstZip")
				'rsInst("BAddress") = CleanMe(Request("txtBillAddr"))
				'rsInst("BCity") = Request("txtBillCity")
				'rsInst("BState") = Request("txtBillState")
				'rsInst("BZip") = Request("txtBillZip")
				rsInst.Update
				'CREATE LOG
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Institution (ID: " & Request("selInst") & ") was edited by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
			Else
				If Request("txtNewInst") = "" Then
					Session("MSG") = Session("MSG") & "<br>Error: Instituion's name cannot be blank."
				End If
			End If
		End If
		rsInst.Close
		Set rsInst = Nothing
		'EDIT DEPARTMENT
		If Request("selDept") <> 0 Then
			If Request("txtNewDept") <> "" Then 
				Set rsDept = Server.CreateObject("ADODB.RecordSet")
				sqlDept = "SELECT * FROM dept_T WHERE [index] = " & Request("selDept")
				rsDept.Open sqlDept, g_strCONN, 1, 3
				If Not rsDept.EOF Then
					tmpDeptID = Request("selDept")
					rsDept("dept") = Request("txtNewDept")
					'rsDept("Class") = Request("selClass")
					rsDept("Address") = CleanMe(Request("txtInstAddr"))
					rsDept("City") = Request("txtInstCity")
					rsDept("State") = Request("txtInstState")
					rsDept("Zip") = Request("txtInstZip")
					rsDept("InstAdrI") = CleanMe(Request("txtInstAddrI"))
					rsDept("Blname") = CleanMe(Request("txtBlname"))
					rsDept("BAddress") = CleanMe(Request("txtBillAddr"))
					rsDept("BCity") = Request("txtBillCity")
					rsDept("BState") = Request("txtBillState")
					rsDept("BZip") = Request("txtBillZip")
					rsDept.Update
					'CREATE LOG
					Set fso = CreateObject("Scripting.FileSystemObject")
					Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
					strLog = Now & vbTab & "Department (ID: " & Request("selDept") & ") was edited by " & Request.Cookies("LBUsrName") & "."
					LogMe.WriteLine strLog
					Set LogMe = Nothing
					Set fso = Nothing
				End If
				rsDept.Close
				Set rsDept = Nothing
			Else
				Session("MSG") = Session("MSG") & "<br>Error: Department's name cannot be blank."
			End If
		End If
		If Request("selInst") <> "-1" And Request("selDept") = 0 And  Request("txtNewDept") <> "" Then
			Set rsNewDept = Server.CreateObject("ADODB.RecordSet")
			sqlNewDept = "SELECT * FROM Dept_T WHERE UCase(dept) = '" &  Ucase(Trim(Request("txtNewDept"))) & "' AND InstID = " & Request("selInst")
			rsNewDept.Open sqlNewDept, g_strCONN, 1, 3
			If rsNewDept.EOF Then
				rsNewDept.AddNew
				tmpDeptID = rsNewDept("index")
				rsNewDept("dept") = Request("txtNewDept")
				rsNewDept("InstID") = Request("selInst")
				'rsNewDept("Class") = Request("selClass")
				rsNewDept("Address") = CleanMe(Request("txtInstAddr"))
				rsNewDept("City") = Request("txtInstCity")
				rsNewDept("State") = Request("txtInstState")
				rsNewDept("Zip") = Request("txtInstZip")
				rsNewDept("InstAdrI") = CleanMe(Request("txtInstAddrI"))
				rsNewDept("Blname") = CleanMe(Request("txtBlname"))
				rsNewDept("BAddress") = CleanMe(Request("txtBillAddr"))
				rsNewDept("BCity") = Request("txtBillCity")
				rsNewDept("BState") = Request("txtBillState")
				rsNewDept("BZip") = Request("txtBillZip")
				rsNewDept.Update
				'CREATE LOG
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Department (ID: " & tmpDeptID & ") was created by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
			Else
				Session("MSG") = Session("MSG") & "<br>Error: Department already exists."
			End If
			rsNewDept.Close
			Set rsNewDept = Nothing
		End If
	End If
	'EDIT INTERPRETER
	If Request("selIntr") <> "-1" Then 
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & Request("selIntr")
		rsIntr.Open sqlIntr, g_strCONN, 1, 3
		If Not rsIntr.EOF Then
			tmpIntr =  Request("selIntr")
			If Request("txtIntrLname") <> "" And Request("txtIntrFname") <> "" Then
				rsIntr("First Name") = CleanMe(Request("txtIntrFname"))
				rsIntr("Last Name") = CleanMe(Request("txtIntrLname"))
				rsIntr("E-mail") = Request("txtIntrEmail")
				rsIntr("Phone1") = Request("txtIntrP1")
				rsIntr("P1Ext") = Request("txtIntrExt")
				rsIntr("Phone2") = Request("txtIntrP2")
				rsIntr("Fax") = Request("txtIntrFax")
				rsIntr("Address1") = CleanMe(Request("txtIntrAddr"))
				rsIntr("City") = Request("txtIntrCity")
				rsIntr("State") = Request("txtIntrState")
				rsIntr("Zip Code") = Request("txtIntrZip")
				rsIntr("IntrAdrI") = CleanMe(Request("txtIntrAddrI"))
				'rsIntr("IntrAdrI") = Request("txtHire")
				rsIntr("InHouse") = False
				If Request("chkInHouse") <> "" Then rsIntr("InHouse") = True
				rsIntr("Stat") = Request("radioStatIntr")
				rsIntr("Rate") = Request("selIntrRate")
				rsIntr("Crime") = False
				If Request("chkCrim") <> "" Then rsIntr("crime") = True
				rsIntr("drive") = False
				If Request("chkdriv") <> "" Then rsIntr("drive") = True
				rsIntr("train") = Request("txttrain")
				rsIntr("Active") = False
				rsIntr("datehired") = Empty
				If Request("txthire") <> "" Then
        	If isDate(Request("txthire")) THen rsIntr("datehired") = Request("txthire")
        End If
        rsIntr("dateterm") = Empty
        If Request("txtterm") <> "" Then
        	If isDate(Request("txtterm")) THen rsIntr("dateterm") = Request("txtterm")
        End If
				If Request("radioStatIntr1") = 0 Then rsIntr("Active") = True
				rsIntr("Comments") = Request("txtIntrCom")	
				If Request("SelIntrLang") <> "0" Then 'SAVE LANGUAGES OF INTERPRETER
					If rsIntr("Language1") = "" Or IsNull(rsIntr("Language1")) Then 
						rsIntr("Language1") = Request("SelIntrLang")
					Else
						If rsIntr("Language2") = ""  Or IsNull(rsIntr("Language2")) Then
							rsIntr("Language2") = Request("SelIntrLang")
						Else
							If rsIntr("Language3") = ""  Or IsNull(rsIntr("Language3")) Then
								rsIntr("Language3") = Request("SelIntrLang")
							Else
								If rsIntr("Language4") = "" Or IsNull(rsIntr("Language4")) Then
									rsIntr("Language4") = Request("SelIntrLang")
								Else
									If rsIntr("Language5") = "" Or IsNull(rsIntr("Language5")) Then rsIntr("Language5") = Request("SelIntrLang")
								End If
							End If
						End If 	
					End If
				End If
				'DELETE LANGUAGES OF INTERPRETER
				If Request("chkLang1") <> "" Then  rsIntr("Language1") = ""
				If Request("chkLang2") <> "" Then  rsIntr("Language2") = ""
				If Request("chkLang3") <> "" Then  rsIntr("Language3") = ""
				If Request("chkLang4") <> "" Then  rsIntr("Language4") = ""
				If Request("chkLang5") <> "" Then  rsIntr("Language5") = ""
				'CREATE LOG
			on error resume next
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "Interpreter (ID: " & Request("selIntr") & ") was edited by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
				rsIntr.Update
			Else
				Session("MSG") = Session("MSG") & "<br>Error: Interpreter's name cannot be blank."
			End If
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	Else
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM interpreter_T"
		rsIntr.Open sqlIntr, g_strCONN, 1, 3
		If (Request("txtIntrLname") <> "" Or Request("txtIntrFname") <> "") Then
			rsIntr.AddNew
			rsIntr("First Name") = CleanMe(Request("txtIntrFname"))
			rsIntr("Last Name") = CleanMe(Request("txtIntrLname"))
			rsIntr("E-mail") = Request("txtIntrEmail")
			rsIntr("Phone1") = Request("txtIntrP1")
			rsIntr("P1Ext") = Request("txtIntrExt")
			rsIntr("Phone2") = Request("txtIntrP2")
			rsIntr("Fax") = Request("txtIntrFax")
			rsIntr("Address1") = CleanMe(Request("txtIntrAddr"))
			rsIntr("City") = Request("txtIntrCity")
			rsIntr("State") = Request("txtIntrState")
			rsIntr("Zip Code") = Request("txtIntrZip")
			rsIntr("IntrAdrI") = CleanMe(Request("txtIntrAddrI"))
			rsIntr("InHouse") = False
			If Request("chkInHouse") <> "" Then rsIntr("InHouse") = True
			rsIntr("Stat") = Request("radioStatIntr")
			rsIntr("Rate") = Request("selIntrRate")
			rsIntr("Language1") = Request("SelIntrLang")
			rsIntr("Crime") = False
			If Request("chkCrim") <> "" Then rsIntr("crime") = True
			rsIntr("drive") = False
			If Request("chkdriv") <> "" Then rsIntr("drive") = True
			rsIntr("train") = Request("txttrain")
			rsIntr("Active") = False
			If Request("radioStatIntr1") = 0 Then rsIntr("Active") = True
			rsIntr.Update
			tmpIntr = rsIntr("index")
			'CREATE LOG
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
			strLog = Now & vbTab & "Interpreter (ID: " & tmpIntr & ") was Created  by " & Request.Cookies("LBUsrName") & "."
			LogMe.WriteLine strLog
			Set LogMe = Nothing
			Set fso = Nothing
		Else
			If Request("txtIntrLname") = "" And Request("txtIntrFname") = "" Then
				'Session("MSG") = "Error: Interpreter's name cannot be blank."
			End If
		End If
	End If
	'EDIT USER
	If tmpUser <> -1 Then 
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM user_T WHERE [index] = " & Request("selUser")
		rsUser.Open sqlUser, g_strCONN, 1, 3
		If Not rsUser.EOF Then
			If Request("txtUserUname") <> "" And Request("txtUserPword") <> "" Then
				intrassign = false
				if Request("selType") = 2 And (Request("selIntr2") <> Request("hidintr")) then
					set rsintr = server.createobject("adodb.recordset")
					sqlintr = "select * from user_t where intrid = " & Request("selIntr2")
					rsintr.open sqlintr, g_strconn, 3, 1
					if not rsintr.eof then intrassign = true	
					rsintr.close
					set rsintr = nothing
				end if
				if intrassign = false then
					rsUser("Fname") = Request("txtUserFname")
					rsUser("Lname") = Request("txtUserLname")
					rsUser("username") = Request("txtUserUname")
					rsUser("password") = Z_DoEncrypt(Request("txtUserPword"))
					rsUser("Type") = Request("selType")
					tmpInst2 = Request("selIntr2")
					If tmpInst2 = -1 Then tmpInst2 = 0
					rsUser("IntrID") = tmpInst2
					rsUser.Update
					'CREATE LOG
					Set fso = CreateObject("Scripting.FileSystemObject")
					Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
					strLog = Now & vbTab & "User (ID: " & Request("selUser") & ") was edited by " & Request.Cookies("LBUsrName") & "."
					LogMe.WriteLine strLog
					Set LogMe = Nothing
					Set fso = Nothing
				else
					Session("MSG") = Session("MSG") & "<br>Error: Interpreter is already assigned for a user."
				end if
			Else
				If Request("txtUserUname") = "" Then
					Session("MSG") = Session("MSG") & "<br>Error: Username cannot be blank."
				ElseIf Request("txtUserPword") = "" Then
					Session("MSG") = Session("MSG") & "<br>Error: Password cannot be blank."
				End If
			End If
		End If
		rsUser.Close
		Set rsUser = Nothing
	Else
	on error resume next
		If Request("txtUserUname") <> "" And Request("txtUserPword") <> "" Then
			intrassign = false
			if Request("selType") = 2 then
				set rsintr = server.createobject("adodb.recordset")
				sqlintr = "select * from user_t where intrid = " & Request("selIntr2")
				rsintr.open sqlintr, g_strconn, 3, 1
				if not rsintr.eof then intrassign = true	
				rsintr.close
				set rsintr = nothing
			end if
			if intrassign = false then
				Set rsUser = Server.CreateObject("ADODB.RecordSet")
				sqlUser = "SELECT * FROM user_T"
				rsUser.Open sqlUser, g_strCONN, 1, 3
				rsUser.AddNew
				rsUser("Fname") = Request("txtUserFname")
				rsUser("Lname") = Request("txtUserLname")
				rsUser("username") = Request("txtUserUname")
				rsUser("password") = Z_DoEncrypt(Request("txtUserPword"))
				rsUser("Type") = Request("selType")
				tmpInst2 = Request("selIntr2")
				If tmpInst2 = -1 Then tmpInst2 = 0
				rsUser("IntrID") = tmpInst2
				tmpUser = rsUser("Index")
				rsUser.Update
				'CREATE LOG
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
				strLog = Now & vbTab & "New user (ID: " & Request("tmpUser") & ") was created by " & Request.Cookies("LBUsrName") & "."
				LogMe.WriteLine strLog
				Set LogMe = Nothing
				Set fso = Nothing
			else
				Session("MSG") = Session("MSG") & "<br>Error: Interpreter is already assigned for a user."
			end if
		End If
	End If
	'EDIT INSTITUION RATE
	tmpRate = Request("SelRate")
	If Request("SelRate") <> 0 Then
		Set rsCancel = Server.CreateObject("ADODB.RecordSet")
		sqlCancel = "UPDATE rate_T SET rate = '" & Request("txtRate") & "' WHERE Rate = " & Request("selRate")
		rsCancel.Open sqlCancel, g_strCONN, 1, 3
		Set rsCancel = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Institution Rate (ID: " & Request("selRate") & ") was edited by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
		Set fso = Nothing
	ElseIf Request("SelRate") = 0 And Request("txtRate") <> "" Then
		Set rsRate = Server.CreateObject("ADODB.RecordSet")
		sqlRate = "SELECT * FROM rate_T WHERE Rate = " &  Request("txtRate")
		rsRate.Open sqlRate, g_strCONN, 1, 3
		If rsRate.EOF Then
			rsRate.AddNew
			rsRate("Rate") = Request("txtRate")
			tmpRate = rsRate("Rate")
			rsRate.Update
			'CREATE LOG
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
			strLog = Now & vbTab & "New Institution rate (ID: " & Request("tmpRate") & ") was created by " & Request.Cookies("LBUsrName") & "."
			LogMe.WriteLine strLog
			Set LogMe = Nothing
			Set fso = Nothing
		Else
			tmpRate = rsRate("Rate")
			Session("MSG") = Session("MSG") & "<br>Error: Institution Rate already exists."
		End If
		rsRate.Close
		Set rsRate = Nothing
	End If
	'EDIT INTERPRETER RATE
	tmpRate2 = Request("SelRate2")
	If Request("SelRate2") <> 0 Then
		Set rsCancel = Server.CreateObject("ADODB.RecordSet")
		sqlCancel = "UPDATE rate2_T SET rate2 = '" & Request("txtRate2") & "' WHERE Rate2 = " & Request("selRate2")
		rsCancel.Open sqlCancel, g_strCONN, 1, 3
		Set rsCancel = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Institution Rate (ID: " & Request("selRate2") & ") was edited by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
		Set fso = Nothing
	ElseIf Request("SelRate2") = 0 And Request("txtRate2") <> "" Then
		Set rsRate = Server.CreateObject("ADODB.RecordSet")
		sqlRate = "SELECT * FROM rate2_T WHERE Rate2 = " &  Request("txtRate2")
		rsRate.Open sqlRate, g_strCONN, 1, 3
		If rsRate.EOF Then
			rsRate.AddNew
			rsRate("Rate2") = Request("txtRate2")
			tmpRate2 = rsRate("Rate2")
			rsRate.Update
			'CREATE LOG
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
			strLog = Now & vbTab & "New Institution rate (ID: " & Request("tmpRate2") & ") was created by " & Request.Cookies("LBUsrName") & "."
			LogMe.WriteLine strLog
			Set LogMe = Nothing
			Set fso = Nothing
		Else
			tmpRate2 = rsRate("Rate2")
			Session("MSG") = Session("MSG") & "<br>Error: Institution Rate already exists."
		End If
		rsRate.Close
		Set rsRate = Nothing
	End If
	'EDIT REASON
	tmpReason = Request("SelCancel")
	If Request("SelCancel") <> 0 Then
		Set rsCancel = Server.CreateObject("ADODB.RecordSet")
		sqlCancel = "UPDATE cancel_T SET reason = '" & Request("txtCancel") & "' WHERE [index] = " & Request("selCancel")
		rsCancel.Open sqlCancel, g_strCONN, 1, 3
		Set rsCancel = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Cancel reason (ID: " & Request("selCancel") & ") was edited by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
		Set fso = Nothing
	ElseIf Request("SelCancel") = 0 And Request("txtCancel") <> "" Then
		Set rsCancel = Server.CreateObject("ADODB.RecordSet")
		sqlCancel = "SELECT * FROM cancel_T" 
		rsCancel.Open sqlCancel, g_strCONN, 1, 3
		rsCancel.AddNew
		rsCancel("Reason") = Request("txtCancel")
		tmpReason = rsCancel("Index")
		rsCancel.Update
		rsCancel.Close
		Set rsCancel = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "New cancel reason (ID: " & Request("tmpReason") & ") was created by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	tmpReason1 = Request("SelMissed")
	If Request("SelMissed") <> 0 Then
		Set rsMissed = Server.CreateObject("ADODB.RecordSet")
		sqlMissed = "UPDATE missed_T SET reason = '" & Request("txtMissed") & "' WHERE [index] = " & Request("SelMissed")
		rsMissed.Open sqlMissed, g_strCONN, 1, 3
		Set rsMissed = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Missed reason (ID: " & Request("SelMissed") & ") was edited by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
		Set fso = Nothing
	ElseIf Request("SelMissed") = 0 And Request("txtMissed") <> "" Then
		Set rsMissed = Server.CreateObject("ADODB.RecordSet")
		sqlMissed = "SELECT * FROM Missed_T" 
		rsMissed.Open sqlMissed, g_strCONN, 1, 3
		rsMissed.AddNew
		rsMissed("Reason") = Request("txtMissed")
		tmpReason1 = rsMissed("Index")
		rsMissed.Update
		rsMissed.Close
		Set rsMissed = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "New missed reason (ID: " & Request("tmpReason1") & ") was created by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	'EDIT MILEAGE RATE interpreter
	Set rsTMRate = Server.CreateObject("ADODB.RecordSet")
	sqlTMRate = "SELECT * FROM MileageRate_T"
	rsTMRate.Open sqlTMRate, g_strCONN, 1, 3
	If Request("txtMR") = "" Or Not IsNumeric(Request("txtMR")) Then 
		rsTMRate("mileageRate") = 0
	Else
		rsTMRate("mileageRate") = Request("txtMR")
	End If
	rsTMRate.Update
	rsTMRate.Close
	Set rsTMRate = Nothing
	'EDIT MILEAGE CAP interpreter
	Set rsTMRate = Server.CreateObject("ADODB.RecordSet")
	sqlTMRate = "SELECT * FROM travel_T"
	rsTMRate.Open sqlTMRate, g_strCONN, 1, 3
	If Request("txtmile") = "" Or Not IsNumeric(Request("txtmile")) Then 
		rsTMRate("milediff") = 0
	Else
		rsTMRate("milediff") = Request("txtmile")
	End If
	rsTMRate.Update
	rsTMRate.Close
	Set rsTMRate = Nothing
	'EDIT MILEAGE CAP institution
	Set rsTMRate = Server.CreateObject("ADODB.RecordSet")
	sqlTMRate = "SELECT * FROM travelinst_T"
	rsTMRate.Open sqlTMRate, g_strCONN, 1, 3
	If Request("txtmileinst") = "" Or Not IsNumeric(Request("txtmileinst")) Then 
		rsTMRate("milediffinst") = 0
	Else
		rsTMRate("milediffinst") = Request("txtmileinst")
	End If
	rsTMRate.Update
	rsTMRate.Close
	Set rsTMRate = Nothing
	'EDIT MILEAGE CAP institution - court
	Set rsTMRate = Server.CreateObject("ADODB.RecordSet")
	sqlTMRate = "SELECT * FROM travelinstcourt_T"
	rsTMRate.Open sqlTMRate, g_strCONN, 1, 3
	If Request("txtmilecourt") = "" Or Not IsNumeric(Request("txtmilecourt")) Then 
		rsTMRate("milediffcourt") = 0
	Else
		rsTMRate("milediffcourt") = Request("txtmilecourt")
	End If
	rsTMRate.Update
	rsTMRate.Close
	Set rsTMRate = Nothing
	'EDIT EMERGENCY RATE
	Set rsTMRate = Server.CreateObject("ADODB.RecordSet")
	sqlTMRate = "SELECT * FROM EmergencyFee_T"
	rsTMRate.Open sqlTMRate, g_strCONN, 1, 3
	If Request("txtFeel") = "" Or Not IsNumeric(Request("txtFeel")) Then 
		rsTMRate("FeeLegal") = 0
	Else
		rsTMRate("FeeLegal") = Request("txtFeel")
	End If
	If Request("txtFeeO") = "" Or Not IsNumeric(Request("txtFeeO")) Then 
		rsTMRate("FeeOther") = 0
	Else
		rsTMRate("FeeOther") = Request("txtFeeO")
	End If
	rsTMRate.Update
	rsTMRate.Close
	Set rsTMRate = Nothing
	Response.redirect "admintools.asp?ReqID=" & tmpReq & "&LangID=" & tmpLang & "&InstID=" & Request("selInst") & _
		"&IntrID=" & tmpIntr & "&UserID=" & tmpUser & "&ReasonID=" & tmpReason & "&Reason1ID=" & tmpReason1 & "&RateID=" & tmpRate & _
		"&DeptID=" & tmpDeptID & "&RateID2=" & tmpRate2
ElseIf Request("ctrl") = 7 Then
On error resume next
	If Request("selReq") <> -1 Then
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "DELETE FROM requester_T WHERE [index] = " & Request("selReq")
		rsReq.Open sqlReq, g_strCONN, 1, 3
		Set rsReq = Nothing
		'DELETE REALATIONSHIP WITH DEPARTMENT
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "DELETE FROM reqdept_T WHERE ReqID = " & Request("selReq")
		rsReq.Open sqlReq, g_strCONN, 1, 3
		Set rsReq = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Requesting person (ID: " & Request("selReq") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	If Request("selLang") <> -1 Then
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang = "DELETE FROM language_T WHERE [index] = " & Request("selLang")
		rsLang.Open sqlLang, g_strCONN, 1, 3
		Set rsLang = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Language (ID: " & Request("selLang") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	'DELETE INSTITUTION AND ALL INFO
	If Request("selInst") <> "-1" And Request("selDept") = "0" Then
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "DELETE FROM institution_T WHERE [index] = " & Request("selInst")
		rsInst.Open sqlInst, g_strCONN, 1, 3
		Set rsInst = Nothing
		Set rsDept = Server.CreateObject("ADODB.RecordSet")
		sqlDept = "SELECT * FROM dept_T WHERE InstID = " & Request("selInst")
		rsDept.Open sqlDept, g_strCONN, 1, 3
		If Not rsDept.EOF Then
			IDDept = rsDept("index")
			Do Until rsDept.EOF
				rsDept.Delete
				rsDept.Update
				rsDept.MoveNext
			Loop
		End If
		rsDept.Close
		Set rsDept = Nothing
		If (Not IsNull(IDDept)) Or IDDept <> 0 Then
			'DELETE REALATIONSHIP WITH DEPARTMENT
			Set rsReq = Server.CreateObject("ADODB.RecordSet")
			sqlReq = "DELETE FROM reqdept_T WHERE DeptID = " & IDDept
			rsReq.Open sqlReq, g_strCONN, 1, 3
			Set rsReq = Nothing
		End If
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Institution (ID: " & Request("selInst") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	'DELETE DEPT ONLY OF INSTITUTION
	If Request("selInst") <> "-1" And Request("selDept") <> "0" Then
		Set rsDept = Server.CreateObject("ADODB.RecordSet")
		sqlDept = "DELETE FROM dept_T WHERE [index] = " & Request("selDept")
		rsDept.Open sqlDept, g_strCONN, 1, 3
		Set rsDept = Nothing
		'DELETE REALATIONSHIP WITH DEPARTMENT
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "DELETE FROM reqdept_T WHERE DeptID = " & Request("selDept")
		rsReq.Open sqlReq, g_strCONN, 1, 3
		Set rsReq = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Department (ID: " & Request("selDept") & ") of Institution (ID: " & Request("selInst") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	If Request("selIntr") <> -1 Then
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "DELETE FROM interpreter_T WHERE [index] = " & Request("selIntr")
		rsIntr.Open sqlIntr, g_strCONN, 1, 3
		Set rsIntr = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Interpreter (ID: " & Request("selIntr") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	If Request("selCancel") <> 0 Then
		Set rsCancel = Server.CreateObject("ADODB.RecordSet")
		sqlCancel = "DELETE FROM cancel_T WHERE [index] = " & Request("selCancel")
		rsCancel.Open sqlCancel, g_strCONN, 1, 3
		Set rsCancel = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Cancel reason (ID: " & Request("selCancel") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	If Request("selMissed") <> 0 Then
		Set rsMissed = Server.CreateObject("ADODB.RecordSet")
		sqlMissed = "DELETE FROM Missed_T WHERE [index] = " & Request("selMissed")
		rsMissed.Open sqlMissed, g_strCONN, 1, 3
		Set rsMissed = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Missed reason(ID: " & Request("selMissed") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	If Request("selRate") <> 0 Then
		Set rsMissed = Server.CreateObject("ADODB.RecordSet")
		sqlMissed = "DELETE FROM rate_T  WHERE rate = " & Request("selRate")
		rsMissed.Open sqlMissed, g_strCONN, 1, 3
		Set rsMissed = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Institution Rate (ID: " & Request("selRate") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	If Request("selRate2") <> 0 Then
		Set rsMissed = Server.CreateObject("ADODB.RecordSet")
		sqlMissed = "DELETE FROM rate2_T  WHERE rate2 = " & Request("selRate2")
		rsMissed.Open sqlMissed, g_strCONN, 1, 3
		Set rsMissed = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "Interpreter Rate (ID: " & Request("selRate") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	If Request("selUser") <> -1 Then
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "DELETE FROM user_T WHERE [index] = " & Request("selUser")
		rsUser.Open sqlUser, g_strCONN, 1, 3
		Set rsUser = Nothing
		'CREATE LOG
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
		strLog = Now & vbTab & "User (ID: " & Request("selUser") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
		LogMe.WriteLine strLog
		Set LogMe = Nothing
	End If
	Response.redirect "admintools.asp"
ElseIf Request("ctrl") = 8 Then
	Response.Cookies("LBREPORT") = Z_DoEncrypt(99)
'	Response.Cookies("LBREPORT").Expires = Now + 0.34
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		response.redirect "calendarview2.asp?rep=1&tmpM='" & Request("Hmonth") & "' "
	Else
		response.redirect "calendarview2.asp?rep=1&tmpM='" & Request("Hmonth") & "'&tmpRP=" & Request.Cookies("LBUSERTYPE")
	End If
ElseIf Request("ctrl") = 9 Then
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT * FROM request_T WHERE [index] = " & Request("ReqID")
	rsReq.Open sqlReq, g_strCONN, 1, 3
	If Not rsReq.EOF Then
		tmpHPID = rsReq("HPID")
		rsReq.Delete
		rsReq.Update
	End If
	rsReq.Close
	Set rsReq = Nothing
	'DELETE IN HP
	If Z_CZero(tmpHPID) <> 0 Then
		Set rsHP  = Server.CreateObject("ADODB.RecordSet")
		sqlHP = "DELETE  FROM Appointment_T WHERE [index] = " & tmpHPID
		rsHp.Open sqlHP, g_strCONNHP, 1, 3
		Set rsHp = Nothing
	End If
	'DELETE IN INTR DB
	Set rsHP  = Server.CreateObject("ADODB.RecordSet")
	sqlHP = "DELETE  FROM appt_T WHERE appID = " & Request("ReqID")
	rsHp.Open sqlHP, g_strCONN, 1, 3
	Set rsHp = Nothing
	'CREATE LOG
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	If  Z_CZero(tmpHPID) = 0 Then
		strLog = Now & vbTab & "Request (ID: " & Request("ReqID") & ") was deleted by " & Request.Cookies("LBUsrName") & "."
	Else
		strLog = Now & vbTab & "Request (ID: " & Request("ReqID") & " -- HP ID: " & tmpHPID & ") was deleted by " & Request.Cookies("LBUsrName") & "."
	End If
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Session("MSG") = "Request deleted. Request ID: " & Request("ReqID")
	Response.Redirect "calendarview2.asp?appdate="& date
ElseIf Request("ctrl") = 10 Then
	'STORE ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBBILL") = Z_DoEncrypt(Request("HID")	& "|" & Request("radioStat")	& "|" & Request("selCancel") & "|" & _
		Request("selMissed")	& "|" & Request("chkPaid") & "|" & Request("txtBilTInst") & "|" & Request("txtBilTIntr") & "|" & _
		Request("txtBilMInst") & "|" & Request("txtBilMIntr") & "|" & Request("txtActTFrom")	& "|" & Request("txtActTTo") & "|" & _
		Request("txtBilHrs") & "|" & Request("hidInstRate")	& "|" & Request("hidIntrRate") & "|" & Request("chkBillInst")	& "|" & _
		Request("txtTTRate") & "|" & Request("txtMRate") & "|" & Request("chkBilTIntr") & "|" & Request("chkBilTInst") & "|" &_
		Request("chkBilMIntr") & "|" & Request("chkBilMIns") & "|" & Request("txtCombil") & "|" & Request("chkEmer") & "|" &  _
		Request("chkEmerFee") & "|" & Request("txtDecTT") & "|" & Request("txtDecMile") & "|" & Request("chkTollCon"))
	'CHECK FOR VALID VALUES
	If Request("txtBilHrs") <> "" Then
		If Not IsNumeric(Request("txtBilHrs")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Billable Hours."
	End If
	If Request("txtActdate") <> "" Then
		If Not IsDate(Request("txtActdate")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Actual date."
	End If
	If Request("txtActTFrom") <> "" Then
		If Not IsDate(Request("txtActTFrom")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Actual Time (From:)."
	End If
	If Request("txtActTTo") <> "" Then
		If Not IsDate(Request("txtActTTo")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Actual Time (To:)."
	End If
	If Request("txtBilTInst") <> "" Then
		If Not IsNumeric(Request("txtBilTInst")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Travel Time (Institution)."
	End If
	If Request("txtBilTIntr") <> "" Then
		If Not IsNumeric(Request("txtBilTIntr")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Travel Time (Interpreter)."
	End If
	If Request("txtBilMInst") <> "" Then
		If Not IsNumeric(Request("txtBilMInst")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Mileage (Institution)."
	End If
	If Request("txtBilMIntr") <> "" Then
		If Not IsNumeric(Request("txtBilMIntr")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Mileage (Interpreter)."
	End If
	If Request("txtTTRate") <> "" Then
		If Not IsNumeric(Request("txtTTRate")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Travel Time Rate."
	End If
	If Request("txtMRate") <> "" Then
		If Not IsNumeric(Request("txtMRate")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Mileage Rate."
	End If
	If Session("MSG") = "" Then	
		'GET COOKIE OF REQUEST
		MyStatNum = GetStatNum(Request("ReqID"))
		tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBBILL")), "|")
		Set rsBill = Server.CreateObject("ADODB.RecordSet")
		sqlBill = "SELECT * FROM request_T WHERE [index] = " & Request("ReqID")
		tmpAppDate = GetAppDate(Request("ReqID"))
		tmpstat = GetStatNum(Request("ReqID"))
		rsBill.Open sqlBill, g_strCONN, 1, 3
		If Not rsBill.EOF Then
			rsBill("Status") = tmpEntry(1)
			If tmpEntry(1) = 3 Or tmpEntry(1) = 4 Then
				If tmpEntry(1) = 4 Then
					syscomstr = Z_fixNull(rsBill("syscom"))
					syscomstr = Replace(syscomstr, "<br>This appointment was cancelled with less than 24 hours.", "")
					rsBill("syscom") = syscomstr & "<br>This appointment was cancelled with less than 24 hours."  
				End If
				rsBill("Cancel") = tmpEntry(2)
				rsBill("Missed") = 0
				If MyStatNum <> 3 And MyStatNum <> 4 Then
					TimeNow = Now 'SAVE IN HISTORY WHEN CANCELED
					Set rsHist = Server.CreateObject("ADODB.RecordSet")
					sqlHist = "SELECT * FROM History_T WHERE ReqID = " & Request("ReqID")
					rsHist.Open sqlHist, g_strCONNHist, 1,3 
					If Not rsHist.EOF Then
						rsHist("cancelTS") = TimeNow
						rsHist("cancelU") = Request.Cookies("LBUsrName")
						rsHist.UPdate
					End If
					rsHist.CLose
					Set rsHist = Nothing
				End if
			Else
				rsBill("Cancel") = 0
			End If
			'response.write tmpEntry(1) & " " & tmpEntry(3)
			If tmpEntry(1) = 2 Then 
				rsBill("Missed") = tmpEntry(3)
				rsBill("Cancel") = 0
			Else
				rsBill("Missed") = 0
			End If
			If tmpEntry(4) <> "" Then rsBill("Paid") = True
				tmpbill = Z_FixNull(tmpEntry(11))
				if tmpbill = "" Then 
					rsBill("Billable") = Null 'Z_FixNull
				else
					rsBill("Billable") = Z_Czero(tmpEntry(11))
				end if
			'rsMain("adate") = Z_DateNull(tmpEntry(23))
			
			If Z_FixNull(tmpEntry(9)) <> "" And Z_FixNull(tmpEntry(10)) <> "" Then
				date1st = Date & " " & cdate(tmpEntry(9))
				date2nd = Date & " " & cdate(tmpEntry(10))
				
				if datediff("n", date1st, date2nd) >= 0 then
					minTime = DateDiff("n", date1st, date2nd)
				else
					minTime = DateDiff("n", date1st, dateadd("d", 1, date2nd))
				end If
				rsBill("totalhrs") = MakeTime(Z_CZero(minTime))
			End If
			If Z_FixNull(tmpEntry(9)) <> "" Then
				rsBill("astarttime") = tmpAppDate & " " & Z_DateNull(tmpEntry(9))
			Else
				rsBill("astarttime") = empty
			End If
			If Z_FixNull(tmpEntry(10)) <> "" Then
				rsBill("aendtime") = tmpAppDate & " " & Z_DateNull(tmpEntry(10))
			Else
				rsBill("aendtime") = empty
			End If
			rsBill("TT_Inst") = Z_CZero(tmpEntry(5))
			rsBill("TT_Intr") = Z_CZero(tmpEntry(6))
			rsBill("M_Inst") = Z_CZero(tmpEntry(7))
			rsBill("M_Intr") = Z_CZero(tmpEntry(8))
			
			rsBill("emerFEE") = False
			rsBill("Emergency") = False
			If Request("chkEmer") <> "" Then rsBill("Emergency") = True 
			If Request("chkEmerFee") <> "" Then rsBill("emerFEE") = True
			
			'save actual TT and Mil
			rsBill("actTT") = Z_CZero(Request("txtDecTT"))
			rsBill("actMil") = Z_CZero(Request("txtDecMile"))
			'
			rsBill("overTTIntr") = false
			If tmpEntry(17) <> "" Then rsBill("overTTIntr") = True
			rsBill("overTTInst") = false
			If tmpEntry(18) <> "" Then rsBill("overTTInst") = True
			rsBill("overMIntr") = false
			If tmpEntry(19) <> "" Then rsBill("overMIntr") = True
			rsBill("overMInst") = false
			If tmpEntry(20) <> "" Then rsBill("overMInst") = True
			
			rsBill("BillInst") = False
			rsBill("TTRate") = 0
			rsBill("MRate") = 0
			If tmpEntry(14) <> "" Then 
				rsBill("BillInst") = True 
				rsBill("TTRate") = Z_CZero(tmpEntry(15))
				rsBill("MRate") = Z_CZero(tmpEntry(16))
			End If
			
			If tmpEntry(9) <> "" And tmpEntry(10) <> "" And (tmpEntry(11) <> "") Then 'CHECK ACTUAL TIME AND BILL. HRS
				If tmpEntry(12) > 0 And  tmpEntry(13) > 0 Then 'CHECK RATES
					If tmpEntry(1) <> 4 Then
						rsBill("Status") = 1 
						rsBill("completed") = Now
					Else
						rsBill("Status") = 4
						'rsBill("actTT") = 0
						'rsBill("actMil") = 0
					End If
				End If
			Else
				If tmpEntry(1) = 2 Then
					rsBill("Status") = 2
				ElseIf tmpEntry(1) = 3 Then
					rsBill("Status") = 3
				ElseIf tmpEntry(1) = 4 Then
					rsBill("Status") = 4
				Else
					rsBill("Status") = 0
				End If
			End If
			rsBill("BilComment") = Request("txtCombil")
			'rsBill("showinst") = False
			'If Request("chkVis2") <> "" Then rsBill("showinst") = True
			rsBill("showintr") = False
			If Request("chkVis") <> "" Then rsBill("showintr") = True
			rsBill("LBconfirm") = False
			If Request("chkCon") <> "" Then rsBill("LBconfirm") = True 
			rsBill("LBconfirmToll") = False
			If Request("chkTollCon") <> "" Then rsBill("LBconfirmToll") = True 
			rsBill("payIntr") = False
			If Request("chkBillIntr") <> "" Then rsBill("payIntr") = True 
			rsBill("overmile") = False
			'If Request("chkOpay") <> "" Then rsBill("overpayhrs") = True 
			'rsBill("PayHrs") = Request("txtPayHrs")
			'If Request("chkOpay") <> "" Then rsBill("overmile") = True 
			'rsBill("actMil") = Z_Czero(Request("txtDecMile"))
			rsBill("ApproveHrs") = False
			If Request("chkHrs") <> "" Then rsBill("ApproveHrs") = True
			rsBill("happen") = Request("selHap")
			If tmpEntry(1) = 4 Then	rsBill("happen") = 1 'canceled billable
			'rsBill("happen") = Request("selHap")
			If Request("selHap") = 1 Or rsBill("happen") = 1 Then
				syscomstr = Z_fixNull(rsBill("syscom"))
				syscomstr = Replace(syscomstr, "<br>Appointment did not happen.", "")
				syscomstr = Replace(syscomstr, "<br>Appointment did happen.", "")
				rsBill("syscom") = syscomstr & "<br>Appointment did not happen." 
				rsBill("vermed") = 2
				rsBill("noreas") = Request("noreasq")
				rsBill("DSnoreas") = Z_CDate(Request("DSnoreas"))
			ElseIf Request("selHap") = 2 Then
				syscomstr = Z_fixNull(rsBill("syscom"))
				syscomstr = Replace(syscomstr, "<br>Appointment did not happen.", "")
				syscomstr = Replace(syscomstr, "<br>Appointment did happen.", "")
				rsBill("syscom") = syscomstr & "<br>Appointment did happen." 
				rsBill("noreas") = 0
				rsBill("DSnoreas") = Empty
			ElseIf Request("selHap") = 0 Then
				syscomstr = Z_fixNull(rsBill("syscom"))
				syscomstr = Replace(syscomstr, "<br>Appointment did not happen.", "")
				syscomstr = Replace(syscomstr, "<br>Appointment did happen.", "")
				rsBill("syscom") = syscomstr & "" 
				rsBill("noreas") = 0
				rsBill("DSnoreas") = Empty
			End If
			rsBill("phoneappt") = False
			If Request("chkFonCall") = 1 Then rsBill("phoneappt") = True
			rsBill.Update
			If tmpEntry(1) = 2 Then 'misseed
				tmpIntr = rsBill("IntrID")
				tmpDate = rsBill("appDate")
				tmpTimeTo = rsBill("appTimeFrom") & " - " & rsBill("appTimeTo")
				tmpReas = GetMisReason(rsBill("Missed"))
				tmpLang = GetLang(rsBill("langID"))
			End If
			If Request("email") = 1 Then 'For cancel email
				tmpIntr = rsBill("IntrID")
				tmpDate = rsBill("appDate")
				tmpTime = rsBill("appTimeFrom")
				tmpCity = GetCity(rsBill("deptID"))
				If rsBill("cliadd") Then tmpCity = rsBill("cCity")
				tmpInst = GetInst(rsBill("InstID"))
				tmpFname = rsBill("Cfname")
			End If
			tmpHPID = Z_CZero(rsBill("HPID"))
			tmpLBStat = rsBill("Status")
		End If
		rsBill.Close
		Set rsBill = Nothing
		'SAVE STATUS IN HP
		If tmpHPID <> 0 Then
			Set rsHPStat = Server.CreateObject("ADODB.RecordSet")
			sqlHPStat = "SELECT * FROM appointment_T WHERE [index] = " & tmpHPID
			rsHPStat.Open sqlHPStat, g_strCONNHP, 1, 3
			If Not rsHPStat.EOF Then
				rsHPStat("Status") = tmpLBStat
				rsHPStat.Update
			End If
			rsHPStat.Close
			Set rsHpStat = Nothing
		End If
		'send cancel email
		If Request("email") = 1 Then
			If tmpEntry(1) = 3 or tmpEntry(1) = 4 Then
				If GetEmailIntr(tmpIntr) <> "" Then
					Set mlMail = CreateObject("CDO.Message")
					strBody = "This is to let you know that appointment on " & _
						 tmpDate & ", " & tmpTime & ", in " & tmpCity & " at " & tmpInst & " for " & tmpFname & " is CANCELED.<br>" & _
						 "If you have any questions please contact the LanguageBank office immediately at 410-6183 or email us at " & _
						 "<a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>.<br>" & _
						 "E-mail about this cancelation was initiated by " & Request.Cookies("LBUsrName") & ".<br><br>" & _
						 "Thanks,<br>" & _
						 "Language Bank"

					retErr = zSendMessage(GetEmailIntr(tmpIntr), "language.services@thelanguagebank.org" _
							, "Appointment Cancellation " & tmpDate & "; " & tmpTime & ", " & tmpCity & " - " &  tmpInst _
							, strBody)
					'save to notes
					IntrName = GetIntr2(tmpIntr)
					Set rsNotes = Server.CreateObject("ADODB.RecordSet")
					sqlNotes = "SELECT LBComment FROM request_T WHERE [index] = " & Request("ReqID")
					rsNotes.Open sqlNotes, g_StrCONN, 1, 3
					If Not rsNotes.EOF Then
						rsNotes("LBComment") = rsNotes("LBComment") & vbCrlF & "Cancelation Email sent to " & IntrName & " on " & now
						rsNotes.Update
					End If
					rsNotes.CLose
					Set rsNotes = Nothing
					Session("MSG") = "Cancelation email sent to " & GetEmailIntr(tmpIntr)
				Else
					Session("MSG") = "ERROR: Interpreter has no email address assigned."
				End If
			End If
		End If
		If Date = cdate(tmpAppDate) Then
			If tmpEntry(1) = 2 And tmpstat <> 2 Then
				'email missed
				retErr = zSendMessage("language.services@thelanguagebank.org", GetPrime2(tmpIntr) _
						, "Missed appointment " & tmpAppDate, strBody)
				Session("MSG") = "Email Sent for Missed Appointment."
			End If
		End If
		If SaveHist(Request("ReqID"), "mainbill.asp") Then
	
		End If
		Response.Redirect "ReqConfirm.asp?ID=" & Request("ReqID")
	Else
		Response.Redirect "mainBill.asp?ID=" & Request("ReqID")
	End If
ElseIf Request("ctrl") = 11 Then 'SAVE ASSIGNED INTERPRETER
	'STORE INTERPRETER ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBINTR") = Z_DoEncrypt(Request("txtIntrLname") & "|" & Request("txtIntrFname") & "|" & Request("txtIntrEmail")	& _
	 	"|" & Request("txtIntrP1")	& "|" & Request("txtIntrFax")	& "|" & Request("txtIntrP2")	& "|" & Request("txtIntrAddr") & "|" & _
	 	Request("txtIntrCity")	& "|" & Request("txtIntrState") & "|" & Request("txtIntrCZip") & "|" & Request("HnewIntr") & _
	 	"|" & Request("chkInHouse") & "|" & Request("radioPrim2") & "|" & Request("txtIntrExt") & "|" & Request("selIntrRate") & "|" & Request("txtIntrAddrI") & "|" & Request("txtcomintr"))
	 'CHECK AVAILABILITY
	'If Request("SelIntr") <> "-1" Then
	'	If Z_CZero(Request("SelIntr")) <> Z_CZero(Request("IntrID")) Then
	'		Set rsAvail = Server.CreateObject("ADODB.RecordSet")
	'		sqlAvail = "SELECT * FROM Request_T WHERE appDate = '" & Request("txtAppDate") & "' AND appTimeFrom = '" & Request("txtAppTFrom") & "' AND IntrID = " & Request("SelIntr")
	'		rsAvail.Open sqlAvail, g_strCONN, 3, 1
	'		If Not rsAvail.EOF Then
	'			Session("MSG") = Session("MSG") & "<br>ERROR: Interpreter is not available for the said date and time."
	'		End If
	'		rsAvail.Close
	'		Set rsAvail = Nothing
	'	End If
	'End If
	If Session("MSG") = "" Then
		If Request("txtIntrLname") <> "" Or Request("txtIntrFname") <> "" Then
				tmpIntr = Split(Z_DoDecrypt(Request.Cookies("LBINTR")), "|")
				Set rsIntr = Server.CreateObject("ADODB.RecordSet")
				sqlIntr = "SELECT * FROM interpreter_T"
				rsIntr.Open sqlIntr, g_strCONN, 1, 3
				rsIntr.AddNew
				tmpIntrID = rsIntr("Index")
				rsIntr("Last Name") = CleanMe(tmpIntr(0))
				rsIntr("First Name") = CleanMe(tmpIntr(1))
				rsIntr("E-mail") = tmpIntr(2)
				rsIntr("Phone1") = tmpIntr(3)
				rsIntr("P1Ext") = tmpIntr(13)
				rsIntr("Fax") = tmpIntr(4)
				rsIntr("Phone2") = tmpIntr(5)
				rsIntr("Address1") = CleanMe(tmpIntr(6))
				rsIntr("City") = tmpIntr(7)
				rsIntr("State") = tmpIntr(8)
				rsIntr("Zip Code") = tmpIntr(9)
				rsIntr("IntrAdrI") = CleanMe(tmpIntr(15))
				rsIntr("Rate") = tmpIntr(14)
				newIntrRate = tmpIntr(14)
				rsIntr("InHouse") = False
				If tmpIntr(11) <> "" Then rsIntr("InHouse") = True
				If IsNull(tmpIntr(12)) Then tmpIntr(12) = 3
				rsIntr("prime") = tmpIntr(12)
				LangKo = LangName(Request("LangID"))
				If rsIntr("Language1") = "" Or IsNull(rsIntr("Language1")) Then 
					rsIntr("Language1") = LangKo
				Else
					If rsIntr("Language2") = ""  Or IsNull(rsIntr("Language2")) Then
						rsIntr("Language2") = LangKo
					Else
						If rsIntr("Language3") = ""  Or IsNull(rsIntr("Language3")) Then
							rsIntr("Language3") = LangKo
						Else
							If rsIntr("Language4") = "" Or IsNull(rsIntr("Language4")) Then
								rsIntr("Language4") = LangKo
							Else
								If rsIntr("Language5") = "" Or IsNull(rsIntr("Language5")) Then rsIntr("Language5") = LangKo
							End If
						End If
					End If 	
				End If
				rsIntr.Update
				rsIntr.Close
				Set rsIntr = Nothing
			End If
		Set rsAss = Server.CreateObject("ADODB.RecordSet")
		sqlAss = "SELECT * FROM Request_T WHERE [index] = " & Request("ReqID")
		rsAss.Open sqlAss,g_strCONN, 1, 3
		If Not rsAss.EOF Then
			'rsAss("intrID") = Request("SelIntr")
			tmpIntr = Request("SelIntr")
			'If tmpIntr = "" Then tmpIntr = tmpIntrID
			rsAss("IntrID") = tmpIntr
			RateIntr = 0
			If newIntrRate <> 0 Then 
				RateIntr = newIntrRate
			Else
				RateIntr =Request("txtIntrRate")
			End If
			rsAss("IntrRate") = Z_CDbl(RateIntr)
			If Request("selInstRate") <> 0 Then rsAss("InstRate") = Z_CDbl(Request("selInstRate"))
			rsAss("intrcomment") = Request("txtcomintr")
			'rsAss("Assignedby") = Response.Cookies("UID")
			'FOR HP
			tmpHPID = rsAss("HPID")
			rsAss.Update
		End If
		rsAss.Close
		Set rsAss = Nothing
			
			'SAVE TO HP
			If Z_CZero(tmpHPID) <> 0 Then
				Set rsHP = Server.CreateObject("ADODB.RecordSet")
				sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
				rsHP.Open sqlHp, g_strCONNHP, 1, 3
				If Not rsHP.EOF Then
					rsHP("intrID") = tmpIntr
					rsHp.Update
				End If
				rsHp.Close
				Set rsHp = Nothing
			End If
			'SAVE FB
			If tmpIntr > 0 And Request("txtintrfeed") <> "" Then
				Set rsFB = Server.CreateObject("ADODB.RecordSet")
				sqlFB = "SELECT * FROM InterpreterEval_T WHERE intrID = " & tmpIntr & " AND appID = " & Request("ReqID") & " AND UID = " & Request.Cookies("UID") 
				rsFB.Open sqlFB, g_strCONN, 1, 3
				If rsFB.EOF Then rsFB.AddNew
				rsFB("IntrID") = tmpIntr
				rsFB("date") = Now
				rsFB("appID") = Request("ReqID")
				rsFB("UID") = Request.Cookies("UID") 
				rsFB("comment") = Request("txtintrfeed")
				rsFB.Update
				rsFB.Close
				Set rsFB = Nothing
			ElseIf  tmpIntr > 0 And Request("txtintrfeed") = "" Then 'look if FB was just deleted
				Set rsFB = Server.CreateObject("ADODB.RecordSet")
				sqlFB = "SELECT * FROM InterpreterEval_T WHERE intrID = " & tmpIntr & " AND appID = " & Request("ReqID") & " AND UID = " & Request.Cookies("UID") 
				rsFB.Open sqlFB, g_strCONN, 1, 3
				If Not rsFB.EOF Then 
					rsFB.Delete
					rsFB.Update
				End If
				rsFB.Close
				Set rsFB = Nothing
			End If
			'SAVE HISTORY
			TimeNow = Now
			Set rsHist = Server.CreateObject("ADODB.RecordSet")
			sqlHist = "SELECT * FROM History_T WHERE ReqID = " & Request("ReqID")
			rsHist.Open sqlHist, g_strCONNHist, 1,3 
			If Not rsHist.EOF Then
				If Z_CZero(rsHist("interID")) <> Z_CZero(tmpIntr) Then
					rsHist("interID") = Z_CZero(tmpIntr)
					rsHist("interTS") = TimeNow
					rsHist("interU") = Request.Cookies("LBUsrName")
				End If
			Else
				rsHist.AddNew
				rsHist("interID") = Z_CZero(tmpIntr)
				rsHist("interTS") = TimeNow
				rsHist("interU") = Request.Cookies("LBUsrName")
			End If
			rsHist.Update
			rsHist.Close
			Set rsHist = Nothing
			If SaveHist(Request("ReqID"), "assign.asp") Then
	
			End If
			'email "yes" interpreters?
			
			Response.Redirect "ReqConfirm.asp?ID=" & Request("ReqID")
		Else
			Response.Redirect "assign.asp?ID=" & Request("ReqID")
		End If
ElseIf Request("ctrl") = 12 Then 'EDIT CONTACT INFORMATION
	'STORE ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBREQUEST") = Z_DoEncrypt(Request("txttstamp")	& "|" & Request("selReq")	& "|" & Request("txtClilname")	& "|" & _
		Request("txtClifname")	& "|" & Request("txtCliAdd")	& "|" & Request("txtCliCity")	& "|" & Request("txtCliState")	& "|" & _
		Request("txtCliZip")	& "|" & Request("txtCliDir") & "|" & Request("txtCliCir") & "|" & Request("txtDOB")	& "|" & _
		Request("selLang") & "|" & Request("txtAppdate")	& "|" & Request("txtAppTFrom")	& "|" & Request("txtAppTTo")	& "|" & _
		Request("txtAppLoc")	& "|" & Request("SelInst") & "|" & Request("selInstRate") & "|" & Request("txtDocNum") & "|" & _
		Request("txtCrtNum") & "|" & Request("chkClient") & "|" & Request("txtCliFon") & "|" & Request("selIntr") & "|" & _
		Request("txtActdate")	& "|" & Request("txtActTFrom")	& "|" & Request("txtActTTo") & "|" & Request("radioStat") & "|" & _
		Request("chkVer") & "|" & Request("chkPaid") & "|" & Request("txtBilHrs") & "|" & Request("txtcom") & "|" & Request("selCancel") & "|" & _
		Request("selIntrRate") & "|" & Request("chkEmer") & "|" & Request("selMissed") & "|" & Request("txtInstRate") & "|" & Request("txtIntrRate") & "|" & _
		Request("selDept") & "|" & Request("txtAlter") & "|" & Request("chkClientAdd") & "|" & Request("txtBilTInst") & "|" & Request("txtBilMInst") & "|" & _
		Request("txtBilMInst") & "|" & Request("txtBilMIntr") & "|" & Request("txtHPID") & "|" & Request("txtCliAddrI") & "|" & Request("chkemerfee"))
	'STORE INSTITUTION ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBINST") = Z_DoEncrypt(Request("txtNewInst") & "|" & Request("txtInstDept")	& "|" & Request("txtInstAddr")	& "|" & _
		Request("txtInstCity")	& "|" & Request("txtInstState")	& "|" & Request("txtInstZip") & "|" & Request("HnewInt") & "|" & _
		Request("selClass") & "|" & Request("chkBill")	& "|" & Request("txtBilAddr")	& "|" & Request("txtBilCity") & "|" & Request("txtBilState") & "|" & Request("txtBIlZip") & "|" & _
		Request("txtBlname") & "|" & Request("txtBfname"))	 
	'STORE DEPARTMENT ENTRIES 
	Response.Cookies("LBDEPT") = Z_DoEncrypt(Request("txtInstDept") & "|" & Request("selDept") & "|" & Request("txtInstAddr")	& "|" & _
		Request("txtInstCity")	& "|" & Request("txtInstState")	& "|" & Request("txtInstZip") & "|" & Request("HnewDept") & "|" & _
		Request("selClass") & "|" & Request("chkBill")	& "|" & Request("txtBilAddr")	& "|" & Request("txtBilCity") & "|" & Request("txtBilState") & "|" & Request("txtBIlZip") & "|" & _
		Request("txtBlname") & "|" & Request("txtBfname") & "|" & Request("selInst") & "|" & Request("txtInstAddrI"))	 
	'STORE REQUESTING PERSON's ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBREQ") = Z_DoEncrypt(Request("txtReqLname") & "|" & Request("txtReqFname")	& "|" & Request("txtphone") & "|" & _
		Request("txtemail")	& "|" & Request("txtfax") & "|" & Request("SelInst") & "|" & Request("HnewReq") & "|" & Request("radioPrim1") & "|" & Request("txtReqExt") & _
		"|" & Request("selDept"))	
	If Request("HnewDept") = "BACK" Then
		If Request("txtInstDept") = "" Then Session("MSG") = "<br>ERROR: Department's Name is required."
	Else
		If Request("selDept") = 0 Then Session("MSG") = "ERROR: Department is required."
	End If
	If Request("btnNewReq") = "BACK" Then
		If Request("txtReqLname") = "" And Request("txtReqFname") = "" Then Session("MSG") = Session("MSG") & "<br>ERROR: Requesting Person's full name is required."
	Else
		If Request("selReq") = "-1" Then Session("MSG") = Session("MSG") & "<br>ERROR: Requesting Person is required."
	End If
	If Request("txtphone") = "" And Request("txtfax") = "" And Request("txtemail") = "" Then Session("MSG") = Session("MSG") & _
		"<br>ERROR: At least one(1) Contact Number is required."
	If Request("txtInstAddr") = "" Or Request("txtInstCity") = "" Or Request("txtInstState") = "" Or Request("txtInstZip") = "" Then Session("MSG") = Session("MSG") & _
		"<br>ERROR: Instituition's full address is required."	
	If Request("txtInstRate") <> "" Then
		If Not IsNumeric(Request("txtInstRate")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Institution Rate."
	End If
	'CHECK INSTITUITION
	If Request("txtNewInst") <> "" Then
		Set rsRP = Server.CreateObject("ADODB.RecordSet")
		sqlRP = "SELECT * FROM institution_T WHERE facility = '" & Request("txtNewInst") & "' "
		rsRP.Open sqlRP, g_strCONN, 3, 1
		If NOT rsRP.EOF Then
			Session("MSG") = Session("MSG") & "<br>ERROR: Institution already exists."	
		End If
		rsRP.Close
		Set rsRP =Nothing
	End If 
	'CHECK DEPARTMENT
	If Request("txtInstDept") <> "" And Request("HnewInt") = "NEW"  And Request("Hnewdept") = "BACK" Then
		Set rsRP = Server.CreateObject("ADODB.RecordSet")
		sqlRP = "SELECT * FROM dept_T WHERE dept = '" & Request("txtInstDept") & "' AND InstID = " & Request("selInst")
		rsRP.Open sqlRP, g_strCONN, 3, 1
		If NOT rsRP.EOF Then
			Session("MSG") = Session("MSG") & "<br>ERROR: Department already exists for this insitution."	
		End If
		rsRP.Close
		Set rsRP =Nothing
	End If 
	If Session("MSG") = "" Then	
		'GET COOKIE OF REQUEST
		tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
		'ADD NEW INSTITUTION
		If Request("txtNewInst") <> "" Then
			tmpInst = Split(Z_DoDecrypt(Request.Cookies("LBINST")), "|")
			Set rsInst = Server.CreateObject("ADODB.RecordSet")
			sqlInst = "SELECT * FROM institution_T"
			rsInst.Open sqlInst, g_strCONN, 1, 3
			rsInst.AddNew
			tmpInstID = rsInst("Index")
			rsInst("Facility") = tmpInst(0)
			rsInst("Date") = Date
			rsInst.Update
			rsInst.Close
			Set rsInsr = Nothing
		End If
		'ADD NEW DEPARTMENT
		If Request("txtInstDept") <> "" Then
			tmpDept = Split(Z_DoDecrypt(Request.Cookies("LBDEPT")), "|")
			Set rsDept = Server.CreateObject("ADODB.RecordSet")
			sqlDept = "SELECT * FROM dept_T"
			rsDept.Open sqlDept, g_strCONN, 1, 3
			rsDept.AddNew
			tmpDeptID = rsDept("index")
			rsDept("dept") = tmpDept(0)
			If  Request("txtNewInst") = "" Then
				rsDept("InstID") = tmpDept(15)
			Else
				rsDept("InstID") =tmpInstID
			End If
			rsDept("Address") = CleanMe(tmpDept(2))
			rsDept("City") = tmpDept(3)
			rsDept("State") = tmpDept(4)
			rsDept("Zip") = tmpDept(5)
			If IsNull(tmpDept(7)) Then tmpDept(7) = 1
			rsDept("Class") = tmpDept(7)
			rsDept("Blname") = tmpDept(13)
			rsDept("InstAdrI") = CleanMe(tmpDept(16))
			If tmpDept(8) = "" Then
				rsDept("BAddress") = CleanMe(tmpDept(9))
				rsDept("BCity") = tmpDept(10)
				rsDept("BState") = tmpDept(11)
				rsDept("BZip") = tmpDept(12)
			Else
				rsDept("BAddress") = CleanMe(tmpDept(2))
				rsDept("BCity") = tmpDept(3)
				rsDept("BState") = tmpDept(4)
				rsDept("BZip") = tmpDept(5)
			End If
			rsDept.Update
			rsDept.Close
			Set rsDept = Nothing	
		End If
		'ADD NEW REQUESTING PERSON
		If Request("txtReqLname") <> "" Or Request("txtReqFname") <> ""Then
			tmpReq = Split(Z_DoDecrypt(Request.Cookies("LBREQ")), "|")
			Set rsReq = Server.CreateObject("ADODB.RecordSet")
			sqlReq = "SELECT * FROM requester_T"
			rsReq.Open sqlReq, g_strCONN, 1, 3
			rsReq.AddNew
			tmpReqID = rsReq("Index")
			rsReq("Lname") = CleanMe(tmpReq(0))
			rsReq("Fname") = CleanMe(tmpReq(1))
			rsReq("phone") = tmpReq(2)
			rsReq("pExt") = tmpReq(8)
			rsReq("eMail") = tmpReq(3)
			rsReq("fax") = tmpReq(4)
			If IsNull(tmpReq(7)) Then tmpReq(7) = 2
			rsReq("prime") = tmpReq(7)
			rsReq.Update
			rsReq.Close
			Set rsReq = Nothing
		End If
		
		'SAVE EDITTED ENTRIES
		Set rsMain = Server.CreateObject("ADODB.RecordSet")
		sqlMain = "SELECT * FROM request_T WHERE [index] = " & Request("HID")
		rsMain.Open sqlMain, g_strCONN, 1, 3
		If Not rsMain.EOF Then
			
			If tmpEntry(1) = "" Then tmpEntry(1) = tmpReqID
			rsMain("reqID") = tmpEntry(1)
			If Request("txtNewInst") = "" Then
				rsMain("InstID") = tmpEntry(16)
			Else
				rsMain("InstID") = tmpInstID
			End If
			If Request("txtInstDept") = "" Then
				rsMain("DeptID") = tmpEntry(37)
			Else
				rsMain("DeptID") = tmpDeptID
			End If
			If Z_CZero(tmpEntry(17)) <> 0 Then rsMain("InstRate") = Z_Cdbl(tmpEntry(17))
			
			'rsMain("Emergency") = False
			'If tmpEntry(33) <> "" Then rsMain("Emergency") = True
			'	
			'rsMain("Emerfee") = False
			'If tmpEntry(46) <> "" Then rsMain("Emerfee") = True
			
			tmpHPID = Z_CZero(rsMain("HPID"))
			rsMain.Update
			tmpLBStat = rsMain("Status")
		End If
		rsMain.Close
		Set rsMain = Nothing
		'SAVE REQUESTING PERSON'S ENTRIES
		If Request("txtReqLname") = "" Or Request("txtReqFname") = "" Then
			Set rsReq = Server.CreateObject("ADODB.RecordSet")
			sqlReq = "SELECT * FROM requester_T WHERE [index] = " & tmpEntry(1)
			rsReq.Open sqlReq, g_strCONN, 1, 3
			If Not rsReq.EOF Then
				rsReq("Phone") = Request("txtphone")
				rsReq("eMail") = Request("txtemail")
				rsReq("Fax") = Request("txtfax")
				rsReq("prime") = Request("radioPrim1")
				rsReq("pExt") = Request("txtReqExt")
				rsReq.Update
			End If
			rsReq.Close
			Set rsReq = Nothing
		End If
		'SAVE INSTITUTION ENTRIES
		If Request("txtNewInst") = "" Then
			Set rsInst = Server.CreateObject("ADODB.RecordSet")
			sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpEntry(16)
			rsInst.Open sqlInst, g_strCONN, 1, 3
			If Not rsInst.EOF Then
			'	rsInst("Address") = CleanMe(Request("txtInstAddr"))
				'rsInst("Department") = Request("txtInstDept")
			'	rsInst("City") = Request("txtInstCity")
			'	rsInst("State") = Request("txtInstState")
			'	rsInst("Zip") = Request("txtInstZip")
			'	rsInst("Rate") = Request("txtInstRate")
			'	rsInst("Class") = Request("selClass")
			'	rsInst("Blname") = Request("txtBlname")
			'	rsInst("Bfname") = Request("txtBfname")
			'	If Request("chkBill") = "" Then
			'		rsInst("BAddress") = CleanMe(Request("txtBilAddr"))
			'		rsInst("BCity") =Request("txtBilCity")
			'		rsInst("BState") = Request("txtBilState")
			'		rsInst("BZip") = Request("txtBilZip")
			'	Else
			'		rsInst("BAddress") = CleanMe(Request("txtInstAddr"))
			'		rsInst("BCity") =Request("txtInstCity")
			'		rsInst("BState") = Request("txtInstState")
			'		rsInst("BZip") = Request("txtInstZip")
			'	End If
			'	rsInst.Update
			End If
			rsInst.Close
			Set rsInst = Nothing
		End If
		'SAVE DEPARTMENT ENTRIES
		If Request("txtInstDept") = "" Then
			Set rsDept = Server.CreateObject("ADODB.RecordSet")
			sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpEntry(37)
			rsDept.Open sqlDept, g_strCONN, 1, 3
			If Not rsDept.EOF Then
				rsDept("Address") = CleanMe(Request("txtInstAddr"))
				rsDept("City") = Request("txtInstCity")
				rsDept("State") = Request("txtInstState")
				rsDept("Zip") = Request("txtInstZip")
				'rsDept("Class") = Request("selClass")
				rsDept("Blname") = Request("txtBlname")
				rsDept("InstAdrI") = CleanMe(Request("txtInstAddrI"))
				If Request("chkBill") = "" Then
					rsDept("BAddress") = CleanMe(Request("txtBilAddr"))
					rsDept("BCity") =Request("txtBilCity")
					rsDept("BState") = Request("txtBilState")
					rsDept("BZip") = Request("txtBilZip")
				Else
					rsDept("BAddress") = CleanMe(Request("txtInstAddr"))
					rsDept("BCity") =Request("txtInstCity")
					rsDept("BState") = Request("txtInstState")
					rsDept("BZip") = Request("txtInstZip")
				End If	
				rsDept.Update
			End If
			rsDept.Close
			Set rsDept = Nothing
		End If
		'SAVE REQUESTER TO DEPARTMENT RELATIONSHIP
		If Request("txtReqLname") = "" Or Request("txtReqFname") = "" Then
			IDReq = tmpEntry(1)
		Else
			IDReq = tmpReqID
		End If
		If Request("txtInstDept") = "" Then
			IDDept = tmpEntry(37)
		Else
			IDDept = tmpDeptID
		End If
		Set rsReqDept = Server.CreateObject("ADODB.RecordSet")
		sqlReqDept = "SELECT * FROM reqdept_T WHERE ReqID = " & IDReq & " AND DeptID = " & IDDept
		rsReqDept.Open sqlReqDept, g_strCONN, 1, 3
		If rsReqDept.EOF Then
			rsReqDept.AddNew
			rsReqDept("ReqID") = IDReq
			rsReqDept("DeptID") = IDDept
			rsReqDept.Update
		End If
		rsReqDept.Close
		Set rsReqDept = Nothing
		If SaveHist(Request("HID"), "editcontact.asp") Then
		 'Session("MSG") = "HIST SAVED"
		End If
		'SAVE REQUEST
		Response.Redirect "reqconfirm.asp?ID=" & Request("HID")
	Else
		Response.Redirect "editcontact.asp?ID=" & Request("HID")
	End If	
ElseIf Request("ctrl") = 13 Then 'EDIT APPOINTMENT INFORMATION
	'STORE ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBREQUEST") = Z_DoEncrypt(Request("txttstamp")	& "|" & Request("selReq")	& "|" & Request("txtClilname")	& "|" & _
		Request("txtClifname")	& "|" & Request("txtCliAdd")	& "|" & Request("txtCliCity")	& "|" & Request("txtCliState")	& "|" & _
		Request("txtCliZip")	& "|" & Request("txtCliDir") & "|" & Request("txtCliCir") & "|" & Request("txtDOB")	& "|" & _
		Request("selLang") & "|" & Request("txtAppdate")	& "|" & Request("txtAppTFrom")	& "|" & Request("txtAppTTo")	& "|" & _
		Request("txtAppLoc")	& "|" & Request("SelInst") & "|" & Request("selInstRate") & "|" & Request("txtDocNum") & "|" & _
		Request("txtCrtNum") & "|" & Request("chkClient") & "|" & Request("txtCliFon") & "|" & Request("selIntr") & "|" & _
		Request("txtActdate")	& "|" & Request("txtActTFrom")	& "|" & Request("txtActTTo") & "|" & Request("radioStat") & "|" & _
		Request("chkVer") & "|" & Request("chkPaid") & "|" & Request("txtBilHrs") & "|" & Request("txtcom") & "|" & Request("selCancel") & "|" & _
		Request("selIntrRate") & "|" & Request("chkEmer") & "|" & Request("selMissed") & "|" & Request("txtInstRate") & "|" & Request("txtIntrRate") & "|" & _
		Request("selDept") & "|" & Request("txtAlter") & "|" & Request("chkClientAdd") & "|" & Request("txtBilTInst") & "|" & Request("txtBilMInst") & "|" & _
		Request("txtBilMInst") & "|" & Request("txtBilMIntr") & "|" & Request("txtHPID") & "|" & Request("txtCliAddrI")& "|" & Request("chkout") & "|" & _
		Request("chkmed") & "|" & Request("MCnum") & "|" & Request("txtPDamount") & "|" & Request("h_tmpfilename") & "|" & Request("sellateres") & "|" & _
		Request("MHPnum") & "|" & Request("NHHFnum") & "|" & Request("WSHPnum") & "|" & Request("chkawk") & "|" & Request("txtjudge") & "|" & _
		Request("txtclaim") & "|" & Request("chkcall") & "|" & Request("chkleave"))
	'CHECK REQUIRED FIELDS
	If Request("txtClilname") = "" Or Request("txtClifname") = "" Then Session("MSG") = Session("MSG") & "<br>ERROR: Client's full name is required."
	If Request("selLang") = "-1" Then Session("MSG") = Session("MSG") & "<br>ERROR: Language is required."
	If Request("txtAppDate") = "" Then Session("MSG") = Session("MSG") & "<br>ERROR: Appointment Date is required."
	If Request("txtAppTFrom") = "" Then Session("MSG") = Session("MSG") & "<br>ERROR: Appointment Time (From:) is required."	
	'CHECK VALID VALUES
	If Request("txtDOB") <> "" Then
		If Not IsDate(Request("txtDOB")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Date of Birth."
	End If
	If Request("txtAppdate") <> "" Then
		If Not IsDate(Request("txtAppdate")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment date."
	End If
	If Request("txtAppTFrom") <> "" Then
		If Not IsDate(Request("txtAppTFrom")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (From:)."
	End If
	If Request("txtAppTTo") <> "" Then
		If Not IsDate(Request("txtAppTTo")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (To:)."
	End If
	'check avalability of interpreter if date/time is changed
	If Request("Intr") <> 1 Then
		If (Request("txtAppDate") <> Request("mydate")) Or (Request("mystime") <> Request("txtAppTFrom")) Then
			If Z_CZero(Request("myint")) <> 0 Then
				Set rsAvail = Server.CreateObject("ADODB.RecordSet")
				sqlAvail = "SELECT * FROM Request_T WHERE appDate = '" & Request("txtAppDate") & "' AND appTimeFrom = '" & Request("txtAppTFrom") & "' AND IntrID = " & Request("myint")
				rsAvail.Open sqlAvail, g_strCONN, 3, 1
				If Not rsAvail.EOF Then
					Session("MSG") = Session("MSG") & "<br>ERROR: Interpreter is not available for the said date and time."
				End If
				rsAvail.Close
				Set rsAvail = Nothing
			End If
				'Session("MSG") = cbool(Request("txtAppDate") <> Request("mydate")) & " OR " & cbool(Request("mystime") <> cdate(Request("txtAppTFrom"))) & _
				'	" " & Request("mystime") & " <> " & Request("txtAppTFrom")
		End If
	End If
	If Session("MSG") = "" Then
		'SAVE EDITTED ENTRIES
		tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
		Set rsMain = Server.CreateObject("ADODB.RecordSet")
		sqlMain = "SELECT * FROM request_T WHERE [index] = " & Request("ReqID")
		rsMain.Open sqlMain, g_strCONN, 1, 3
		If Not rsMain.EOF Then
			appdate = rsMain("appdate")
			stime = rsMain("appTimeFrom")
			etime = rsMain("appTimeTo")
			lang = rsMain("langID")
			rsMain("clname") = tmpEntry(2)
			rsMain("cfname") = tmpEntry(3)
			rsMain("Client") = False
			If tmpEntry(20) <> "" Then rsMain("Client") = True
			rsMain("Caddress") = CleanMe(tmpEntry(4))
			rsMain("Ccity") = tmpEntry(5)
			rsMain("Cstate") = Ucase(tmpEntry(6))
			rsMain("Czip") = tmpEntry(7)
			rsMain("directions") = tmpEntry(8)
			rsMain("spec_cir") = tmpEntry(9)
			rsMain("DOB") = Z_DateNull(tmpEntry(10))
			rsMain("LangID") = tmpEntry(11)
			rsMain("appDate") = Z_DateNull(tmpEntry(12))
			rsMain("appTimeFrom") = Z_DateNull(tmpEntry(12)) & " " & Z_DateNull(tmpEntry(13))
			rsMain("appTimeTo") = Z_DateNull(tmpEntry(12)) & " " & Z_DateNull(tmpEntry(14))
			rsMain("appLoc") = tmpEntry(15)
			rsMain("judge") = Request("txtjudge")
			rsMain("claimant") = Request("txtclaim")
			rsMain("docNum") = tmpEntry(18)
			rsMain("CrtRumNum") = tmpEntry(19)
			rsMain("Comment") = tmpEntry(30)
			rsMain("Cphone") = tmpEntry(21)
			rsMain("CAphone") = tmpEntry(38)
			rsMain("CliAdd") = False
			If tmpEntry(39) <> "" Then rsMain("CliAdd") = True
			rsMain("CliAdrI") = tmpEntry(45)
			tmpHPID = Z_CZero(rsMain("HPID"))
			If Request("Intr") = 1 Then rsMain("IntrID") = "-1"
			rsMain("Gender") = Request("selGender")
			rsMain("Child") = False
			If Request("chkMinor") <> "" Then rsMain("Child") = True
			rsMain("late") = Request("sellate")
			rsMain("lateres") = Z_CZero(Request("sellateres"))
			rsMain("outpatient") = False
			If Request("chkout") <> "" Then rsMain("outpatient") = True
			rsMain("hasmed") = False
			If Request("chkmed") <> "" Then rsMain("hasmed") = True
			rsMain("medicaid") = Request("MCnum")
			rsMain("meridian") = Request("MHPnum")
			rsMain("nhhealth") = Request("NHHFnum")
			rsMain("wellsense") = Request("WSHPnum")
			rsMain("acknowledge") = false
			If Request("chkawk") <> "" Then rsMain("acknowledge") = True
			rsMain("vermed") = 0
			If Request("chkAppMed") <> "" Then rsMain("vermed") = 1
			If Request("chkDisAppMed") <> "" Then rsMain("vermed") = 2
			rsMain("autoacc") = False
			If Request("chkacc") <> "" Then rsMain("autoacc") = True
			rsMain("wcomp") = False
			If Request("chkcomp") <> "" Then rsMain("wcomp") = True
			rsMain("secins") = Request("selIns")
			rsMain("PDamount") = Z_CZero( Request("txtPDamount")) 'PD
			If FileUpload(Request("h_tmpfilename")) Then 
				rsMain("UploadFile") = True
				rsMain("filename") = Request("h_tmpfilename") & ".PDF"
			End If
			If rsMain("InstID") = 479 Then rsMain("training") = Request("selTrain")
			rsMain("mrrec") = Request("mrrec")
			rsMain("blocksched") = False
			If Request("chkblock") <> "" Then rsMain("blocksched") = True
			rsMain("courtcall") = False
			If Request("chkcall") <> "" Then rsMain("courtcall") = True
			rsMain("leavemsg") = False
			If Request("chkleave") <> "" Then rsMain("leavemsg") = True
			rsMain.Update
			block = 0
			If rsMain("blocksched") Then block = 1
			tmpLBStat = rsMain("Status")
			tmpIntrID = rsMain("IntrID")
			tmpInstID = rsMain("InstID")
			vermed = rsMain("vermed")
			clilname = Ucase(Trim(rsMain("clname")))
			clifname = Ucase(Trim(rsMain("cfname")))
			medicaid = Ucase(Trim(rsMain("medicaid")))
			If medicaid = "" Then medicaid = Ucase(Trim(rsMain("meridian")))
			If medicaid = "" Then medicaid = Ucase(Trim(rsMain("nhhealth")))
			If medicaid = "" Then medicaid = Ucase(Trim(rsMain("wellsense")))
			dob = Trim(rsMain("dob"))
			gender = 0
			If rsMain("gender") Then gender = 1
			
			If vermed = 1 Then 'add in client list
				Set rsCli = Server.CreateObject("ADODB.RecordSet")
				sqlCli = "SELECT * FROM clientuploaded_T WHERE lname = '" & clilname & "' AND fname = '" & clifname & "' AND medicaid = '" & medicaid & _
					"' AND dob = '" & dob & "' AND gender = " & gender 
				rsCli.Open sqlCli, g_strCONN, 3, 1
				If rsCli.EOF Then
					Set rsCliList = Server.CreateObject("ADODB.RecordSet")
					rsCliList.Open "SELECT * FROM clientList_T", g_strCONN, 1, 3
					rsCliList.AddNew
					rsCliList("lname") = clilname
					rsCliList("fname") = clifname
					rsCliList("medicaid") = medicaid
					rsCliList("dob") = dob
					rsCliList("gender") = False
					If gender = 1 Then rsCliList("gender") = True
					rsCliList("timestamp") = now
					rsCliList.Update
					rsCliList.Close
					Set rsCliList = Nothing
				End If
				rsCli.Close
				Set rsCli = Nothing
			ElseIf vermed = 2 Then 'remove in client list
				Set rsCli = Server.CreateObject("ADODB.RecordSet")
				sqlCli = "SELECT * FROM clientList_T WHERE lname = '" & clilname & "' AND fname = '" & clifname & "' AND medicaid = '" & medicaid & _
					"' AND dob = '" & dob & "' AND gender = " & gender 
				rsCli.Open sqlCli, g_strCONN, 1, 3
				If Not rsCli.EOF Then
					rsCli.Delete
					rsCli.Update
				End If
				rsCli.Close
				Set rsCli = Nothing
			End if 
		End If
		rsMain.Close
		Set rsMain = Nothing
		'response.write "LANGID: " & lang & " = " & tmpentry(11) & " : " & cbool(Z_Czero(lang) =  Z_Czero(tmpentry(11))) & "<br>" & _
		'	"DATE: " & Z_DateNull(appdate) & " = " & Z_DateNull(tmpEntry(12)) & " : " & cbool(Z_DateNull(appdate) = Z_DateNull(tmpEntry(12))) & "<br>" & _
		'	"TIME: " & stime & " = " & Z_DateNull(tmpEntry(12)) & " " & Z_DateNull(tmpEntry(13)) & " : " & cbool(Z_DateNull(stime) = Z_DateNull(tmpEntry(12) & " " & tmpEntry(13))) & "<br>"
		If block <> 1 Then 'not include block
			If Z_Czero(tmpEntry(11)) <> 52 And Z_Czero(tmpEntry(11)) <> 78 And Z_Czero(tmpEntry(11)) <> 81 And Z_Czero(tmpEntry(11)) <> 85 And Z_Czero(tmpEntry(11)) <> 90 Then 'do not include ASL
				If Not (Z_DateNull(appdate) = Z_DateNull(tmpEntry(12)) And Z_DateNull(stime) = Z_DateNull(tmpEntry(12) & " " & tmpEntry(13))) And Z_Czero(lang) = Z_Czero(tmpEntry(11)) Then
					Call Z_ResetIntr(Request("ReqID"))
				ElseIf Z_Czero(lang) <> Z_Czero(tmpEntry(11)) Then
					Call Z_ResetIntr2(Request("ReqID"))
				End If
			End If
		End If
		'LOG UPLOAD
  	If FileUpload(Request("h_tmpfilename")) Then  'save Form on DB
			Set rsFile = Server.CreateObject("ADODB.RecordSet")
			sqlFile = "SELECT * FROM pdf_T"
			rsFile.Open sqlFile, g_strCONN, 1, 3
			rsFile.AddNew
			rsFile("appID") = Request("ReqID")
			rsFile("filename") = Request("h_tmpfilename") & ".PDF"
			rsFile("datestamp") = Now
			rsFile.Update
			rsFile.Close
			Set rsFile = Nothing
		End If
		If Date = cdate(tmpEntry(12)) Then
			'response.write date & " = " & Z_DateNull(tmpEntry(12)) & " = " & cbool(Date = cdate(tmpEntry(12)))
			If Request("sellate") > 0 And tmpIntrID > 0 Then
				'email late
				strMSG = "<p>" & vbCrLf & GetIntr2(tmpIntrID) & " running " & Request("sellate") & _
						" minutes late, " & GetInst(tmpInstID) & vbCrLf & "<br>Reason: " & _
						GetReasonTardy(Request("sellateres")) & "</p>"
				set mlMail=nothing
				retErr = zSendMessage("language.services@thelanguagebank.org", GetPrime2(tmpIntrID) _
						, GetIntr2(tmpIntrID) & " running " & Request("sellate") & " minutes late, " & GetInst(tmpInstID) _
						, strMSG)
				Session("MSG") = "Email Sent for tardiness."
			End If
		End If
		'SAVE INTERPRETER AND OTHER INFO TO HOSPITAL PILOT SITE
		If tmpEntry(44) <> "" Then
			Set rsHP = Server.CreateObject("ADODB.RecordSet")
			sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpEntry(44)
			rsHp.Open sqlHp, g_strCONNHP, 1, 3
			If Not rsHp.EOF Then
				rsHp("clname") = Z_DoEncrypt(tmpEntry(2))
				rsHp("cfname") =  Z_DoEncrypt(tmpEntry(3))
				rsHp("appdate") = tmpEntry(12)
				rsHp("DOB") = Z_DateNull(tmpEntry(10))
				If tmpEntry(13) <> "" Then
					rsHp("TimeFrom") = tmpEntry(12) & " " & tmpEntry(13)
				Else
					rsHp("TimeFrom") = Empty
				End If
				If tmpEntry(14) <> "" Then 
					rsHp("TimeTo") = tmpEntry(12) & " " & tmpEntry(14)
				Else
					rsHp("TimeTo") = Empty
				End If
				rsHp("docknum") = tmpEntry(18)
				rsHp("crtroom") = tmpEntry(19)
				rsHp("langID") = GetHPLang(tmpEntry(11))
				rsHp("phone") =  Z_DoEncrypt(tmpEntry(21))
				rsHp("comment") = tmpEntry(30)
				rsHp("mobile") =  Z_DoEncrypt(tmpEntry(38))
				'rsHp("IntrID") = tmpEntry(22)
				rsHp("Minor") = False
				If Request("chkMinor") <> "" Then rsHp("Minor") = True
				If tmpEntry(39) <> "" Then
					rsHp("mwhere")  = 0
					rsHp("maddr")  = CleanMe(tmpEntry(4))
					rsHp("mcity")  = tmpEntry(5)
					rsHp("mstate")  = Ucase(tmpEntry(6))
					rsHp("mzip")  = tmpEntry(7)
					rsHp("mlocation")  = 0
					rsHp("mother")  = ""	
				End If
				' save PD
				rsHp("PDamount") = Z_CZero( Request("txtPDamount")) 'PD
				If FileUpload(Request("h_tmpfilename")) Then 
					rsHp("UploadFile") = True
					rsHp("filename") = Request("h_tmpfilename") & ".PDF"
				End If
				'save medicaid
				rsHp("outpatient") = False
				If Request("chkout") <> "" Then rsHp("outpatient") = True
				rsHp("hasmed") = False
				If Request("chkmed") <> "" Then rsHp("hasmed") = True
				rsHp("medicaid") = Request("MCnum")
				rsHp("meridian") = Request("MHPnum")
				rsHp("nhhealth") = Request("NHHFnum")
				rsHp("wellsense") = Request("WSHPnum")
				rsHp("acknowledge") = false
				If Request("chkawk") <> "" Then rsHp("acknowledge") = True
				rsHp("autoacc") = False
				If Request("chkacc") <> "" Then rsHp("autoacc") = True
				rsHp("wcomp") = False
				If Request("chkcomp") <> "" Then rsHp("wcomp") = True
				rsHp("secins") = Request("selIns")
				rsHp("mrrec") = Request("mrrec")
				rsHP("block") = false
				If Request("chkblock") <> "" Then rsHP("block") = True
				rsHp.Update
			End If
			rsHp.Close
			Set rsHp = Nothing
		End If
		'SAVE HISTORY
	
		TimeNow = Now
		Set rsHist = Server.CreateObject("ADODB.RecordSet")
		sqlHist = "SELECT * FROM History_T WHERE ReqID = " & Request("ReqID")
		rsHist.Open sqlHist, g_strCONNHist, 1,3 
		If rsHist.EOF Then 
			rsHist.AddNew
			rsHist("ReqID") = Request("ReqID")
			rsHist("Creator") = Request.Cookies("LBUsrName")
			rsHist("date") = tmpEntry(12)
			rsHist("dateTS") = TimeNow
			rsHist("dateU") = Request.Cookies("LBUsrName")
			rsHist("Stime") = Z_dateNull(tmpEntry(12) & " " & tmpEntry(13))
			rsHist("StimeTS") = TimeNow
			rsHist("StimeU") = Request.Cookies("LBUsrName")
			If tmpEntry(39) <> "" Then
				tmpHistAdr = tmpEntry(4) & "|" & tmpEntry(5) & "|" & tmpEntry(6) & "|" & tmpEntry(7)
			Else
				tmpHistAdr = Request("hidHistaddr")
			End If
			rsHist("location") = tmpHistAdr
			rsHist("locationTS") = TimeNow
			rsHist("locationU") = Request.Cookies("LBUsrName")
			'If tmpEntry(22) <> "-1" Then
			'	rsHist("interID") = tmpEntry(22)
			'	rsHist("interTS") = TimeNow
			'	rsHist("interU") = Request.Cookies("LBUsrName")
			'End If
		Else
			
			If z_datenull(rsHist("date")) <> z_datenull(tmpEntry(12)) Then
				rsHist("date") = tmpEntry(12)
				rsHist("dateTS") = TimeNow
				rsHist("dateU") = Request.Cookies("LBUsrName")
			End If
			If z_datenull(rsHist("Stime")) <> Z_dateNull(tmpEntry(12) & " " & tmpEntry(13)) Then
				rsHist("Stime") = Z_dateNull(tmpEntry(12) & " " & tmpEntry(13))
				rsHist("StimeTS") = TimeNow
				rsHist("StimeU") = Request.Cookies("LBUsrName")
			End If
			If tmpEntry(39) <> "" Then
				tmpHistAdr = tmpEntry(4) & "|" & tmpEntry(5) & "|" & tmpEntry(6) & "|" & tmpEntry(7)
			Else
				tmpHistAdr = Request("hidHistaddr")
			End If
			'response.write cbool(rsHist("location") <> tmpHistAdr)
			If Z_fixnull(rsHist("location")) <> tmpHistAdr Then
				rsHist("location") = tmpHistAdr
				rsHist("locationTS") = TimeNow
				rsHist("locationU") = Request.Cookies("LBUsrName")
			End If
			'If tmpEntry(22) <> "-1" Then 
			'	If rsHist("interID") <> Cint(tmpEntry(22)) Then
			'		rsHist("interID") = tmpEntry(22)
			'		rsHist("interTS") = TimeNow
			'		rsHist("interU") = Request.Cookies("LBUsrName")
			'	End If
			'End If
		End If
		rsHist.Update
		rsHist.Close
		Set rsHist = Nothing
		If SaveHist(Request("ReqID"), "editapp.asp") Then
		 'Session("MSG") = "HIST SAVED"
		End If
		Response.Redirect "reqconfirm.asp?ID=" & Request("ReqID")
	Else
		Response.Redirect "editapp.asp?ID=" & Request("ReqID")
	End If	
ElseIf Request("ctrl") = 14 Then 'INTERPRETER CONFIRMATION
	Set rsConfirm = server.createobject("ADODB.RecordSet")
	sqlConfirm = "SELECT * FROM request_T WHERE [index] = " & Request("ReqID")
	rsConfirm.Open sqlConfirm, g_strCONN, 1, 3
	If Not rsConfirm.EOF Then
		tmpdate = rsConfirm("appdate")
		rsConfirm("verified") = True
		rsConfirm.Update
	End If
	rsConfirm.CLose
	Set rsConfirm = Nothing
	Session("MSG") = "Appointment: " & Request("ReqID") & " has been confirmed."
	response.Redirect "calendarview2.asp?selMonth=" & Month(tmpdate) & "&txtday=" & Day(tmpdate) & "&txtyear=" & Year(tmpdate)
ElseIf Request("ctrl") = 15 Then 'today
	Response.Cookies("LBREPORT") = Z_DoEncrypt(100)
'	Response.Cookies("LBREPORT").Expires = Now + 0.34
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		response.redirect "calendarview2.asp?rep=25&tmpdate='" & Request("tmpDate") & "' "
	Else
		response.redirect "calendarview2.asp?rep=25&tmpdate='" & Request("tmpDate") & "'&tmpRP=" & Request.Cookies("LBUSERTYPE")
	End If
ElseIf Request("ctrl") = 16 Then 'edit notes
	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	sqlMain = "SELECT * FROM request_T WHERE [index] = " & Request("ReqID")
	rsMain.Open sqlMain, g_strCONN, 1, 3
	If Not rsMain.EOF Then
		rsMain("LBcomment") = Request("txtLBcom")
		rsMain.Update
	End If
	rsMain.Close
	Set rsMain = Nothing
	If SaveHist(Request("ReqID"), "LBcom.asp") Then
		 'Session("MSG") = "HIST SAVED"
		End If
	Response.Redirect "reqconfirm.asp?ID=" & Request("ReqID")
ElseIf Request("ctrl") = 17 Then'timsheet/mileage
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT * FROM request_T WHERE NOT timestamp IS NULL"
	If Request("selIntr") > 0 Then
		sqlTBL = sqlTBL & " AND IntrID = " & Request("selIntr")
	End If
	If Request("txtFromd8") <> "" Then
		If IsDate(Request("txtFromd8")) Then
			sqlTBL = sqlTBL & " AND appdate >= '" & Request("txtFromd8") & "'"
		End If
	End If
	If Request("txtTod8") <> "" Then
		If IsDate(Request("txtTod8")) Then
			sqlTBL = sqlTBL & " AND appdate <= '" & Request("txtTod8") & "'"
		End If
	End If
	rsTBL.Open sqlTBL, g_strCONN, 1, 3 
	If Not rsTBL.EOF Then 
		y = Request("Hctr")
		For ctr = 1 To y - 1
			tmpID = Request("ID" & ctr)
			tmpIndex = "Index= " & tmpID
			rsTBL.MoveFirst
			rsTBL.Find(tmpIndex)
			If Not rsTBL.EOF Then
				If Request("ctrlX") = 1 Then
					If Request("txtstime" & ctr) <> "" Then
						If Not IsDate(Request("txtstime" & ctr)) Then
							Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Actual Start Time in Request ID " & tmpID & "."
						End If
					End If
					If Request("txtetime" & ctr) <> "" Then
						If Not IsDate(Request("txtetime" & ctr)) Then
							Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Actual End Time in Request ID " & tmpID & "."
						End If
					End If
					If Request("txtPhrs" & ctr) <> "" Then
						If Not IsNumeric(Request("txtPhrs" & ctr)) Then
							Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Payable Hours in Request ID " & tmpID & "."
						End If
					End If
				ElseIf Request("ctrlX") = 2 Then
					If Request("txtTol" & ctr) <> "" Then
						If Not IsNumeric(Request("txtTol" & ctr)) Then
							Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Tolls & parking in Request ID " & tmpID & "."
						End If
					End If
					If Request("txtmile" & ctr) <> "" Then
						If Not IsNumeric(Request("txtmile" & ctr)) Then
							Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Mileage in Request ID " & tmpID & "."
						End If
					End If
				End If
				If Session("MSG") = "" Then
					If Request("ctrlX") = 1 Then
						If Z_FixNull(Request("txtstime" & ctr)) <> "" Then 
							rsTBL("AStarttime") = Z_DateNull(rsTBL("appdate") & " " & Request("txtstime" & ctr))
						Else
							rsTBL("AStarttime") = Empty
						End If
						If Z_FixNull(Request("txtetime" & ctr)) <> "" Then 
							rsTBL("AEndtime") = Z_DateNull(rsTBL("appdate") & " " & Request("txtetime" & ctr))
						Else
							rsTBL("AEndtime") = Empty
						End If
						rsTBL("ShowIntr") = False
						If Request("chkshow" & ctr) = 1 Then rsTBL("ShowIntr") = True
						'If Request("chkpay" & ctr) = 1 Then 
						'	rsTBL("PayIntr") = True
						'Else
						'	rsTBL("PayIntr") = False
						'	rsTBL("actTT") = 0
						'	rsTBL("actMil") = 0
						'	rsTBL("TT_Intr") = 0
						'	rsTBL("M_Intr") = 0
						'End If
						rsTBL("overpayhrs") = False
						If Request("chkOverPhrs" & ctr) <> "" Then rsTBL("overpayhrs") = True
						tmpPayHrs = Request("txtPhrs" & ctr)
						rsTBL("payhrs") = tmpPayHrs
						rsTBL("actTT") =  Z_Czero(Request("txtTT" & ctr))
						If Not rsTBL("LBconfirm") Then
							rsTBL("LBconfirm") = False
							If Request("chkTS" & ctr) <> "" Then
								If tmpPayHrs <> "" Then 
									rsTBL("LBconfirm") = True
									Set fso = CreateObject("Scripting.FileSystemObject")
									Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
									strLog = Now & vbTab & "Timesheet confirmed (ID: " & tmpID & " PayHrs: " & tmpPayHrs & " Actual Time: " &  Request("txtstime" & ctr) & " - " &  Request("txtetime" & ctr) & ") by " & Request.Cookies("LBUsrName") & "."
									LogMe.WriteLine strLog
									Set LogMe = Nothing
									Set fso = Nothing
								End If
							End If
						End If
						pageused = "reqtable2.asp(TS)"
					ElseIf Request("ctrlX") = 2 Then
						rsTBL("overmile") = False
						If Request("chkOverMile" & ctr) <> "" Then rsTBL("overmile") = True
						rsTBL("actMil") = Z_Czero(Request("txtmile" & ctr))
						rsTBL("toll") = Z_Czero(Request("txtTol" & ctr))
						'If Request("chkpay" & ctr) = 1 Then 
						'	rsTBL("PayIntr") = True
						'Else
						'	rsTBL("PayIntr") = False
						'	rsTBL("actTT") = 0
						'	rsTBL("actMil") = 0
						'	rsTBL("TT_Intr") = 0
						'	rsTBL("M_Intr") = 0
						'End If
						If Not rsTBL("LbconfirmToll") Then
							rsTBL("LbconfirmToll") = False
							If Request("chkM" & ctr) <> "" Then
								rsTBL("LbconfirmToll") = True
								Set fso = CreateObject("Scripting.FileSystemObject")
								Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
								strLog = Now & vbTab & "Mileage confirmed (ID: " & tmpID & ") by " & Request.Cookies("LBUsrName") & "."
								LogMe.WriteLine strLog
								Set LogMe = Nothing
								Set fso = Nothing
							End If
						End If
						pageused = "reqtable2.asp(M)"
					End If
					rsTBL.Update
					If SaveHist(tmpID, pageused) Then
	
						End If
				End If
			End If
		Next
	End If
	rsTBL.Close
	Set rsTBL = Nothing
	Response.Redirect "reqtable2.asp?radioStat=" & Request("radioStat") & "&txtFromd8=" & Request("txtFromd8") & "&txtTod8=" & Request("txtTod8") & _
		"&txtFromID=" & Request("txtFromID") & "&txtToID=" & Request("txtToID") & "&selInst=" & Request("selInst") & "&selLang=" & Request("selLang") & "&tmpclilname=" & Request("txtclilname") & "&tmpclifname=" & Request("txtclifname") & _
		"&selIntr=" & Request("selIntr") & "&selClass=" & Request("selClass") & "&selAdmin=" & Request("selAdmin") & "&action=3&ctrlX=" & Request("ctrlX")
ElseIf Request("ctrl") = 18 Then 'save timsheet/mileage
	Set rsMain = Server.CreateObject("ADODB.RecordSet")
	sqlMain = "SELECT TT_Intr, M_Intr, actMil, actTT, InstActMil,InstActTT, RealTT, RealM FROM request_T WHERE [index] = " & Request("ReqID")
	rsMain.Open sqlMain, g_strCONN, 1, 3
	If Not rsMain.EOF Then
		rsMain("actMil") = Z_CZero(Request("txtMile"))
		rsMain("actTT") = Z_CZero(Request("txtTravel"))
		rsMain("InstActMil") = Z_CZero(Request("txtMileInst"))
		rsMain("InstActTT") = Z_CZero(Request("txtTravelInst"))
		rsMain("RealTT") = Z_CZero(Request("txtRTravel"))
		rsMain("RealM") = Z_CZero(Request("txtRMile"))
		rsMain("TT_Intr") = 0
		rsMain("M_Intr") = 0
		rsMain.Update
	End If
	rsMain.Close
	Set rsMain = Nothing
	Session("MSG") = "Travel Time and Mileage Saved."
	If SaveHist(Request("ReqID"), "reqconfirm.asp") Then
	
	End If
	Response.Redirect "reqconfirm.asp?ID=" & Request("ReqID")
ElseIf Request("ctrl") = 19 Then'apprve hrs bill inst
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT * FROM request_T"
	If Request("txtFromd8") <> "" Then
		If IsDate(Request("txtFromd8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' "
		End If
	End If
	If Request("txtTod8") <> "" Then
		If IsDate(Request("txtTod8")) Then
			sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
			sqlTBL = "SELECT * FROM request_T WHERE appDate <= '" & Request("txtTod8") & "' "
		End If
	End If
	If Request("txtFromd8") <> "" And Request("txtTod8") <> "" Then
		If IsDate(Request("txtFromd8")) And IsDate(Request("txtTod8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' AND appDate <= '" & Request("txtTod8") & "' "
		End If
	End If
	rsTBL.Open sqlTBL, g_strCONN, 1, 3 
	strNow = Now
	If Not rsTBL.EOF Then 
		y = Request("Hctr")
		For ctr = 1 To y - 1
			tmpID = Request("chkM" & ctr)
			If tmpID <> "" Then
				tmpIndex = "Index= " & tmpID
				rsTBL.MoveFirst
				rsTBL.Find(tmpIndex)
				If Not rsTBL.EOF Then
						If Request("selInstRate" & ctr) <> 0 And Request("txtbilHrs" & ctr) <> "" Then
							'response.write "> " & Request("selInstRate" & ctr) & " ---- " & Request("txtbilHrs" & ctr) & " ____ " & Request("ctrlX")
							If rsTBL("InstRate") <> 0 And NOT isNull(rsTBL("Billable")) Then 'not saving med due to Billable1
								If Request("ctrlX") = 1 Then
									rsTBL("ApproveHrs") = True
									Set fso = CreateObject("Scripting.FileSystemObject")
									Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
									strLog = Now & vbTab & "Institution Hours approved (ID: " & tmpID & ") by " & Request.Cookies("LBUsrName") & "."
									LogMe.WriteLine strLog
									Set LogMe = Nothing
									Set fso = Nothing
								Else
									'rsTBL("vermed") = True
									'Set fso = CreateObject("Scripting.FileSystemObject")
									'Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
									'strLog = Now & vbTab & "Medicaid approved (ID: " & tmpID & ") by " & Session("UsrName") & "."
									'LogMe.WriteLine strLog
									'Set LogMe = Nothing
									'Set fso = Nothing
								End If
								If GetStatus(tmpID) = 4 Then
									' do nothing
								Else
									rsTBL("Status") = 1
									rsTBL("completed") = strNow
								End If
							End If
							rsTBL.Update
						End If
				End If
			End If
		Next
	End If
	rsTBL.Close
	Set rsTBL = Nothing
	Session("MSG") = "Checked Entries Approved."
	Response.Redirect "reqtable3.asp?radioStat=" & Request("radioStat") & "&txtFromd8=" & Request("txtFromd8") & "&txtTod8=" & Request("txtTod8") & _
		"&txtFromID=" & Request("txtFromID") & "&txtToID=" & Request("txtToID") & "&selInst=" & Request("selInst") & "&selLang=" & Request("selLang") & "&tmpclilname=" & Request("txtclilname") & "&tmpclifname=" & Request("txtclifname") & _
		"&selIntr=" & Request("selIntr") & "&selClass=" & Request("selClass") & "&selAdmin=" & Request("selAdmin") & "&action=3&ctrlX=" & Request("ctrlX")
ElseIf Request("ctrl") = 20 Then'save bill inst
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT * FROM request_T WHERE "
		If Request("txtFromd8") <> "" Then
			If IsDate(Request("txtFromd8")) Then
				sqlTBL = sqlTBL & "appDate >= '" & Request("txtFromd8") & "' AND "
			End If
		End If
		If Request("txtTod8") <> "" Then
			If IsDate(Request("txtTod8")) Then
				'sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
				sqlTBL = sqlTBL & "appDate <= '" & Request("txtTod8") & "' AND "
			End If
		End If
		If Request("txtFromd8") <> "" And Request("txtTod8") <> "" Then
			If IsDate(Request("txtFromd8")) And IsDate(Request("txtTod8")) Then
				sqlTBL = sqlTBL & "appDate >= '" & Request("txtFromd8") & "' AND appDate <= '" & Request("txtTod8") & "' AND "
			End If
		End If
		sqlTBL = sqlTBL &  "Processed IS NULL AND NOT AStarttime IS NULL AND NOT AEndtime IS NULL AND instID <> 479"
	rsTBL.Open sqlTBL, g_strCONN, 1, 3 
	If Not rsTBL.EOF Then 
		y = Request("Hctr")
		For ctr = 1 To y - 1
			tmpID = Request("ID" & ctr)
			tmpIndex = "Index= " & tmpID
			rsTBL.MoveFirst
			rsTBL.Find(tmpIndex)
			If Not rsTBL.EOF Then
				rsTBL("InstRate") = Z_Czero(Request("selInstRate" & ctr))
				If Request("chkBilInst" & ctr) = 1 Then
					rsTBL("BillInst") = True
					rsTBL("TTRate") = Z_Czero(Request("txtTTrate" & ctr))
					rsTBL("MRate") = Z_Czero(Request("txtmrate" & ctr))
				Else
					rsTBL("BillInst") = False
					rsTBL("TTRate") = 0
					rsTBL("MRate") = 0
				End If
				rsTBL("TT_Inst") = Z_Czero(Request("txtTT" & ctr))
				rsTBL("M_Inst") = Z_Czero(Request("txtMile" & ctr))
				rsTBL("overTTInst") = False
				If Request("chkOverTT" & ctr) = 1 Then rsTBL("overTTInst") = True
				rsTBL("overMInst") = False
				If Request("chkOverM" & ctr) = 1 Then rsTBL("overMInst") = True	
				rsTBL("Billable") = Z_Czero(Request("txtbilHrs" & ctr)) 
				If SaveHist(tmpID, "reqtable3.asp") Then
				 'Session("MSG") = "HIST SAVED"
				End If
				rsTBL.Update
	
			End If
		Next
	End If
	rsTBL.Close
	Set rsTBL = Nothing
	Session("MSG") = "Entries Saved."
	Response.Redirect "reqtable3.asp?radioStat=" & Request("radioStat") & "&txtFromd8=" & Request("txtFromd8") & "&txtTod8=" & Request("txtTod8") & _
		"&txtFromID=" & Request("txtFromID") & "&txtToID=" & Request("txtToID") & "&selInst=" & Request("selInst") & "&selLang=" & Request("selLang") & "&tmpclilname=" & Request("txtclilname") & "&tmpclifname=" & Request("txtclifname") & _
		"&selIntr=" & Request("selIntr") & "&selClass=" & Request("selClass") & "&selAdmin=" & Request("selAdmin") & "&action=3&ctrlX=" & Request("ctrlX")
ElseIf Request("ctrl") = 21 Then'approve 604a
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT approvePDF, HPID FROM Request_T WHERE [index] = " & Request("h_ID")
	rsTBL.Open sqlTBL, g_strCONN, 1, 3
	If Not rsTBL.EOF Then
		rsTBL("approvePDF") = True
		rsTBL.Update
		tmpHPID = rsTBL("HPID")
	End If
	rsTBL.Close
	Set rsTBL = Nothing
	If SaveHist(Request("h_ID"), "f603a.asp") Then
	 'Session("MSG") = "HIST SAVED"
	End If
	If tmpHPID > 0 Then
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT approvePDF FROM Appointment_T WHERE [index] = " & tmpHPID
		rsTBL.Open sqlTBL, g_strCONNHP, 1, 3
		If Not rsTBL.EOF Then
			rsTBL("approvePDF") = True
			rsTBL.Update
		End If
		rsTBL.Close
	End If
	Session("MSG") = "Form 604A Apporved."
	Response.Redirect "f603A.asp?ID=" & Request("h_ID") 
ElseIf Request("ctrl") = 22 Then 'Approve Medicaid
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT * FROM request_T"
	If Request("txtFromd8") <> "" Then
		If IsDate(Request("txtFromd8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' AND autoacc <> 1 AND wcomp <> 1 AND (medicaid <> '' OR NOT medicaid IS NULL OR meridian <> '' OR NOT meridian IS NULL)"
		End If
	End If
	If Request("txtTod8") <> "" Then
		If IsDate(Request("txtTod8")) Then
			sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
			sqlTBL = "SELECT * FROM request_T WHERE appDate <= '" & Request("txtTod8") & "'  AND autoacc <> 1 AND wcomp <> 1 AND (medicaid <> '' OR NOT medicaid IS NULL OR meridian <> '' OR NOT meridian IS NULLL)"
		End If
	End If
	If Request("txtFromd8") <> "" And Request("txtTod8") <> "" Then
		If IsDate(Request("txtFromd8")) And IsDate(Request("txtTod8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' AND appDate <= '" & Request("txtTod8") & "'  AND autoacc <> 1 AND wcomp <> 1 AND (medicaid <> '' OR NOT medicaid IS NULL OR meridian <> '' OR NOT meridian IS NULL)"
		End If
	End If
rsTBL.Open sqlTBL, g_strCONN, 1, 3 
ts = now
If Not rsTBL.EOF Then 
	y = Request("Hctr")
	For ctr = 1 To y - 1
		tmpID = Request("ID" & ctr)
		tmpIndex = "Index= " & tmpID
		rsTBL.MoveFirst
		rsTBL.Find(tmpIndex)
		If Not rsTBL.EOF Then
			rsTBL("vermed") = 0
			If Request("chkM" & ctr) <> "" Then 
				syscomstr = Z_fixNull(rsTBL("syscom"))
				syscomstr = Replace(syscomstr, "<br>Medicaid approved.", "")
				syscomstr = Replace(syscomstr, "<br>Medicaid disapproved.", "")
				rsTBL("syscom") = syscomstr & "<br>Medicaid approved." 
				rsTBL("vermed") = 1
			End If
			If Request("chkX" & ctr) <> "" Then 
				rsTBL("vermed") = 2
				If Z_Czero(Instr(rsTBL("syscom"), "Medicaid disapproved.")) = 0 Then rsTBL("syscom") = rsTBL("syscom") & "<br>Medicaid disapproved." 
				syscomstr = Z_fixNull(rsTBL("syscom"))
				syscomstr = Replace(syscomstr, "<br>Medicaid approved.", "")
				syscomstr = Replace(syscomstr, "<br>Medicaid disapproved.", "")
				rsTBL("syscom") = syscomstr & "<br>Medicaid disapproved." 
			End If
			rsTBL.Update
			If rsTBL("vermed") = 1 Then
				clilname = CleanMeSingleQuote(Ucase(Trim(rsTBL("clname"))))
				clifname = CleanMeSingleQuote(Ucase(Trim(rsTBL("cfname"))))
				'medicaid = Ucase(Trim(rsTBL("medicaid")))
				'If medicaid = "" Then medicaid = Ucase(Trim(rsTBL("meridian")))
				'If medicaid = "" Then medicaid = Ucase(Trim(rsTBL("nhhealth")))
				'If medicaid = "" Then medicaid = Ucase(Trim(rsTBL("wellsense")))
				medicaid = Z_FixNull(Ucase(Trim(rsTBL("medicaid")))) 
				If Z_FixNull(rsTBL("meridian")) <> "" Then 
					medicaid = Z_FixNull(Ucase(Trim(rsTBL("meridian"))))
				End If
				If Z_FixNull(rsTBL("nhhealth")) <> "" Then 
					medicaid = Z_FixNull(Ucase(Trim(rsTBL("nhhealth"))))
				End If
				If Z_FixNull(rsTBL("wellsense")) <> "" Then 
					medicaid = Z_FixNull(Ucase(Trim(rsTBL("wellsense")))) 
				End If
				dob = Trim(rsTBL("dob"))
				gender = 0
				If rsTBL("gender") Then gender = 1
				Set rsCli = Server.CreateObject("ADODB.RecordSet")
				sqlCli = "SELECT * FROM clientuploaded_T WHERE lname = '" & clilname & "' AND fname = '" & clifname & "' AND medicaid = '" & medicaid & _
					"' AND dob = '" & dob & "' AND gender = " & gender 
				rsCli.Open sqlCli, g_strCONN, 3, 1
				If rsCli.EOF Then
					Set rsCliList = Server.CreateObject("ADODB.RecordSet")
					rsCliList.Open "SELECT * FROM clientList_T", g_strCONN, 1, 3
					rsCliList.AddNew
					rsCliList("lname") = Ucase(Trim(rsTBL("clname")))
					rsCliList("fname") = Ucase(Trim(rsTBL("cfname")))
					rsCliList("medicaid") = medicaid
					rsCliList("dob") = dob
					rsCliList("gender") = False
					If gender = 1 Then rsCliList("gender") = True
					rsCliList("timestamp") = ts
					rsCliList.Update
					rsCliList.Close
					Set rsCliList = Nothing
					'Set fso = CreateObject("Scripting.FileSystemObject")
					'Set oFileNew = fso.OpenTextFile(clientList, 8)
					'oFileNew.Write """" & cliname & """,""" & clifname & """,""" & medicaid & """,""" & dob & """,""" & gender & """"
					'oFileNew.Close
					'Set oFileNew = Nothing
					'Set fso = Nothing
				End If
				rsCli.Close
				Set rsCli = Nothing
				
			End If
			If SaveHist(tmpID, "reqtable4.asp") Then
			 'Session("MSG") = "HIST SAVED"
			End If
		End If
	Next
End If
rsTBL.Close
Set rsTBL = Nothing
Session("MSG") = "Medicaid Approved/Disapproved."
Response.Redirect "reqtable4.asp?radioStat=" & Request("radioStat") & "&txtFromd8=" & Request("txtFromd8") & "&txtTod8=" & Request("txtTod8") & _
	"&txtFromID=" & Request("txtFromID") & "&txtToID=" & Request("txtToID") & "&selInst=" & Request("selInst") & "&selLang=" & Request("selLang") & "&tmpclilname=" & Request("txtclilname") & "&tmpclifname=" & Request("txtclifname") & _
	"&selIntr=" & Request("selIntr") & "&selClass=" & Request("selClass") & "&selAdmin=" & Request("selAdmin") & "&action=3"
ElseIf Request("ctrl") = 23 Then 'Medicaid check
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	sqlTBL = "SELECT * FROM request_T WHERE autoacc <> 1 AND wcomp <> 1 AND (medicaid <> '' OR NOT medicaid IS NULL)"
	If Request("txtFromd8") <> "" Then
		If IsDate(Request("txtFromd8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' AND autoacc <> 1 AND wcomp <> 1 AND (medicaid <> '' OR NOT medicaid IS NULL)"
		End If
	End If
	If Request("txtTod8") <> "" Then
		If IsDate(Request("txtTod8")) Then
			'sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
			sqlTBL = "SELECT * FROM request_T WHERE appDate <= '" & Request("txtTod8") & "'  AND autoacc <> 1 AND wcomp <> 1 AND (medicaid <> '' OR NOT medicaid IS NULL)"
		End If
	End If
	If Request("txtFromd8") <> "" And Request("txtTod8") <> "" Then
		If IsDate(Request("txtFromd8")) And IsDate(Request("txtTod8")) Then
			sqlTBL = "SELECT * FROM request_T WHERE appDate >= '" & Request("txtFromd8") & "' AND appDate <= '" & Request("txtTod8") & "'  AND autoacc <> 1 AND wcomp <> 1 AND (medicaid <> '' OR NOT medicaid IS NULL)"
		End If
	End If
	rsTBL.Open sqlTBL, g_strCONN, 1, 3 
	If Not rsTBL.EOF Then 
		dte = FormatDateTime(date, 2)
		dteYr = Year(dte)
		dteYr2 = Right("0" & Year(dte), 2)
		dteMn = Right("0" & Month(dte), 2)
		dteDy = Right("0" & Day(dte), 2)
		dtetime = FormatDateTime(Now, 4)
		tme = Replace(dtetime, ":", "")
		strMedHdr = "ISA*00*          *00*          *ZZ*NH100496       *ZZ*026000618      *" & dteYr2 & dteMn & dteDy & "*" & tme & "*^*00501*000000007*0*P*:~" & _ 
			"GS*HS*NH100496*026000618*" & dteYr & dteMn & dteDy & "*" & tme & "*1*X*005010X279A1~"
		y = Request("Hctr")
		For ctr = 1 To y - 1
			tmpID = Request("ID" & ctr)
			tmpIndex = "Index= " & tmpID
			rsTBL.MoveFirst
			rsTBL.Find(tmpIndex)
			If Not rsTBL.EOF Then
				'get med info
				dte = FormatDateTime(date, 2)
				dteYr = Year(dte)
				dteYr2 = Right("0" & Year(dte), 2)
				dteMn = Right("0" & Month(dte), 2)
				dteDy = Right("0" & Day(dte), 2)
				dtetime = FormatDateTime(Now, 4)
				tme = Replace(dtetime, ":", "")
				medicaid = Ucase(Trim(rsTBL("medicaid")))
				If medicaid = "" Then medicaid = Ucase(Trim(rsTBL("meridian")))
				If medicaid = "" Then medicaid = Ucase(Trim(rsTBL("nhhealth")))
				If medicaid = "" Then medicaid = Ucase(Trim(rsTBL("wellsense")))
				'change to intepreter - "NM1*1P*1*" & Z_IntrInfo(rsTBL("intrID")) & "~" & _
				STnum = Right("0000" & ctr, 4)
				strMedBody = strMedBody & "ST*270*" & STnum & "*005010X279A1~" & _
					"BHT*0022*13*10001234*" & dteYr & dteMn & dteDy & "*" & tme & "~" & _
					"HL*1**20*1~" & _
					"NM1*PR*2*NH MEDICAID*****PI*026000618~" & _
					"HL*2*1*21*1~" & _
					"NM1*1P*1*OMERBEGOVIC*ALEN****SV*30849597~" & _
					"REF*EO*820000243~" & _
					"HL*3*2*22*0~" & _
					"NM1*IL*1*" & Z_NameMed(rsTBL("index")) & "****MI*" & medicaid & "~" & _
					"DMG*D8*" & Z_DOBMed(rsTBL("index")) & "*" & Z_GenderMed(rsTBL("index")) & "~" & _
					"DTP*291*RD8*" & Z_DateMed(rsTBL("index")) & "-" & Z_DateMed(rsTBL("index")) & "~" & _
					"EQ*30~" & _
					"SE*13*" & STnum & "~"
			End If
		Next
		strMedFtr = "GE*" & ctr - 1 & "*1~IEA*1*000000007~"
	End If
	rsTBL.Close
	Set rsTBL = Nothing
	strMed = Trim(strMedHdr & strMedBody & strMedFtr)
	response.write strMed
	'CREATE x12
	tmpdate = replace(date, "/", "") 
	tmpTime = replace(FormatDateTime(time, 3), ":", "")
	Repx12 =  "x12270-" & tmpdate & tmpTime & ".x12" 
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set ofilex12 = fso.CreateTextFile(x12Path & Repx12, 8, True) 
	ofilex12.Write strMed
	Set ofilex12 = Nothing
	fso.CopyFile x12Path & Repx12, x12pathbackup
	Set fso = Nothing
	Response.Redirect "reqtable4.asp?radioStat=" & Request("radioStat") & "&txtFromd8=" & Request("txtFromd8") & "&txtTod8=" & Request("txtTod8") & _
	"&txtFromID=" & Request("txtFromID") & "&txtToID=" & Request("txtToID") & "&selInst=" & Request("selInst") & "&selLang=" & Request("selLang") & "&tmpclilname=" & Request("txtclilname") & "&tmpclifname=" & Request("txtclifname") & _
	"&selIntr=" & Request("selIntr") & "&selClass=" & Request("selClass") & "&selAdmin=" & Request("selAdmin") & "&action=3" & "&dload=1" & "&filname=" & Repx12
ElseIf Request("ctrl") = 24 Then 'finance billing change 'TODO: FIX medicaid info
	Set rsBill = Server.CreateObject("ADODB.RecordSet")
	rsBill.Open "SELECT * FROM request_T WHERE [index] = " & request("ReqID"), g_strCONN, 1, 3
	If Not rsBill.EOF Then
		If Request("radioFin") = 1 Then
			rsBill("med_billInst") = Request("med_billInst")
			'rsBill("status") = 0
			rsBill("med_billmed") = Empty
			rsBill("billmedreas") = ""
			rsBill("processedMedicaid") = Empty
			rsBill("vermed") = 2
			rsBill("billingTrail") = rsBill("billingTrail") & "<br>Billing change to Institution " & Request("med_billInst")
		ElseIf Request("radioFin") = 2 Then
			rsBill("vermed") = 0
			rsBill("med_billmed") = Request("med_billmed")
			rsBill("billmedreas") = Request("billmedreas")
			rsBill("med_billInst") = Empty
			rsBill("Processed") = Empty
			rsBill("billingTrail") = rsBill("billingTrail") & "<br>Billing change to Medicaid " & Request("med_billmed")
		End If
		rsBill.Update
		If SaveHist(request("ReqID"), "reqconfirm.asp") Then
		 'Session("MSG") = "HIST SAVED"
		End If
	End If
	rsBill.Close
	Set rsBill = Nothing
	Session("MSG") = "Billing status saved."
	Response.Redirect "reqconfirm.asp?ID=" & request("ReqID")
ElseIf Request("ctrl") = 25 Then 'nohours appointment
	arrIntr = Split(Request("selIntr"), ",")
	x = 0
	tmpNow = Now
	Do Until x = Ubound(arrIntr) + 1
		Set rsApp = Server.CreateObject("ADODB.RecordSet")
		sqlApp = "SELECT * FROM request_T WHERE timestamp = '" & date & "'"
		rsApp.Open sqlApp, g_strCONN, 1, 3
		rsApp.AddNew
		rsApp("timestamp") = tmpNow
		rsApp("reqID") = Request("selRP")
		rsApp("appdate") = Request("txtAppDate")
		rsApp("apptimefrom") = Request("txtAppDate") & " " & Request("txtAppTFrom")
		rsApp("apptimeto") = Request("txtAppDate") & " " & Request("txtAppTTo")
		rsApp("langID") = 95 'change to live
		rsApp("clname") = GetReqlname(Request("selRP"))
		rsApp("cfname") = GetReqfname(Request("selRP"))
		rsApp("InstID") = 479
		rsApp("InstRate") = 1
		rsApp("IntrRate") = 20
		rsApp("deptID") = Request("selDept")
		rsApp("IntrID") = arrIntr(x)
		rsApp("training") = Request("selTrain")
		rsApp.Update
		rsApp.Close
		Set rsApp = Nothing
		x = x + 1
	Loop
	Session("MSG") = "Appointment Saved."
	Response.Redirect "nohours.asp"
ElseIf Request("ctrl") = 26 Then 'today
	Response.Cookies("LBREPORT") = Z_DoEncrypt(100)
'	Response.Cookies("LBREPORT").Expires = Now + 0.34
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		response.redirect "calendarview2.asp?mytype=1&rep=26&tmpdate='" & Request("tmpDate") & "' "
	Else
		response.redirect "calendarview2.asp?mytype=1&rep=26&tmpdate='" & Request("tmpDate") & "'&tmpRP=" & Request.Cookies("LBUSERTYPE")
	End If
ElseIf Request("ctrl") = 27 Then 'Remove inter
	Set rsRem = Server.CreateObject("ADODB.RecordSet")
	rsRem.Open "UPDATE request_T SET intrid = 0 WHERE [index] = " & Request("ReqID"), g_strCONN, 1, 3
	Set rsRem = Nothing
	If GetReqHPID2(Request("ReqID")) > 0 Then 'remove in HP
		Set rsRem = Server.CreateObject("ADODB.RecordSet")
		rsRem.Open "UPDATE appointment_T SET intrid = 0 WHERE [index] = " & GetReqHPID2(Request("ReqID")), g_strCONNHP, 1, 3
		Set rsRem = Nothing
	End If
	Call SaveHist(request("ReqID"), "reqconfirm.asp")
	Session("MSG") = "Interpreter Removed."
	Response.Redirect "reqconfirm.asp?ID=" & Request("ReqID") 
ElseIf Request("ctrl") = 28 Then 'cancel appt
	'SET STATUS inLB
	Set rsStat = Server.CreateObject("ADODB.RecordSet")
	sqlStat = "SELECT appTimeFrom, appTimeTo, intrID, deptID, InstID, status, Cancel, Missed, Astarttime, Aendtime, payHrs, Billable, happen, " & _
		"payintr, M_Intr, TT_Intr, LBcomment, HPID, appdate, clname, cfname FROM request_T WHERE [index] = " & Request("ID")
	rsStat.Open sqlStat, g_strCONN, 1, 3
	If Not rsStat.EOF Then
		tmpAppDateTime = rsStat("appTimeFrom")
		tmpAppDateTime2 = rsStat("appTimeTo")
		tmpIntr = Z_CZero(rsStat("intrID"))
		myClass = ClassInt(rsStat("deptID"))
		tmpInstID = rsStat("InstID")
		tmpInst = GetInst(rsStat("InstID"))
		tmpDate = rsStat("appDate")
		tmpTime = rsStat("appTimeFrom") & " - " & rsStat("appTimeTo")
		tmpTimeFrom = rsStat("appTimeFrom")
		tmpTimeTo = rsStat("appTimeTo")
		tmpDept = GetDept(rsStat("deptID"))
		tmpName = rsStat("clname") & ", " & rsStat("cfname")
		tmpCity = GetCity(rsStat("DeptID"))
		tmpFname = rsStat("cfname")
		If tmpIntr > 0  Then IntrName = GetIntr2(tmpIntr)
		If DateDiff("n", Now, tmpAppDateTime) < 1440 and tmpIntr > 0 Then
			rsStat("status") = 4
			rsStat("Cancel") = 5
			rsStat("Missed") = 0
			rsStat("Astarttime") = tmpAppDateTime
			rsStat("Aendtime") = tmpAppDateTime
			rsStat("payHrs") = 2
			rsStat("Billable") = 2
			If tmpInstID = 273 Or myClass = 3 Then 'darth leb - court
				rsStat("Billable") = Z_GetBillhrs(tmpAppDateTime, tmpAppDateTime2)
			End If
			rsStat("happen") = 1
			rsStat("payintr") = false
			rsStat("M_Intr") = 0
			rsStat("TT_Intr") = 0
		Else
			rsStat("status") = 3
			rsStat("Cancel") = 5
			rsStat("Missed") = 0
		End If
		If myClass = 3 And DateDiff("n", Now, tmpAppDateTime) < 2880 And DateDiff("n", Now, tmpAppDateTime) > 1440 and tmpIntr > 0 Then 'courts and 24-48 hours cancel
			rsStat("status") = 4
			rsStat("Cancel") = 5
			rsStat("Missed") = 0
			rsStat("Astarttime") = tmpAppDateTime
			rsStat("Aendtime") = tmpAppDateTime
			rsStat("payHrs") = 0
			rsStat("showIntr") = 0
			rsStat("Billable") = Z_GetBillhrsCourt(tmpAppDateTime, tmpAppDateTime2)	
		End If
		If tmpIntr > 0 Then 
			IntrName = GetIntr2(tmpIntr)
			rsStat("LBcomment") = rsStat("LBcomment") & vbCrlF & "Cancelation Email sent to " & IntrName & " on " & now
		Else
			rsStat("LBcomment") = rsStat("LBcomment") & vbCrlF & "Cancelation Email sent to Langbank on " & now
		End If
		
		rsStat.Update
		tmpHPID = rsStat("HPID")
	End If
	rsStat.Close
	Set rsStat = Nothing
	If tmpHPID > 0 Then
		'SET STATUS in HP
		Set rsStat = Server.CreateObject("ADODB.RecordSet")
		'sqlStat = "SELECT InstID, appDate, TimeFrom, TimeTo, deptID, clname, cfname, IntrID, DeptID  FROM Appointment_T WHERE [index] = " & tmpHPID
		sqlStat = "UPDATE Appointment_T SET [status] = 3 WHERE [index] = " & tmpHPID
		rsStat.Open sqlStat, g_strCONNHP, 1, 3
		Set rsStat = Nothing
	End If
	If SaveHist(Request("ID"), "reqconfirm.asp") Then
		'log edit
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile("C:\work\lss-lbis\log\editlogs.txt", 8, True)
		strLog = Now & vbtab & "Appointment Cancelled - ID: " & Request("ID") & " by " & Request.Cookies("LBUsrName") & " DATE: " & tmpDate & " TIME: " & tmpTime
		LogMe.WriteLine strLog
		Set LogMe = Nothing
		Set fso = Nothing
	End If
	'SEND EMAIL TO NOTIFY CANCEL TO LB
	strBody = "<img src='http:https://staff.thelanguagebank.org/lsslbis/images/LBISLOGOBandW.jpg'><br><br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Request ID: " & Request("ID") & " has been CANCELED by " & Request.Cookies("LBUsrName") & ".<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Date: " & tmpDate & "<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Time: " & tmpTime & "<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Department: " & tmpDept & "<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Client: " & tmpName & "<br><br>" & vbCrLf & _ 
	 "<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email.</font>" & vbCrLf
	
	retErr = zSendMessage("language.services@thelanguagebank.org", "", "Request Cancellation - Request ID: " & tmpLBID, strBody)	 
	 If tmpIntr > 0  Then
	 	' SEND EMAIL TO NOTIFY CANCEL TO INTR
	 	strBody = "This is to let you know that appointment on " & _
			 tmpDate & ", " & tmpTime & ", in " & tmpCity & " at " & tmpInst & " for " & tmpFname & " is CANCELED.<br>" & _
			 "If you have any questions please contact the LanguageBank office immediately at 410-6183 or email us at " & _
			 "<a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>.<br>" & _
			 "E-mail about this cancelation was initiated by " & Request.Cookies("LBUsrName") & ".<br><br>" & _
			 "Thanks,<br>" & _
			 "Language Bank"
		retErr = zSendMessage(GetPrime2(tmpIntr), "language.services@thelanguagebank.org" _
				, "Appointment Cancellation " & tmpDate & "; " & tmpTime & ", " & tmpCity & " - " &  tmpInst _
				, strBody)
	End If
	Session("MSG") = "NOTICE: Request has been cancelled." 
	Response.Redirect "reqconfirm.asp?ID=" & Request("ID")
End If
%>
<!-- #include file="_closeSQL.asp" -->

<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
DIM tmpIntr(), tmpTown(), tmpIntrName(), tmpLang(), tmpClass(), tmpBill(), tmpAhrs(), tmpApp(), tmpInst(), tmpDept(), tmpAmt(), tmpFac(), tmpMonthYr(), tmpCtr(), tmpMonthYr2(), tmpMonthYr3()
DIM tmpMonthYr4(), tmpHrs(), tmpHHrs(), tmpMile(), tmpToll(), arrTS(), arrAuthor(), arrPage(), tmpTrain(), tmpbhrs()
server.scripttimeout = 360000
Function Z_MinRate()
	Z_MinRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT MinWage FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		Z_MinRate = rsRate("MinWage")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function AmtRate(xxx)
	AmtRate = 0
	If Z_Czero(xxx) = 0 Then
		AmtRate = 0
		Exit Function
	End If
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM MileageRate_T"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		AmtRate = rsRate("mileageRate") * xxx
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function
Function EFee(bln, myClass, EmerFeeL, EmerFeeO)
	If bln Then 
		If myClass = 3 Or MyClass = 5 Then
			EFee = EmerFeeL
		Else
			EFee = EmerFeeO
		End If
	Else
		EFee = 0
	End If
End Function
tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
'	tmpReport = (		Request("selRep")
'						Request("txtRepFrom")
'						Request("txtRepTo")	
'						Request("selInst")	
'						Request("selIntr")	
'						Request("selTown")
'						Request("selLang")
'						Request("selCli")
'						Request("selClass")
'						Request("chkAddnl")
'						Request("selIntrStat"))
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")
ctr = 13
If tmpReport(0) = 99 Then
	RepCSV =  "InstCalendar" & tmpdate & ".csv" 
	If Request("Hmonth") <> "" Then Response.Cookies("LB-CALENDAR") = Request("Hmonth")
	myMonth = Request.Cookies("LB-CALENDAR")
	myRP = GetInst(Session("UInst"))
	If myRP = "N/A" Then	
		strMSG = "Interpreter	request for the month of " & myMonth & "." 
	Else
		strMSG = "Interpreter	request for the month of " & myMonth & " for " & myRP & "."
	End If
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	tmpMonth = Replace(myMonth, " - ", " 1, ")
	tmpMonth = Replace(tmpMonth, "'", "")
	tmpMonth = Month(tmpMonth)
	If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
		strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Time</td>" & vbCrlf & _
			"<td class='tblgrn'>Client</td>" & vbCrlf & _
			"<td class='tblgrn'>Language</td>" & vbCrlf & _
			"<td class='tblgrn'>Interpreter</td>" & vbCrlf
		CSVHead = "Date, Time, Client First Name, Client Last Name, Language, Interpreter Last Name, Interpreter First Name"
	Else
		strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Time</td>" & vbCrlf & _
			"<td class='tblgrn'>Language</td>" & vbCrlf & _
			"<td class='tblgrn'>Interpreter</td>" & vbCrlf
		CSVHead = "Date, Time,Language, Interpreter Last Name, Interpreter First Name"
	End If	
	
	
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		sqlRep = "SELECT * FROM request_T WHERE request_T.[instID] <> 479 AND Month(appDate) = " & tmpMonth & " ORDER BY appDate, appTimeFrom"
	Else
		sqlRep = "SELECT * FROM request_T WHERE request_T.[instID] <> 479 AND InstID = " & Session("UInst") & " AND Month(appDate) = " & tmpMonth & " ORDER BY appDate, appTimeFrom"
	End If
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appTimeFrom") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("CLname") & ", " & rsRep("CFname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetStat(rsRep("status")) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("appDate") & "," & rsRep("appTimeFrom") & "," & rsRep("CLname") & "," & _
					rsRep("CFname") & "," & GetLang(rsRep("LangID")) & "," & GetIntr(rsRep("IntrID")) & "," & GetStat(rsRep("status")) & vbCrLf
		Else
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appTimeFrom") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetStat(rsRep("status")) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("appDate") & "," & rsRep("appTimeFrom") & "," & GetLang(rsRep("LangID")) & "," & GetIntr(rsRep("IntrID")) & "," & GetStat(rsRep("status")) & vbCrLf
		End If
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 100 Then
	RepCSV =  "InstCalendarToday" & tmpdate & ".csv" 
	If Request("Hdate") <> "" Then Response.Cookies("LB-CALENDARDATE") = Request("HDate")
	myDate = Request.Cookies("LB-CALENDARDATE")
	myRP = GetInst(Session("UInst"))
	If myRP = "N/A" Then	
		strMSG = "Interpreter	request for " & myDate & "." 
	Else
		strMSG = "Interpreter	request for " & myDate & " for " & myRP & "."
	End If
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	myDate = Replace(myDate, "'", "")
	'tmpMonth = Replace(tmpMonth, "'", "")
	'tmpMonth = Month(tmpMonth)
	If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
		strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Time</td>" & vbCrlf & _
			"<td class='tblgrn'>Institution</td>" & vbCrlf & _
			"<td class='tblgrn'>Client</td>" & vbCrlf & _
			"<td class='tblgrn'>Language</td>" & vbCrlf & _
			"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
			"<td class='tblgrn'>Status</td>" & vbCrlf
		CSVHead = "Date, Time, Institution-Department, Client First Name, Client Last Name, Language, Interpreter Last Name, Interpreter First Name, Status"
	Else
		strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Time</td>" & vbCrlf & _
			"<td class='tblgrn'>Institution</td>" & vbCrlf & _
			"<td class='tblgrn'>Language</td>" & vbCrlf & _
			"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
			"<td class='tblgrn'>Status</td>" & vbCrlf
		CSVHead = "Date, Time, Institution, Language, Interpreter Last Name, Interpreter First Name, Status"
	End If	
	
	If Request.Cookies("LBUSERTYPE") <> 2 Then
		sqlRep = "SELECT * FROM request_T WHERE request_T.[instID] <> 479  AND appDate = '" & myDate & "' ORDER BY appDate, appTimeFrom"
	Else
		sqlRep = "SELECT * FROM request_T WHERE request_T.[instID] <> 479  AND InstID = " & Session("UInst") & " AND appDate = '" & myDate & "' ORDER BY appDate, appTimeFrom"
	End If
	If Request.Cookies("LBUSERTYPE") <> 2 And Z_CZero(Request("mytype")) = 1 Then
		sqlRep = "SELECT * FROM request_T WHERE request_T.[instID] <> 479  AND IntrID > 0 AND status <> 2 AND status <> 3 AND Status <> 4 AND appDate = '" & myDate & "' ORDER BY appDate, appTimeFrom"
	ElseIf Request.Cookies("LBUSERTYPE") = 2 And Z_CZero(Request("mytype")) = 1 Then
		sqlRep = "SELECT * FROM request_T WHERE request_T.[instID] <> 479  AND IntrID > 0 AND status <> 2 AND status <> 3 AND Status <> 4 AND InstID = " & Session("UInst") & " AND appDate = '" & myDate & "' ORDER BY appDate, appTimeFrom"
	End If
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		tmpcliadd = ""
		if rsRep("cliadd") = true then tmpcliadd = "*"
		tmpInsti = tmpcliadd & GetInst(rsRep("InstID"))
		If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom") )& "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpInsti & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("CLname") & ", " & rsRep("CFname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetStat(rsRep("status")) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("appDate") & "," & ctime(rsRep("appTimeFrom")) & ","  & tmpInsti & "," & rsRep("CLname") & "," & _
					rsRep("CFname") & "," & GetLang(rsRep("LangID")) & "," & GetIntr(rsRep("IntrID")) & "," & GetStat(rsRep("status")) & vbCrLf
		Else
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpInsti & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetStat(rsRep("status")) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("appDate") & "," & ctime(rsRep("appTimeFrom")) & "," & GetLang(rsRep("LangID")) & "," & GetIntr(rsRep("IntrID")) & "," & GetStat(rsRep("status")) & vbCrLf
		End If
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 1 Then
	RepCSV =  "InvReq" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Billable Hrs.</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Amount</td>" & vbCrlf 
	CSVHead = "Institution, Department, Language, Billable Hrs., Total Amount"
	sqlRep = "SELECT * FROM request_T, institution_T, language_T, Dept_T  WHERE request_T.[instID] <> 479  AND DeptID = dept_T.[index] And request_T.InstID = institution_T.[index] AND LangID = language_T.[index] "
	strMSG = "Invoice request report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If tmpReport(3) = "" Then tmpReport(3) = 0
	If tmpReport(3) <> 0 Then
		sqlRep = sqlRep & " AND institution_T.[index] = " & tmpReport(3) 
		strMSG = strMSG & " for " & GetInst(tmpReport(3))
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	'sqlRep = sqlRep & " AND showInst = 1 ORDER BY [Facility], [Dept],[Language]"
	sqlRep = sqlRep & " ORDER BY [Facility], [Dept],[Language]"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF
			strInst = rsRep("Facility")
			strDept = rsRep("Dept")
			strLang = rsRep("Language")
			strBill = rsRep("Billable")
			strAmt = rsRep("Billable") * rsRep("InstRate")
			lngIdx = SearchArraysInst2(strInst, strDept, strLang, tmpInst, tmpDept, tmpLang)
			If lngIdx < 0 Then 
					ReDim Preserve tmpInst(x)
					ReDim Preserve tmpDept(x)
					ReDim Preserve tmpLang(x)
					ReDim Preserve tmpBill(x)
					ReDim Preserve tmpAmt(x)
										
					tmpInst(x) = strInst
					tmpDept(x) = strDept
					tmpLang(x) = strLang
					tmpBill(x) = strBill
					tmpAmt(x) = strAmt
					x = x + 1
				Else
					tmpBill(lngIdx) = tmpBill(lngIdx) + strBill
					tmpAmt(lngIdx) = tmpAmt(lngIdx) + strAmt
				End If
			rsRep.MoveNext
		Loop
		y = 0
		Do Until y = x
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpInst(y) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpDept(y) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpLang(y) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpBill(y), 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpAmt(y), 2) & "</td></tr>" & vbCrLf 
								
			CSVBody = CSVBody & tmpInst(y) & "," & tmpDept(y) & "," & tmpLang(y) & "," & _
				tmpBill(y) & "," & tmpAmt(y) & vbCrLf
			y = y + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 2 Then
	RepCSV =  "CanRequest" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Status</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Classification</td>" & vbCrlf & _
		"<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Remarks</td>" & vbCrlf 
		CSVHead = "Status, Interpreter Last Name, Interpreter First Name, Language, Institution, Classification, Date, Remarks"
	sqlRep = "SELECT * FROM request_T, interpreter_T, dept_T WHERE request_T.[instID] <> 479  AND DeptID = dept_T.[index] AND Status = 3 AND IntrID = interpreter_T.[index]"
	strMSG = "Canceled request report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY Status, [Last Name], [First Name]"	
	strMSG = strMSG & "."
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			tmpStat = GetStat(rsRep("Status"))
			tmpIntName = rsRep("Last Name") & ", " & rsRep("First Name")
			tmpLng = GetLang(rsRep("LangID"))
			tmpInsti = GetInst(rsRep("request_T.InstID"))
			'tmpFac = tmpInsti
			tmpClas = GetClass(rsRep("Class"))
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("request_T.[index]") & ")'><td class='tblgrn2'><nobr>" & tmpStat & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpIntName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpLng & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpInsti & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpClas & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td></tr>" & vbCrLf
			'CSVBody = CSVBody & tmpStat & "," & tmpIntName & "," & tmpLng & "," & tmpInsti & "," & tmpClas & "," & rsRep("appDate") & ",""" & rsRep("Comment") & "" & vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='7' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 3 Then
	'INSTITUTION BILLING
	RepCSV =  "InstBillReq" & tmpdate & "-" & tmpTime & ".csv" 
	RepCSVBill = "InstBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
	RepCSVBillL = "InstBillReqNewL" & tmpdate & "-" & tmpTime & ".csv"
	RepCSVBillSigma = "InstBillReqNewSigma" & tmpdate & "-" & tmpTime & ".csv"
	
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf & _
		"<td class='tblgrn'>DOB</td>" & vbCrlf & _
		"<td class='tblgrn'>Customer ID</td>" & vbCrlf & _
		"<td class='tblgrn'>DHHS</td>" & vbCrlf 
	
	CSVHead = "Request ID,Institution, Department, Appointment Date, Client Last Name, Client First Name, Language, Interpreter Last Name, Interpreter First Name, Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Emergency Surcharge, Total, Comments, DOB, Customer ID, DHHS"	
	
	sqlRep = "SELECT DocNum, billingTrail, ReqID, HPID, syscom, processed, hasmed, outpatient, medicaid, vermed, autoacc, wcomp, drg, pid, request_T.[index] as myindex, status, [Last Name], [First Name], Clname, Cfname, AStarttime, AEndtime, " & _
		"Billable, DOB, emerFEE, class, TT_Inst, M_Inst, request_T.InstID as myinstID, DeptID, LangID, appDate, InstRate, bilComment, custID, ccode, billgroup, IntrID " & _
		"FROM request_T, interpreter_T , dept_T WHERE request_T.[instID] <> 479 AND request_T.deptID =  dept_T.[index] AND IntrID = interpreter_T.[index]  AND " & _
		"(Status = 1 OR Status = 4) AND Processed IS NULL "
	
	strMSG = "Institution Billing request report"
	
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & ". * - Cancelled Billable."
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY CustID ASC, AppDate DESC" '" ORDER BY AppDate DESC"
	rsRep.Open sqlRep, g_strCONN, 1, 3
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	If Not rsRep.EOF Then 
		x = 0
		tmpCID = ""
		Do Until rsRep.EOF
			BillHours = 0
			IncludeReq = True
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			CB = ""
			If rsRep("status") = 4 Then CB = "*"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
			'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
			BillHours =  rsRep("Billable")
			PayMed = PaytoMedicaid(rsRep("outpatient"), rsRep("hasmed"), rsRep("vermed"), rsRep("autoacc"), rsRep("wcomp"), rsRep("drg"), rsRep("IntrID"))
			If PayMed AND BillHours <= 4 Then IncludeReq = False
			If PayMed AND BillHours > 4 And rsRep("vermed") = 1 Then 
				BillHours = BillHours - 4
				If Instr(BillHours, ".") > 0 Then
					If Z_GetDec(BillHours) > 0 Then
						tmpMin = CDbl("0." & Z_GetDec(billhours))
						If tmpMin > 0 And tmpMin <= 0.25 Then billhours = Int(billhours) + 0.25
            If tmpMin > 0.25 And tmpMin <= 0.5 Then billhours = Int(billhours) + 0.5
            If tmpMin > 0.5 And tmpMin <= 0.75 Then billhours = Int(billhours) + 0.75
            If tmpMin > 0.75 And tmpMin <= 0.99 Then billhours = Int(billhours) + 1
					End If
				End If
				If Z_Czero(Instr(rsRep("syscom"), "hours billed to institution.")) = 0 Then rsRep("syscom") = rsRep("syscom") & "<br>" & billhours & " hours billed to institution. 4 hours billed to medicaid"
				'syscomstr = Z_fixNull(rsRep("syscom"))
				'syscomstr = Replace(syscomstr, "<br>Medicaid approved.", "")
				'syscomstr = Replace(syscomstr, "<br>Medicaid disapproved.", "")
				'rsRep("syscom") = syscomstr & "<br>Medicaid approved." 
			End If
			If PayMed And rsRep("vermed") = 0 Then IncludeReq = False
			If IncludeReq Then
				If rsRep("emerFEE") = True Then
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
					ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
						tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst")) + tmpFeeO
					End If
				Else
					tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
				End If
				totalPay = Z_FormatNumber(tmpPay, 2)
				strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
					"<td class='tblgrn2'><nobr>" & CB & rsRep("myindex") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetInst2(rsRep("myinstID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strIntrName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
					If rsRep("emerFEE") = True Then 
							If rsRep("class") = 3 Or rsRep("class") = 5 Then
								strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeL & "</td>" & vbCrLf
							Else
								strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
							End If
					Else
						strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
					End If
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
					If rsRep("emerFEE") = True Then 
						If rsRep("class") = 3 Or rsRep("class") = 5 Then
							strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
						ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
							strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeO & "</td>" & vbCrLf
						End If
					Else
						strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
					End If
					bilcomment = Z_fixNull(rsRep("bilComment") & rsRep("syscom"))
					strBody = strBody & "<td class='tblgrn2'><nobr><b>$" & totalPay & "</b></td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & bilcomment & "</td><td class='tblgrn2'><nobr>" & rsRep("DOB") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("CustID") & "</td>"
					If rsRep("myinstID") = 108 Then
						strBody = strBody & "<td class='tblgrn2'><nobr>" & GetUserID(rsRep("DeptID")) & "</td><tr>" & vbCrLf 
					Else
						strBody = strBody & "<td class='tblgrn2'><nobr>&nbsp;</td><tr>" & vbCrLf 
					End If
			
				CSVBody = CSVBody & CB & rsRep("myindex") & "," & GetInst2(rsRep("myinstID")) & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & rsRep("appDate") & "," & rsRep("Clname") & "," & rsRep("Cfname") &  "," & GetLang(rsRep("LangID")) & "," & rsRep("Last Name") & _
					"," & rsRep("First Name") & ","  & cTime(rsRep("AStarttime")) & "," & cTime(rsRep("AEndtime")) & "," & BillHours
					
				'''''MIP IMPORT'''''
				If rsRep("myinstID") <> 108 And rsRep("myinstID") <> 30 Then 'exclude DHHS and LSS
					'If rsRep("class") <> 3 Then
					tmpCID = Trim(UCase(rsRep("custID")))
					tmpccode = "LB " & rsRep("InstRate") & " Rate"
					If tmpCID <> tmpCID2 Or tmpCID2 = "" Then
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & "HOTC" & """,""" & rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
							Else
								CSVBodyBill = CSVBodyBill & """" & "HOTC" & """,""" & rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & "HOTC" & """,""" & rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
						End If
					End If	
					If rsRep("class") <> 5 Then
						If rsRep("myinstID") = 229 Or rsRep("myinstID") = 27 Or rsRep("myinstID") = 273 Or rsRep("myinstID") = 308 Or rsRep("myinstID") = 19 _
							Or rsRep("myinstID") = 300 Or rsRep("myinstID") = 427 Or rsRep("myinstID") = 289 Or rsRep("myinstID") = 268 Or rsRep("myinstID") = 265 _
							Or rsRep("myinstID") = 22 Or rsRep("myinstID") = 398 Or rsRep("myinstID") = 269 Or rsRep("myinstID") = 431 Then 'DHMC and Elliot
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - " & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						ElseIf rsRep("myinstID") =  41 Or rsRep("myinstID") =  15 Or rsRep("myinstID") =  70 Or rsRep("DeptID") = 645 Then 'concord hosp, fam and mchc and com council nashua
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - " & GetLang(rsRep("LangID")) & """,""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						ElseIf rsRep("deptID") = 1058 Then 'man welfare
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - MW" & rsRep("myindex") & """,""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						ElseIf rsRep("myinstID") = 39 Then
							reqp = GetReq(rsRep("reqID"))
							If Z_Czero(rsRep("HPID")) > 0 Then reqp = reqp & " / " & GetReqHPID(rsRep("HPID"))
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " (Req: " & reqp & ")"",""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						Else
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & "DOTC" & """,""" & _
									"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " – Interpretation" &  """,""" & date & """,""" & _
									"rfglang2" & """,""" & BillHours & """" & vbCrLf
							Else
								If rsRep("custID") = "New England Heart In" Then
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
										"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " " & rsRep("DOB") & """,""" & date & """,""" & _
										"rfglang2" & """,""" & BillHours & """" & vbCrLf
								Else
									If rsRep("class") = 3 Then
										If rsRep("emerFEE") = True Then tmpccode = "LB 60 Rate"
									End If
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
										"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " – Interpretation" &  """,""" & date & """,""" & _
										"rfglang2" & """,""" & BillHours & """" & vbCrLf
								End If
							End If
						End If
					Else
						If rsRep("emerFEE") = True Then tmpccode = "LB 60 Rate"
						docnum = ""
						If Z_Fixnull(rsRep("docnum")) <> "" Then docnum = " - " & rsrep("docnum")
						CSVBodyBillL = CSVBodyBillL & """" & "DOTC" & """,""" & _
							"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & docnum & " – Interpretation" & """,""" & date & """,""" & _
							"rfglang2" & """,""" & BillHours & """" & vbCrLf
					End If
						
					If Z_CZero(rsRep("TT_Inst")) <> 0 Then 'new billing rules for Travel
						strTT = rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname")
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
							Else
								If rsRep("myInstID") = 273 Or _
									ClassInt(rsRep("DeptID")) = 3 Then 'dart and court
										CSVBodyBill = CSVBodyBill & """" & MyTravelTimeDarthCourt(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
								Else
									CSVBodyBill = CSVBodyBill & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
								End If
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
						End If
					End If
					
					If Z_CZero(rsRep("M_Inst")) <> 0 Then 'new billing rules for Mileage
						strM = rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname")
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM) & """" & vbCrLf
							Else
								CSVBodyBill = CSVBodyBill & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM) & """" & vbCrLf
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM) & """" & vbCrLf
						End If
					End If
					
					If rsRep("emerFEE") = True Then 
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & "DOTC" & """,""" & _
									"0" & """,""" & "LB Emer Fee" & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - Emer Fee" & """,""" & date & """,""" & _
									"rfglang2" & """,""" & "1" & """" & vbCrLf
							Else 
								If rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
										"0" & """,""" & "LB Emer Fee" & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - Emer Fee" & """,""" & date & """,""" & _
										"rfglang2" & """,""" & "1" & """" & vbCrLf
								End If
							End If
						'Else
						'	CSVBodyBillL = CSVBodyBillL & """" & "DOTC" & """,""" & _
						'		"0" & """,""" & "LB Emer Fee" & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - Emer Fee" & """,""" & date & """,""" & _
						'		"rfglang2" & """,""" & "1" & """" & vbCrLf
						End If
					End If
					tmpCID2 = tmpCID
				End If
				
				'''''''''''''''''''
					
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						CSVBody = CSVBody & "," & tmpFeeL
					Else
						CSVBody = CSVBody & "," & rsRep("InstRate")
					End If
				Else
					CSVBody = CSVBody & "," & rsRep("InstRate")
				end if
				
				CSVBody = CSVBody & ",""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & ""","
				
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						CSVBody = CSVBody & "0.00"
					ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
						CSVBody = CSVBody & tmpFeeO
					End If
				Else
					CSVBody = CSVBody & "0.00"
				end if
				bilcommentcsv = Replace(Z_fixNull(rsRep("bilComment") & rsRep("syscom")), "<br>", " / ")
				CSVBody = CSVBody & ",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB") & """,""" & rsRep("CustID")
				If rsRep("myinstID") = 108 Then
					CSVBody = CSVBody & """,""" & GetUserID(rsRep("DeptID")) & """" & vbCrLf 
				Else
					CSVBody = CSVBody & """" & vbCrLf 
				End If
			
				'If Request("bill") = 1 Then
					rsRep("billingTrail") = rsRep("billingTrail") & "<br>Billed to Institution " & Date
					rsRep("Processed") = Date
				'End If
				x = x + 1
				rsRep.Update
			End If
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	
	End If
	rsRep.Close
	Set rsRep = Nothing	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	strLog = Now & vbTab & "Billing ran by " & Session("UsrName") & "."
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
ElseIf tmpReport(0) = 4 Then
	RepCSV =  "PerInstReq" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Time of Appointment</td>" & vbCrlf & _
		"<td class='tblgrn'>Duration (mins)</td>" & vbCrlf & _
		"<td class='tblgrn'>Instituion - Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Billed Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf
	CSVHead = "Appointment Date,Client's Last Name,Client's First Name,Actual Start Time,Actual End Time,Duration (mins),Instituion,Department," & _
		"Language,Interpreter's Last Name, Interpreter's First Name,Billed Hours,Total Amount,Travel Time,Mileage"
	sqlRep = "SELECT * FROM request_T, interpreter_T, institution_T, language_T, dept_T WHERE request_T.[instID] <> 479 AND Dept_T.[index] = [DeptID] AND IntrID = interpreter_T.[index] " & _
		"AND request_T.InstID = institution_T.[index] AND LangID = language_T.[index] AND (request_T.Status = 1 OR request_T.Status = 4)"
	strMSG = "Per-institution request report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If tmpReport(3) = "" Then tmpReport(3) = 0
	If tmpReport(3) <> 0 Then
		sqlRep = sqlRep & " AND institution_T.[index] = " & tmpReport(3) 
		strMSG = strMSG & " for " & GetInst(tmpReport(3))
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY appDate, AStarttime, Facility, dept, Clname, Cfname"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF 
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			tmpCliName = rsRep("Clname") & ", " & rsRep("Cfname")
			appTime = ctime(rsRep("AStarttime")) & " - " & ctime(rsRep("AEndtime"))
			appmin = DateDiff("n", rsRep("AStarttime"), rsRep("AEndtime"))
			tmpFacil = rsRep("Facility") & " - " & rsRep("Dept")
			tmpIName = rsRep("Last name") & ", " & rsRep("first name")
			tmpPay = (rsRep("billable") * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & appTime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & appmin & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpFacil & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpIName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("billable") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpPay, 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td></tr>" & vbCrLf
				
			CSVBody = CSVBody & rsRep("appDate") & ",""" & rsRep("Clname") & """,""" & rsRep("Cfname") & """," & rsRep("AStarttime") & _
				"," & rsRep("AEndtime") & "," & appmin & ",""" & rsRep("Facility") & """,""" & rsRep("Dept") & """," & GetLang(rsRep("LangID")) & _
				",""" & rsRep("Last name") & """,""" & rsRep("first name") & """," & rsRep("billable") & ",""" & Z_FormatNumber(tmpPay, 2) & _
				"""," & Z_CZero(rsRep("TT_Inst")) & "," & Z_CZero(rsRep("M_Inst")) & vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='11' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 5 Then
	RepCSV =  "PerTownReq" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Town</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointments</td>" & vbCrlf & _
		"<td class='tblgrn'>Billable Hrs.</td>" & vbCrlf & _
		"<td class='tblgrn'>Actual Hrs.</td>" & vbCrlf & _
		"<td class='tblgrn'>Classification</td>" & vbCrlf
	CSVHead = "Town, Interpreter Last Name, Interpreter First Name, Language, Appointments, Billable Hrs., Actual Hrs., Classification"
	sqlRep = "SELECT * FROM request_T, interpreter_T, institution_T, language_T, dept_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND IntrID = interpreter_T.[index] " & _
		"AND request_T.InstID = institution_T.[index] AND LangID = language_T.[index] "
	strMSG = "Per-town request report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If tmpReport(5) = "" Then tmpReport(5) = 0
	If tmpReport(5) <> 0 Then
		sqlRep = sqlRep & " AND [dept_T.City] = '" & tmpReport(5) & "'"
		strMSG = strMSG & " for " & tmpReport(5)
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY dept_T.City, [Last Name], [First Name], " & _
		"[Language], [Class]"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF
			strTown = rsRep("City")
			strIntrName = rsRep("Last Name") & ", " & rsRep("First Name")
			strLang = rsRep("Language")
			strClass = rsRep("Class")
			strBill = rsRep("Billable")
			strAhrs = DateDiff("h", rsRep("AStarttime"), rsRep("AEndtime"))
			lngIdx = SearchArraysTown(strTown, strIntrName, strLang, strClass, tmpTown, tmpIntrName, tmpLang, tmpClass)
			If lngIdx < 0 Then 
					ReDim Preserve tmpTown(x)
					ReDim Preserve tmpIntrName(x)
					ReDim Preserve tmpLang(x)
					ReDim Preserve tmpClass(x)
					ReDim Preserve tmpBill(x)
					ReDim Preserve tmpAhrs(x)
					ReDim Preserve tmpApp(x)
					
					tmpTown(x) = strTown
					tmpIntrName(x) = strIntrName
					tmpLang(x) = strLang
					tmpClass(x) = strClass
					tmpBill(x) = strBill
					tmpAhrs(x) = strAhrs
					tmpApp(x) = 1
					x = x + 1
				Else
					tmpBill(lngIdx) = tmpBill(lngIdx) + strBill
					tmpAhrs(lngIdx) = tmpAhrs(lngIdx) + strAhrs
					tmpApp(lngIdx) = tmpApp(lngIdx) + 1
				End If
			rsRep.MoveNext
		Loop
		y = 0
		Do Until y = x
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpTown(y) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpIntrName(y) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpLang(y) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpApp(y) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpBill(y), 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpAhrs(y), 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetClass(tmpClass(y)) & "</td></tr>" & vbCrLf
				
			CSVBody = CSVBody & tmpTown(y) & "," & tmpIntrName(y) & "," & tmpLang(y) & "," & tmpApp(y) & "," & tmpBill(y) & "," & _
				tmpAhrs(y) & "," & GetClass(tmpClass(y)) & vbCrLf
			y = y + 1
		Loop
	Else
		strBody = "<tr><td colspan='7' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 6 Then
On Error Resume Next
	RepCSV =  "UsageReq" & tmpdate & ".csv" 
	strMSG = "Usage Report"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	'GET LIST
	'GET SOCIAL SERVICE
	sqlRep = "SELECT billable FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND class = 1 "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "AND NOT processed IS NULL "

	rsRep.Open sqlRep, g_strCONN, 3, 1
	NumEnc1 = 0
	Do Until rsRep.EOF
	    HrPaid1 = HrPaid1 + rsRep("billable") 'CONVERT TO MIN.
	    NumEnc1 = NumEnc1 + 1
	    rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
	MinPaid1 = HrPaid1 * 60
	AvgMinEnc1 = MinPaid1 / NumEnc1
	strBODY = "<tr bgcolor='#F5F5F5'><td class='tblgrn2'>Social Service</td><td class='tblgrn2'>" & NumEnc1 & "</td><td class='tblgrn2'>" & MinPaid1 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc1, 2) & "</td><td class='tblgrn2'>" & MinPaid1 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc1, 2) & "</td></tr>" & vbCrLf 
	CSVBody = "Social Service," & NumEnc1 & "," & MinPaid1 & "," & Z_FormatNumber(AvgMinEnc1, 2) & "," & MinPaid1 & "," & Z_FormatNumber(AvgMinEnc1, 2) & vbCrLf
	'GET PRIVATE
	Set rsRep2 = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT billable FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND class = 2 "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		'strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		'strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "AND NOT (processed IS NULL)"
	rsRep2.Open sqlRep, g_strCONN, 3, 1
	NumEnc2 = 0
	Do Until rsRep2.EOF
	    HrPaid2 = HrPaid2 + rsRep2("billable") 'CONVERT TO MIN.
	    NumEnc2 = NumEnc2 + 1
	    rsRep2.MoveNext
	Loop
	rsRep2.Close
	Set rsRep2 = Nothing
	MinPaid2 = HrPaid2 * 60
	AvgMinEnc2 = MinPaid2 / NumEnc2
	strBODY = strBODY & "<tr bgcolor='#FFFFFF'><td class='tblgrn2'>Private</td><td class='tblgrn2'>" & NumEnc2 & "</td><td class='tblgrn2'>" & MinPaid2 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc2, 2) & "</td><td class='tblgrn2'>" & MinPaid2 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc2, 2) & "</td></tr>" & vbCrLf 
	CSVBody = CSVBody & "Private," & NumEnc2 & "," & MinPaid2 & "," & Z_FormatNumber(AvgMinEnc2, 2) & "," & MinPaid2 & "," & Z_FormatNumber(AvgMinEnc2, 2) & vbCrLf
	'GET MEDICAL
	Set rsRep3 = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT billable FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND class = 4 "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		'strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		'strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "AND (NOT Processed IS NULL OR NOT ProcessedMedicaid IS NULL)"
	rsRep3.Open sqlRep, g_strCONN, 3, 1
	NumEnc4 = 0
	Do Until rsRep3.EOF
	    HrPaid4 = HrPaid4 + rsRep3("billable") 'CONVERT TO MIN.
	    NumEnc4 = NumEnc4 + 1
	    rsRep3.MoveNext
	Loop
	rsRep3.Close
	Set rsRep3 = Nothing
	MinPaid4 = HrPaid4 * 60
	AvgMinEnc4 = MinPaid4 / NumEnc4
	strBODY = strBODY & "<tr bgcolor='#F5F5F5'><td class='tblgrn2'>Medical</td><td class='tblgrn2'>" & NumEnc4 & "</td><td class='tblgrn2'>" & MinPaid4 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc4, 2) & "</td><td class='tblgrn2'>" & MinPaid4 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc4, 2) & "</td></tr>" & vbCrLf 
	CSVBody = CSVBody & "Medical," & NumEnc4 & "," & MinPaid4 & "," & Z_FormatNumber(AvgMinEnc4, 2) & "," & MinPaid4 & "," & Z_FormatNumber(AvgMinEnc4, 2) & vbCrLf
	'GET COURT
	Set rsRep4 = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT billable FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND class = 3 "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		'strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		'strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "AND NOT (processed IS NULL)"
	rsRep4.Open sqlRep, g_strCONN, 3, 1
	NumEnc3 = 0
	Do Until rsRep4.EOF
	    HrPaid3 = HrPaid3 + rsRep4("billable") 'CONVERT TO MIN.
	    NumEnc3 = NumEnc3 + 1
	    rsRep4.MoveNext
	Loop
	rsRep4.Close
	Set rsRep4 = Nothing
	MinPaid3 = HrPaid3 * 60
	AvgMinEnc3 = MinPaid3 / NumEnc3
	strBODY = strBODY & "<tr bgcolor='#FFFFFF'><td class='tblgrn2'>Court</td><td class='tblgrn2'>" & NumEnc3 & "</td><td class='tblgrn2'>" & MinPaid3 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc3, 2) & "</td><td class='tblgrn2'>" & MinPaid3 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc3, 2) & "</td></tr>" & vbCrLf 
	CSVBody = CSVBody & "Court," & NumEnc3 & "," & MinPaid3 & "," & Z_FormatNumber(AvgMinEnc3, 2) & "," & MinPaid3 & "," & Z_FormatNumber(AvgMinEnc3, 2) & vbCrLf
	'GET LEGAL
	Set rsRep5 = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT billable FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND class = 5 "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		'strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		'strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "AND NOT (processed IS NULL)"
	'response.write sqlrep
	rsRep5.Open sqlRep, g_strCONN, 3, 1
	NumEnc5 = 0
	Do Until rsRep5.EOF
	    HrPaid5 = HrPaid5 + rsRep5("billable") 'CONVERT TO MIN.
	    NumEnc5 = NumEnc5 + 1
	    rsRep5.MoveNext
	Loop
	rsRep5.Close
	Set rsRep5 = Nothing
	MinPaid5 = HrPaid5 * 60
	AvgMinEnc5 = MinPaid5 / NumEnc5
	strBODY = strBODY & "<tr bgcolor='#FFFFFF'><td class='tblgrn2'>Legal</td><td class='tblgrn2'>" & NumEnc5 & "</td><td class='tblgrn2'>" & MinPaid5 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc5, 2) & "</td><td class='tblgrn2'>" & MinPaid5 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc5, 2) & "</td></tr>" & vbCrLf 
	CSVBody = CSVBody & "Legal," & NumEnc5 & "," & MinPaid5 & "," & Z_FormatNumber(AvgMinEnc5, 2) & "," & MinPaid5 & "," & Z_FormatNumber(AvgMinEnc5, 2) & vbCrLf
	'GET TOTALS
	TotNumEnc = NumEnc1 + NumEnc2 + NumEnc3 + NumEnc4 + NumEnc5
	TotMinPaid = MinPaid1 + MinPaid2 + MinPaid3 + MinPaid4 + MinPaid5
	TotAvgMinEnc = TotMinPaid / TotNumEnc
	strBODY = strBODY & "<tr bgcolor='#F5F5F5'><td class='tblgrn2'>Total</td><td class='tblgrn2'>" & TotNumEnc & "</td><td class='tblgrn2'>" & TotMinPaid & "</td><td class='tblgrn2'>" & Z_FormatNumber(TotAvgMinEnc, 2) & "</td><td class='tblgrn2'>" & TotMinPaid & "</td><td class='tblgrn2'>" & Z_FormatNumber(TotAvgMinEnc, 2) & "</td></tr>" & vbCrLf 	
	CSVBody = CSVBody & "Total," & TotNumEnc & "," & TotMinPaid & "," & Z_FormatNumber(TotAvgMinEnc, 2) & "," & TotMinPaid & "," & Z_FormatNumber(TotAvgMinEnc, 2) & vbCrLf
	'GET LEGAL W/0 MDC and NDC
	Set rsRep6 = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT billable FROM request_T, dept_T, institution_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] " & _
	    "AND request_T.InstID = institution_T.[index] AND class = 3 " & _
	    "AND request_T.instID <> 1 AND request_T.instID <> 12 "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		'strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		'strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "AND NOT (processed IS NULL)"
	rsRep6.Open sqlRep, g_strCONN, 3, 1
	NumEnc6 = 0
	Do Until rsRep6.EOF
	    HrPaid6 = HrPaid6 + rsRep6("billable") 'CONVERT TO MIN.
	    NumEnc6 = NumEnc6 + 1
	    rsRep5.MoveNext
	Loop
	rsRep6.Close
	Set rsRep6 = Nothing
	MinPaid6 = HrPaid6 * 60
	AvgMinEnc6 = MinPaid6 / NumEnc6
	strBODY = strBODY & "<tr bgcolor='#FFFFFF'><td class='tblgrn2'>Court without Manchester District Court or Nashua District Court</td><td class='tblgrn2'>" & NumEnc6 & _
		"</td><td class='tblgrn2'>" & MinPaid6 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc6, 2) & "</td><td class='tblgrn2'>" & MinPaid6 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc6, 2) & "</td></tr>" & vbCrLf 	
	CSVBody = CSVBody & "Court without Manchester District Court or Nashua District Court," & _
	    NumEnc6 & "," & MinPaid6 & "," & Z_FormatNumber(AvgMinEnc6, 2) & "," & MinPaid6 & "," & Z_FormatNumber(AvgMinEnc6, 2) & vbCrLf
	'GET MEDICAL W/0 SNHMC
	Set rsRep7 = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT billable FROM request_T, dept_T, institution_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] " & _
	    "AND request_T.InstID = institution_T.[index] AND class = 4 " & _
			"AND request_T.instID <> 93 " '- LIVE
	    If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		'strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		'strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "AND NOT (processed IS NULL)"
	rsRep7.Open sqlRep, g_strCONN, 3, 1
	NumEnc7 = 0 
	Do Until rsRep7.EOF
	    HrPaid7 = HrPaid7 + rsRep7("billable") 'CONVERT TO MIN.
	    NumEnc7 = NumEnc7 + 1
	    rsRep7.MoveNext
	Loop
	rsRep7.Close
	Set rsRep7 = Nothing
	MinPaid7 = HrPaid7 * 60
	AvgMinEnc7 = MinPaid7 / NumEnc7
	strBODY = strBODY & "<tr bgcolor='#F5F5F5'><td class='tblgrn2'>Medical without Southern NH Medical Center</td><td class='tblgrn2'>" & NumEnc7 & _
		"</td><td class='tblgrn2'>" & MinPaid7 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc7, 2) & "</td><td class='tblgrn2'>" & MinPaid7 & "</td><td class='tblgrn2'>" & Z_FormatNumber(AvgMinEnc6, 2) & "</td></tr>" & vbCrLf 	
	CSVBody = CSVBody & "Medical without Southern NH Medical Center," & _
	    NumEnc7 & "," & MinPaid7 & "," & Z_FormatNumber(AvgMinEnc7, 2) & "," & MinPaid7 & "," & Z_FormatNumber(AvgMinEnc7, 2) & vbCrLf

	strHead = "<td class='tblgrn'>Options for Filtering</td>" & vbCrlf & _
		"<td class='tblgrn'>Number of Encounters</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Number of Minutes paid to interpreteres</td>" & vbCrlf & _
		"<td class='tblgrn'>Average Number of Interpreter Minutes/Encounter</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Number of Minutes billed to Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Average Number of Minutes billed to Institution</td>" & vbCrlf
	CSVHEAD = "Options for Filtering,Number of Encounters," & _
	    "Total Number of Minutes paid to interpreteres," & _
	    "Average Number of Interpreter Minutes/Encounter," & _
	    "Total Number of Minutes billed to Institution," & _
	    "Average Number of Minutes billed to Institution"
ElseIf tmpReport(0) = 7 Then
	RepCSV =  "ReqPer" & tmpdate & ".csv" 
	strMSG = "Requesting Person report"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Address</td>" & vbCrlf
	CSVHead = "Last Name, First Name, Institution, Address, City, State, Zip"
	sqlRep = "SELECT * FROM requester_T, reqdept_T WHERE ReqID = requester_T.[index] ORDER BY Lname, Fname"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		tmpName = rsRep("Lname") & ", " & rsRep("Fname")
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetInstDept(rsRep("DeptID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetDeptAdr(rsRep("DeptID")) & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & tmpName & "," & GetInstDept(rsRep("DeptID")) & "," & GetDeptAdr(rsRep("DeptID")) & vbCrLf
		y = y + 1
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 8 Then
	RepCSV =  "Inter" & tmpdate & ".csv" 
	strMSG = "Interpreter report"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Address</td>" & vbCrlf & _
		"<td class='tblgrn'>Phone No.</td>" & vbCrlf & _
		"<td class='tblgrn'>Mobile No.</td>" & vbCrlf & _
		"<td class='tblgrn'>Email</td>" & vbCrlf 
	CSVHead = "Last Name, First Name, Address, City, State, Zip, Phone No, Mobile No, Email"
	sqlRep = "SELECT * FROM interpreter_T ORDER BY Active DESC, [Last Name], [First Name]"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	inact = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		tmpName = rsRep("Last Name") & ", " & rsRep("First Name")
		tmpIntrAddr = rsRep("Address1") & ", " & rsRep("City") & ", " & rsRep("State") & ", " & rsRep("Zip Code")
		If inact = 0 Then
			If Not rsRep("Active") Then 
				inact = 1
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td colspan='4'>&nbsp;</td></tr><tr bgcolor='" & kulay & "'><td colspan='4'>&nbsp;</td></tr>" & vbCrLf
				CSVBody = CSVBody & vbCrLf & vbCrLf
			End If
		End If
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & tmpIntrAddr & "</td><td class='tblgrn2'><nobr>" & rsRep("phone1") & "</td><td class='tblgrn2'><nobr>" & rsRep("phone2") & _
			"</td><td class='tblgrn2'><nobr>" & rsRep("e-mail") & "</td></tr>" & vbCrLf 
		CSVBody = CSVBody & tmpName & ",""" & rsRep("Address1") & """," & rsRep("City") & "," & rsRep("State") & "," & rsRep("Zip Code") & "," &  _
			rsRep("phone1") & "," & rsRep("phone2") & "," & rsRep("e-mail") & vbCrLf
		y = y + 1
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 9 Then
	RepCSV =  "MisRequest" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Status</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Classification</td>" & vbCrlf & _
		"<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Remarks</td>" & vbCrlf 
	CSVHead = "Status, Interpreter Last Name, Interpreter First Name, Language, Institution, Classification, Date,Remarks"
	sqlRep = "SELECT * FROM request_T, interpreter_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] AND Status = 2 AND IntrID = interpreter_T.[index]"
	strMSG = "Missed request report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY Status, [Last Name], [First Name]"	
	strMSG = strMSG & "."
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			tmpStat = GetStat(rsRep("Status"))
			tmpIntName = rsRep("Last Name") & ", " & rsRep("First Name")
			tmpLng = GetLang(rsRep("LangID"))
			tmpInsti = GetInst(rsRep("request_T.InstID"))
			'tmpFac = tmpInsti
			tmpClas = GetClass(rsRep("Class"))
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("request_T.[index]") & ")'><td class='tblgrn2'><nobr>" & tmpStat & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpIntName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpLng & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpInsti & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpClas & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & tmpStat & "," & tmpIntName & "," & tmpLng & "," & tmpInsti & "," & tmpClas & "," & rsRep("appDate") & ",""" & rsRep("Comment") & "" & vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='7' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 10 Then
	RepCSV =  "Stats" & tmpdate & ".csv" 
	'FACILITY
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT DISTINCT(Facility), InstID FROM request_T, institution_T WHERE request_T.[instID] <> 479 AND InstID = institution_T.[index]"
	strMSG = "Language Bank Statistics report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY Facility"
	rsRep.Open sqlRep, g_strCONN, 1, 3
	If Not rsRep.EOF Then 
		strBody = "<tr><td class='tblgrn'>Institution</td>" & vbCrlf
		'GET MONTHS
		MonthCtr = DateDiff("m", tmpReport(1), tmpReport(2))
		ReDim Preserve tmpMonthYr2(MonthCtr)
		ReDim Preserve tmpMonthYr3(MonthCtr)
		ReDim Preserve tmpMonthYr4(MonthCtr + 1)
		MonthNum = Month(tmpReport(1))
		YearNum = Year(tmpReport(1))
		YearHead = YearNum
		MonthHead = MonthNum
		Ctr = 0
		Ctr2 = 0
		Do Until Ctr = MonthCtr + 1
			MonthHead = MonthHead + Ctr2
			If MonthHead > 12 Then 
				MonthHead = 1
				YearHead = YearHead + 1
			End If
			tmpMonth = MonthName(MonthHead, True)
			strBody = strBody & "<td class='tblgrn'>" & tmpMonth & " " & Right(YearHead, 2) & "</td>" & vbCrlf
			tmpMonthYr2(Ctr) = MonthHead
			tmpMonthYr3(Ctr) = YearHead
			Ctr2 = 1
			Ctr = Ctr + 1
		Loop
		strBody = strBody & "<td class='tblgrn'>YTD TOTAL</td></tr>" & vbCrLf
		w = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(w) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetInst2(rsRep("InstID")) & "</td>" & vbCrLf 
			x = 0
			ytdctr = 0
			Do Until x = Ubound(tmpMonthYr2) + 1
				Set rsCtr = Server.CreateObject("ADODB.RecordSet")
				sqlCtr = "SELECT Count(appDate) AS tmpCtr FROM request_T WHERE request_T.[instID] <> 479 AND InstID = " & rsRep("InstID") & " AND Month(appDate) = " & tmpMonthYr2(x) & " AND Year(appDate) = " & _
					tmpMonthYr3(x) 
				rsCtr.Open sqlCtr, g_strCONN, 1, 3
				strBody = strBody & "<td class='tblgrn2'><nobr>" & rsCtr("tmpCtr") & "</td>" & vbCrLf 
				tmpMonthYr4(x) = tmpMonthYr4(x) + rsCtr("tmpCtr")
				ytdctr = ytdctr + rsCtr("tmpCtr")
				x = x + 1
				rsCtr.Close
				Set rsCtr = Nothing	
			Loop
			strBody = strBody & "<td class='tblgrn4'><nobr>" & ytdctr & "</td></tr>" & vbCrLf
			tmpMonthYr4(Ubound(tmpMonthYr4)) = tmpMonthYr4(Ubound(tmpMonthYr4)) + ytdctr
			w = w + 1
			rsRep.MoveNext
		Loop
		z = 0
		strBody = strBody & "<tr><td class='tblgrn4'>TOTAL</td>" & vbCrLf
		Do Until z  = Ubound(tmpMonthYr4) + 1
			strBody = strBody &" <td class='tblgrn4'>" & tmpMonthYr4(z) & "</td>" & vbCrLf
			z = z + 1
		Loop
		strBody = strBody & "</tr>" & vbCrLf
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
		rsRep.Close
		Set rsRep =Nothing
ElseIf tmpReport(0) = 11 Then 'pending
	ctr = 9
	RepCSV =  "Pending" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, Facility, dept, DeptID, LangID, Clname, Cfname, [Last Name], [First Name], appDate, appTimeFrom, appTimeTo, Comment, InstRate FROM request_T, Interpreter_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND IntrID = interpreter_T.[index] AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 0"
	strMSG = "Pending appointment report"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
	CSVHead = "Request ID, Institution, Department,Language, Client Last Name, Client First Name, Interpreter Last Name, Interpreter First Name, Appointment Date, Appointment Start Time, " & _
		"Appointment End Time, Rate, Comments"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY Facility, appDate, Clname, Cfname"
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & " - " & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom")) & " - " & ctime(rsRep("appTimeTo")) &"</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("InstRate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td>" & _
				"</tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & GetLang(rsRep("LangID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & rsRep("Last Name") & _
				"," & rsRep("First Name") & ","  & rsRep("appDate") & "," & ctime(rsRep("appTimeFrom")) & "," & ctime(rsRep("appTimeTo")) & "," & Z_FormatNumber(rsRep("InstRate"), 2) & ",""" & rsRep("Comment") & """" &  vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 12 Then 'completed
	ctr = 9
	RepCSV =  "Completed" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, Facility, dept, DeptID, LangID, Clname, Cfname, [Last Name], [First Name], appDate, appTimeFrom, appTimeTo, Comment FROM request_T, Interpreter_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND IntrID = interpreter_T.[index] AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 1"
	strMSG = "Completed appointment report"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
	CSVHead = "Request ID, Institution, Department,Language, Client Last Name, Client First Name, Interpreter Last Name, Interpreter First Name, Appointment Date, Appointment Start Time, " & _
		"Appointment End Time, Comments"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY Facility, appDate, Clname, Cfname"
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & " - " & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom")) & " - " & ctime(rsRep("appTimeTo")) &"</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td>" & _
				"</tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & GetLang(rsRep("LangID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & rsRep("Last Name") & _
				"," & rsRep("First Name") & ","  & rsRep("appDate") & "," & ctime(rsRep("appTimeFrom")) & "," & ctime(rsRep("appTimeTo")) & ",""" & rsRep("Comment") & """" &  vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 13 Then 'missed'
	ctr = 10
	RepCSV =  "Missed" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, Facility, dept, intrID, DeptID, LangID, Clname, Cfname, missed, appDate, appTimeFrom, appTimeTo, Comment FROM request_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 2"
	strMSG = "Missed appointment report"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Reason</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
	CSVHead = "Request ID, Institution, Department,Language, Client Last Name, Client First Name, Interpreter Last Name, Interpreter First Name, Appointment Date, Appointment Start Time, " & _
		"Appointment End Time, Reason, Comments"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY Facility, appDate, Clname, Cfname"
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & " - " & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("intrID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
								"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom")) & " - " & ctime(rsRep("appTimeTo")) &"</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetMisReason(rsRep("Missed")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td></tr>" & vbCrLf
			If GetIntr(rsRep("intrID")) = "N/A" Then 
				intrName = "N/A,"
			Else
				intrName = GetIntr(rsRep("intrID"))
			End If
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & GetLang(rsRep("LangID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & _
				intrName & ","  & rsRep("appDate") & ","  & ctime(rsRep("appTimeFrom")) & "," & ctime(rsRep("appTimeTo")) & ",""" & GetMisReason(rsRep("Missed")) & """,""" & _
				rsRep("Comment") &"""" &  vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 14 Then 'canceled'
	ctr = 10
	RepCSV =  "Canceled" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, cancel, Facility, dept, DeptID, LangID, Clname, Cfname, [Last Name], [First Name], appDate, appTimeFrom, appTimeTo, Comment FROM request_T, Interpreter_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND IntrID = interpreter_T.[index] AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 3"
	strMSG = "Canceled appointment report"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Reason</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
	CSVHead = "Request ID, Institution, Department,Language, Client Last Name, Client First Name, Interpreter Last Name, Interpreter First Name, Appointment Date, Appointment Start Time, " & _
		"Appointment End Time, Reason, Comments"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY Facility, appDate, Clname, Cfname"
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & " - " & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom")) & " - " & ctime(rsRep("appTimeTo")) &"</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetCanReason(rsRep("Cancel")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & GetLang(rsRep("LangID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & rsRep("Last Name") & _
				"," & rsRep("First Name") & ","  & rsRep("appDate") & "," & ctime(rsRep("appTimeFrom")) & "," & ctime(rsRep("appTimeTo"))  & ",""" & GetCanReason(rsRep("Cancel")) & """,""" & _
				rsRep("Comment") &"""" &  vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 15 Then 'canceled- billable
	ctr = 10
	RepCSV =  "CanceledBill" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, Facility, dept, DeptID, LangID, Clname, Cfname, [Last Name], [First Name], appDate, AStarttime, AEndtime, Comment, cancel  FROM request_T, Interpreter_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND IntrID = interpreter_T.[index] AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 4"
	strMSG = "Canceled (Billable) appointment report"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Reason</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
	CSVHead = "Request ID, Institution, Department,Language, Client Last Name, Client First Name, Interpreter Last Name, Interpreter First Name, Appointment Date, Appointment Start Time, " & _
		"Appointment End Time, Reason, Comments"		
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY Facility, appDate, Clname, Cfname"
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & " - " & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("AStarttime")) & " - " & ctime(rsRep("AEndtime")) &"</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetCanReason(rsRep("Cancel")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & GetLang(rsRep("LangID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & rsRep("Last Name") & _
				"," & rsRep("First Name") & ","  & rsRep("appDate") & "," & ctime(rsRep("AStarttime")) & "," & ctime(rsRep("AEndtime")) & ",""" & GetCanReason(rsRep("Cancel")) & """,""" & _
				rsRep("Comment") &"""" &  vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 16 Then 'billing w/o tagging
	'INSTITUTION BILLING
	RepCSV =  "InstXBillReq" & tmpdate & "-" & tmpTime & ".csv" 
	RepCSVBill = "InstXBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
	RepCSVBillL = "InstXBillReqNewL" & tmpdate & "-" & tmpTime & ".csv"
	RepCSVBillSigma = "InstXBillReqNewSigma" & tmpdate & "-" & tmpTime & ".csv"
	
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf & _
		"<td class='tblgrn'>DOB</td>" & vbCrlf & _
		"<td class='tblgrn'>Customer ID</td>" & vbCrlf & _
		"<td class='tblgrn'>DHHS</td>" & vbCrlf 
	
	CSVHead = "Request ID,Institution, Department, Appointment Date, Client Last Name, Client First Name, Language, Interpreter Last Name, Interpreter First Name, Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Emergency Surcharge, Total, Comments, DOB, Customer ID, DHHS"	
	'add  vermed = 0 AND  if medicaid billing is go
	sqlRep = "SELECT Docnum, ReqID, HPID, syscom, hasmed, outpatient, medicaid, vermed, autoacc, wcomp, drg, pid, request_T.[index] as myindex, status, [Last Name], [First Name], Clname, Cfname, AStarttime, AEndtime, " & _
		"Billable, DOB, emerFEE, class, TT_Inst, M_Inst, request_T.InstID as myinstID, DeptID, LangID, appDate, InstRate, bilComment, custID, ccode, billgroup, IntrID " & _
		"FROM request_T, interpreter_T , dept_T WHERE request_T.[instID] <> 479 AND request_T.deptID =  dept_T.[index] AND IntrID = interpreter_T.[index]  AND " & _
		"(Status = 1 OR Status = 4) AND Processed IS NULL "
	strMSG = "Institution Billing request report (simulated)"
	
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & ". * - Cancelled Billable."
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY CustID ASC, AppDate DESC" '" ORDER BY AppDate DESC"
	'response.write sqlRep
	rsRep.Open sqlRep, g_strCONN, 1, 3
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	If Not rsRep.EOF Then 
		x = 0
		tmpCID = ""
		Do Until rsRep.EOF
			BillHours = 0
			IncludeReq = True
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			CB = ""
			If rsRep("status") = 4 Then CB = "*"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
			'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
			BillHours =  rsRep("Billable")
			PayMed = PaytoMedicaid(rsRep("outpatient"), rsRep("hasmed"), rsRep("vermed"), rsRep("autoacc"), rsRep("wcomp"), rsRep("drg"), rsRep("IntrID"))
			'PayMed = PaytoMedicaid(rsRep("outpatient"), rsRep("hasmed"), rsRep("vermed"), rsRep("autoacc"), rsRep("wcomp"), deptdrg, rsRep("IntrID"), rsRep("medicaid"))
			'If PayMed And rsRep("vermed") = 0 Then IncludeReq = False
			If PayMed AND BillHours <= 4 Then IncludeReq = False 'causing LAG
			If PayMed AND BillHours > 4 Then 
				BillHours = BillHours - 4
				If Instr(BillHours, ".") > 0 Then
					If Z_GetDec(BillHours) > 0 Then
						tmpMin = CDbl("0." & Z_GetDec(billhours))
						If tmpMin > 0 And tmpMin <= 0.25 Then billhours = Int(billhours) + 0.25
            If tmpMin > 0.25 And tmpMin <= 0.5 Then billhours = Int(billhours) + 0.5
            If tmpMin > 0.5 And tmpMin <= 0.75 Then billhours = Int(billhours) + 0.75
            If tmpMin > 0.75 And tmpMin <= 0.99 Then billhours = Int(billhours) + 1
					End If
				End If
			End If
			If PayMed And rsRep("vermed") = 0 Then IncludeReq = False
			If IncludeReq Then
				If rsRep("emerFEE") = True Then
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
					ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
						tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst")) + tmpFeeO
					End If
				Else
					tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
				End If
				totalPay = Z_FormatNumber(tmpPay, 2)
				strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
					"<td class='tblgrn2'><nobr>" & CB & rsRep("myindex") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetInst2(rsRep("myinstID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strIntrName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
					If rsRep("emerFEE") = True Then 
							If rsRep("class") = 3 Or rsRep("class") = 5 Then
								strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeL & "</td>" & vbCrLf
							Else
								strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
							End If
					Else
						strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
					End If
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
					If rsRep("emerFEE") = True Then 
						If rsRep("class") = 3 Or rsRep("class") = 5 Then
							strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
						ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
							strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeO & "</td>" & vbCrLf
						End If
					Else
						strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
					End If
					bilcomment = Z_fixNull(rsRep("bilComment") & rsRep("syscom"))
					strBody = strBody & "<td class='tblgrn2'><nobr><b>$" & totalPay & "</b></td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("bilComment") & "</td><td class='tblgrn2'><nobr>" & rsRep("DOB") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("CustID") & "</td>"
					If rsRep("myinstID") = 108 Then
						strBody = strBody & "<td class='tblgrn2'><nobr>" & GetUserID(rsRep("DeptID")) & "</td><tr>" & vbCrLf 
					Else
						strBody = strBody & "<td class='tblgrn2'><nobr>&nbsp;</td><tr>" & vbCrLf 
					End If
			
				CSVBody = CSVBody & CB & rsRep("myindex") & "," & GetInst2(rsRep("myinstID")) & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & rsRep("appDate") & "," & rsRep("Clname") & "," & rsRep("Cfname") &  "," & GetLang(rsRep("LangID")) & "," & rsRep("Last Name") & _
					"," & rsRep("First Name") & ","  & cTime(rsRep("AStarttime")) & "," & cTime(rsRep("AEndtime")) & "," & BillHours
				'''''MIP IMPORT'''''
				If rsRep("myinstID") <> 108 And rsRep("myinstID") <> 30 Then 'exclude DHHS and LSS
					
					'If rsRep("class") <> 3 Then
					tmpCID = Trim(UCase(rsRep("custID")))
					tmpccode = "LB " & rsRep("InstRate") & " Rate"
					If tmpCID <> tmpCID2 Or tmpCID2 = "" Then
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & "HOTC" & """,""" & rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
							Else
								CSVBodyBill = CSVBodyBill & """" & "HOTC" & """,""" & rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & "HOTC" & """,""" & rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
						End If
					End If	
					If rsRep("class") <> 5 Then
						If rsRep("myinstID") = 229 Or rsRep("myinstID") = 27 Or rsRep("myinstID") = 273 Or rsRep("myinstID") = 308 Or rsRep("myinstID") = 19 _
							Or rsRep("myinstID") = 300 Or rsRep("myinstID") = 427 Or rsRep("myinstID") = 289 Or rsRep("myinstID") = 268 Or rsRep("myinstID") = 265 _
							Or rsRep("myinstID") = 22 Or rsRep("myinstID") = 398 Or rsRep("myinstID") = 269 Or rsRep("myinstID") = 431 Then 'DHMC and Elliot
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - " & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						ElseIf rsRep("myinstID") =  41 Or rsRep("myinstID") =  15 Or rsRep("myinstID") =  70 Or rsRep("DeptID") = 645 Then 'concord hosp, fam and mchc and com council nashua
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - " & GetLang(rsRep("LangID")) & """,""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						ElseIf rsRep("deptID") = 1058 Then 'man welfare
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - MW" & rsRep("myindex") & """,""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						ElseIf rsRep("myinstID") = 39 Then
							reqp = GetReq(rsRep("reqID"))
							If Z_Czero(rsRep("HPID")) > 0 Then reqp = reqp & " / " & GetReqHPID(rsRep("HPID"))
							CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " (Req: " & reqp & ")"",""" & date & """,""" & _
								"rfglang2" & """,""" & BillHours & """" & vbCrLf
						Else
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & "DOTC" & """,""" & _
									"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " – Interpretation" &  """,""" & date & """,""" & _
									"rfglang2" & """,""" & BillHours & """" & vbCrLf
							Else
								If rsRep("custID") = "New England Heart In" Then
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
										"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " " & rsRep("DOB") & """,""" & date & """,""" & _
										"rfglang2" & """,""" & BillHours & """" & vbCrLf
								Else
									If rsRep("class") = 3 Then
										If rsRep("emerFEE") = True Then tmpccode = "LB 60 Rate"
									End If
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
										"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " – Interpretation" &  """,""" & date & """,""" & _
										"rfglang2" & """,""" & BillHours & """" & vbCrLf
								End If
							End If
						End If
					Else
						If rsRep("emerFEE") = True Then tmpccode = "LB 60 Rate"
						docnum = ""
						If Z_Fixnull(rsRep("docnum")) <> "" Then docnum = " - " & rsrep("docnum")
						CSVBodyBillL = CSVBodyBillL & """" & "DOTC" & """,""" & _
							"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & docnum & " – Interpretation" & """,""" & date & """,""" & _
							"rfglang2" & """,""" & BillHours & """" & vbCrLf
					End If
						
					If Z_CZero(rsRep("TT_Inst")) <> 0 Then 'new billing rules for Travel
						strTT = rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname")
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
							Else
								If rsRep("myInstID") = 273 Or _
									ClassInt(rsRep("DeptID")) = 3 Then 'dart and court rsRep("myInstID") = 457 Or rsRep("myInstID") = 229 Or rsRep("myInstID") = 373 Or rsRep("myInstID") = 27 Or 
										CSVBodyBill = CSVBodyBill & """" & MyTravelTimeDarthCourt(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
								Else
									CSVBodyBill = CSVBodyBill & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
								End If
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT) & """" & vbCrLf
						End If
					End If
					
					If Z_CZero(rsRep("M_Inst")) <> 0 Then 'new billing rules for Mileage
						strM = rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname")
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM) & """" & vbCrLf
							Else
								CSVBodyBill = CSVBodyBill & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM) & """" & vbCrLf
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM) & """" & vbCrLf
						End If
					End If
					
					If rsRep("emerFEE") = True Then 
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & "DOTC" & """,""" & _
									"0" & """,""" & "LB Emer Fee" & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - Emer Fee" & """,""" & date & """,""" & _
									"rfglang2" & """,""" & "1" & """" & vbCrLf
							Else 
								If rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""" & _
										"0" & """,""" & "LB Emer Fee" & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - Emer Fee" & """,""" & date & """,""" & _
										"rfglang2" & """,""" & "1" & """" & vbCrLf
								End If
							End If
						'Else
						'	CSVBodyBillL = CSVBodyBillL & """" & "DOTC" & """,""" & _
						'		"0" & """,""" & "LB Emer Fee" & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - Emer Fee" & """,""" & date & """,""" & _
						'		"rfglang2" & """,""" & "1" & """" & vbCrLf
						End If
					End If
					tmpCID2 = tmpCID
				End If
				
				'''''''''''''''''''
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						CSVBody = CSVBody & "," & tmpFeeL
					Else
						CSVBody = CSVBody & "," & rsRep("InstRate")
					End If
				Else
					CSVBody = CSVBody & "," & rsRep("InstRate")
				end if
				
				CSVBody = CSVBody & ",""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & ""","
				
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						CSVBody = CSVBody & "0.00"
					ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
						CSVBody = CSVBody & tmpFeeO
					End If
				Else
					CSVBody = CSVBody & "0.00"
				end if
				bilcommentcsv = Replace(Z_FixNull(rsRep("bilComment") & rsRep("syscom")), "<br>", " / ")
				CSVBody = CSVBody & ",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB") & """,""" & rsRep("CustID")
				If rsRep("myinstID") = 108 Then
					CSVBody = CSVBody & """,""" & GetUserID(rsRep("DeptID")) & """" & vbCrLf 
				Else
					CSVBody = CSVBody & """" & vbCrLf 
				End If
				x = x + 1
			End If
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	
	End If
	rsRep.Close
	Set rsRep = Nothing	
ElseIf tmpReport(0) = 17 Then 'KPI
	RepCSV =  "KPI" & tmpdate & ".csv" 
	strMSG = "KPI report "
	If tmpReport(1) <> "" Then
		strMSG = strMSG & " from " & tmpReport(1)
		
	End If
	If tmpReport(2) <> "" Then
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strHead = "<td class='tblgrn'>Classification</td>" & vbCrlf & _
		"<td class='tblgrn'>Status</td>" & vbCrlf & _
		"<td class='tblgrn'>" & tmpReport(1) & " - " & tmpReport(2) & "</td>" & vbCrlf
	CSVHead = "Classification,Status," & tmpReport(1) & " - " & tmpReport(2)
	tmpRef = 0
	tmpCan = 0
	tmpCanB = 0
	tmpMis = 0
	tmpMis2 = 0
	tmpPen = 0
	tmpCom = 0
	'''''''''''COURT'''''''''''''''
	strBody = "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Court</td>" & vbCrLf
	CSVBody = "Court,"
	'REFERRALS
	strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td>" & vbCrLf
	CSVBody = CSVBody & "# of Referrals,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 3"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpRef = tmpRef + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 3 AND Status = 3"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCan = tmpCan + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED BILLABLE
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments (Billable),"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 3 AND Status = 4"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCanB = tmpCanB + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED
	strBody = strBody & "<tr><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments Missed by Interpreters,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 3 AND Status = 2 AND Missed <> 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis = tmpMis + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED 2
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 3 AND Status = 2 AND Missed = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis2 = tmpMis2 + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'PENDING
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Pending Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 3 AND Status = 0"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpPen = tmpPen + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'COMLPETED
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Completed Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 3 AND Status = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCom = tmpCom + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	strBody = strBody & "<tr><td>&nbsp;</td></tr>"
	CSVBody = CSVBody &  vbCrLf
	'''''''''''LEGAL'''''''''''''''
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Legal</td>" & vbCrLf
	CSVBody = CSVBody & "Legal,"
	'REFERRALS
	strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td>" & vbCrLf
	CSVBody = CSVBody & "# of Referrals,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 5"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpRef = tmpRef + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 5 AND Status = 3"'"AND Month(appDate) = " & Month(tmpReport(1)) & "AND Year(appDate) = " & Year(tmpReport(1)) & _
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCan = tmpCan + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED BILLABLE
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments (Billable),"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 5 AND Status = 4"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCanB = tmpCanB + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED
	strBody = strBody & "<tr><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments Missed by Interpreters,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 5 AND Status = 2 AND Missed <> 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis = tmpMis + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED 2
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 5 AND Status = 2 AND Missed = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis2 = tmpMis2 + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'PENDING
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Pending Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 5 AND Status = 0"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpPen = tmpPen + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'COMLPETED
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Completed Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 5 AND Status = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCom = tmpCom + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	strBody = strBody & "<tr><td>&nbsp;</td></tr>"
	CSVBody = CSVBody &  vbCrLf
	'''''''''''MEDICAL'''''''''''''''
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Medical</td>" & vbCrLf
	CSVBody = CSVBody &  "Medical,"
	'REFERRALS
	strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td>" & vbCrLf
	CSVBody = CSVBody &  "# of Referrals,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 4"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpRef = tmpRef + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 4 AND Status = 3"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCan = tmpCan + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED BILLABLE
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments (Billable),"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		" AND Class = 4 AND Status = 4"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCanB = tmpCanB + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED
	strBody = strBody & "<tr><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments Missed by Interpreters,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		" AND Class = 4 AND Status = 2 AND Missed <> 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis = tmpMis + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED2
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 4 AND Status = 2 AND Missed = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis2 = tmpMis2 + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'PENDING
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Pending Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 4 AND Status = 0"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpPen = tmpPen + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'COMPLETED
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Completed Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND Class = 4 AND Status = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCom = tmpCom + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	strBody = strBody & "<tr><td>&nbsp;</td></tr>"
	CSVBody = CSVBody &  vbCrLf
	'''''''''''OTHERS'''''''''''''''
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Other</td>" & vbCrLf
	CSVBody = CSVBody &  "Other,"
	'REFERRALS
	strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td>" & vbCrLf
	CSVBody = CSVBody &  "# of Referrals,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND (Class = 1 OR Class = 2)"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpRef = tmpRef + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND (Class = 1 OR Class = 2) AND Status = 3"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCan = tmpCan + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'CANCELLED BILLABLE
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments (Billable),"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND (Class = 1 OR Class = 2) AND Status = 4"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCanB = tmpCanB + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED
	strBody = strBody & "<tr><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments Missed by Interpreters,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND (Class = 1 OR Class = 2) AND Status = 2 AND Missed <> 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis = tmpMis + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'MISSED2
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND (Class = 1 OR Class = 2) AND Status = 2 AND Missed = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpMis2 = tmpMis2 + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'PENDING
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Pending Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND (Class = 1 OR Class = 2) AND Status = 0"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpPen = tmpPen + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	'COMPLETED
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Completed Appointments,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(appDate) AS CTR FROM request_T, dept_T WHERE request_T.[instID] <> 479 AND DeptID = dept_T.[index] " & _
		"AND (Class = 1 OR Class = 2) AND Status = 1"
	If tmpReport(1) <> "" Then
		sqlRef = sqlRef & " AND appDate >= '" & tmpReport(1) & "'"
	End If
	If tmpReport(2) <> "" Then
		sqlRef = sqlRef & " AND appDate <= '" & tmpReport(2) & "'"
	End If
	rsRef.Open sqlRef, g_strCONN, 1, 3
		strBody = strBody & "<td class='tblgrn4'>" & rsRef("CTR") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody &  rsRef("CTR") & "," & vbCrLf
		tmpCom = tmpCom + rsRef("CTR")
	rsRef.Close
	Set rsRef = Nothing
	strBody = strBody & "<tr><td>&nbsp;</td></tr>"
	CSVBody = CSVBody &  vbCrLf
	'''''''''''TOTALS'''''''''''''''
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>TOTALS</td>" & vbCrLf
	CSVBody = CSVBody &  "TOTALS,"
	'REFERRALS
	strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td><td class='tblgrn4'>" & tmpRef & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  "# of Referrals," & tmpRef & vbCrLf
	'CANCELLED
	strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td><td class='tblgrn4'>" & tmpCan & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments," & tmpCan & vbCrLf
	'CANCELLED BILLABLE
	strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td><td class='tblgrn4'>" & tmpCanB & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments (Billable)," & tmpCanB & vbCrLf
	'MISSED
	strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td><td class='tblgrn4'>" & tmpMis & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments Missed by Interpreter," & tmpMis & vbCrLf
	'MISSED 2
	strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td><td class='tblgrn4'>" & tmpMis2 & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter," & tmpMis2 & vbCrLf
	'PENDING
	strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td><td class='tblgrn4'>" & tmpPen & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  ",# of Pending Appointments," & tmpPen & vbCrLf
	'COMLPETED
	strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td><td class='tblgrn4'>" & tmpCom & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  ",# of Completed Appointments," & tmpCom & vbCrLf
ElseIf tmpReport(0) = 18 Then 'court request 30 days
	ctr = 8
	RepCSV =  "CourtReq" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, Facility, dept, DeptID, LangID, Clname, Cfname, [Last Name], [First Name], appDate, appTimeFrom, appTimeTo, Comment, cancel FROM request_T, Interpreter_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND IntrID = interpreter_T.[index] AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 0 AND Class = 3 AND appDate <= '" & Date & "' AND appDate >= '" & DateAdd("d", -30, Date) & "' ORDER BY appDate, Facility, dept, Clname, CFname"
	strMSG = "Court pending appointment report"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
	CSVHead = "Request ID, Institution, Department,Language, Client Last Name, Client First Name, Interpreter Last Name, Interpreter First Name, Appointment Date, Appointment Start Time, " & _
		"Appointment End Time, Comments"		
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & " - " & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom")) & " - " & ctime(rsRep("appTimeTo")) &"</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & GetLang(rsRep("LangID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & rsRep("Last Name") & _
				"," & rsRep("First Name") & ","  & rsRep("appDate") & "," & ctime(rsRep("appTimeFrom")) & "," & ctime(rsRep("appTimeTo")) & ",""" &  _
				rsRep("Comment") &"""" &  vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 19 Then 'court request
	ctr = 10
	RepCSV =  "CourtReqMonth" & tmpdate & ".csv" 
	strMSG = "Court appointment"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")	
	sqlRep = "SELECT * FROM request_T, Interpreter_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND IntrID = interpreter_T.[index] AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND (Status = 1 OR Status = 4) AND Class = 3"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & " report."	
	sqlRep = sqlRep & " ORDER BY appDate, Facility, dept, Clname, CFname"		
	
	strHead = "<td class='tblgrn'>Appt. Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf
	CSVHead = "Appt. Date, Institution,Department,Client Last Name, Client First Name, Language, Hours, Rate, Travel, Mileage, Total"	
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			tmpTotal =  Z_CZero(rsRep("TT_Inst")) +  Z_CZero(rsRep("M_Inst")) + (rsRep("billable") * rsRep("instRate"))
			strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("billable") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("instRate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(tmpTotal, 2) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("appDate") & "," & rsRep("Facility") & "," & rsRep("dept") & "," & rsRep("Clname") & "," & rsRep("Cfname") & "," & GetLang(rsRep("LangID")) & "," & rsRep("billable") & _
				"," & rsRep("instRate") & "," &  Z_CZero(rsRep("TT_Inst")) & "," &  Z_CZero(rsRep("M_Inst")) & ",""" & Z_FormatNumber(tmpTotal, 2) & """" & vbCrLf
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='10' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 20 Then 'audit report
	ctr = 10
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	RepCSV =  "Audit" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")	
	sqlRep = "SELECT request_T.[index] as myindex, Facility, deptID, Clname, Cfname, appDate, billable, TT_Inst, M_Inst, " & _
		"emerFEE, Class, instRate, processedmedicaid FROM request_T, institution_T, dept_T WHERE request_T.[instID] <> 479 AND request_T.InstID = institution_T.[index] AND DeptID = dept_T.[index] " & _
		"AND NOT processed IS NULL" 
	strMSG = "Audit report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If tmpReport(3) = "" Then tmpReport(3) = 0
	If tmpReport(3) <> 0 Then
		sqlRep = sqlRep & " AND institution_T.[index] = " & tmpReport(3) 
		strMSG = strMSG & " for " & GetInst(tmpReport(3))
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY Facility"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Fee</td>" & vbCrlf
	CSVHead = "Request ID, Institution,Department,Client Last Name, Client First Name, Date, Hours, Rate, Travel, Mileage, Emergency Fee"	
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			myhours = Z_CZero(rsRep("billable"))
			If Not IsNull(rsRep("processedmedicaid")) Then myhours = 4 - myhours
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" &  GetDept(rsRep("deptID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_formatNumber(myhours, 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("instRate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & EFee(rsRep("emerFEE"), rsRep("Class"), tmpFeeL, tmpFeeO) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  GetDept(rsRep("deptID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") & "," &  rsRep("appDate") & _
				"," & Z_formatNumber(myhours, 2) & "," & rsRep("instRate") & "," & Z_CZero(rsRep("TT_Inst")) & "," & Z_CZero(rsRep("M_Inst")) & "," & EFee(rsRep("emerFEE"), rsRep("Class"), tmpFeeL, tmpFeeO) & vbCrLf
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='11' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 21 Then 'payroll report
	'INTERPRETER BILLING
	RepCSV =  "IntrBillReq" & tmpdate & "-" & tmpTime & ".csv"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>&nbsp;</td>" & vbCrlf & _
		"<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf 
	CSVHead = ",Request ID,Institution, Department,Client Last Name, Client First Name, Language,  Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Total, Comments"	
	'sqlRep = "SELECT * FROM request_T, interpreter_T, dept_T WHERE request_T.deptID =  dept_T.index AND IntrID = interpreter_T.index  AND (Status = 1 OR Status = 4) AND IsNull(ProcessedPR)" 
	sqlRep = "SELECT * FROM request_T, interpreter_T, dept_T WHERE request_T.[instID] <> 479 AND request_T.deptID =  dept_T.[index] AND IntrID = interpreter_T.[index]  AND (Status = 1 OR Status = 4 or Status = 0) AND ProcessedPR IS NULL" 
	strMSG = "Payroll report (simulated)"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
		
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
		
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	If tmpReport(10) <> 0 Then
		If tmpReport(10) = 1 Then 
			sqlRep = sqlRep & " AND (interpreter_T.stat = 0 OR interpreter_T.stat IS NULL)"
			strMSG = strMSG & " (Employee)"
		End If
		If tmpReport(10) = 2 Then 
			sqlRep = sqlRep & " AND interpreter_T.stat = 1"
			strMSG = strMSG & " (Outside Consultant)"
		End If
	End If
	sqlRep = sqlRep & " ORDER BY [last name], [first name], appdate"
	rsRep.Open sqlRep, g_strCONN, 1, 3
	If Not rsRep.EOF Then 
		x = 0
		tmpIid = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  rsRep("AStarttime") & " -  " & rsRep("AEndtime")
			'BillHours =  rsRep("Billable") 'CHANGE
			BillHours = IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime"), rsRep("request_T.InstID"))
			'tmpBilHrs = tmpBilHrs + BillHours
			tmpPay2 = (BillHours * rsRep("IntrRate")) + Z_CZero(rsRep("TT_Intr")) + Z_CZero(rsRep("M_Intr"))
			totalPay2 = Z_FormatNumber(tmpPay2, 2)	
			If tmpIid <> rsRep("intrID") Then
				If tmpIid <> 0 Then
					CSVBody = CSVBody & "," & vbCrLf
					strBody = strBody & "<tr><td colspan='13'>&nbsp;</td></tr>" & vbCrLf
				End If
				strBody = strBody & "<tr bgcolor='#FFFFCE' onclick='PassMe(" & rsRep("request_T.[index]") & ")'>" & _
					"<td class='tblgrn2'><nobr><b>" & strIntrName & "</b></td><td class='tblgrn2' colspan='12'>&nbsp;</td></tr>" & vbCrLf 
				CSVBody = CSVBody & """" & strIntrName & """" & vbCrLf
			End If
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("request_T.[index]") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("request_T.[index]") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" &  GetInst2(rsRep("request_T.InstID"))  & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & rsRep("IntrRate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Intr")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Intr")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr><b>$" & totalPay2 & "</b></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td><tr>" & vbCrLf 
				
			tmpIid = rsRep("intrID")
			
			CSVBody = CSVBody & rsRep("appDate") &"," & rsRep("request_T.[index]") & ","  & GetInst2(rsRep("request_T.InstID")) & ",""" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & rsRep("Clname") & """,""" & rsRep("Cfname") & _
				"""," & GetLang(rsRep("LangID")) & ","  & rsRep("AStarttime") & "," & rsRep("AEndtime") & "," & BillHours & _
				"," & rsRep("IntrRate") & ",""" & Z_CZero(rsRep("TT_Intr")) & """,""" & Z_CZero(rsRep("M_Intr")) & """,""" &  totalPay2 &""",""" & rsRep("Comment") & """" & vbCrLf
			rsRep("ProcessedPR") = Date
			rsRep.Update
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	strLog = Now & vbTab & "Payroll ran by " & Session("UsrName") & "."
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
ElseIf tmpReport(0) = 22 Then 'pre payroll report
	'INTERPRETER BILLING
	RepCSV =  "IntrXBillReq" & tmpdate & "-" & tmpTime & ".csv"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>&nbsp;</td>" & vbCrlf & _
		"<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf 
	CSVHead = ",Request ID,Institution, Department,Client Last Name, Client First Name, Language,  Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Total, Comments"	
	'sqlRep = "SELECT * FROM request_T, interpreter_T, dept_T WHERE request_T.deptID =  dept_T.index AND IntrID = interpreter_T.index  AND (Status = 1 OR Status = 4) AND IsNull(ProcessedPR)" 
	sqlRep = "SELECT * FROM request_T, interpreter_T, dept_T WHERE request_T.[instID] <> 479 AND request_T.deptID =  dept_T.[index] AND IntrID = interpreter_T.[index]  AND (Status = 1 OR Status = 4 or Status = 0) AND ProcessedPR IS NULL" 
	strMSG = "Payroll report (simulated)"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
		
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
		
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	If tmpReport(10) <> 0 Then
		If tmpReport(10) = 1 Then 
			sqlRep = sqlRep & " AND (interpreter_T.stat = 0 OR interpreter_T.stat IS NULL)"
			strMSG = strMSG & " (Employee)"
		End If
		If tmpReport(10) = 2 Then 
			sqlRep = sqlRep & " AND interpreter_T.stat = 1"
			strMSG = strMSG & " (Outside Consultant)"
		End If
	End If
	sqlRep = sqlRep & " ORDER BY [last name], [first name], appdate"
	rsRep.Open sqlRep, g_strCONN, 1, 3
	If Not rsRep.EOF Then 
		x = 0
		tmpIid = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  rsRep("AStarttime") & " -  " & rsRep("AEndtime")
			'BillHours =  rsRep("Billable") 'CHANGE
			BillHours = IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime"), rsRep("request_T.InstID"))
			'tmpBilHrs = tmpBilHrs + BillHours
			tmpPay2 = (BillHours * rsRep("IntrRate")) + Z_CZero(rsRep("TT_Intr")) + Z_CZero(rsRep("M_Intr"))
			totalPay2 = Z_FormatNumber(tmpPay2, 2)	
			If tmpIid <> rsRep("intrID") Then
				If tmpIid <> 0 Then
					CSVBody = CSVBody & "," & vbCrLf
					strBody = strBody & "<tr><td colspan='13'>&nbsp;</td></tr>" & vbCrLf
				End If
				strBody = strBody & "<tr bgcolor='#FFFFCE' onclick='PassMe(" & rsRep("request_T.[index]") & ")'>" & _
					"<td class='tblgrn2'><nobr><b>" & strIntrName & "</b></td><td class='tblgrn2' colspan='12'>&nbsp;</td></tr>" & vbCrLf 
				CSVBody = CSVBody & """" & strIntrName & """" & vbCrLf
			End If
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("request_T.[index]") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("request_T.[index]") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" &  GetInst2(rsRep("request_T.InstID"))  & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & rsRep("IntrRate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Intr")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Intr")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr><b>$" & totalPay2 & "</b></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td><tr>" & vbCrLf 
				
			tmpIid = rsRep("intrID")
			
			CSVBody = CSVBody & rsRep("appDate") &"," & rsRep("request_T.[index]") & ","  & GetInst2(rsRep("request_T.InstID")) & ",""" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & rsRep("Clname") & """,""" & rsRep("Cfname") & _
				"""," & GetLang(rsRep("LangID")) & ","  & rsRep("AStarttime") & "," & rsRep("AEndtime") & "," & BillHours & _
				"," & rsRep("IntrRate") & ",""" & Z_CZero(rsRep("TT_Intr")) & """,""" & Z_CZero(rsRep("M_Intr")) & """,""" &  totalPay2 &""",""" & rsRep("Comment") & """" & vbCrLf
			'rsRep("ProcessedPR") = Date
			'rsRep.Update
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing	
ElseIf tmpReport(0) = 23 Then 'cancelled courts appts.
		ctr = 5
	RepCSV =  "CanceledCourtReqMonth" & tmpdate & ".csv" 
	strMSG = "Canceled Court appointment"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")	
	sqlRep = "SELECT * FROM request_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 3 AND Class = 3"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & " report."	
	sqlRep = sqlRep & " ORDER BY appDate, Facility, dept, Clname, CFname"	
	strHead = "<td class='tblgrn'>Appt. Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf
	CSVHead = "Appt. Date, Institution,Department,Client Last Name, Client First Name, Language"	
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("appDate") & "," & rsRep("Facility") & "," & rsRep("dept") & "," & rsRep("Clname") & "," & rsRep("Cfname") & ",""" & GetLang(rsRep("LangID")) & """" & vbCrLf
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='5' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 24 Then 'active interpreters
	RepCSV =  "ActiveInter" & tmpdate & ".csv" 
	strMSG = "Active Interpreter report"
	strHead = "<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Name</td>" & vbCrlf & _
		"<td class='tblgrn'>E-mail</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf & _
		"<td class='tblgrn'>Home Phone</td>" & vbCrlf & _
		"<td class='tblgrn'>Mobile Phone</td>" & vbCrlf & _
		"<td class='tblgrn'>Address</td>" & vbCrlf
	CSVHead = "Language,Last Name, First Name, Email,Comments,Home Phone,Mobile Phone,Address, City, State, Zip"

	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM Language_T ORDER BY [Language]"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	Do Until rsRep.EOF
		IntrLang = UCase(rsRep("Language"))
		strBody = strBody & "<tr bgcolor='#FFFFCE'><td colspan='7' align='left'>" & IntrLang & "</td></tr>"
		CSVBody = CSVBody & IntrLang & vbCrLf
		Set rsRep2 = Server.CreateObject("ADODB.RecordSet")
		sqlRep2 = "SELECT * FROM interpreter_T WHERE (Upper(Language1) = '" & IntrLang & "' OR Upper(Language2) = '" & IntrLang & _
			"' OR Upper(Language3) = '" & IntrLang & "' OR Upper(Language4) = '" & IntrLang & "' OR Upper(Language5) = '" & IntrLang & _
			"') AND Active = 1 ORDER BY [Last Name], [First Name]" 
		rsRep2.Open sqlRep2, g_strCONN, 3, 1
		y = 0
		Do Until rsRep2.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			tmpName = Trim(rsRep2("Last Name") & ", " & rsRep2("First Name"))
			tmpphone = rsRep2("phone1")
			If rsRep2("P1Ext") <> "" Then tmpphone = tmpphone & " ext. " & rsRep2("P1Ext")
			tmpAdr = rsRep2("Address1") & ", " & rsRep2("IntrAdrI") & ", " & rsRep2("city") & ", " & rsRep2("state") & ", " & rsRep2("zip code")
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td>&nbsp;</td><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep2("E-mail") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep2("Comments") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpphone & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep2("phone2") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpAdr & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & ",""" & rsRep2("Last Name") & """,""" & rsRep2("First Name") & """,""" & rsRep2("e-mail") & """,""" & _
				rsRep2("comments") & """,""" & tmpphone & """,""" & rsRep2("phone2") & """,""" & rsRep2("Address1") & ", " & _
				rsRep2("IntrAdrI") & """,""" & rsRep2("city") & """,""" & rsRep2("state") & """,""" & rsRep2("zip code") & """" & vbCrLf
			y = y + 1
			rsRep2.MoveNext
		Loop
		rsRep2.Close
		Set rsRep2 = Nothing
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 25 Then 'weekly report
	RepCSV =  "Weekly Report" & tmpdate & ".csv" 
	tmpDate = tmpReport(1)
	If WeekDay(tmpDate) <> 1 Then
		Do Until WeekDay(tmpDate) = 1
			tmpDate = DateAdd("d", "-1", tmpDate)
		Loop
	End If
	tmpSun = tmpDate
	tmpSat = DateAdd("d", 6, tmpDate)
	strMSG = "Weekly report for the week of " & tmpSun & " - " & tmpSat
	strHead = "<td class='tblgrn'>Classification</td>" & vbCrlf & _
		"<td class='tblgrn'>#</td>" & vbCrlf 
	tmpMedCom = 0
	tmpMedMis = 0
	tmpMedMisX = 0
	tmpMedCan = 0
	tmpLegCom = 0
	tmpLegMis = 0
	tmpLegMisX = 0
	tmpLegCan = 0
	tmpOthCom = 0
	tmpOthMis = 0
	tmpOthMisX = 0
	tmpOthCan = 0
	tmpNew = 0
	tmpBilHrs = 0
	'MEDICAL completed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND Class = 4 AND (status = 4 OR status = 1) " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		If rsMed("appTimeTo") <> "" Then
			tmpBilHrs = tmpBilHrs + DateDiff("n", rsMed("appTimeFrom"), rsMed("appTimeTo"))		
		End If
		tmpMedCom = tmpMedCom + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor=''><td class='tblgrn2'><nobr>Medical Appts Completed</td><td class='tblgrn2'>" & tmpMedCom & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Medical Appts Completed," & tmpMedCom & vbCrLf
	'MEDICAL missed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND Class = 4 AND status = 2 " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpMedMis = tmpMedMis + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Medical Appts Missed</td><td class='tblgrn2'>" & tmpMedMis & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Medical Appts Missed," & tmpMedMis & vbCrLf
	'MEDICAL missed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND Class = 4 AND status = 2 " & _
		"AND missed = 1 AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpMedMisX = tmpMedMisX + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor=''><td class='tblgrn2'><nobr>Medical Appts Unable to Send Interpreter</td><td class='tblgrn2'>" & tmpMedMisX & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Medical Appts Unable to Send Interpreter," & tmpMedMisX & vbCrLf
	'MEDICAL canceled
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND Class = 4 AND status = 3 " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpMedCan = tmpMedCan + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Medical Appts Cancelled</td><td class='tblgrn2'>" & tmpMedCan & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Medical Appts Cancelled," & tmpMedCan & vbCrLf
	'LEGAL completed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 3 OR Class = 5) AND (status = 4 OR status = 1) " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		If rsMed("appTimeTo") <> "" Then
			tmpBilHrs = tmpBilHrs + DateDiff("n", rsMed("appTimeFrom"), rsMed("appTimeTo"))		
		End If
		tmpLegCom = tmpLegCom + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor=''><td class='tblgrn2'><nobr>Legal Appts Completed</td><td class='tblgrn2'>" & tmpLegCom & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Legal Appts Completed," & tmpLegCom & vbCrLf
	'LEGAL missed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 3 OR Class = 5) AND status = 2 " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpLegMis = tmpLegMis + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Legal Appts Missed</td><td class='tblgrn2'>" & tmpLegMis & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Legal Appts Missed," & tmpLegMis & vbCrLf
	'LEGAL missed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 3 OR Class = 5) AND status = 2 " & _
		"AND missed = 1 AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpLegMisX = tmpLegMisX + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor=''><td class='tblgrn2'><nobr>Legal Appts Unable to Send Interpreter</td><td class='tblgrn2'>" & tmpLegMisX & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Legal Appts Unable to Send Interpreter," & tmpLegMisX & vbCrLf
	'LEGAL canceled
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 3 OR Class = 5) AND status = 3 " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpLegCan = tmpLegCan + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Legal Appts Cancelled</td><td class='tblgrn2'>" & tmpLegCan & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Legal Appts Cancelled," & tmpLegCan & vbCrLf
	'OTHERS completed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 1 OR Class = 2) AND (status = 4 OR status = 1) " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		If rsMed("appTimeTo") <> "" Then
			tmpBilHrs = tmpBilHrs + DateDiff("n", rsMed("appTimeFrom"), rsMed("appTimeTo"))		
		End If
		tmpOthCom = tmpOthCom + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor=''><td class='tblgrn2'><nobr>Other Appts Completed</td><td class='tblgrn2'>" & tmpOthCom & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Other Appts Completed," & tmpOthCom & vbCrLf
	'OTHERS missed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 1 OR Class = 2) AND status = 2 " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpOthMis = tmpOthMis + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Other Appts Missed</td><td class='tblgrn2'>" & tmpOthMis & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Other Appts Missed," & tmpOthMis & vbCrLf
	'OTHERS missed
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 1 OR Class = 2) AND status = 2 " & _
		"AND missed = 1 AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpOthMisX = tmpOthMisX + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor=''><td class='tblgrn2'><nobr>Other Appts Unable to Send Interpreter</td><td class='tblgrn2'>" & tmpOthMisX & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Other Appts Unable to Send Interpreter," & tmpOthMisX & vbCrLf
	'OTHERS canceled
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	sqlMed = "SELECT * FROM dept_T, request_T WHERE request_T.[instID] <> 479 AND deptID = dept_T.[index] AND (Class = 1 OR Class = 2) AND status = 3 " & _
		"AND appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' "
	rsMed.Open sqlMed, g_strCONN, 3, 1
	Do Until rsMed.EOF
		tmpOthCan = tmpOthCan + 1
		rsMed.MoveNext
	Loop
	rsMed.Close
	Set rsMed = Nothing
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Other Appts Cancelled</td><td class='tblgrn2'>" & tmpOthCan & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Other Appts Cancelled," & tmpOthCan & vbCrLf
	'NEW INST
	Set rsNew = Server.CreateObject("ADODB.RecordSet")
	sqlNew = "SELECT * FROM institution_T WHERE Date >= '" & tmpSun & "' AND Date <= '" & tmpSat & "' "
	rsNew.Open sqlNew, g_strCONN, 3, 1
	Do Until rsNew.EOF
		tmpNew = tmpNew + 1
		rsNew.MoveNext
	Loop
	rsNew.Close
	Set rsNew = Nothing
	strBody = strBody & "<tr  bgcolor=''><td class='tblgrn2'><nobr>New Institution</td><td class='tblgrn2'>" & tmpNew & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "New Institution," & tmpNew & vbCrLf
	'BILLABLE HOURS
	tmpBilHrs = Z_FormatNumber((tmpBilHrs / 60), 2)
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>Billable Hours</td><td class='tblgrn2'>" & tmpBilHrs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & "Billable Hours," & tmpBilHrs & vbCrLf
ElseIf tmpReport(0) = 26 Then 'mileage report
	RepCSV =  "Mileage" & tmpdate & "-" & tmpTime & ".csv" 
	tmpMonthYear = MonthName(Month(tmpReport(1))) & " - " & Year(tmpReport(1))
	strMSG = "Mileage report for the month of " & tmpMonthYear
	strHead = "<td class='tblgrn'>File Number</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Miles</td>" & vbCrlf & _
		"<td class='tblgrn'>Miles Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Receipts Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf 
	CSVHead = "File Number,Last Name,First Name,Miles,Miles Amount,Receipts Amount,Total"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT FileNum, actmil, overmile, appDate, Interpreter_T.[index] as myIntrIndex, Toll, mileageproc " & _
		" FROM Request_T, Interpreter_T WHERE request_T.[instID] <> 479 AND IntrID = Interpreter_T.[index] AND Month(appDate) = " & Month(tmpReport(1)) & " AND Year(appDate) = " & _
		Year(tmpReport(1)) & " "
	If Z_CZero(tmpReport(4)) > 0 Then
		sqlRep = sqlRep & "AND IntrID = " & tmpReport(4) & " "
		strMSG = strMSG & " for " & GetIntr(tmpReport(4)) & "."
	End If
	sqlRep = sqlRep & "AND LbconfirmToll = 1 AND mileageproc IS NULL ORDER BY [last name], [first name], appDate"
	rsRep.Open sqlRep, g_strCONN, 1, 3
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			'IntrName = rsRep("Last Name") & ", " & rsRep("First Name")
			strMile = Z_Czero(rsRep("actmil"))
			IntrID = rsRep("myIntrIndex")
			strTol = Z_Czero(rsRep("Toll"))
			lngIDx = SearchArraysHours(IntrID, tmpIntr)
			If lngIdx < 0 Then
				ReDim Preserve tmpIntr(x)
				ReDim Preserve tmpMile(x)
				ReDim Preserve tmpToll(x)
				
				tmpIntr(x) = IntrID
				tmpMile(x) = strMile
				tmpToll(x) = strTol
				x = x + 1
			Else	
				tmpMile(lngIdx) = tmpMile(lngIdx) + strMile
				tmpToll(lngIdx) = tmpToll(lngIdx) + strTol
			End If
			rsRep("mileageproc") = date
			rsRep.Update
			rsRep.MoveNext
		Loop
		y = 0
		Do Until y = x
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			myMile = Z_Czero(tmpMile(y))
			myToll = Z_Czero(tmpToll(y))
	
			mileTOT = MileageAmt(myMile) + myToll
			mileAmt = MileageAmt(myMile)
			strBody = strBody & "<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center' style='width: 80%;'>" & vbCrLf & _
				"<tr><td align='center' colspan='14'>LUTHERAN COMMUNITY SERVICES OF NEW HAMPSHIRE, INC.</td></tr>" & vbCrLf & _
				"<tr><td align='center' colspan='14'>Monthly Staff Expense Report</td></tr>" & vbCrLf & _
				"<tr><td colspan='14'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Name:</td><td align='left'><nobr><b><u>" & GetIntr(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Address:</td><td align='left'><nobr><b><u>" & Z_IntrAdr(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>City:</td><td align='left'><nobr><b><u>" & Z_IntrCity(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Zip:</td><td align='left'><nobr><b><u>" & Z_IntrZip(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Month:</td><td align='left'><nobr><b><u>" & Month(tmpReport(1)) & "/" & Year(tmpReport(1)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr><td align='center' colspan='14'>" & vbCrLf & _
					"<table bgColor='white' border='0' style='width: 80%;' cellSpacing='0' cellPadding='0' align='center'>" & vbCrLf & _
						"<tr><td>Non Client Miles:</td><td align='left'>" & Z_FormatNumber(myMile, 2) & "</td><td align='center'>&nbsp;&nbsp;X&nbsp;&nbsp;</td><td> " & Z_FormatNumber(MileageRate(), 2) & "</td><td align='left'>$" & Z_FormatNumber(mileAmt, 2) & "</td></tr>" & vbCrLf & _
						"<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td align='left'>Parking/Tolls:</td><td align='left'>$" & Z_FormatNumber(myToll, 2) & "</td></tr>" & vbCrLf & _
						"<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td align='left'>Fares:</td><td align='left'>$0.00</td></tr>" & vbCrLf & _
						"<tr><td>&nbsp;</td><td align='left' colspan='3'>Total Travel Expense <b>(Feeds to Travel line below)</b>:</td><td align='left'>$" & Z_FormatNumber(mileTOT, 2) & "</td></tr>" & vbCrLf & _
					"</table>" & vbCrLf & _					
				"</td></tr>" & vbCrLf & _	
				"<tr><td colspan='14'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr><td colspan='14'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Entity</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Region</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Site</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Department</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Service</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid; width: 200px;'><b>G/L</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid; width: 300px;'><b>Fiscal Yr</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Payer Source</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Person Code</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Position Code</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>FAS 117</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>TBD2</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Amount</b></td>" & vbCrLf & _	
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Mileage</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65010</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _	
				"</tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Meals</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65020</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Lodging</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65000</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Other</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65020</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Postage &amp; Delivery</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63190</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Telephone</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63210</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Seminars &amp; Workshops</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63220</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Program Supplies</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>60350</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Office Supplies</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63170</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='right' style='border: 1px solid;' colspan='13'><b>Total</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _	
				"</tr>" & vbCrLf & _	
				"<tr><td align='right' colspan='2'><nobr><b>Payment is hearby requested by:</b></td></tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='right' colspan='2'>Signature:</td>" & vbCrLf & _
					"<td align='left' colspan='4'><nobr><b><u><i>ELECTRONICALLY APPROVED</i></u></b>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf & _
					"Date:<u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></td>" & vbCrLf & _
				"</tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='right' colspan='2'>Approved By:</td>" & vbCrLf & _
					"<td align='left' colspan='4'><nobr><b><u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></b></u></td>" & vbCrLf & _
				"</tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"</table><div class='page-break'><br></div>" & vbCrLf
				
				
		'	strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetFileNum(tmpIntr(y)) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>" & GetIntr(tmpIntr(y)) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myMile,2) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(AmtRate(myMile),2) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(myToll,2) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(mileTOT,2) & "</td></tr>" & vbCrLf
								
			CSVBody = CSVBody & GetFileNum(tmpIntr(y)) & "," & GetIntr(tmpIntr(y)) & "," & Z_FormatNumber(myMile,2) & "," & _
				Z_FormatNumber(AmtRate(myMile),2) & "," & Z_FormatNumber(myToll,2) & "," & Z_FormatNumber(mileTOT,2) & vbCrLf
			y = y + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	
	rsRep.Close
	Set rsRep = Nothing
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	strLog = Now & vbTab & "Mileage ran by " & Session("UsrName") & "."
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
ElseIf tmpReport(0) = 27 Then 'timsheet report
	RepCSV =  "Timesheet" & tmpdate & ".csv"
	tmpMonthYear = MonthName(Month(tmpReport(1))) & " - " & Year(tmpReport(1))
	mysundate = GetSun(tmpReport(1))
	mysatdate = GetSat(tmpReport(1))
	strMSG = "Timsheet report "'for the week of " & mysundate & " - " & mysatdate
	strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Activity</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Appt. Start Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Appt. End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Payable Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Final Payable Hours</td>" & vbCrlf 
	CSVHead = "Date,Last Name,First Name,Activity, Travel Time, Appt. Start Time,Appt. End Time,Total Hours, Payable Hours, Final Payable Hours"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT [Last Name], [First Name], InstID, Cfname, totalhrs, actTT, overpayhrs, AStarttime, AEndtime, appDate, payhrs, Interpreter_T.[index] as myintrID " & _
		"FROM Request_T, Interpreter_T WHERE IntrID = Interpreter_T.[index]" ' AND appDate >= '" & mysundate & "' AND appDate <= '" & _
		'mysatdate & "' "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If Z_CZero(tmpReport(4)) > 0 Then
		sqlRep = sqlRep & "AND IntrID = " & tmpReport(4) & " "
		strMSG = strMSG & " for " & GetIntr(tmpReport(4)) & "."
	End If
	sqlRep = sqlRep & "AND [LBconfirm] = 1 ORDER BY [last name], [first name], appDate"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	IntrID2 = ""
	totHrs = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		IntrName = rsRep("Last Name") & ", " & rsRep("First Name")
		CliName = GetInst(rsRep("InstID")) & " - " & rsRep("Cfname")
		tmpAMTs = rsRep("totalhrs")
		TT = Z_FormatNumber(rsRep("actTT"), 2)
		If rsRep("overpayhrs") Then 
			PHrs = Z_FormatNumber(rsRep("payhrs"), 2)
			OvrHrs = "*"
		Else
			PHrs = Z_FormatNumber(IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime")), 2)
			OvrHrs = ""
		End If
		FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
		If Z_CZero(tmpReport(4)) > 0 Then
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & IntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & TT & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AEndtime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpAMTs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(PHrs, 2) & OvrHrs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(FPHrs, 2) & "</td></tr>" & vbCrLf 
		Else
			IntrID = rsRep("myintrID")
			
			If IntrID <> IntrID2 And IntrID2 <> "" Then
				strBody = strBody & "<tr bgcolor='#FFFFCE'><td colspan='8' class='tblgrn2'>&nbsp;</td><td class='tblgrn2'>" & Z_FormatNumber(totHrs,2) & "</td></tr>"
				If IntrID2 <> "" Then strBody = strBody & "<P CLASS='pagebreakhere'>"
				totHrs = 0
			End If
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & IntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & TT & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AEndtime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpAMTs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(PHrs, 2) & OvrHrs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(FPHrs, 2) & "</td></tr>" & vbCrLf 
			IntrID2 = IntrID
		End If
		totHrs = totHrs + Z_CZero(FPHrs)
		CSVBody = CSVBody & ",""" & rsRep("appDate") & """,""" & rsRep("Last Name") & """,""" & rsRep("First Name") & """,""" & _
				CliName & """,""" & TT & """,""" & CTime(rsRep("AStarttime")) & """,""" & CTime(rsRep("AEndtime")) & _
				""",""" & tmpAMTs & """,""" & Z_FormatNumber(PHrs, 2) & OvrHrs & """,""" & Z_FormatNumber(FPHrs, 2) & """" & vbCrLf
		y = y + 1
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
	strBody = strBody & "<tr bgcolor='#FFFFCE'><td colspan='8' class='tblgrn2'>&nbsp;</td><td class='tblgrn2'>" & Z_FormatNumber(totHrs,2) & "</td></tr>"
	
ElseIf tmpReport(0) = 28 Then 'Total hours report
	RepCSV =  "TotalHours" & tmpdate & ".csv"
	strMSG = "Total Hours report"
	strHead = "<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>File Number</td>" & vbCrlf & _
		"<td class='tblgrn'>Regular Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Holiday Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Over Time Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Back Hours</td>" & vbCrlf 
	CSVHead = "Co Code,Batch ID,Last Name,First Name,File #,temp dept,temp rate,reg hours,o/t hours,hours 3 code,hours 3 amount,hours 4 code,hours 4 amount,earnings 3 code,earnings 3 amount,earnings 4 code,earnings 4 amount,earnings 5 code,earnings 5 amount,memo code,memo amount"'"Last Name,First Name,File Number,Regular Hours,Holiday Hours,Over Time Hours"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM request_T, interpreter_T WHERE intrID = interpreter_T.[index] AND (IntrID <> 0 OR intrID = -1) AND STATUS <> 2 AND STATUS <> 3 AND showintr = 1 "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY [last name], [first name]"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF
			strIntr = rsRep("IntrID")
			TT = Z_FormatNumber(rsRep("actTT"), 2)
			If rsRep("overpayhrs") Then 
				PHrs = Z_FormatNumber(rsRep("payhrs"), 2)
			Else
				PHrs = Z_FormatNumber(IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime")), 2)
			End If
			If rsRep("deptID") = 1876 Then 'back hours
				FPHHrs = 0
				FPHrs = 0
				thours = 0
				bhrs = Z_Czero(PHrs) + Z_Czero(TT)
			Else
				bhrs = 0
				If Not rsRep("training") Then
					thours = 0
					If IsHoliday(rsRep("appdate")) Then
						FPHHrs = Z_Czero(PHrs) + Z_Czero(TT)
						FPHrs = 0
					Else
						FPHHrs = 0
						FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
					End If
				Else
					FPHHrs = 0
					FPHrs = 0
					thours = Z_Czero(PHrs) + Z_Czero(TT)
				End If
			End If
			lngIDx = SearchArraysHours(strIntr, tmpIntr)
			If lngIdx < 0 Then
				ReDim Preserve tmpIntr(x)
				ReDim Preserve tmpHrs(x)
				ReDim Preserve tmpHHrs(x)
				ReDim Preserve tmpTrain(x)
				ReDim Preserve tmpbhrs(x)
				
				tmpIntr(x) = strIntr
				tmpHrs(x) = FPHrs
				tmpHHrs(x) = FPHHrs
				tmpTrain(x) = thours
				tmpbhrs(x) = bhrs
				x = x + 1
			Else	
				tmpHrs(lngIdx) = tmpHrs(lngIdx) + FPHrs
				tmpHHrs(lngIdx) = tmpHHrs(lngIdx) + FPHHrs
				tmpTrain(lngIdx) = tmpTrain(lngIdx) + thours
				tmpbhrs(lngIdx) = tmpbhrs(lngIdx) + bhrs
			End If
			rsRep.MoveNext
		Loop
		y = 0
		Do Until y = x
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			TotHours = tmpHrs(y) + tmpHHrs(y) + tmpTrain(y)
			myTrain = Z_Czero(tmpTrain(y))
			myBhrs = Z_Czero(tmpbhrs(y))
			If TotHours <= 80 Then
				myHrs = Z_Czero(tmpHrs(y))
				myHHrs = Z_Czero(tmpHHrs(y))
				myOTHrs = 0
			Else
				myHrs = 80'tmpHrs(y)
				myHHrs = tmpHHrs(y)
				myOTHrs = tmpHrs(y) - 80
			End If	
			If TotHours <> 0 Or myBhrs <> 0 Then
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(tmpIntr(y)) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetFileNum(tmpIntr(y)) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myHrs,2) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myHHrs,2) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myOTHrs,2) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myBhrs,2) & "</td></tr>" & vbCrLf
									
				CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",,," & Z_FormatNumber(myHrs,2) & "," & Z_FormatNumber(myOTHrs,2) & vbCrLf
				If 	myHHrs > 0 Then
					CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",,,0,0" & ",HWK," & Z_FormatNumber(myHHrs,2) & vbCrLf
					'CSVBody = CSVBody & ",HWK," & Z_FormatNumber(myHHrs,2)
				End If
				If myTrain > 0 Then
					CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",," & Z_MinRate() & "," & Z_FormatNumber(myTrain,2) & "," & Z_FormatNumber(myOTHrs,2) & vbCrLf
				End If
				If myBhrs > 0 Then
					CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",,,0,0" & ",BCK," & Z_FormatNumber(myBhrs,2) & vbCrLf
				End If
				'CSVBody = CSVBody & vbCrLf
				'GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & "," & Z_FormatNumber(myHrs,2) & "," & _
				'	Z_FormatNumber(myHHrs,2) & "," & Z_FormatNumber(myOTHrs,2) & vbCrLf
			End If
			y = y + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 29 Then 'Billable hours report 
	RepCSV =  "BillHours" & tmpdate & ".csv"
	strMSG = "Billable Hours report "	
	strHead = "<td class='tblgrn'>ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours Billed</td>" & vbCrlf
	CSVHead = "ID,Instituition,Language,Interpreter Last Name,Interpreter First Name,Rate,Hours Billed"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index], InstID, LangID, IntrID, InstRate, Billable, AppTimeFrom, AppTimeTo, DeptID FROM request_T, Institution_T WHERE request_T.[instID] <> 479 AND InstID = Institution_T.[index] AND (status = 0 OR status = 1 OR status = 4) " & _
		"AND IntrID > 0 "
		If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY Facility"		
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("index") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetInst(rsRep("InstID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("InstRate") & "</td>" & vbCrLf
			billhrs = Z_fixNull(rsRep("Billable"))
			If billhrs = "" Then billHrs = InstBillHrs(rsRep("AppTimeFrom"), rsRep("AppTimeTo"), rsRep("InstID"), rsRep("DeptID"))
			strBody = strBody & "<td class='tblgrn2'><nobr>" & billhrs & "</td></tr>" & vbCrLf
									
			CSVBody = CSVBody & rsRep("index") & "," & GetInst(rsRep("InstID")) & "," & GetLang(rsRep("LangID")) & "," & _
				GetIntr(rsRep("IntrID")) & "," & rsRep("InstRate") & "," & billhrs & vbCrLf	
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If	 
ElseIf tmpReport(0) = 30 Then 'activity report
	RepCSV =  "Activity" & tmpdate & ".csv" 
	
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf & _
		"<td class='tblgrn'>STATUS</td>" & vbCrlf 
	
	CSVHead = "Request ID,Institution, Department, Appointment Date, Client Last Name, Client First Name, Language, Interpreter Last Name, Interpreter First Name, Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Emergency Surcharge, Total, Comments, STATUS"	
	
	sqlRep = "SELECT request_T.[index] as myindex, status, [Last Name], [First Name], Clname, Cfname, AStarttime, AEndtime, " & _
		"Billable, emerFEE, class, TT_Inst, M_Inst, request_T.InstID as myinstID, DeptID, LangID, appDate, InstRate, bilComment, Processed, ProcessedMedicaid FROM request_T, interpreter_T , dept_T WHERE request_T.[instID] <> 479 AND request_T.deptID =  dept_T.[index] AND IntrID = interpreter_T.[index]  AND (Status = 1 OR Status = 4 Or Status = 0)" 
	strMSG = "All Activity Report"
	
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & "."
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY AppDate DESC"
	rsRep.Open sqlRep, g_strCONN, 1, 3
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
			'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
			BillHours =  rsRep("Billable")
			'totHrs2 = Z_FormatNumber( totHrs/60, 2)
			If rsRep("emerFEE") = True Then
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
				ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
					tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst")) + tmpFeeO
				End If
			Else
				tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			End If
			totalPay = Z_FormatNumber(tmpPay, 2)
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetInst2(rsRep("myinstID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strIntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
				If rsRep("emerFEE") = True Then 
						If rsRep("class") = 3 Or rsRep("class") = 5 Then
							strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeL & "</td>" & vbCrLf
						Else
							strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
						End If
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
				End If
				strBody = strBody & "<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
					ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
						strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeO & "</td>" & vbCrLf
					End If
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
				End If
				strBody = strBody & "<td class='tblgrn2'><nobr><b>$" & totalPay & "</b></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("bilComment") & "</td>"
				If IsNull(rsRep("Processed")) And IsNull(rsRep("ProcessedMedicaid")) Then
					strBody = strBody & "<td class='tblgrn2'><nobr>" & GetStat(rsRep("status")) & "</td><tr>" & vbCrLf 
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>Billed</td><tr>" & vbCrLf 
				End If
		
			CSVBody = CSVBody & """" & rsRep("myindex") & """,""" & GetInst2(rsRep("myinstID")) & """,""" &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & rsRep("appDate") & """,""" & rsRep("Clname") & """,""" & rsRep("Cfname") &  """,""" & GetLang(rsRep("LangID")) & """,""" & rsRep("Last Name") & _
				""",""" & rsRep("First Name") & ""","""  & cTime(rsRep("AStarttime")) & """,""" & cTime(rsRep("AEndtime")) & """,""" & BillHours
				
			If rsRep("emerFEE") = True Then 
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					CSVBody = CSVBody & """,""" & tmpFeeL
				Else
					CSVBody = CSVBody & """,""" & rsRep("InstRate")
				End If
			Else
				CSVBody = CSVBody & """,""" & rsRep("InstRate")
			end if
			
			CSVBody = CSVBody & """,""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & ""","""
			
			If rsRep("emerFEE") = True Then 
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					CSVBody = CSVBody & "0.00"
				ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
					CSVBody = CSVBody & tmpFeeO
				End If
			Else
				CSVBody = CSVBody & "0.00"
			end if
			
			CSVBody = CSVBody & """,""" & totalPay & """,""" & rsRep("bilComment")
			If IsNull(rsRep("Processed")) And IsNull(rsRep("ProcessedMedicaid")) Then
				CSVBody = CSVBody & """,""" & GetStat(rsRep("status")) & """" & vbCrLf 
			Else
				CSVBody = CSVBody & """,""" & "Billed" & """" & vbCrLf 
			End If
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 31 Then 'emergency report
	RepCSV =  "Emergency" & tmpdate & ".csv"
	strMSG = "Emergency Appointments report "	
	strHead = "<td class='tblgrn'>ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>STATUS</td>" & vbCrlf
	
	CSVHead = "ID,Instituition,Department,Date,Language,Interpreter Name,Status"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, request_T.[InstID], LangID, IntrID, Emergency, missed, appdate, DeptID, facility, dept, status FROM request_T, Institution_T, Dept_T " & _
		"WHERE request_T.[instID] <> 479 AND request_T.[InstID] = Institution_T.[index] AND DeptID = Dept_T.[index] And [Emergency] = 1 "
		If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY AppDate, facility, Dept"		
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			myStat = ""
			If rsRep("Status") = 2 And rsRep("Missed") = 1 Then 
				myStat = "unable to send interpreter"
			Else
				myStat = "covered"
			End If
			If rsRep("IntrID") > 0 Or (rsRep("Status") = 2 And rsRep("Missed") = 1)  Then
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("facility") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("Dept") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("AppDate") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & myStat & "</td></tr>" & vbCrLf
	
				CSVBody = CSVBody & """" & rsRep("myindex") & """,""" & rsRep("facility") & """,""" & rsRep("Dept") & """,""" & _
					rsRep("AppDate") & """,""" & GetLang(rsRep("LangID")) & """,""" & GetIntr(rsRep("IntrID")) & """,""" &  myStat & """" & vbCrLf	
			End If
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If	 
ElseIf tmpReport(0) = 32 Then 'duration report
	RepCSV =  "Duration" & tmpdate & ".csv"
	strMSG = "Duration complete report "	
	strHead = "<td class='tblgrn'>ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Appt. Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Completed Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Bill Date</td>" & vbCrlf & _
		"<td class='tblgrn'>No. of Days to complete</td>" & vbCrlf
	
	CSVHead = "ID,Instituition,Department,Language,Date,Completed Date,Billed on,No. of Days to complete"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, request_T.[InstID], LangID, appdate, DeptID, facility, dept, timestamp, completed, processed, processedMedicaid FROM request_T, Institution_T, Dept_T " & _
		"WHERE request_T.[instID] <> 479 AND NOT completed IS NULL AND request_T.[InstID] = Institution_T.[index] AND DeptID = Dept_T.[index] "
		If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY AppDate, facility, Dept"		
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			myDays = DateDiff("d", DateValue(rsRep("AppDate")), DateValue(rsRep("completed")))
			myPDate = ""
			If Not Z_FixNull(rsRep("processed")) = "" Then 
				myPDate = DateValue(rsRep("processed"))
			Else
				If Not Z_FixNull(rsRep("processedMedicaid")) = "" Then myPDate = DateValue(rsRep("processedMedicaid"))
			End If
			
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("facility") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("Dept") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & DateValue(rsRep("AppDate")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & DateValue(rsRep("completed")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & myPDate & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & myDays & "</td></tr>" & vbCrLf

			CSVBody = CSVBody & """" & rsRep("myindex") & """,""" & rsRep("facility") & """,""" & rsRep("Dept") & """,""" & _
				rsRep("AppDate") & """,""" & GetLang(rsRep("LangID")) & """,""" &  DateValue(rsRep("completed")) & """,""" &  myPDate & """,""" &  myDays & """" & vbCrLf	

			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If	
ElseIf tmpReport(0) = 33 Then 'Institution report
	RepCSV =  "Institution.csv"
	strMSG = "Institution report "	
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Address</td>" & vbCrlf & _
		"<td class='tblgrn'>Billing Address</td>" & vbCrlf & _
		"<td class='tblgrn'>Class</td>" & vbCrlf & _
		"<td class='tblgrn'>Billing Group</td>" & vbCrlf & _
		"<td class='tblgrn'>Customer ID</td>" & vbCrlf
	
	CSVHead = "Instituition,Department,Address,Apartment/Suite Number,City,State,ZIP,Billing Address,City,State,Zip,Class,Billing Group,Customer ID"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT Facility, dept, InstAdrI, address, city, state, zip, baddress, bcity, bstate, bzip, class, billgroup, custID FROM Institution_T, Dept_T " & _
		"WHERE InstID = institution_T.[index]"
	If tmpReport(3) = "" Then tmpReport(3) = 0
	If tmpReport(3) <> 0 Then
		sqlRep = sqlRep & " AND institution_T.[index] = " & tmpReport(3) 
		strMSG = strMSG & " for " & GetInst(tmpReport(3))
	End If
	sqlRep = sqlRep & " ORDER BY facility, dept"		
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			myAddr = rsRep("address") & ", " & rsRep("InstAdrI") & ", " & rsRep("city") & ", " & rsRep("state") & ", " & rsRep("zip")
			myBAddr = rsRep("baddress") & ", " & rsRep("bcity") & ", " & rsRep("bstate") & ", " & rsRep("bzip")
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("facility") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("Dept") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & myAddr & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & myBAddr & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetClass(rsRep("class")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("billgroup") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("custID") & "</td></tr>" & vbCrLf
					
			CSVBody = CSVBody & """" & rsRep("facility") & """,""" & rsRep("Dept") & """,""" & _
				rsRep("address") & """,""" & rsRep("InstAdrI") & """,""" &  rsRep("city") & """,""" &  rsRep("state") & """,""" &  rsRep("zip") & _
				""",""" & rsRep("baddress") & """,""" & rsRep("bcity") & """,""" &  rsRep("bstate") & """,""" & rsRep("bzip") & _
				""",""" & GetClass(rsRep("class")) & """,""" & rsRep("billgroup") & """,""" &  rsRep("custID") & """" & vbCrLf	

			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If	
ElseIf tmpReport(0) = 34 Then 'oncall
	RepCSV =  "Oncall" & tmpdate & ".csv" 
	strMSG = "On call schedule for the month of " & MonthName(Month(tmpReport(1))) & " " & Year(tmpReport(1))
	CSVHead = "Date,Last Name,First Name"
	strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM oncall_T WHERE Month(OCdate) = " & Month(tmpReport(1)) & " AND Year(Ocdate) = " & Year(tmpReport(1)) & _
		" ORDER BY InstID, OCdate, pm"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		tmpInstID2 = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			tmpInstID = rsRep("InstID")
			If tmpInstID <> tmpInstID2 Or tmpInstID2 = 0 Then
				strBody = strBody & "<tr bgcolor='#FFFFCE'><td colspan='7' align='left'>" & GetInst(rsRep("InstID")) & "</td></tr>"
				CSVBody = CSVBody & """" & GetInst(rsRep("InstID")) & """" & vbCrLf
			End If
			ampm = "(am)"
			If rsRep("pm") Then ampm = "(pm)"
			strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("OCDate") & " " & ampm & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & """" & rsRep("OCDate") & " " & ampm & """,""" & GetIntr(rsRep("IntrID")) & """" & vbCrLf
			tmpInstID2 = tmpInstID
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 35 Then 'Billablehours
	RepCSV =  "HrsBillable" & tmpdate & ".csv" 
	strMSG = "Hours Billable for the month of " & MonthName(Month(tmpReport(1))) & " " & Year(tmpReport(1))
	strHead = "<td class='tblgrn'>Type</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf 
	CSVHead = "Type,Hours"
	totsum = 0
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT SUM(billable) AS SumBill FROM Request_T WHERE request_T.[instID] <> 479 AND Month(appdate) = " & Month(tmpReport(1)) & " AND Year(appdate) = " & Year(tmpReport(1)) & _
		" AND (NOT Processed IS NULL OR NOT ProcessedMedicaid IS NULL)"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
				"<td class='tblgrn2'><nobr>Billed Appointments</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_Czero(rsRep("SumBill")) & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & """" & "Billed Appointments" & """,""" & Z_Czero(rsRep("SumBill")) & """" & vbCrLf
		totsum = totsum +  Z_Czero(rsRep("SumBill"))
	End If
	rsRep.Close
	Set rsRep = Nothing
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT SUM(billable) AS SumBill FROM Request_T WHERE request_T.[instID] <> 479 AND Month(appdate) = " & Month(tmpReport(1)) & " AND Year(appdate) = " & Year(tmpReport(1)) & _
		" AND (Processed IS NULL OR ProcessedMedicaid IS NULL) AND Status = 1"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		strBody = strBody & "<tr bgcolor='#F5F5F5'>" & _
				"<td class='tblgrn2'><nobr>Completed Appointments</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_Czero(rsRep("SumBill")) & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & """" & "Completed Appointments" & """,""" & Z_Czero(rsRep("SumBill")) & """" & vbCrLf
		totsum = totsum +  Z_Czero(rsRep("SumBill"))
	End If
	rsRep.Close
	Set rsRep = Nothing
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT SUM(billable) AS SumBill FROM Request_T WHERE request_T.[instID] <> 479 AND Month(appdate) = " & Month(tmpReport(1)) & " AND Year(appdate) = " & Year(tmpReport(1)) & _
		" AND  (Processed IS NULL OR ProcessedMedicaid IS NULL) AND Status = 4"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
				"<td class='tblgrn2'><nobr>Canceled-Billable Appointments</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_Czero(rsRep("SumBill")) & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & """" & "Canceled-Billable Appointments" & """,""" & Z_Czero(rsRep("SumBill")) & """" & vbCrLf	
		totsum = totsum +  Z_Czero(rsRep("SumBill"))
	End If
	totestHrs = 0
	Set rsRep = Server.CreateObject("ADODB.RecordSet") '''not done
	sqlRep = "SELECT appTimeFrom, appTimeTo, InstID, deptID FROM Request_T WHERE request_T.[instID] <> 479 AND Month(appdate) = " & Month(tmpReport(1)) & " AND Year(appdate) = " & Year(tmpReport(1)) & _
		" AND Status = 0"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		Do Until rsRep.EOF
			estHrs = InstBillHrs2(rsRep("appTimeFrom"), rsRep("appTimeTo"), rsRep("InstID"), rsRep("deptID"))
			totestHrs = totestHrs + estHrs
			rsRep.MoveNext
		Loop
	End If
	rsRep.Close
	Set rsRep = Nothing
	totsum = totsum +  totestHrs
	strBody = strBody & "<tr bgcolor='#F5F5F5'>" & _
				"<td class='tblgrn2'><nobr>Pending Appointments</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & totestHrs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & """" & "Pending Appointments" & """,""" & totestHrs & """" & vbCrLf
	strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
			"<td class='tblgrn4'><nobr>TOTAL</td>" & vbCrLf & _
			"<td class='tblgrn4'><nobr>" & totsum & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & """" & "TOTAL" & """,""" & totsum & """" & vbCrLf
ElseIf tmpReport(0) = 36 Then 'pre mileage report
	RepCSV =  "XMileage" & tmpdate & "-" & tmpTime & ".csv" 
	tmpMonthYear = MonthName(Month(tmpReport(1))) & " - " & Year(tmpReport(1))
	strMSG = "Pre-Mileage report for the month of " & tmpMonthYear
	strHead = "<td class='tblgrn'>File Number</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Miles</td>" & vbCrlf & _
		"<td class='tblgrn'>Miles Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Receipts Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf 
	CSVHead = "File Number,Last Name,First Name,Miles,Miles Amount,Receipts Amount,Total"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT FileNum, actmil, overmile, appDate, Interpreter_T.[index] as myIntrIndex, Toll " & _
		" FROM Request_T, Interpreter_T WHERE request_T.[instID] <> 479 AND IntrID = Interpreter_T.[index] AND Month(appDate) = " & Month(tmpReport(1)) & " AND Year(appDate) = " & _
		Year(tmpReport(1)) & " "
	If Z_CZero(tmpReport(4)) > 0 Then
		sqlRep = sqlRep & "AND IntrID = " & tmpReport(4) & " "
		strMSG = strMSG & " for " & GetIntr(tmpReport(4)) & "."
	End If
	sqlRep = sqlRep & "AND LbconfirmToll = 1 AND mileageproc IS NULL ORDER BY [last name], [first name], appDate"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			'IntrName = rsRep("Last Name") & ", " & rsRep("First Name")
			strMile = Z_Czero(rsRep("actmil"))
			IntrID = rsRep("myIntrIndex")
			strTol = Z_Czero(rsRep("Toll"))
			lngIDx = SearchArraysHours(IntrID, tmpIntr)
			If lngIdx < 0 Then
				ReDim Preserve tmpIntr(x)
				ReDim Preserve tmpMile(x)
				ReDim Preserve tmpToll(x)
				
				tmpIntr(x) = IntrID
				tmpMile(x) = strMile
				tmpToll(x) = strTol
				x = x + 1
			Else	
				tmpMile(lngIdx) = tmpMile(lngIdx) + strMile
				tmpToll(lngIdx) = tmpToll(lngIdx) + strTol
			End If
			'rsRep("mileageproc") = date
			rsRep.MoveNext
		Loop
		y = 0
		Do Until y = x
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			myMile = Z_Czero(tmpMile(y))
			myToll = Z_Czero(tmpToll(y))
	
			mileTOT = MileageAmt(myMile) + myToll
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetFileNum(tmpIntr(y)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntr(tmpIntr(y)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myMile,2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(AmtRate(myMile),2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(myToll,2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(mileTOT,2) & "</td></tr>" & vbCrLf
								
			CSVBody = CSVBody & GetFileNum(tmpIntr(y)) & "," & GetIntr(tmpIntr(y)) & "," & Z_FormatNumber(myMile,2) & "," & _
				Z_FormatNumber(AmtRate(myMile),2) & "," & Z_FormatNumber(myToll,2) & "," & Z_FormatNumber(mileTOT,2) & vbCrLf
			y = y + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 37 Then
	RepCSV =  "InactiveInter" & tmpdate & ".csv" 
	strMSG = "Inactive Interpreter report"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
		"<td class='tblgrn'>eMail</td>" & vbCrlf
	CSVHead = "Last Name, First Name, eMail"
	sqlRep = "SELECT * FROM interpreter_T WHERE Active = 0 ORDER BY [Last Name], [First Name]"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		tmpName = rsRep("Last Name") & ", " & rsRep("First Name")
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("E-mail") & "</td></tr>" & vbCrLf 
		CSVBody = CSVBody & tmpName & ",""" & rsRep("E-mail") & """" & vbCrLf 
		y = y + 1
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 38 Then
	RepCSV =  "tardy" & tmpdate & ".csv" 
	strMSG = "Tardiness report"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Late(Mins.)</td>" & vbCrlf & _
		"<td class='tblgrn'>Reason</td>" & vbCrlf
	CSVHead = "Last Name, First Name, Date, TimeFrom, TimeTo, Insitution, Department, Late"
	sqlRep = "SELECT * FROM Request_T, Interpreter_T WHERE request_T.[instID] <> 479 AND IntrID = Interpreter_T.[index] AND (Not Late IS NULL And Late > 0)"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If Z_CZero(tmpReport(4)) > 0 Then
		sqlRep = sqlRep & "AND IntrID = " & tmpReport(4) & " "
		strMSG = strMSG & " for " & GetIntr(tmpReport(4)) & "."
	End If
	sqlRep = sqlRep & "ORDER BY [Last Name], [First Name]"
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		tmpName = rsRep("Last Name") & ", " & rsRep("First Name")
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & Ctime(rsRep("apptimefrom")) & " - " & Ctime(rsRep("apptimeto")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetInst(rsRep("InstID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetDept(rsRep("DeptID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("Late") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetReasonTardy(rsRep("lateres")) & "</td></tr>" & vbCrLf 
		CSVBody = CSVBody & tmpName & ",""" & rsRep("appdate") & """,""" & Ctime(rsRep("apptimefrom")) & """,""" & Ctime(rsRep("apptimeto")) & _
			""",""" & GetInst(rsRep("InstID")) & """,""" & GetDept(rsRep("DeptID")) & """,""" & rsRep("Late") & """,""" & GetReasonTardy(rsRep("lateres")) & """" & vbCrLf 
		y = y + 1
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 39 Then 'medicaid sim
	RepCSV =  "MedicaidXBillReq" & tmpdate & "-" & tmpTime & ".csv" 
	RepCSVBill = "MedicaidXBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
	CSVHeadBill = """" & FixDateFormat(date) & "LB Billing" & """"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Medicaid</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf & _
		"<td class='tblgrn'>DOB</td>" & vbCrlf & _
		"<td class='tblgrn'>DHHS</td>" & vbCrlf 
	
	CSVHead = "Request ID,Institution, Department, Appointment Date, Client Last Name, Client First Name, Medicaid, Language, Interpreter Last Name, Interpreter First Name, Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Emergency Surcharge, Total, Comments, DOB, DHHS"	
	' add vermed = 1 AND if medicaid billing is go 
	sqlRep = "SELECT syscom, wid, medicaid, vermed, autoacc, wcomp, drg, pid, request_T.[index] as myindex, status, [Last Name], [First Name], Clname, Cfname, AStarttime, AEndtime, " & _
		"Billable, DOB, emerFEE, class, TT_Inst, M_Inst, request_T.InstID as myinstID, DeptID, LangID, appDate, InstRate, bilComment, custID, ccode, billgroup, IntrID " & _
		"FROM request_T, interpreter_T , dept_T WHERE request_T.[instID] <> 479 AND outpatient = 1 AND hasmed = 1 AND vermed = 1 AND autoacc <> 1 AND wcomp <> 1 AND drg = 1 " & _
		"AND (medicaid <> '' OR NOT medicaid IS NULL) " & _
		"AND request_T.deptID = dept_T.[index] AND IntrID = interpreter_T.[index] " & _
		"AND (Status = 1 OR Status = 4) AND ProcessedMedicaid IS NULL" 
	strMSG = "Medicaid Billing request report (simulated)"
	
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & ". * - Cancelled Billable."
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY CustID ASC, AppDate DESC"
	'response.write sqlRep
	rsRep.Open sqlRep, g_strCONN, 1, 3
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	If Not rsRep.EOF Then 
		x = 0
		tmpCID = ""
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			CB = ""
			If rsRep("status") = 4 Then CB = "*"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
			'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
			BillHours =  rsRep("Billable")
			If BillHours > 4 Then BillHours = 4
			'totHrs2 = Z_FormatNumber( totHrs/60, 2)
			'If rsRep("emerFEE") = True Then
			'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
			'		tmpPay = (BillHours * tmpFeeL) '+ Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			'	ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
			'		tmpPay = (BillHours * rsRep("InstRate")) + tmpFeeO '+ Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			'	End If
			'Else
				tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			'End If
			totalPay = Z_FormatNumber(tmpPay, 2)
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & CB & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetInst2(rsRep("myinstID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("medicaid") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strIntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
				'If rsRep("emerFEE") = True Then 
				'		If rsRep("class") = 3 Or rsRep("class") = 5 Then
				'			strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeL & "</td>" & vbCrLf
				'		Else
				'			strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
				'		End If
				'Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
				'End If
				strBody = strBody & "<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
				'If rsRep("emerFEE") = True Then 
				'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
				'		strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
				'	ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
				'		strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeO & "</td>" & vbCrLf
				'	End If
				'Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
				'End If
				bilcomment = Z_fixNull(rsRep("bilComment") & rsRep("syscom"))
				strBody = strBody & "<td class='tblgrn2'><nobr><b>$" & totalPay & "</b></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & bilcomment & "</td><td class='tblgrn2'><nobr>" & rsRep("DOB") & "</td>"
				If rsRep("myinstID") = 108 Then
					strBody = strBody & "<td class='tblgrn2'><nobr>" & GetUserID(rsRep("DeptID")) & "</td><tr>" & vbCrLf 
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>&nbsp;</td><tr>" & vbCrLf 
				End If
		
			CSVBody = CSVBody & CB & rsRep("myindex") & "," & GetInst2(rsRep("myinstID")) & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & _
				rsRep("appDate") & "," & rsRep("Clname") & "," & rsRep("Cfname") & "," & rsRep("medicaid") &  "," & GetLang(rsRep("LangID")) & "," & rsRep("Last Name") & _
				"," & rsRep("First Name") & ","  & cTime(rsRep("AStarttime")) & "," & cTime(rsRep("AEndtime")) & "," & BillHours
		
			'If rsRep("emerFEE") = True Then 
			'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
			'		CSVBody = CSVBody & "," & tmpFeeL
			'	Else
			'		CSVBody = CSVBody & "," & rsRep("InstRate")
			'	End If
			'Else
				CSVBody = CSVBody & "," & rsRep("InstRate")
			'end if
			
			CSVBody = CSVBody & ",""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & ""","
			
			'If rsRep("emerFEE") = True Then 
			'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
			'		CSVBody = CSVBody & "0.00"
			'	ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
			'		CSVBody = CSVBody & tmpFeeO
			'	End If
			'Else
				CSVBody = CSVBody & "0.00"
			'end if
			bilcommentcsv = Replace(Z_fixNull(rsRep("bilComment") & rsRep("syscom")), "<br>", " / ")
			CSVBody = CSVBody & ",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB")
			If rsRep("myinstID") = 108 Then
				CSVBody = CSVBody & """,""" & GetUserID(rsRep("DeptID")) & """" & vbCrLf 
			Else
				CSVBody = CSVBody & """" & vbCrLf 
			End If
			If BillHours >= 2 Then
				'If BillHours = 2 Then
					CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LBHR" & """,""" & "" & """,""" & _
						FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
				'Else
				'	CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LBHR" & """,""" & "" & """,""" & _
				'		FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
				'	CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LB1QH" & """,""" & "" & """,""" & _
				'		FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
					tmpLbhrs = BillHours - 2
					CSVBodyBill = CSVBodyBill & GetLBCode(tmpLbhrs, rsRep("medicaid"), rsRep("wid"), rsRep("appDate"))
				'End If
			Else
				tmpLbhrs = BillHours
				CSVBodyBill = CSVBodyBill & GetLBCode(tmpLbhrs, rsRep("medicaid"), rsRep("wid"), rsRep("appDate"))
			End If
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	
	End If
	rsRep.Close
	Set rsRep = Nothing	
ElseIf tmpReport(0) = 40 Then 'medicaid billing
	RepCSV =  "MedicaidBillReq" & tmpdate & "-" & tmpTime & ".csv" 
	RepCSVBill = "MedicaidBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
	CSVHeadBill = """" & FixDateFormat(date) & "LB Billing" & """"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Medicaid</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf & _
		"<td class='tblgrn'>DOB</td>" & vbCrlf & _
		"<td class='tblgrn'>DHHS</td>" & vbCrlf 
	
	CSVHead = "Request ID,Institution, Department, Appointment Date, Client Last Name, Client First Name, Medicaid, Language, Interpreter Last Name, Interpreter First Name, Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Emergency Surcharge, Total, Comments, DOB, DHHS"	
	' add vermed = 1 AND if medicaid billing is go 
	sqlRep = "SELECT billingTrail, syscom ,wid, medicaid, vermed, autoacc, wcomp, drg, pid, request_T.[index] as myindex, status, [Last Name], [First Name], Clname, Cfname, AStarttime, AEndtime, " & _
		"Billable, DOB, emerFEE, class, TT_Inst, M_Inst, request_T.InstID as myinstID, DeptID, LangID, appDate, InstRate, bilComment, custID, ccode, billgroup, IntrID, ProcessedMedicaid " & _
		"FROM request_T, interpreter_T , dept_T WHERE request_T.[instID] <> 479 AND outpatient = 1 AND hasmed = 1 AND vermed = 1 AND autoacc <> 1 AND wcomp <> 1 AND drg = 1 " & _
		"AND (medicaid <> '' OR NOT medicaid IS NULL) " & _
		"AND request_T.deptID = dept_T.[index] AND IntrID = interpreter_T.[index] " & _
		"AND (Status = 1 OR Status = 4) AND ProcessedMedicaid IS NULL" 
	strMSG = "Medicaid Billing request report "
	
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & ". * - Cancelled Billable."
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY CustID ASC, AppDate DESC"
	'response.write sqlRep
	rsRep.Open sqlRep, g_strCONN, 1, 3
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	If Not rsRep.EOF Then 
		x = 0
		tmpCID = ""
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			CB = ""
			If rsRep("status") = 4 Then CB = "*"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
			'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
			BillHours =  rsRep("Billable")
			If BillHours > 4 Then BillHours = 4
			'totHrs2 = Z_FormatNumber( totHrs/60, 2)
			'If rsRep("emerFEE") = True Then
			'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
			'		tmpPay = (BillHours * tmpFeeL) '+ Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			'	ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
			'		tmpPay = (BillHours * rsRep("InstRate")) + tmpFeeO '+ Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			'	End If
			'Else
				tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			'End If
			totalPay = Z_FormatNumber(tmpPay, 2)
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & CB & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetInst2(rsRep("myinstID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("medicaid") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strIntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
				'If rsRep("emerFEE") = True Then 
				'		If rsRep("class") = 3 Or rsRep("class") = 5 Then
				'			strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeL & "</td>" & vbCrLf
				'		Else
				'			strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
				'		End If
				'Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
				'End If
				strBody = strBody & "<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
				'If rsRep("emerFEE") = True Then 
				'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
				'		strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
				'	ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
				'		strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeO & "</td>" & vbCrLf
				'	End If
				'Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
				'End If
				bilcomment = Z_fixNull(rsRep("bilComment") & rsRep("syscom"))
				strBody = strBody & "<td class='tblgrn2'><nobr><b>$" & totalPay & "</b></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & bilcomment & "</td><td class='tblgrn2'><nobr>" & rsRep("DOB") & "</td>"
				If rsRep("myinstID") = 108 Then
					strBody = strBody & "<td class='tblgrn2'><nobr>" & GetUserID(rsRep("DeptID")) & "</td><tr>" & vbCrLf 
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>&nbsp;</td><tr>" & vbCrLf 
				End If
		
		CSVBody = CSVBody & CB & rsRep("myindex") & "," & GetInst2(rsRep("myinstID")) & "," &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & _
				rsRep("appDate") & "," & rsRep("Clname") & "," & rsRep("Cfname") & "," & rsRep("medicaid") &  "," & GetLang(rsRep("LangID")) & "," & rsRep("Last Name") & _
				"," & rsRep("First Name") & ","  & cTime(rsRep("AStarttime")) & "," & cTime(rsRep("AEndtime")) & "," & BillHours
		
			'If rsRep("emerFEE") = True Then 
			'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
			'		CSVBody = CSVBody & "," & tmpFeeL
			'	Else
			'		CSVBody = CSVBody & "," & rsRep("InstRate")
			'	End If
			'Else
				CSVBody = CSVBody & "," & rsRep("InstRate")
			'end if
			
			CSVBody = CSVBody & ",""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & ""","
			
			'If rsRep("emerFEE") = True Then 
			'	If rsRep("class") = 3 Or rsRep("class") = 5 Then
			'		CSVBody = CSVBody & "0.00"
			'	ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
			'		CSVBody = CSVBody & tmpFeeO
			'	End If
			'Else
				CSVBody = CSVBody & "0.00"
			'end if
			bilcommentcsv = Replace(Z_fixNull(rsRep("bilComment") & rsRep("syscom")), "<br>", " / ")
			CSVBody = CSVBody & ",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB")
			If rsRep("myinstID") = 108 Then
				CSVBody = CSVBody & """,""" & GetUserID(rsRep("DeptID")) & """" & vbCrLf 
			Else
				CSVBody = CSVBody & """" & vbCrLf 
			End If
			'CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LBHR" & """,""" & "" & """,""" & _
			'	FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
			'If BillHours > 2 Then
			'	CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LB1QH" & """,""" & "" & """,""" & _
			'		FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
			'	tmpLbhrs = BillHours - 2
			'	CSVBodyBill = CSVBodyBill & GetLBCode(tmpLbhrs, rsRep("medicaid"), rsRep("wid"), rsRep("appDate"))
			'End If 
			If BillHours >= 2 Then
				'If BillHours = 2 Then
					CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LBHR" & """,""" & "" & """,""" & _
						FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
				'Else
				'	CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LBHR" & """,""" & "" & """,""" & _
				'		FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
				'	CSVBodyBill = CSVBodyBill & """" & rsRep("medicaid") & """,""" & rsRep("wid") & """,""" & Progcode(rsRep("wid")) & """,""" & "LB1QH" & """,""" & "" & """,""" & _
				'		FixDateFormat(rsRep("appDate")) & """,""" & FixDateFormat(rsRep("appDate")) & """,""" & "1" & """,""" & """" & vbCrLf 
					tmpLbhrs = BillHours - 2
					CSVBodyBill = CSVBodyBill & GetLBCode(tmpLbhrs, rsRep("medicaid"), rsRep("wid"), rsRep("appDate"))
				'End If
			Else
				tmpLbhrs = BillHours
				CSVBodyBill = CSVBodyBill & GetLBCode(tmpLbhrs, rsRep("medicaid"), rsRep("wid"), rsRep("appDate"))
			End If
			x = x + 1
			rsRep("billingTrail") = rsRep("billingTrail") & "<br>Billed to Medicaid " & Date
			rsRep("processedmedicaid") = Date
			rsRep.Update
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	
	End If
	rsRep.Close
	Set rsRep = Nothing	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	strLog = Now & vbTab & "Medicaid Billing ran by " & Session("UsrName") & "."
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
ElseIf tmpReport(0) = 41 Then 'timesheet and mileage data
	RepCSV =  "timemile" & tmpdate & ".csv" 
	strMSG = "Timesheet and Mileage report"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage(m)</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time(hrs)</td>" & vbCrlf
	CSVHead = "Date, Institution, Department, Last Name, First Name, Mileage(m), Travel Time(hrs)"
	sqlRep = "SELECT * FROM Request_T, Interpreter_T WHERE request_T.[instID] <> 479 AND IntrID = Interpreter_T.[index] AND (actTT > 0 OR actMil > 0) "
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	'If tmpReport(4) <> 0 Then
	'	sqlRep = sqlRep & "AND IntrID = " & tmpReport(4) & " "
	'	strMSG = strMSG & " for " & GetIntr(tmpReport(4)) & "."
	'End If
	sqlRep = sqlRep & "AND status = 1 ORDER BY [Last Name], [First Name]"
	'response.write sqlRep
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		tmpName = rsRep("Last Name") & ", " & rsRep("First Name")
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetInst(rsRep("InstID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetDept(rsRep("DeptID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("actMil") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("actTT") & "</td></tr>" & vbCrLf 
		CSVBody = CSVBody & rsRep("appdate") & ",""" & GetInst(rsRep("InstID")) & """,""" & GetDept(rsRep("DeptID")) & """," & GetIntr(rsRep("IntrID")) & _
			",""" &  rsRep("actMil") & """,""" & rsRep("actTT") & """" & vbCrLf 
		y = y + 1
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 42 Then 'billed hours
	RepCSV =  "BilledHours" & tmpdate & ".csv"
	strMSG = "Billed Hours report "	
	strHead = "<td class='tblgrn'>ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours Billed</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf
	CSVHead = "ID,Instituition,Language,Interpreter Last Name,Interpreter First Name,Rate,Hours Billed,Travel Time,Total"
	tothrsbill = 0
	totTT = 0
	tottot = 0
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index], InstID, LangID, IntrID, InstRate, Billable, AppTimeFrom, AppTimeTo, InstActTT FROM request_T, Institution_T WHERE request_T.[instID] <> 479 AND InstID = Institution_T.[index] AND (status = 1 OR status = 4) " & _
		"AND IntrID > 0 AND (NOT Processed IS NULL OR NOT ProcessedMedicaid IS NULL)"
		If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY Facility"		
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			biltot =  rsRep("Billable") + Z_Czero(rsRep("InstActTT"))
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("index") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetInst(rsRep("InstID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetIntr(rsRep("IntrID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("InstRate") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("Billable") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_Czero(rsRep("InstActTT")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & biltot & "</td></tr>" & vbCrLf
									
			CSVBody = CSVBody & rsRep("index") & "," & GetInst(rsRep("InstID")) & "," & GetLang(rsRep("LangID")) & "," & _
				GetIntr(rsRep("IntrID")) & "," & rsRep("InstRate") & "," & rsRep("Billable") & "," &  Z_Czero(rsRep("InstActTT")) & "," & _
				biltot & vbCrLf	
			tothrsbill = tothrsbill + rsRep("Billable")
			totTT = totTT + Z_Czero(rsRep("InstActTT"))
			tottot = tottot + biltot
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If	
	strBody = strBody & "<tr><td colspan='4'>&nbsp;</td><td bgcolor='#FFFFCE'><nobr>TOTAL:</td>" & vbCrLf & _
		"<td class='tblgrn2' bgcolor='#FFFFCE'><nobr>" & Z_formatNumber(tothrsbill, 2) & "</td>" & vbCrLf & _
		"<td class='tblgrn2' bgcolor='#FFFFCE'><nobr>" & Z_formatNumber(totTT, 2) & "</td>" & vbCrLf & _
		"<td class='tblgrn2' bgcolor='#FFFFCE'><nobr>" & Z_formatNumber(tottot, 2) & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & ",,,,,TOTAL:,""" & Z_formatNumber(tothrsbill, 2) & """,""" & Z_formatNumber(totTT, 2) & """,""" & Z_formatNumber(tottot, 2) & """" & vbCrLf	
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 43 Then 'weekly TS
	RepCSV =  "WeekTimesheet" & tmpdate & ".csv"
	strMSG = "Weekly Timesheet report from " & tmpReport(1) & " to " & tmpReport(2) & " for " & GetIntr(tmpReport(4)) & "."
	CSVHead = "Date,Last Name,First Name,Activity, Travel Time, Appt. Start Time,Appt. End Time,Total Hours, Payable Hours, Final Payable Hours"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	OrgDate = Cdate(tmpReport(1))
	x = 0
	FirstDate = OrgDate
	LastDate = DateAdd("d", 1, CDate(tmpReport(2)))
	Do Until FirstDate = LastDate
		strBody = strBody 
		sqlRep = "SELECT [Last Name], [First Name], InstID, Cfname, totalhrs, actTT, overpayhrs, AStarttime, AEndtime, appDate, payhrs, Interpreter_T.[index] as myintrID " & _
        "FROM Request_T, Interpreter_T WHERE IntrID = Interpreter_T.[index] AND IntrID = " & tmpReport(4) & " " & _
        "AND [LBconfirm] = 1 AND AppDate = '" & FirstDate & "'"
    rsRep.Open sqlRep, g_strCONN, 3, 1
    If Not rsRep.EOF Then
    	Do Until rsRep.EOF
    		IntrName = rsRep("Last Name") & ", " & rsRep("First Name")
	      CliName = GetInst(rsRep("InstID")) & " - " & rsRep("Cfname")
	      tmpAMTs = rsRep("totalhrs")
	      TT = Z_FormatNumber(rsRep("actTT"), 2)
	      If rsRep("overpayhrs") Then
	          PHrs = Z_FormatNumber(rsRep("payhrs"), 2)
	          OvrHrs = "*"
	      Else
	          PHrs = Z_FormatNumber(IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime")), 2)
	          OvrHrs = ""
	      End If
	      FPHrs = Z_CZero(PHrs) + Z_CZero(TT)
	      kulay = "#FFFFFF"
				If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
	      strBodytmp = strBodytmp & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & IntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & TT & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AEndtime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpAMTs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(PHrs, 2) & OvrHrs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(FPHrs, 2) & "</td></tr>" & vbCrLf 
	      CSVBody = CSVBody & """" & rsRep("appDate") & """,""" & rsRep("Last Name") & """,""" & rsRep("First Name") & """,""" & _
	              CliName & """,""" & TT & """,""" & CTime(rsRep("AStarttime")) & """,""" & CTime(rsRep("AEndtime")) & _
	              """,""" & tmpAMTs & """,""" & Z_FormatNumber(PHrs, 2) & OvrHrs & """,""" & Z_FormatNumber(FPHrs, 2) & """" & vbCrLf
	      totFPHrs = Z_CZero(totFPHrs) + Z_CZero(FPHrs)
    		y = y + 1
    		rsRep.MoveNext
    	Loop
    End If
    rsRep.Close
    If Weekday(FirstDate) = 7 Then
        If totFPHrs > 0 Then
        	strHeader = "<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>" & vbCrLf & _
        	"<tr>" & vbCrLf & _
						"<td valign='top' colspan='9'>" & vbCrLf & _
							"<table bgColor='white' border='0' cellSpacing='0' cellPadding='0' align='center'>" & vbCrLf & _
								"<tr>" & vbCrLf & _
									"<td>" & vbCrLf & _
										"<img src='images/LBISLOGO.jpg' align='center'>" & vbCrLf & _
									"</td>" & vbCrLf & _
								"</tr>" & vbCrLf & _
								"<tr>" & vbCrLf & _
									"<td align='center'>" & vbCrLf & _
										"261&nbsp;Sheep&nbsp;Davis&nbsp;Road,&nbsp;Concord,&nbsp;NH&nbsp;03301<br>" & vbCrLf & _
										"Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186" & vbCrLf & _
									"</td>" & vbCrLf & _
								"</tr>" & vbCrLf & _
							"</table>" & vbCrLf & _
						"</td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
        	"<tr bgcolor='#C2AB4B'><td colspan='9' align='center'>Weekly Timesheet report from " & GetSun(FirstDate) & " to " & FirstDate & " for " & GetIntr(tmpReport(4)) & ".</td></tr>" & vbCrLf & _
					"<tr><td class='tblgrn'>Date</td>" & vbCrlf & _
					"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
					"<td class='tblgrn'>Activity</td>" & vbCrlf & _
					"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
					"<td class='tblgrn'>Appt. Start Time</td>" & vbCrlf & _
					"<td class='tblgrn'>Appt. End Time</td>" & vbCrlf & _
					"<td class='tblgrn'>Total Hours</td>" & vbCrlf & _
					"<td class='tblgrn'>Payable Hours</td>" & vbCrlf & _
					"<td class='tblgrn'>Final Payable Hours</td></tr>" & vbCrlf 
        	strBody = strBody & strHeader & strBodytmp & "<tr bgcolor='#FFFFCE'><td colspan='8' class='tblgrn2'>&nbsp;</td><td class='tblgrn2'>" & Z_FormatNumber(totFPHrs,2) & "</td></tr></table><div class='page-break'><br></div>"
          csvbody = csvbody & ",,,,,,,,TOTAL:,""" & Z_FormatNumber(totFPHrs, 2) & """" & vbCrLf
          strBodytmp = ""
          totFPHrs = 0
        End If
    End If
		FirstDate = DateAdd("d", 1, FirstDate)
	Loop
	Set rsRep = Nothing
ElseIf tmpReport(0) = 44 Then 'court cost
	RepCSV =  "CourtCost" & tmpdate & ".csv"
	strMSG = "Court Appointment cost report "
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf
	CSVHead = "Institution,Department,Appointment Date,Client Last Name,Client First Name,Language," & _
    "Interpreter Last Name,Interpreter First Name,Appointment Start Time,Appointment End Time,Hours," & _
    "Rate,Travel Time,Mileage,Emergency Surcharge,Total"
  Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3, 1
	If Not rsRate.EOF Then
	    tmpFeeL = rsRate("FeeLegal")
	End If
	rsRate.Close
	Set rsRate = Nothing
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM request_T, institution_T, dept_T, interpreter_T " & _
    "WHERE request_T.[instID] = institution_T.[index] AND deptID = dept_T.[index] AND " & _
    "intrID = interpreter_T.[index] AND class = 3 AND NOT processed IS NULL AND Billable > 0 "
   If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY facility, dept, appdate"
	rsRep.Open sqlRep, g_strCONN, 3, 1		
	InstID2 = 0
	tottotalPay = 0
	x = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
    If InstID2 <> rsRep("InstID") And InstID2 <> 0 Then
    	strBody = strBody & "<tr><td colspan='12'>&nbsp;</td><td class='tblgrn2'><nobr>" & Z_formatNumber(tottotalPay, 2) & "</td></tr>" & vbCrLf
      CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay & vbCrLf
      tottotalPay = 0
    End If
    BillHours = rsRep("Billable")
    If rsRep("emerFEE") Then
        tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    Else
        tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    End If
    totalPay = Z_FormatNumber(tmpPay, 2)
    strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("facility") & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & rsRep("dept") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("clname") & ", " & rsRep("cfname") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetLang(rsRep("langID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & " - " & CTime(rsRep("AEndtime"))  & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
        
    CSVBody = CSVBody & """" & rsRep("facility") & """,""" & rsRep("dept") & """,""" & rsRep("appdate") & _
	    """,""" & rsRep("clname") & """,""" & rsRep("cfname") & """,""" & GetLang(rsRep("langID")) & _
	    """,""" & rsRep("Last Name") & """,""" & rsRep("First Name") & """,""" & _
	    CTime(rsRep("AStarttime")) & """,""" & CTime(rsRep("AEndtime")) & """,""" & _
	    BillHours & ""","""
	    
    If rsRep("emerFEE") Then
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & tmpFeeL & "</td>" & vbCrLf
      CSVBody = CSVBody & tmpFeeL & ""","""
    Else
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & rsRep("InstRate") & "</td>" & vbCrLf
      CSVBody = CSVBody & rsRep("InstRate") & ""","""
    End If
    strBody = strBody & "<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>0.00</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & totalPay & "</td></tr>" & vbCrLf	
    CSVBody = CSVBody & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & """,""" & _
        "0.00" & """,""" & totalPay & """" & vbCrLf
    tottotalPay = tottotalPay + tmpPay
    InstID2 = rsRep("InstID")
    x = x + 1
    rsRep.MoveNext
	Loop
	CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 45 Then 'court cost by lang
	RepCSV =  "CourtCostLang" & tmpdate & ".csv"
	strMSG = "Court Appointment cost report by Language "
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf
	CSVHead = "Institution,Department,Appointment Date,Client Last Name,Client First Name,Language," & _
    "Interpreter Last Name,Interpreter First Name,Appointment Start Time,Appointment End Time,Hours," & _
    "Rate,Travel Time,Mileage,Emergency Surcharge,Total"
  Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3, 1
	If Not rsRate.EOF Then
	    tmpFeeL = rsRate("FeeLegal")
	End If
	rsRate.Close
	Set rsRate = Nothing
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM request_T, institution_T, dept_T, interpreter_T, language_T " & _
    "WHERE request_T.[instID] = institution_T.[index] AND deptID = dept_T.[index] AND " & _
    "intrID = interpreter_T.[index] AND class = 3 AND NOT processed IS NULL AND Billable > 0 " & _
    "AND langID = language_T.[index] "
   If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY language, facility, dept, appdate"
	rsRep.Open sqlRep, g_strCONN, 3, 1		
	InstID2 = 0
	tottotalPay = 0
	x = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
    If InstID2 <> rsRep("LangID") And InstID2 <> 0 Then
    	strBody = strBody & "<tr><td colspan='12'>&nbsp;</td><td class='tblgrn2'><nobr>" & Z_formatNumber(tottotalPay, 2) & "</td></tr>" & vbCrLf
      CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay & vbCrLf
      tottotalPay = 0
    End If
    BillHours = rsRep("Billable")
    If rsRep("emerFEE") Then
        tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    Else
        tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    End If
    totalPay = Z_FormatNumber(tmpPay, 2)
    strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("facility") & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & rsRep("dept") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("clname") & ", " & rsRep("cfname") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetLang(rsRep("langID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & " - " & CTime(rsRep("AEndtime"))  & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
        
    CSVBody = CSVBody & """" & rsRep("facility") & """,""" & rsRep("dept") & """,""" & rsRep("appdate") & _
	    """,""" & rsRep("clname") & """,""" & rsRep("cfname") & """,""" & GetLang(rsRep("langID")) & _
	    """,""" & rsRep("Last Name") & """,""" & rsRep("First Name") & """,""" & _
	    CTime(rsRep("AStarttime")) & """,""" & CTime(rsRep("AEndtime")) & """,""" & _
	    BillHours & ""","""
	    
    If rsRep("emerFEE") Then
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & tmpFeeL & "</td>" & vbCrLf
      CSVBody = CSVBody & tmpFeeL & ""","""
    Else
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & rsRep("InstRate") & "</td>" & vbCrLf
      CSVBody = CSVBody & rsRep("InstRate") & ""","""
    End If
    strBody = strBody & "<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>0.00</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & totalPay & "</td></tr>" & vbCrLf	
    CSVBody = CSVBody & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & """,""" & _
        "0.00" & """,""" & totalPay & """" & vbCrLf
    tottotalPay = tottotalPay + tmpPay
    InstID2 = rsRep("LangID")
    x = x + 1
    rsRep.MoveNext
	Loop
	CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 46 Then 'court cost by freq
	RepCSV =  "CourtCostFreq" & tmpdate & ".csv"
	strMSG = "Court Appointment by Language Frequency "
	Set rsRepA = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf
	CSVHead = "Institution,Language,Count"
	sqlRepA = "SELECT distinct(institution_T.[index]) as IDinst, request_T.[instID], facility FROM request_T, institution_T, dept_T, interpreter_T " & _
	  "WHERE request_T.[instID] = institution_T.[index] AND deptID = dept_T.[index] AND " & _
	  "intrID = interpreter_T.[index] AND class = 3 AND NOT processed IS NULL AND Billable > 0 "
	If tmpReport(1) <> "" Then
		sqlRepA = sqlRepA & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRepA = sqlRepA & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRepA = sqlRepA & "ORDER BY facility"
	rsRepA.Open sqlRepA, g_strCONN, 3, 1
	Do Until rsRepA.EOF
		kulay = "#F5F5F5"
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRepA("facility") & "</td></tr>" & vbCrLf
	  CSVBody = CSVBody & """" & rsRepA("facility") & """" & vbCrLf
	  Set rsRep = Server.CreateObject("ADODB.RecordSet")
    sqlrep = "SELECT distinct(language_T.[index]) AS myLang, langID, language FROM request_T, dept_T, language_T WHERE " & _
	    "deptID = dept_T.[index] AND langID = language_T.[index] AND " & _
	    "request_T.[instID] = " & rsRepA("IDinst") & " AND class = 3 AND NOT processed IS NULL " & _
	    "AND Billable > 0 " 
	  If tmpReport(1) <> "" Then
			sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		End If
		If tmpReport(2) <> "" Then
			sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		End If
	  sqlrep = sqlrep & "order by [language]"
	 ' response.write sqlRep & "<br>"
    rsRep.Open sqlrep, g_strCONN, 3, 1
    Do Until rsRep.EOF
    	strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'><nobr>" & rsRep("language") & "</td>" & vbCrLf
      CSVBody = CSVBody & """" & """,""" & rsRep("language") & ""","""
      Set rsNum = Server.CreateObject("ADODB.RecordSet")
      sqlNum = "SELECT COUNT(request_T.[index]) AS myCount FROM request_T, Dept_T WHERE langID = " & rsRep("langID") & " AND " & _
	      "request_T.[instID] = " & rsRepA("IDinst") & " AND deptID = dept_T.[index] AND class = 3 AND " & _
	      "NOT processed IS NULL AND Billable > 0 "
	    If tmpReport(1) <> "" Then
				sqlNum = sqlNum & "AND appDate >= '" & tmpReport(1) & "' "
			End If
			If tmpReport(2) <> "" Then
				sqlNum = sqlNum & "AND appDate <= '" & tmpReport(2) & "' "
			End If
      rsNum.Open sqlNum, g_strCONN, 3, 1
      strBody = strBody & "<td class='tblgrn2'><nobr>" & rsNum("myCount") & "</td></tr>" & vbCrLf
      CSVBody = CSVBody & rsNum("myCount") & """" & vbCrLf
      rsNum.Close
      Set rsNum = Nothing
      rsRep.MoveNext
    Loop
    rsRep.Close
    Set rsRep = Nothing
    rsRepA.MoveNext
	Loop
	rsRepA.Close
	Set rsRepA = Nothing	
ElseIf tmpReport(0) = 47 Then 'Language frequency
	RepCSV =  "LangFreq" & tmpdate & ".csv"
	strMSG = "Language Frequency: " & tmpReport(1) & " to " & tmpReport(2)
	Set rsRepA = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf
	CSVHead = "Language,Count"
	sqlRepA = "SELECT [language], COUNT([langid]) AS [f] " & _
				"FROM [request_t] AS r " & _
				"INNER JOIN [language_t] AS l ON r.[langID]=l.[index] " & _
				"WHERE (status = 1 OR status = 0) AND instID <> 479 " & _
				"AND [appdate] >='" & tmpReport(1) & "' " & _
				"AND [appdate] <= '" & tmpReport(2) & "' " & _
				"GROUP BY [language] " & _
				"ORDER BY [language] ASC"
	rsRepA.Open sqlRepA, g_strCONN, 3, 1
	y = 0
	Do Until rsRepA.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		'strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRepA("language") & "</td>" & vbCrLf
    strBody = strBody & "<tr  bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRepA("language") & "</td>" & vbCrLf
		CSVBody = CSVBody & """" & rsRepA("language") & ""","""
		strBody = strBody & "<td class='tblgrn2'><nobr>" & rsRepA("f") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & rsRepA("f") & """" & vbCrLf
		rsRepA.MoveNext
		y = y + 1
	Loop
	rsRepA.Close
	Set rsRepA = Nothing		
ElseIf tmpReport(0) = 48 Then 'Interpreter Inactivity
	RepCSV =  "InterpreterInactivity" & tmpdate & "-" & tmpTime & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT [index] as myIntrID, [last name], [first name], IntrAdrI, address1, city, state, [zip code] FROM interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
	rsRep.Open sqlRep, g_strCONN, 1, 3	
	Do Until rsRep.EOF
		Set rsApp = Server.CreateObject("ADODB.RecordSet")
		sqlApp = "SELECT [index] FROM Request_T WHERE IntrID = " & rsRep("myIntrID")
		If tmpReport(1) <> "" Then
			sqlApp = sqlApp & " AND appDate >= '" & tmpReport(1) & "'"
		End If
		If tmpReport(2) <> "" Then
			sqlApp = sqlApp & " AND appDate <= '" & tmpReport(2) & "'"
		End If
		rsApp.Open sqlApp, g_strCONN, 1, 3
		If rsApp.EOF Then 'create letter
			strBody = strBody & "<table border='0'>" & _
				"<tr>" & _
					"<td align='center'>" & _
						"<img src='images/LBISLOGOtrans.gif' align='center' height='60px'>" & _
						"<br>" & _
						"261&nbsp;Sheep&nbsp;Davis&nbsp;Road,&nbsp;Concord,&nbsp;NH&nbsp;03301<br>" & _
						"Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186" & _
						"<br>" & _
					"</td>" & _
				"</tr>" & _
				"<tr><td>&nbsp;</td></tr>" & _ 
				"<tr><td>" & date & "</td></tr>" & _
				"<tr><td>&nbsp;</td></tr>" & _ 
				"<tr><td>" & rsRep("first name") & " " & rsRep("last name") & "</td></tr>" & _
				"<tr><td>" & rsRep("address1") & ", " & rsRep("IntrAdrI") & ", " & rsRep("city") & ", " & rsRep("state") & ", " & rsRep("zip code") & "</td></tr>" & _
				"<tr><td>&nbsp;</td></tr>" & _ 
				"<tr><td>Dear " & rsRep("first name") & ",</td></tr>" & _
				"<tr><td>&nbsp;</td></tr>" & _ 
				"<tr><td><p>According to our records, your last appointment was <u>" & LastApp(rsRep("myIntrID")) & "</u>.You have not accepted any appointment since your last appointment.<br><br>" & _
				"We are currently in the process of updating our Interpreter list and are inquiring as to your employment status with our agency. Please indicate below your interest and return this letter to me.<br><br>" & _
				"<b>_____ Yes, I am still interested in working with LSS Language Bank.<br><br>" & _
				"_____ No, I am no longer interested in employment with LSS Language Bank. Please take my name off the Interpreter List.****</b><br><br>" & _
				"Your Signature:______________________________________   Date:____________________________<br><br>" & _
				"<i>****Please note that if a return response is not received within 14 days, we will assume that you are no longer interested in employment with LSS Language Bank.</i><br><br>" & _
				"Thank you for your interest in this matter.<br><br>" & _
				"Sincerely,<br><br>" & _
				"<font face='Script MT Bold'><i>Alen Omerbegovic</i></font><br><br>" & _
				"<i><b>Alen Omerbegovic</b></i><br>" & _
				"Program Manager<br>" & _
				"LanguageBank<br>" & _
				"Lutheran Social Services<br>" & _ 
				"261 Sheep Davis Road, Suite A-1<br>" & _ 
				"Concord, NH 03301<br>" & _ 
				"603-410-6183  603-410-6186 fax</p>" & _
				"</td></tr>" & _
				"</table><div class='page-break'><br></div>" & vbCrLf
		End If	
		rsApp.Close
		Set rsApp = Nothing
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 49 Then 'alen's report
	countMiss = 0
	countAssign = 0
	countNew = 0
	RepCSV =  "AlenReport" & tmpdate & ".csv"
	strMSG = "Alen's report "
	Set rsRepMiss = Server.CreateObject("ADODB.RecordSet")
	Set rsRepAssign = Server.CreateObject("ADODB.RecordSet")
	Set rsRepNew = Server.CreateObject("ADODB.RecordSet")
	sqlRepMiss = "SELECT COUNT([index]) as countMiss FROM Request_T WHERE missed = 1 "
	sqlRepAssign = "SELECT DISTINCT(LBID) FROM Hist_T WHERE LBID > 0 "
	sqlRepNew = "SELECT COUNT([index]) as countNew FROM Request_T WHERE NOT timestamp IS NULL "
	strHead = "<td class='tblgrn'>Criteria</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf
	CSVHead = "Criteria,Count"
	If tmpReport(1) <> "" Then
		sqlRepMiss = sqlRepMiss & "AND appDate >= '" & tmpReport(1) & "' "
		sqlRepAssign = sqlRepAssign & "AND timestamp >= '" & tmpReport(1) & "' "
		sqlRepNew = sqlRepNew & "AND timestamp >= '" & tmpReport(1) & "' "
		sqlRepNew = sqlRepNew & "AND appdate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRepMiss = sqlRepMiss & "AND appDate <= '" & tmpReport(2) & "' "
		sqlRepAssign = sqlRepAssign & "AND timestamp <= '" & tmpReport(2) & "' "
		sqlRepNew = sqlRepNew & "AND timestamp <= '" & tmpReport(2) & "' "
		sqlRepNew = sqlRepNew & "AND appdate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	rsRepMiss.Open sqlRepMiss, g_strCONN, 3, 1
	If Not rsRepMiss.EOF Then
		countMiss = Z_Czero(rsRepMiss("countMiss"))
	End If
	rsRepMiss.Close
	Set rsRepMiss = Nothing
	strBody = strBody & "<tr><td class='tblgrn2'>No interpreter Available (Missed)</td><td class='tblgrn2'>" & countMiss & "</td></tr>" & vbCrLf
	csvBody = csvBody & "No interpreter Available (Missed)," & countMiss & vbCrLf
	rsRepAssign.Open sqlRepAssign, g_strCONNHIST2, 3, 1
	If Not rsRepAssign.EOF Then
		Do Until rsRepAssign.EOF
			tmpindex = ""
			Set rsHist = Server.CreateObject("ADODB.RecordSet")
			rsHist.Open "SELECT * FROM Hist_T WHERE LBID = " & rsRepAssign("LBID") & " ORDER BY [index] DESC", g_strCONNHIST2, 3, 1
			ctr = 0
			Do Until rsHist.EOF
				ReDim Preserve arrTS(ctr)
        ReDim Preserve arrAuthor(ctr)
        ReDim Preserve arrPage(ctr)
        arrTS(ctr) = rsHist("timestamp")
        arrAuthor(ctr) = rsHist("author")
        arrPage(ctr) = rsHist("pageused")
				arrHist = Split(rsHist("hist"), """,""")
				If Ubound(arrHist) > 0 Then
					arrIntr = Split(arrhist(22), "|")
          intrID = arrIntr(0)
					If intrID = "-1" Or intrID = "0" Then
						If ctr > 0 Then
	            If arrTS(ctr - 1) >= cdate(tmpReport(1)) And arrTS(ctr - 1) <= cdate(tmpReport(2)) Then
	              countAssign = countAssign + 1
	              Exit Do
	            End If
	            Exit Do
            End If
          End If
				End If
				rsHist.MoveNext
				ctr = ctr + 1
			Loop
			rsHist.Close
			Set rsHist = Nothing
			rsRepAssign.MoveNext
		Loop
	End If
	rsRepAssign.Close
	Set rsRepAssign = Nothing
	strBody = strBody & "<tr><td class='tblgrn2'>Assigned Interpreters</td><td class='tblgrn2'>" & countAssign & "</td></tr>" & vbCrLf
	csvBody = csvBody & "Assigned Interpreters," & countAssign & vbCrLf
	rsRepNew.Open sqlRepNew, g_strCONN, 3, 1
	If Not rsRepNew.EOF Then
		countNew = Z_Czero(rsRepNew("countNew"))
	End If
	rsRepNew.Close
	Set rsRepNew = Nothing
	strBody = strBody & "<tr><td class='tblgrn2'>New Appointments</td><td class='tblgrn2'>" & countNew & "</td></tr>" & vbCrLf
	csvBody = csvBody & "New Appointments," & countNew & vbCrLf
ElseIf tmpReport(0) = 50 Then 'alen's report 2
	countMiss = 0
	countAssign = 0
	countNew = 0
	countnointr = 0
	RepCSV =  "AlenReport2" & tmpdate & ".csv"
	strMSG = "Alen's report 2 "
	Set rsRepMiss = Server.CreateObject("ADODB.RecordSet")
	Set rsRepAssign = Server.CreateObject("ADODB.RecordSet")
	Set rsRepNew = Server.CreateObject("ADODB.RecordSet")
	Set rsRepNoIntr = Server.CreateObject("ADODB.RecordSet")
	sqlRepMiss = "SELECT COUNT([index]) as countMiss FROM Request_T WHERE missed = 1 "
	sqlRepAssign = "SELECT DISTINCT(LBID) FROM Hist_T WHERE LBID > 0 "
	sqlRepNew = "SELECT COUNT([index]) as countNew FROM Request_T WHERE NOT timestamp IS NULL "
	sqlRepNoIntr = "SELECT COUNT([index]) as countNoIntr FROM Request_T WHERE (intrID = 0 Or intrID = -1) "
	strHead = "<td class='tblgrn'>Criteria</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf
	CSVHead = "Criteria,Count"
	If tmpReport(1) <> "" Then
		sqlRepMiss = sqlRepMiss & "AND appDate >= '" & tmpReport(1) & "' "
		sqlRepAssign = sqlRepAssign & "AND timestamp >= '" & tmpReport(1) & "' "
		sqlRepNew = sqlRepNew & "AND timestamp >= '" & tmpReport(1) & "' "
		sqlRepNoIntr = sqlRepNoIntr & "AND timestamp >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRepMiss = sqlRepMiss & "AND appDate <= '" & tmpReport(2) & "' "
		sqlRepAssign = sqlRepAssign & "AND timestamp <= '" & tmpReport(2) & "' "
		sqlRepNew = sqlRepNew & "AND timestamp <= '" & tmpReport(2) & "' "
		sqlRepNoIntr = sqlRepNoIntr & "AND timestamp <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	rsRepNew.Open sqlRepNew, g_strCONN, 3, 1
	If Not rsRepNew.EOF Then
		countNew = Z_Czero(rsRepNew("countNew"))
	End If
	rsRepNew.Close
	Set rsRepNew = Nothing
	strBody = strBody & "<tr><td class='tblgrn2'>New Appointments</td><td class='tblgrn2'>" & countNew & "</td></tr>" & vbCrLf
	csvBody = csvBody & "New Appointments," & countNew & vbCrLf
	rsRepAssign.Open sqlRepAssign, g_strCONNHIST2, 3, 1
	If Not rsRepAssign.EOF Then
		Do Until rsRepAssign.EOF
			tmpindex = ""
			Set rsHist = Server.CreateObject("ADODB.RecordSet")
			rsHist.Open "SELECT * FROM Hist_T WHERE LBID = " & rsRepAssign("LBID") & " ORDER BY [index] DESC", g_strCONNHIST2, 3, 1
			ctr = 0
			Do Until rsHist.EOF
				ReDim Preserve arrTS(ctr)
        ReDim Preserve arrAuthor(ctr)
        ReDim Preserve arrPage(ctr)
        arrTS(ctr) = rsHist("timestamp")
        arrAuthor(ctr) = rsHist("author")
        arrPage(ctr) = rsHist("pageused")
				arrHist = Split(rsHist("hist"), """,""")
				If Ubound(arrHist) > 0 Then
					arrIntr = Split(arrhist(22), "|")
          intrID = arrIntr(0)
					If intrID = "-1" Or intrID = "0" Then
						If ctr > 0 Then
	            If arrTS(ctr - 1) >= cdate(tmpReport(1)) And arrTS(ctr - 1) <= cdate(tmpReport(2)) Then
	              countAssign = countAssign + 1
	              Exit Do
	            End If
	            Exit Do
            End If
          End If
				End If
				rsHist.MoveNext
				ctr = ctr + 1
			Loop
			rsHist.Close
			Set rsHist = Nothing
			rsRepAssign.MoveNext
		Loop
	End If
	rsRepAssign.Close
	Set rsRepAssign = Nothing
	strBody = strBody & "<tr><td class='tblgrn2'>Assigned Interpreters</td><td class='tblgrn2'>" & countAssign & "</td></tr>" & vbCrLf
	csvBody = csvBody & "Assigned Interpreters," & countAssign & vbCrLf
	rsRepNoIntr.Open sqlRepNoIntr, g_strCONN, 3, 1
	If Not rsRepNoIntr.EOF Then
		countNoIntr = Z_Czero(rsRepNoIntr("countNoIntr"))
	End If
	rsRepNoIntr.Close
	Set rsRepNoIntr = Nothing
	strBody = strBody & "<tr><td class='tblgrn2'>Open Appointments</td><td class='tblgrn2'>" & countNoIntr & "</td></tr>" & vbCrLf
	csvBody = csvBody & "Open Appointments," & countNoIntr & vbCrLf
	rsRepMiss.Open sqlRepMiss, g_strCONN, 3, 1
	If Not rsRepMiss.EOF Then
		countMiss = Z_Czero(rsRepMiss("countMiss"))
	End If
	rsRepMiss.Close
	Set rsRepMiss = Nothing
	strBody = strBody & "<tr><td class='tblgrn2'>No interpreter Available (Missed)</td><td class='tblgrn2'>" & countMiss & "</td></tr>" & vbCrLf
	csvBody = csvBody & "No interpreter Available (Missed)," & countMiss & vbCrLf
	
ElseIf tmpReport(0) = 51 Then
	
ElseIf tmpReport(0) = 52 Then 'Lynda's report
	RepCSV =  "LyndaReport" & tmpdate & ".csv"
	strMSG = "Lynda's Report: " & tmpReport(1) & " to " & tmpReport(2)
	strHead = "<td class='tblgrn'>Type</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf
	CSVHead = "Type,Count"
	'new request
	Set rsNew = Server.CreateObject("ADODB.RecordSet")
	sqlNew = "SELECT Count([index]) AS myctr FROM Request_T WHERE timestamp >= '" & tmpReport(1) & "' AND timestamp <= '" & tmpReport(2) & "'"
	rsNew.Open sqlNew, g_strCONN, 3, 1
	If Not rsNew.EOF Then
		strBody = strBody & "<tr bgcolor='#FFFFFF'><td class='tblgrn2'>New Requests</td><td class='tblgrn2'><nobr>" & rsNew("myctr") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & "New Requests," & rsNew("myctr") & vbCrLf
	End If
	rsNew.Close
	Set rsNew = Nothing
	'All request
	Set rsNew = Server.CreateObject("ADODB.RecordSet")
	sqlNew = "SELECT Count([index]) AS myctr FROM Request_T WHERE appdate >= '" & tmpReport(1) & "' AND appdate <= '" & tmpReport(2) & "'"
	rsNew.Open sqlNew, g_strCONN, 3, 1
	If Not rsNew.EOF Then
		strBody = strBody & "<tr bgcolor='#F5F5F5'><td class='tblgrn2'>All Requests</td><td class='tblgrn2'><nobr>" & rsNew("myctr") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & "All Requests," & rsNew("myctr") & vbCrLf
	End If
	rsNew.Close
	Set rsNew = Nothing
	'with interpreters
	Set rsNew = Server.CreateObject("ADODB.RecordSet")
	sqlNew = "SELECT Count([index]) AS myctr FROM Request_T WHERE appdate >= '" & tmpReport(1) & "' AND appdate <= '" & tmpReport(2) & "' AND IntrID > 0"
	rsNew.Open sqlNew, g_strCONN, 3, 1
	If Not rsNew.EOF Then
		strBody = strBody & "<tr bgcolor='#FFFFFF'><td class='tblgrn2'>Requests with interpreters</td><td class='tblgrn2'><nobr>" & rsNew("myctr") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & "Requests with interpreters," & rsNew("myctr") & vbCrLf
	End If
	rsNew.Close
	Set rsNew = Nothing
	'without interpreters
	Set rsNew = Server.CreateObject("ADODB.RecordSet")
	sqlNew = "SELECT Count([index]) AS myctr FROM Request_T WHERE appdate >= '" & tmpReport(1) & "' AND appdate <= '" & tmpReport(2) & "' AND IntrID < 1"
	rsNew.Open sqlNew, g_strCONN, 3, 1
	If Not rsNew.EOF Then
		strBody = strBody & "<tr bgcolor='#F5F5F5'><td class='tblgrn2'>Requests without interpreters</td><td class='tblgrn2'><nobr>" & rsNew("myctr") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & "Requests without interpreters," & rsNew("myctr") & vbCrLf
	End If
	rsNew.Close
	Set rsNew = Nothing
ElseIf tmpReport(0) = 53 Then 'audit medicaid
	ctr = 10
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	RepCSV =  "AuditMedicaid" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")	
	sqlRep = "SELECT request_T.[index] as myindex, Facility, deptID, Clname, Cfname, appDate, billable, TT_Inst, M_Inst, " & _
		"emerFEE, Class, instRate, processed FROM request_T, institution_T, dept_T WHERE request_T.[instID] <> 479 AND request_T.InstID = institution_T.[index] " & _
		"AND NOT processedMedicaid IS NULL AND DeptID = dept_T.[index]"
	strMSG = "Audit report (medicaid)"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If tmpReport(3) = "" Then tmpReport(3) = 0
	If tmpReport(3) <> 0 Then
		sqlRep = sqlRep & " AND institution_T.[index] = " & tmpReport(3) 
		strMSG = strMSG & " for " & GetInst(tmpReport(3))
	End If
	If tmpReport(9) = "" Then tmpReport(9) = 0
	If tmpReport(9) <> 0 Then
		If tmpReport(6) = "" Then tmpReport(6) = 0
		If tmpReport(6) <> 0 Then 
			sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
		End If
		If tmpReport(7) = "" Then tmpReport(7) = 0
		If tmpReport(7) <> "0" Then
			tmpCli = Split(tmpReport(7), ",")
			sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
		End If
		If tmpReport(8) = "" Then tmpReport(8) = 0
		If tmpReport(8) <> 0 Then 
			sqlRep = sqlRep & " AND Class = " & tmpReport(8)
		End If
	End If
	sqlRep = sqlRep & " ORDER BY Facility"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Fee</td>" & vbCrlf
	CSVHead = "Request ID, Institution,Department,Client Last Name, Client First Name, Date, Hours, Rate, Travel, Mileage, Emergency Fee"	
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			myhours = Z_CZero(rsRep("billable"))
			If Not IsNull(rsRep("processed")) Then 
				If myhours > 4 Then myhours = 4
			End If
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Facility") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" &  GetDept(rsRep("deptID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_formatNumber(myhours, 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("instRate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & EFee(rsRep("emerFEE"), rsRep("Class"), tmpFeeL, tmpFeeO) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("myindex") & "," & rsRep("Facility") & "," &  GetDept(rsRep("deptID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") & "," &  rsRep("appDate") & _
				"," & Z_formatNumber(myhours, 2) & "," & rsRep("instRate") & "," & Z_CZero(rsRep("TT_Inst")) & "," & Z_CZero(rsRep("M_Inst")) & "," & EFee(rsRep("emerFEE"), rsRep("Class"), tmpFeeL, tmpFeeO) & vbCrLf
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='11' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
End If
tmpBills = Request("Bill")
If Request("csv") <> 1 Then
	'CONVERT TO CSV
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
	Prt.WriteLine "LANGUAGE BANK - REPORT"
	Prt.WriteLine strMSG
	Prt.WriteLine CSVHead
	Prt.WriteLine CSVBody
	Prt.Close	
	Set Prt = Nothing
	
	If Z_CZero(tmpReport(0)) = 3 Or Z_CZero(tmpReport(0)) = 16 Then 'additional csv for billing
		Set Prt2 = fso.CreateTextFile(RepPath &  RepCSVBill, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		'Prt.WriteLine strMSG
		Prt2.WriteLine CSVBodyBill
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBill, BackupStr
		
		
		Set Prt2 = fso.CreateTextFile(RepPath &  RepCSVBillL, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		'Prt.WriteLine strMSG
		Prt2.WriteLine CSVBodyBillL
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBillL, BackupStr
		
		Set Prt3 = fso.CreateTextFile(RepPath &  RepCSVBillSigma, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		'Prt.WriteLine strMSG
		Prt3.WriteLine CSVBodyBillSigma
		Prt3.Close	
		Set Prt3 = Nothing
		fso.CopyFile RepPath & RepCSVBillSigma, BackupStr

	End If
	If Z_CZero(tmpReport(0)) = 39  Or Z_CZero(tmpReport(0)) = 40 Then 'medicaid CSV
		Set Prt2 = fso.CreateTextFile(RepPath &  RepCSVBill, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		Prt2.WriteLine CSVHeadBill
		Prt2.Write CSVBodyBill
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBill, BackupStr
	End If
	'COPY FILE TO BACKUP
	
	fso.CopyFile RepPath & RepCSV, BackupStr
	'If Request("bill") = 1 Then
	'	Set Prt2 = fso.CreateTextFile(RepPath2 & RepCSV2, True)
	'	Prt2.WriteLine "LANGUAGE BANK - REPORT"
	'	Prt2.WriteLine strMSG2
	'	Prt2.WriteLine CSVHead2
	'	Prt2.WriteLine CSVBody2
	'	Prt2.Close	
	'	Set Prt2 = Nothing
		'COPY FILE TO BACKUP
	'	fso.CopyFile RepPath2 &  RepCSV2, BackupStr
	'End If
	Set fso = Nothing
	'EXPORT CSV
	'If Request("bill") <> 1 Then
		tmpstring = "CSV/" & repCSV 'add for RepCSVBill
		tmpstring2 = "CSV/" & RepCSVBill
		tmpstring3 = "CSV/" & RepCSVBillL
		tmpstring4 = "CSV/" & RepCSVBillSigma
	'Else
	'	tmpstring= "CSV/" & repCSV2
	'End IF
Else
	'EXPORT CSV
	
	'Set dload = Server.CreateObject("SCUpload.Upload")

	'If Request("bill") <> 1 Then
		'dload.Download RepCSV
	'	tmpstring = "CSV/InstBillReq262007.csv"
	'Else
		'tmpstring= "CSV/IntrBillReq262007.csv"
	'	'dload.Download RepCSV2
	'End IF
	'Set dload = Nothing
End If
%>
<html>
	<head>
		<title>Language Bank - Report Result</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function exportMe()
		{
			document.frmResult.action = "printreport.asp?csv=1"
			document.frmResult.submit();
		}
		function PassMe(xxx)
		{
			window.opener.document.frmReport.hideID.value = xxx;
			window.opener.SubmitAko();
			self.close();
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmResult'>
			<% If tmpReport(0) <> 43 And tmpReport(0) <> 26 And tmpReport(0) <> 48 Then %>
				<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
					<tr>
						<td valign='top'>
							<table bgColor='white' border='0' cellSpacing='0' cellPadding='0' align='center'>
							<tr>
								<td>
									<img src='images/LBISLOGO.jpg' align='center'>
								</td>
							</tr>
							<tr>
								<td align='center'>
									261&nbsp;Sheep&nbsp;Davis&nbsp;Road,&nbsp;Concord,&nbsp;NH&nbsp;03301<br>
									Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
								</td>
							</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr bgcolor='#C2AB4B'>
									<td colspan='<%=ctr + 7%>' align='center'>
										<% If Request("bill") <> 1 Then %>
											<b><%=strMSG%></b>
										<% Else %>
											<b><%=strMSG2%></b>
										<% End If%>
									</td>
								</tr>
								<tr>
									<% If Request("bill") <> 1 Then %>
										<%=strHead%>
									<% Else %>
										<%=strHead2%>
									<% End If%>
								</tr>
								<% If Request("bill") <> 1 Then %>
									<%=strBody%>
								<% Else %>
									<%=strBody2%>
								<% End If%>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<%'<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='exportMe();'>%>
										<% If Z_CZero(tmpReport(0)) <> 34  Then %>
											<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
										<% End If %>
										<% If Z_CZero(tmpReport(0)) = 3 Or Z_CZero(tmpReport(0)) = 16 Then %>
											<input class='btn' type='button' value='LB Billing CSV' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring2%>';">
											<input class='btn' type='button' value='LB Legal CSV' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring3%>';">
											<input class='btn' type='button' value='LB Comp Sigma CSV' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring4%>';">
										<% End If %>
										<% If Z_CZero(tmpReport(0)) = 39 Or Z_CZero(tmpReport(0)) = 40 Then %>
											<input class='btn' type='button' value='Medicaid Billing CSV' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring2%>';">
										<% End If %>
									</td>
								</tr>
									<td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
										* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
									</td>
								<tr>
								</tr>
							</table>	
						</td>
					</tr>
				</table>
			<% Else %>
				<p align="center">
					<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
				</p>
				<%=strBody%>
			<% End If %>
		</form>
	</body>
</html>

<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_googleDMA.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
Function Z_YMDDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = Z_CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_YMDDate = "0"
	Z_YMDDate = Z_YMDDate & lngTmp & "-"
	lngTmp = Z_CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp
	strTmp = DatePart("yyyy", dtDate)
	Z_YMDDate = strTmp & "-" & Z_YMDDate
End Function

Function Z_hoursAfter(rid, intrid, InstID, adate, timeto)
	' Z_hoursAfter(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimeto")) Or _
	Z_hoursAfter = False
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT [index] AS myid FROM request_T WHERE [index] <> " & rid & " AND intrID = " & intrID & " AND InstID = " & InstID & " AND " & _
		"appdate = '" & adate & "' AND apptimefrom >= '" & timeto & "' AND apptimefrom <= '" & DateAdd("n", 120, timeto) & "' AND status <> 3 "
	rsApp.Open sqlApp, g_strCONN, 3, 1
	If Not rsApp.EOF Then Z_hoursAfter = True
	rsApp.Close
	Set rsApp = Nothing
End Function
Function Z_hoursBefore(rid, intrid, InstID, adate, timefrom)
	' Z_hoursBefore(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimefrom")) Or _
	Z_hoursBefore = False
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT [index] AS myid FROM request_T WHERE [index] <> " & rid & " AND intrID = " & intrID & " AND InstID = " & InstID & " AND " & _
		"appdate = '" & adate & "' AND apptimeto <= '" & timefrom & "' AND apptimeto >= '" & DateAdd("n", -120, timefrom) & "' AND status <> 3 "
	rsApp.Open sqlApp, g_strCONN, 3, 1
	If Not rsApp.EOF Then Z_hoursBefore = True
	rsApp.Close
	Set rsApp = Nothing
End Function
Function Z_hoursSame(rid, intrid, InstID, adate, timefrom)
	' Z_hoursSame(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimefrom")) Or _
	Z_hoursSame = False
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT [index] AS myid FROM request_T WHERE [index] <> " & rid & " AND intrID = " & intrID & " AND InstID = " & InstID & " AND " & _
		"appdate = '" & adate & "' AND apptimefrom = '" & timefrom & "' AND status <> 3 "
	rsApp.Open sqlApp, g_strCONN, 3, 1
	If Not rsApp.EOF Then Z_hoursSame = True
	rsApp.Close
	Set rsApp = Nothing
End Function
Function Z_hoursOverFrom(rid, intrid, InstID, adate, timefrom)
	' Z_hoursOverFrom(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimefrom")) Or _
	Z_hoursOverFrom = False
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT [index] AS myid FROM request_T WHERE [index] <> " & rid & " AND intrID = " & intrID & " AND InstID = " & InstID & " AND " & _
		"appdate = '" & adate & "' AND apptimefrom <= '" & timefrom & "' AND apptimeto >= '" & timefrom & "' AND status <> 3 "
	rsApp.Open sqlApp, g_strCONN, 3, 1
	If Not rsApp.EOF Then Z_hoursOverFrom = True
	rsApp.Close
	Set rsApp = Nothing
End Function
Function Z_hoursOverTo(rid, intrid, InstID, adate, timeto)
	' Z_hoursOverTo(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimeto")) Then 
	Z_hoursOverTo = False
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT [index] AS myid FROM request_T WHERE [index] <> " & rid & " AND intrID = " & intrID & " AND InstID = " & InstID & " AND " & _
		"appdate = '" & adate & "' AND apptimefrom <= '" & timeto & "' AND apptimeto>= '" & timeto & "' AND status <> 3 "
	rsApp.Open sqlApp, g_strCONN, 3, 1
	If Not rsApp.EOF Then Z_hoursOverTo = True
	rsApp.Close
	Set rsApp = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
server.scripttimeout = 360000
Function MyStatus(xxx)
	Select Case xxx
		Case 1
			MyStatus = "<font color='#000000' size='+3'>•</font>"
		Case 2
			MyStatus = "<font color='#0000FF' size='+3'>•</font>"
		Case 3
			MyStatus = "<font color='#FF0000' size='+3'>•</font>"
		Case 4
			MyStatus = "<font color='#FF00FF' size='+3'>•</font>"
		Case Else
			MyStatus = ""
	End Select
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
tmpPage = "document.frmTbl."
radioApp = ""
radioID = ""
radioAll = "checked"
radioAss = "checked"
radioUnass = ""
radioUnass2 = ""
radioappall = ""
If Z_fixNull(Request("ctrlX")) = "" Then
	Session("MSG") = "ERROR: Please open page again."
	response.redirect "admin.asp"
End If
If Request("ctrlX") = 1 Then
	mybtn = "Save Timesheet"
Else
	mybtn = "Save Mileage"
End If
x = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST"  Or Request("action") = 3 Then

		sqlReq = "SELECT req.[index]" & _
				", req.[appDate], req.[apptimefrom], req.[apptimeto], req.[astarttime], req.[aendtime]" & _
				", req.[phoneappt], req.[showintr], req.[payintr], req.[overmile], req.[payhrs], req.[overpayhrs]" & _
				", req.InstID AS myInstID, req.[IntrID], req.[LangID], req.[InstRate], req.[SentReq]" & _
				", req.[Processed], req.[Status], req.[DeptID], req.[clname], req.[cfname]" & _
				", req.[actTT], req.[actMil], req.[toll]" & _
				", req.[LBconfirm], req.[LbconfirmToll] " & _
				", ins.[Facility], lan.[language], dep.[dept]" & _
				", CASE WHEN itr.[last name] IS NOT NULL THEN itr.[last name] " & _
				"ELSE 'N/A' " & _
				"END AS [last name] " & _
				", CASE WHEN itr.[first name] IS NOT NULL THEN itr.[first name] " & _
				"ELSE 'N/A' " & _
				"END AS [first name] " & _
				"FROM [request_T] AS req " & _
				"INNER JOIN [institution_T] AS ins	ON req.[InstID]=ins.[index] " & _
				"INNER JOIN [dept_T] AS dep			ON req.[DeptID]=dep.[index] " & _
				"INNER JOIN [language_T] AS lan		ON req.[LangId]=lan.[index] " & _
				"LEFT JOIN [interpreter_T] AS itr	ON req.[IntrId]=itr.[index] " & _
				"WHERE req.[showintr] = 1 " 
		If Request("ctrlX") = 1 Then 			' TIMESHEET'
			If Request("radioAss") = 0 Then	' unapproved '
				sqlReq = sqlReq & "AND req.[lbconfirm] = 0 "
				radioAss = "checked"
				radioUnass = ""
				radioUnass2 = ""
				radioall = ""
				btnSave = ""
			ElseIf Request("radioAss") = 1 Then	
				sqlReq = sqlReq & "AND req.[lbconfirm] = 1 "
				radioAss = ""
				radioUnass = "checked"
				radioUnass2 = ""
				radioappall = "disabled"
				btnSave = "disabled"
			Else
				radioAss = ""
				radioUnass = ""
				radioUnass2 = "checked"
				radioappall = "disabled"
				btnSave = "disabled"
			End If
		Else 									' MILEAGE
			sqlReq = "SELECT req.[index]" & _
					", req.[appDate], req.[apptimefrom], req.[apptimeto], req.[astarttime], req.[aendtime]" & _
					", req.[phoneappt], req.[showintr], req.[payintr], req.[overmile], req.[payhrs], req.[overpayhrs]" & _
					", req.InstID AS myInstID, req.[IntrID], req.[LangID], req.[InstRate], req.[SentReq]" & _
					", req.[Processed], req.[Status], req.[DeptID], req.[clname], req.[cfname]" & _
					", req.[actTT], req.[actMil], req.[toll]" & _
					", req.[LBconfirm], req.[LbconfirmToll] " & _
					", ins.[Facility], lan.[language], dep.[dept]" & _
					", CASE WHEN itr.[last name] IS NOT NULL THEN itr.[last name] " & _
					"ELSE 'N/A' " & _
					"END AS [last name] " & _
					", CASE WHEN itr.[first name] IS NOT NULL THEN itr.[first name] " & _
					"ELSE 'N/A' " & _
					"END AS [first name] " & _
					", gdt.[reqid], gdt.[dstval], gdt.[durval], itr.[index] AS interpretindex " & _
					"FROM [request_T] AS req " & _
					"INNER JOIN [institution_T] AS ins	ON req.[InstID]=ins.[index] " & _
					"INNER JOIN [dept_T] AS dep			ON req.[DeptID]=dep.[index] " & _
					"INNER JOIN [language_T] AS lan		ON req.[LangId]=lan.[index] " & _
					"LEFT JOIN [interpreter_T] AS itr	ON req.[IntrId]=itr.[index] " & _
					"LEFT JOIN [tmpGoogleDist] AS gdt 	ON req.[index]=gdt.[reqid] " & _
					"WHERE req.[showintr] = 1 " 
			AMchkbox = ""
			sqlReq = sqlReq & "AND req.[instID] <> 479 "
			If Request("radioAss") = 0 Then			' Unapproved
				sqlReq = sqlReq & "AND ( req.[LbconfirmToll] = 0 OR req.[LbconfirmToll] IS NULL) "
				radioAss = "checked"
				radioUnass = ""
				radioUnass2 = ""
				btnSave = ""
			ElseIf Request("radioAss") = 1 Then 	' Approved
				sqlReq = sqlReq & "AND req.[LbconfirmToll] = 1 "
				radioAss = ""
				radioUnass = "checked"
				radioUnass2 = ""
				btnSave = "disabled style=""display: none;"""
				AMchkbox = " disabled "
				mybtn = " - - - "
			Else 									' ALL
				radioAss = ""
				radioUnass = ""
				radioUnass2 = "checked"
				btnSave = "disabled style=""display: none;"""
				AMchkbox = " disabled "
				mybtn = " - - - "
			End If
		End If
	'FIND
	If Request("radioStat") = 0 Then
		radioApp = "checked"
		radioID = ""
		radioAll = ""
		If Request("txtFromd8") <> "" Then
			If IsDate(Request("txtFromd8")) Then
				sqlReq = sqlReq & " AND req.[appDate] >= '" & Z_YMDDate(Request("txtFromd8")) & "' "
				tmpFromd8 = Request("txtFromd8") 
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (From)."
				Response.Redirect "reqtable2.asp?ctrlX=" & Request("ctrlX")
			End If
		End If
		If Request("txtTod8") <> "" Then
			If IsDate(Request("txtTod8")) Then
				sqlReq = sqlReq & " AND req.[appDate] <= '" & Z_YMDDate(Request("txtTod8")) & "' "
				tmpTod8 = Request("txtTod8")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (To)."
				Response.Redirect "reqtable2.asp?ctrlX=" & Request("ctrlX")
			End If
		End If
	ElseIf Request("radioStat") = 1 Then
		radioApp = ""
		radioID = "checked"
		radioAll = ""
		If Request("txtFromID") <> "" Then
			If IsNumeric(Request("txtFromID")) Then
				sqlReq = sqlReq & " AND req.[index] >= " & Request("txtFromID") & " "
				tmpFromID = Request("txtFromID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (From)."
				Response.Redirect "reqtable2.asp?ctrlX=" & Request("ctrlX")
			End If
		End If
		If Request("txtToID") <> "" Then
			If IsNumeric(Request("txtToID")) Then
				sqlReq = sqlReq & " AND req.[index] <= " & Request("txtToID") & " "
				tmpToID = Request("txtToID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (To)."
				Response.Redirect "reqtable2.asp?ctrlX=" & Request("ctrlX")
			End If
		End If
	Else
		radioApp = ""
		radioID = ""
		radioAll = "checked"
	End If
	'FILTER
	xInst = Z_CLng(Request("selInst"))
	If xInst <> -1 Then 
		sqlReq = sqlReq & " AND req.[InstID] = " & xInst & " "
	End If
	xLang = Z_CLng(Request("selLang"))
	If xLang <> -1 Then 
		sqlReq = sqlReq & " AND req.[LangID] = " & xLang & " "
	End If
	If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
		If Trim(Request("txtclilname")) <> "" Then
			sqlReq = sqlReq & " AND Upper(req.[Clname]) LIKE '" & CleanMe2(Ucase(Trim(Request("txtclilname")))) & "%' "
		End If
		If Trim(Request("txtclifname")) <> "" Then
			sqlReq = sqlReq & " AND Upper(req.[Cfname]) LIKE '" & CleanMe2(Ucase(Trim(Request("txtclifname")))) & "%' "
		End If
	End If
	xIntr = Cint(Request("selIntr"))
	If xIntr <> -1 Then 
		sqlReq = sqlReq & " AND req.[IntrID] = " & xIntr & " "
	End If
	xClass = Cint(Request("selClass"))
	If xClass <> -1 Then 
		sqlReq = sqlReq & " AND dep.[Class] = " & xClass & " "
	End If
	'ADMIN ONLY
	xAdmin = Z_CZero(Request("selAdmin"))
	If xAdmin = 1 Then
		sqlReq = sqlReq & " AND (req.[Status] = 1) AND req.[Processed] IS NULL "
		meUnBilled = "selected"
	ElseIf xAdmin = 2 Then
		sqlReq = sqlReq & " AND (req.[Status] = 1 OR req.[Status] = 4) AND NOT req.[Processed] IS NULL "
		meBilled = "selected"
	ElseIf xAdmin = 3 Then
		sqlReq = sqlReq & " AND (req.[Status] = 2) "
		meMisded = "selected"
	ElseIf xAdmin = 4 Then
		sqlReq = sqlReq & " AND (req.[Status] = 3) "
		meCanceled = "selected"
	ElseIf xAdmin = 5 Then
		sqlReq = sqlReq & " AND (req.[Status] = 4) "
		meCanceledBill = "selected"
	ElseIf xAdmin = 6 Then
		sqlReq = sqlReq & " AND (req.[Status] = 0) "
		mePending = "selected"
	Else
		sqlReq = sqlReq & "AND req.[status] NOT IN (2,3) "
	End If
	sqlReq = sqlReq & " ORDER BY itr.[last name], itr.[first name], req.[appDate], req.[appTimeFrom], req.[appTimeto], ins.[Facility]"
'End If
'GET REQUESTS
'Response.write "<code>" & vbCrLf & sqlReq & vbCrLf & "</code>"
Set rsReq = Server.CreateObject("ADODB.RecordSet")
rsReq.Open sqlReq, g_strCONN, 3, 1
x = 1
totFPHours = 0
If Not rsReq.EOF Then
	Do Until rsReq.EOF
		kulay = ""
		If Not Z_IsOdd(x) Then kulay = "#FBEEB7"
		If rsReq("phoneappt") Then kulay = "#99ff99" 'Phone Call Appt
		If Z_hoursAfter(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimeto")) Or _
				Z_hoursBefore(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimefrom")) Or _
				Z_hoursSame(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimefrom")) Or _
				Z_hoursOverFrom(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimefrom")) Or _
				Z_hoursOverTo(rsReq("Index"), rsReq("IntrID"), rsReq("myInstID"), rsReq("appdate"), rsReq("apptimeto")) Then 
			b2b = true
			kulay = "#ff9999"
		Else
			b2b = false
		End If
		'GET INSTITUTION
		tmpIname = rsReq("Facility")
		myDept = " - " & rsReq("dept") ' GetMyDept(rsReq("DeptID"))
		'GET INTERPRETER INFO
		tmpInName = rsReq("last name") & ", " & rsReq("first name")
		'GET LANGUAGE
		tmpSalita = rsReq("Language")
		Stat = MyStatus(rsReq("Status") )
		
		TT = Z_FormatNumber(rsReq("actTT"), 2)
		If rsReq("overpayhrs") Then 
			BlnOver = "checked"
			PHrs = Z_FormatNumber(rsReq("payhrs"), 2)
		Else
			BlnOver = ""
			PHrs = Z_FormatNumber(IntrBillHrs(rsReq("AStarttime"), rsReq("AEndtime")), 2)
		End If
		tmpPayHrs = "0:00"
		If Z_FixNull(rsReq("AStarttime")) <> "" And Z_FixNull(rsReq("AEndtime")) <> "" Then
			date1st = rsReq("AStarttime")
			date2nd = rsReq("AEndtime")
			if datediff("n", date1st, date2nd) >= 0 then
				minTime = DateDiff("n", date1st, date2nd)
			else
				minTime = DateDiff("n", date1st, dateadd("d", 1, date2nd))
			end If
			tmpPayHrs = MakeTime(Z_CZero(minTime))
		End If
		FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
		tmpAMT = Z_FormatNumber(rsReq("actMil"), 2)

		If rsReq("overmile") Then
			BlnOver2 = "checked"
		Else
			BlnOver2 = ""
		End If
		If rsReq("LBconfirm") = True Then 
			LBcon = "Checked disabled"
			LBconx = "readonly"
			LBconxx = "disabled"
			LBconxxx = "readonly"
		Else
			LBcon = ""
			LBconx = ""
			LBconxx = ""
			LBconxxx = ""
		End If
		If rsReq("LBconfirmToll") = True Then 
			LBcon2 = "checked disabled"
			LBconx2 = "readonly"
			LBconxx2 = "disabled"
		Else
			LBcon2 = ""
			LBconx2 = ""
			LBconxx2 = ""
		End If
		showintr = ""
		showintr2 = ""
		If rsReq("showintr") Then 
			showintr = "checked"
			showintr2 = "X"
		End If
		payintr = ""
		if rsReq("payintr") Then payintr = "Checked"
		strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
				"<td class='tblgrn2' width='10px'>" & Stat & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsReq("Index") & "'><a class='link2' href='reqconfirm.asp?ID=" & rsReq("Index") & "'><b>" & rsReq("Index") & "</b></a></td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & tmpIname & myDept
		If (b2b) Then strtbl = strtbl & " <b>(b2b)</b>"
		strtbl = strtbl & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("clname") & ", " & rsReq("cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpInName & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & Z_FormatTime(rsReq("apptimefrom")) & " - " & Z_FormatTime(rsReq("apptimeto")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2' valign=""middle""><input class='main2 edmil' name='txtstime" & x & "' maxlength='5' size='7' " & LBconx & " value='" & Z_FormatTime(rsReq("astarttime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf & _
				"<td class='tblgrn2' valign=""middle""><input class='main2 edmil' name='txtetime" & x & "' maxlength='5' size='7' " & LBconx & " value='" & Z_FormatTime(rsReq("aendtime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf 
				' ENDS AT THE "Actual End Time" COLUMN
		If Request("ctrlX") = 1 Then
			kulay2 = "#080F0D"
			If Z_CZero(Phrs) > 6 Then kulay2 = "#FF0000"
			strtbl = strtbl & "<td class='tblgrn2' >" & tmpPayHrs & "</td><td class='tblgrn2' ><nobr><input class='main2' name='txtTT" & x & "' maxlength='6' size='7' " & LBconxxx & " value='" & TT & "'></td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr><input class='main2' name='txtPhrs" & x & "' maxlength='6' size='7' " & LBconx & " style=""color:" & kulay2 & ";"" value='" & PHrs & "'>" & vbCrLf
			If LBconxx = "" Then
				strtbl = strtbl & "<input type='checkbox' name='chkOverPhrs" & x & "' value='1' " & LBconxx & " " & BlnOver & " ></td>" & vbCrLf
			Else
				strtbl = strtbl & "<input type='checkbox' value='1' " & LBconxx & " " & BlnOver & " ><input type='hidden' name='chkOverPhrs" & x & "' value='1' " & BlnOver & "></td>" & vbCrLf
			End If
			strtbl = strtbl & "<td class='tblgrn2' >" & FPHrs & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='checkbox' ID='chkshow" & x & "' name='chkshow" & x & "' value='1' " & showintr & "></td>" & vbCrLf
			if PHrs > 0 Then
				strtbl = strtbl & "<td class='tblgrn2' ><input type='checkbox' ID='chkTS" & x & "' name='chkTS" & x & "' value='1' " & LBcon & "></td>" & vbCrLf
			Else
				strtbl = strtbl & "<td class='tblgrn2' ><input type='checkbox' ID='chkTS" & x & "' name='chkTS" & x & "' value='1' disabled ></td>" & vbCrLf
			End If
			csvHEAD = "ID,Institution,Department,Language,Client Last Name, Client First Name, Interpreter,Date," & _
				"Actual Time (from), Actual Time (to), Total Time, Travel Time, Payable Hours, Final Payable Hours, Show to Interpreter"
			csvTBL = csvTBL & """" & rsReq("Index") & """,""" & tmpIname & """,""" & myDept & """,""" & tmpSalita & """,""" & rsReq("clname") & """,""" & rsReq("cfname") & _
				""",""" & tmpInName & """,""" & rsReq("appDate") & """,""" & Z_FormatTime(rsReq("astarttime")) & """,""" & Z_FormatTime(rsReq("aendtime")) & """,""" & Z_FormatNumber(tmpPayHrs, 2) & """,""" & Z_FormatNumber(TT, 2) & _
				""",""" & Z_FormatNumber(PHrs, 2) & """,""" & Z_FormatNumber(FPHrs, 2) & """,""" & showintr2 & """" & vbCrLf
		ElseIf Request("ctrlX") = 2 Then
			strtbl = strtbl & "<td class='tblgrn2' ><nobr><input class='main2 edmil' name='txtmile" & x & "' maxlength='6' size='7' " & LBconx2 & " value='" & tmpAMT & "' />"
			If AMchkbox = "" Then
				If rsReq("LBconfirmToll") <> True Then 
					strtbl = strtbl & "<input type=""checkbox"" name=""chkOverMile" & x & """ value=""1"" " & BlnOver2 & " />" & vbCrLf 
				End If
			Else
				'If rsReq
				'strtbl = strtbl & "<input type=""hidden"" name=""chkOverMile" & x & """ value='1' " & LBconxx2 & " " & />" & vbCrLf 
				If BlnOver2 = "checked" Then
					strtbl = strtbl & "<img src=""images/ok.gif"" title=""YES"" alt=""X"" />" & _
							"<input type=""hidden"" ID=""chkOverMile" & x & """ name=""chkOverMile" & x & """ value=""1"" />"
				Else
					strtbl = strtbl & "<img src=""images/nok.gif"" title=""NO"" alt=""O"" />" & _
							"<input type=""hidden"" ID=""chkOverMile" & x & """ name=""chkOverMile" & x & """ value="""" />"
				End If
			End If
			strtbl = strtbl & "</nobr>"
			' optional automated mileage information...
			If Z_FixNull( rsReq("reqid") ) <> "" Then
				' mileage was fetched; display it here: dstval * 2, durval / 30
				tmpMil = Z_CDbl(rsReq("dstval")) * 2
				If (tmpMil > 40) Then
					tmpBilMil = tmpMil - 40
					strTbl = strTbl & "<div class=""ggl tooltip"">" & Round(tmpBilMil, 2) & " mi"
				Else
					strTbl = strTbl & "<div class=""ggl tooltip"">&lt; 40 mi"
				End If
				tmpTT = Z_CDbl(rsReq("durval")) / 30
				strTbl = strTbl & "<span class=""tooltiptext"">Fetched: " & Round(tmpMil, 2) & " mi (r/t);<br />" & Round(tmpTT, 2) & " hrs travel</span></div>"
			Else
				If Z_CDbl(rsReq("actMil")) > 0 Then
					If Z_FixNull( rsReq("interpretindex") ) <> "" Then
						Set oGDM = New acaDistanceMatrix
						oGDM.DBCONN = g_strCONN
						Call oGDM.FetchMileageFromReqID(rsReq("index"), TRUE)
						fltRealTT	= oGDM.fltRealTT
						fltRealM	= oGDM.fltRealM
						fltActTT	= oGDM.fltActTT
						If oGDM.fltRealM > 40 Then
							strTbl = strTbl & "<div class=""ggl tooltip"">" & Round(oGDM.fltActMil, 2) & " mi"
						Else
							strTbl = strTbl & "<div class=""ggl tooltip"">&lt; 40 mi"
						End If
						strTbl = strTbl & "<span class=""tooltiptext"">Fetched (now)<br />" & Round(oGDM.fltRealM, 2) & " mi (r/t);<br />" & _
								Round(oGDM.fltRealTT, 2) & " hrs travel</span></div>"
					End If
				End If
			End If
			' TOOLS & PARKING COLUMN'
			strtbl = strtbl & "</td><td class='tblgrn2' ><nobr>$<input class='main2 edmil' name='txtTol" & x & "' maxlength='5' size='7' " & _
					LBconx2 & " value='" & Z_FormatNumber(rsReq("toll"), 2) & "' /></td>" & vbCrLf
			' Approve Mileage checkbox
			If AMchkbox = "" Then
				strtbl = strtbl & "<td class='tblgrn2' ><input type='checkbox' ID='chkM" & x & _
						"' name='chkM" & x & "' value='1' " & LBcon2  & "/></td></tr>" & vbCrLf
			Else
				strtbl = strtbl & "<td class='tblgrn2' >"
				If rsReq("LBconfirmToll") = True Then 
					strtbl = strtbl & "<img src=""images/ok.gif"" title=""YES"" alt=""X"" />" & _
							"<input type=""hidden"" ID=""chkM" & x & """ name=""chkM" & x & """ value=""1"" />"
				Else
					strtbl = strtbl & "<img src=""images/nok.gif"" title=""NO"" alt=""O"" />" & _
							"<input type=""hidden"" ID=""chkM" & x & """ name=""chkM" & x & """ value="""" />"
				End If
				strtbl = strtbl & "</td></tr>" & vbCrLf
			End If
			csvHEAD = "ID,Institution,Department,Language,Client Last Name, Client First Name, Interpreter,Date," & _
				"Actual Time (from), Actual Time (to),Mileage,Tolls and Parking"
			csvTBL = csvTBL & """" & rsReq("Index") & """,""" & tmpIname & """,""" & myDept & """,""" & tmpSalita & """,""" & rsReq("clname") & """,""" & rsReq("cfname") & _
				""",""" & tmpInName & """,""" & rsReq("appDate") & """,""" & Z_FormatTime(rsReq("astarttime")) & """,""" & Z_FormatTime(rsReq("aendtime")) & """,""" & Z_FormatNumber(tmpAMT, 2) & """,""" & Z_FormatNumber(rsReq("toll"), 2) & """" & vbCrLf
		End If
		RepCSV =  "myHours.csv"
		if Request("ctrlX") = 2 Then RepCSV =  "myMileage.csv"
		'CONVERT TO CSV
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
		Prt.WriteLine "LANGUAGE BANK - REPORT"
		Prt.WriteLine csvHEAD
		Prt.WriteLine csvTBL
		Prt.Close	
		Set Prt = Nothing
		fso.CopyFile RepPath & RepCSV, BackupStr
		Set fso = Nothing
		tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)
		strtbl = strtbl & "</tr>" & vbCrLf
		x = x + 1
		totFPHours = totFPHours + FPHrs
		rsReq.MoveNext
	Loop
Else
	strtbl = "<tr><td colspan='14' align='center'><i>&lt -- No records found. -- &gt</i></td></tr>"
End If
rsReq.Close
Set rsReq = Nothing
End If
'SORT
If Request("sType") <> "" Then
	If Request("stype") = 1 Then stype = 2
	If Request("stype") = 2 Then stype = 1
Else
	stype = 1
End If
'FILTER CRITERIA
tmpclilname = Request("txtclilname")
tmpclifname = Request("txtclifname")
'GET INSTITUTION LIST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT Facility, [Index] FROM institution_T ORDER BY [Facility]"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	InstSel = ""
	If Cint(Request("selInst")) = rsInst("Index") Then InstSel = "selected"
	InstName = rsInst("Facility")
	strInst = strInst	& "<option value='" & rsInst("Index") & "' " & InstSel & ">" &  InstName & "</option>" & vbCrlf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT [Index], [language] FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	LangSel = ""
	If Cint(Request("selLang")) = rsLang("Index") Then LangSel = "selected"
	strLang = strLang	& "<option value='" & rsLang("Index") & "' " & LangSel & ">" &  rsLang("language") & "</option>" & vbCrlf
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'GET INTERPRETER LIST
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT [Index], [last name], [first name] FROM interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	IntrSel = ""
	If Cint(Request("selIntr")) = rsIntr("Index") Then IntrSel = "selected"
	strIntr = strIntr	& "<option value='" & rsIntr("Index") & "' " & IntrSel & ">" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbCrlf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then 
	
End If
'FOR CLASSIFICATION
tmpClass = Cint(Request("selClass"))
Select Case tmpClass
	Case 1 SocSer = "selected"
    Case 2 Priv = "selected"
	Case 3 Legal = "selected"	
	Case 4 Med = "selected"
End Select
'FOR ADMIN
tmpAdmin = Z_CZero(Request("selAdmin"))
Select Case tmpAdmin
	Case 1 meUnBilled = "selected"
    Case 2 meBilled = "selected"
	Case 3 meMisded = "selected"
	Case 4 meCanceled = "selected"
	Case 5 meCanceledBill = "selected"
End Select
%>
<html>
	<head>
		<title>Language Bank - Timesheet/Mileage</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
<style>
td.tblgrn2 { font-family: Trebuchet MS, Trebuchet, Tahoma, Verdana, Arial, Helvetica, Sans-Serif; text-align: center; vertical-align: middle; }
input.edmil { border-color: #01a1af; padding: 1px 3px; margin: 4px 3px 1px 3px; }
.ggl { color: #EC7063; font-style: italic; }
.tooltip {  position: relative;  display: inline-block;  border-bottom: 1px dotted black; /* If you want dots under the hoverable text */ }
/* Tooltip text */
.tooltip .tooltiptext {  visibility: hidden; width: 120px; background-color: black; color: #fff; text-align: center; padding: 5px 0; border-radius: 6px; 
  /* Position the tooltip text - see examples below! */
  position: absolute; z-index: 1;
}
/* Show the tooltip text when you mouse over the tooltip container */
.tooltip:hover .tooltiptext { visibility: visible; }
</style>		
		<script language='JavaScript'>
		<!--
		function SaveMe()
		{
			var ans = window.confirm("This action will save all entries inside the table to the database. Please double check your enties.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmTbl.action = "action.asp?ctrl=17";
				document.frmTbl.submit();
			}
		}
		function SortMe(sortnum)
		{
			document.frmTbl.action = "reqtable2.asp?sort=" + sortnum + "&sType=" + <%=stype%>;
			document.frmTbl.submit();
		}
		function FindMe(xxx)
		{
			document.frmTbl.action = "reqtable2.asp?ctrlX=" + xxx;
			document.frmTbl.submit();
		}
		function FixSort()
		{
			document.frmTbl.txtFromd8.disabled = true;
			document.frmTbl.txtTod8.disabled = true;
			document.frmTbl.txtFromID.disabled = true;
			document.frmTbl.txtToID.disabled = true;
			if (document.frmTbl.radioStat[0].checked == true)
			{
				document.frmTbl.txtFromd8.disabled = false;
				document.frmTbl.txtTod8.disabled = false;
			}
			if (document.frmTbl.radioStat[1].checked == true)
			{
				document.frmTbl.txtFromID.disabled = false;
				document.frmTbl.txtToID.disabled = false;
			}
		}
		function CalendarView(strDate)
		{
			document.frmTbl.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmTbl.submit();
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function checkme(xxx)
		{
			var tmpElem;
			var z;
			if (document.frmTbl.chkall.checked == true)
			{
				for(z = 1; z <= xxx; z ++)
				{
					<% If Request("ctrlX") = 1 Then %>
						tmpElem = "chkTS" + z;
					<% Else %>
						tmpElem = "chkM" + z;
					<% End If %>
					if (document.getElementById(tmpElem) != null ) {
						if (document.getElementById(tmpElem).disabled == false) {
							document.getElementById(tmpElem).checked = true;
						}
					}
				}	
			}
			else
			{
				for(z = 1; z <= xxx; z ++)
				{
					<% If Request("ctrlX") = 1 Then %>
						tmpElem = "chkTS" + z;
					<% Else %>
						tmpElem = "chkM" + z;
					<% End If %>
					if (document.getElementById(tmpElem) != null) {
						document.getElementById(tmpElem).checked = false;
					}
				}	
			}
		}
		function exportme(){
			
		}
		-->
		</script>
		<style type="text/css">
	 	.container
	      {
	          border: solid 1px black;
	          overflow: auto;
	      }
	      .noscroll
	      {
	          position: relative;
	          background-color: white;
	          top:expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
		<body onload='FixSort();'>
			<form method='post' name='frmTbl' action='reqtable2.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top'>
							<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
								<!-- #include file="_greetme.asp" -->
								<tr>
									<td>
										<table cellpadding='0' cellspacing='0' width='100%' border='0'>
											<tr>
												<td align='left'>
													Legend: <font color='#000000' size='+3'>•</font>&nbsp;-&nbsp;completed&nbsp;<font color='#0000FF' size='+3'>•</font>&nbsp;-&nbsp;missed&nbsp;<font color='#FF0000 ' size='+3'>•</font>&nbsp;-&nbsp;Canceled&nbsp;
													<font color='#FF00FF' size='+3'>•</font>&nbsp;-&nbsp;Canceled (billable)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<% If Cint(Request.Cookies("LBUSERTYPE")) = 1 Or Cint(Request.Cookies("LBUSERTYPE")) = 3 Or Cint(Request.Cookies("LBUSERTYPE")) = 0 Then %>
														Admin Sort:
														<select class='seltxt' style='width:100px;' name='selAdmin'>
															<option value='0'>&nbsp;</option>
															<option <%=mePending%> value='6'>Pending</option>
															<option <%=meUnBIlled%> value='1'>Completed (Unbilled)</option>
															<option <%=meCanceledBill%> value='5'>Canceled (Billable)</option>
															<option <%=meBilled%> value='2'>BILLED</option>
														</select>
														<input class='btntbl' type='button' value='GO' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe(<%=Request("ctrlX")%>);'>
													<% End If %>
												</td>
												<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
													<% If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then %> 
														<td align='right'>
															<input type='hidden' name='Hctr' value='<%=x%>'>
															<input class='btntbl' type='button' value='Save Table' style='height: 25px; width: 200px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveMe();'>
														</td>
													<% Else %>
													<td align='right'>
															<input type='hidden' name='ctrlX' value='<%=Request("ctrlX")%>'>
															<input type='hidden' name='Hctr' value='<%=x%>'>
															<input class='btntbl' type='button' value='Export Table' style='height: 25px; width: 100px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick="document.location='<%=tmpstring%>';">
<% ' If Len(mybtn) > 7 Then %>
															<input class="btntbl" type="button" value="<%=mybtn%>" style="height: 25px; width: 100px;"
																	onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'"
																	onclick="SaveMe();" <%=btnSave%> />
<% ' End IF %>															
														</td>
													<% End If %>
												<% Else %>
													<td>&nbsp;</td>
												<% End If %>
											</tr>
										</table>
									</td>
								</tr>
								<% If Session("MSG") <> "" Then %>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='14' align='left'>
											<div name="dErr" style="width:300px; height:40px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
								<% End If %>
								<tr>
									<td colspan='11' align='left'>
										<div class='container' style='height: 500px; width:100%; position: relative;'>
											<table class="reqtble" width='100%'>	
												<thead>
													<tr class="noscroll">	
														<td colspan='2' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" class='tblgrn'>Request ID</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Institution</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Language</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Client</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Interpreter</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Appointment Date</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Requested Time</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual Start Time</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual End Time</td>
														<% If Request("ctrlX") = 1 Then %>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Total Time</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Travel Time</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Payable Hours</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Final Payable Hours</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Show to Interpreter</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">
																Approve Timesheet<br>
																<input type='checkbox' name='chkall' onclick='checkme(<%=x%>);' <%=radioappall%>>
															</td>
														<% Else %>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Mileage</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Tolls & parking</td>
															<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">
																Approve Mileage<br>
																<input type='checkbox' name='chkall' onclick='checkme(<%=x%>);'>
															</td>
														<% End If %>
													</tr>
												</thead>
												<tbody style="OVERFLOW: auto;">
													<%=strtbl%>
												</tbody>
											</table>
										</div>	
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table width='100%'  border='0'>
								<tr>
									<td align='left'>
										&nbsp;
									</td>
									<td align='right'>
										<% If x <> 0 Then %>
											<b><u><%=x - 1%></u></b> Records &nbsp;&nbsp;|&nbsp;&nbsp;
										<% End If %>
										<% If totFPHours > 0 Then %>
											<b><u><%=totFPHours%></u></b> Total Hours
										<% End If %>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='82%' border='0' style='border: solid 1px;'>
								<tr bgcolor='#FBEEB7'>
									<td align='right' style='border-bottom: solid 1px;'><b>Sort:</b></td>
									<td style='border-right: solid 1px;border-bottom: solid 1px;'>
										<input type='radio' name='radioStat' value='0' <%=radioApp%> onclick='FixSort();'>&nbsp;<b>App. Date Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='10' maxlength='10' name='txtFromd8' value='<%=tmpFromd8%>'>
										&nbsp;-&nbsp;
										<input class='main' size='10' maxlength='10' name='txtTod8' value='<%=tmpTod8%>'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='1' <%=radioID%> onclick='FixSort();'>&nbsp;<b>Request ID Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='7' maxlength='7' name='txtFromID' value='<%=tmpFromID%>'>
										&nbsp;-&nbsp;
										<input class='main' size='7' maxlength='7' name='txtToID' value='<%=tmpToID%>'>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='2' <%=radioAll%> onclick='FixSort();'>&nbsp;<b>All</b>
									</td>
									<td align='right' style='border-bottom: solid 1px;'><b>&nbsp;&nbsp;</b></td>
									<td style='border-bottom: solid 1px;'>
										<input type='radio' name='radioAss' value='0' <%=radioAss%> onclick='FixSort();'>&nbsp;<b>Unapproved</b>
										&nbsp;&nbsp;
										<input type='radio' name='radioAss' value='1' <%=radioUnAss%> onclick='FixSort();'>&nbsp;<b>Approved</b>
										&nbsp;&nbsp;
										<input type='radio' name='radioAss' value='2' <%=radioUnAss2%> onclick='FixSort();'>&nbsp;<b>ALL</b>
										&nbsp;&nbsp;
									</td>
									<td align='right' style='border-left: solid 1px;' rowspan='3'>
										<input class='btntbl' type='button' value='Find' style='height: 35px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe(<%=Request("ctrlX")%>);'>
									</td>
									</td>
								</tr>
								<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='4'>
										Institution:
										<select class='seltxt' style='width: 285px;' name='selInst'>
											<option value='-1'>&nbsp;</option>
											<%=strInst%>
										</select>
										&nbsp;Language:
										<select class='seltxt' style='width: 150px;' name='selLang'>
											<option value='-1'>&nbsp;</option>
											<%=strLang%>
										</select>
										<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
											&nbsp;Client:
											<input class='main' size='20' maxlength='20' name='txtclilname' value="<%=tmpclilname%>">
											&nbsp;,&nbsp;&nbsp;
											<input class='main' size='20' maxlength='20' name='txtclifname' value="<%=tmpclifname%>">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">Last name, First name</span>
										<% End If %>
										
										&nbsp;
									</td>
								</tr>
								<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='4'>
										Interpreter:
										<select class='seltxt' name='selIntr'>
											<option value='-1'>&nbsp;</option>
											<%=strIntr%>
										</select>
										&nbsp;Classification:
										<select class='seltxt' style='width: 100px;' name='selClass'>
											<option value='-1'>&nbsp;</option>
											<option value='1' <%=SocSer%>>Social Services</option>
											<option value='2' <%=Priv%>>Private</option>
											<option value='3' <%=Legal%>>Legal</option>
											<option value='4' <%=Med%>>Medical</option>
										</select>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td height='50px' valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
<!-- code><%= sqlReq %></code -->
		</body>
	</head>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
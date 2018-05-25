<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function CleanMe(xxx)
	' clean string
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", "''")
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
		GetMyDept = rsDept("Dept")
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
x = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST"  Or Request("action") = 3 Then
	If Request("radioAss") = 0 Then
		sqlReq = "SELECT r.InstID, IntrID, LangID, InstRate, SentReq, Processed, Status, DeptID, r.[index]," & _
			" clname, cfname, appDate, appTimeFrom, appTimeTo, astarttime, aendtime, Billable, IntrRate, comment, dept, hpid " & _
			"FROM request_T AS r " & _
			"INNER JOIN institution_T AS i ON r.InstID = i.[index] " & _
			"INNER JOIN language_T AS l ON r.LangId = l.[index] " & _
			"INNER JOIN interpreter_T AS t ON r.IntrId = t.[index] " & _
			"INNER JOIN dept_T AS d ON r.DeptId = d.[index] " & _
			"WHERE 1=1 "
		radioAss = "checked"
		radioUnass = ""
	Else
		sqlReq = "SELECT r.InstID, IntrID, LangID, InstRate, SentReq, Processed, Status, DeptID, r.[index]," & _
			" clname, cfname, appDate, appTimeFrom, appTimeTo, astarttime, aendtime, Billable, IntrRate, comment, dept, hpid " & _
			"FROM request_T AS r " & _
			"INNER JOIN institution_T AS i ON r.InstID = i.[index] " & _
			"INNER JOIN language_T AS l ON r.LangId = l.[index] " & _
			"INNER JOIN dept_T AS d ON r.DeptId = d.[index] " & _
			"WHERE (IntrId = -1 OR IntrId = 0) "
		radioAss = ""
		radioUnass = "checked"
	End If
	'FIND
	If Request("radioStat") = 0 Then
		radioApp = "checked"
		radioID = ""
		radioAll = ""
		radioHPID = ""
		If Request("txtFromd8") <> "" Then
			If IsDate(Request("txtFromd8")) Then
				sqlReq = sqlReq & " AND appDate >= '" & Request("txtFromd8") & "' "
				tmpFromd8 = Request("txtFromd8") 
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (From)."
				Response.Redirect "reqtable.asp"
			End If
		End If
		If Request("txtTod8") <> "" Then
			If IsDate(Request("txtTod8")) Then
				sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
				tmpTod8 = Request("txtTod8")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (To)."
				Response.Redirect "reqtable.asp"
			End If
		End If
	ElseIf Request("radioStat") = 1 Then
		radioApp = ""
		radioID = "checked"
		radioAll = ""
		radioHPID = ""
		If Request("txtFromID") <> "" Then
			If IsNumeric(Request("txtFromID")) Then
				sqlReq = sqlReq & " AND r.[index] >= " & Request("txtFromID")
				tmpFromID = Request("txtFromID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (From)."
				Response.Redirect "reqtable.asp"
			End If
		Else
			If Request("txtToID") <> "" Then
				Session("MSG") = "ERROR: Invalid Appointment ID Range (From)."
				Response.Redirect "reqtable.asp"
			End If
		End If
		If Request("txtToID") <> "" Then
			If IsNumeric(Request("txtToID")) Then
				sqlReq = sqlReq & " AND r.[index] <= " & Request("txtToID")
				tmpToID = Request("txtToID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (To)."
				Response.Redirect "reqtable.asp"
			End If
		Else
			If Request("txtFromID") <> "" Then
				Session("MSG") = "ERROR: Invalid Appointment ID Range (To)."
				Response.Redirect "reqtable.asp"
			End If
		End If
	ElseIf Request("radiostat") = 3 Then
		radioApp = ""
		radioID = ""
		radioAll = ""
		radioHPID = "checked"
		If Request("txtFromHPID") <> "" Then
			If IsNumeric(Request("txtFromHPID")) Then
				sqlReq = sqlReq & " AND hpid >= " & Request("txtFromHPID")
				tmpFromHPID = Request("txtFromHPID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Vendor Site ID Range (From)."
				Response.Redirect "reqtable.asp"
			End If
		End If
		If Request("txtToHPID") <> "" Then
			If IsNumeric(Request("txtToHPID")) Then
				sqlReq = sqlReq & " AND HPID <= " & Request("txtToHPID")
				tmpToHPID = Request("txtToHPID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Vendor Site ID Range (To)."
				Response.Redirect "reqtable.asp"
			End If
		End If
	Else
		radioApp = ""
		radioID = ""
		radioAll = "checked"
		radioHPID = ""
	End If
	'FILTER
	xInst = Cint(Request("selInst"))
	If xInst > 0 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "r.InstID = " & xInst
	End If
	xDept = Cint(Request("selDept"))
	If xDept > 0 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "r.DeptID = " & xDept
	End If
	xLang = Cint(Request("selLang"))
	If xLang > 0 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "LangID = " & xLang
	End If
	If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
			If Trim(Request("txtclilname")) <> "" Then
				sqlReq = sqlReq & " AND upper(Clname) LIKE '" & Ucase(Trim(CleanMe(Request("txtclilname")))) & "%'"
			End If
			If Trim(Request("txtclifname")) <> "" Then
				sqlReq = sqlReq & " AND upper(Cfname) LIKE '" & Ucase(Trim(CleanMe(Request("txtclifname")))) & "%'"
			End If

	End If
	xIntr = Cint(Request("selIntr"))
	If xIntr > 0 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "IntrID = " & xIntr
	End If
	xClass = Cint(Request("selClass"))
	If xClass <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "Class = " & xClass
	End If
	'ADMIN ONLY
	xAdmin = Z_CZero(Request("selAdmin"))
	If xAdmin = 1 Then
		sqlReq = sqlReq & " AND (Status = 1) AND Processed IS NULL"
		meUnBilled = "selected"
	ElseIf xAdmin = 2 Then
		sqlReq = sqlReq & " AND (Status = 1 OR Status = 4) AND NOT (Processed IS NULL)"
		meBilled = "selected"
	ElseIf xAdmin = 3 Then
		sqlReq = sqlReq & " AND (Status = 2)"
		meMisded = "selected"
	ElseIf xAdmin = 4 Then
		sqlReq = sqlReq & " AND (Status = 3)"
		meCanceled = "selected"
	ElseIf xAdmin = 5 Then
		sqlReq = sqlReq & " AND (Status = 4)"
		meCanceledBill = "selected"
	ElseIf xAdmin = 6 Then
		sqlReq = sqlReq & " AND (Status = 0)"
		mePending = "selected"
	Else
		sqlReq = sqlReq & " AND Processed IS NULL"
	End If
	'SORT
	If Request("sort") <> "" Then
		If Request("sort") = 1 Then sqlReq = sqlReq & " ORDER BY r.[index]"
		If Request("sort") = 2 Then sqlReq = sqlReq & " ORDER BY Facility"
		If Request("sort") = 3 Then sqlReq = sqlReq & " ORDER BY [Language]"
		If Request("sort") = 4 Then sqlReq = sqlReq & " ORDER BY Clname"
		If Request("sort") = 5 And Request("radioAss") = 0 Then sqlReq = sqlReq & " ORDER BY [Last Name]"
		If Request("sort") = 6 Then sqlReq = sqlReq & " ORDER BY appDate"
		If Request("sort") = 7 Then sqlReq = sqlReq & " ORDER BY appTimeFrom"
		If Request("sort") = 8 Then sqlReq = sqlReq & " ORDER BY AStarttime"
		If Request("sort") = 9 Then sqlReq = sqlReq & " ORDER BY AEndtime"
		If Request("sort") = 10 Then sqlReq = sqlReq & " ORDER BY Billable"
		If Request("sort") = 11 Then sqlReq = sqlReq & " ORDER BY r.InstRate"
		If Request("sort") = 14 Then sqlReq = sqlReq & " ORDER BY r.IntrRate"
		If Request("sort") = 12 Then sqlReq = sqlReq & " ORDER BY SentReq"
		If Request("sort") = 13 Then sqlReq = sqlReq & " ORDER BY Paid"
		If Request("sort") = 15 Then sqlReq = sqlReq & " ORDER BY dept"
		If Request("sort") = 5 And Request("radioAss") = 1 Then
			
		Else
			If Request("style") = 1 Then sqlReq = sqlReq & " DESC"
			If Request("style") = 2 Then sqlReq = sqlReq & " ASC"
		End If
		'FIX SORT
		
		If Request("sort") = 4 Then sqlReq = sqlReq & ", Cfname ASC"
		If Request("sort") = 5 And Request("radioAss") = 0 Then sqlReq = sqlReq & ", [First Name] ASC"
	Else
		sqlReq = sqlReq & " ORDER BY Facility"
	End If	

'GET REQUESTS
Server.ScriptTimeout=300
Set rsReq = Server.CreateObject("ADODB.RecordSet")
'Response.Write "<code>" & sqlReq & "</code>"
rsReq.Open sqlReq, g_strCONN, 3, 1
x = 1
If Not rsReq.EOF Then
	Do Until rsReq.EOF
		kulay = ""
		If Not Z_IsOdd(x) Then kulay = "#FBEEB7"
		'GET INSTITUTION
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & rsReq("InstID")
		rsInst.Open sqlInst, g_strCONN, 3, 1
		If Not rsInst.EOF Then
			tmpIname = rsInst("Facility")  
			
		Else
			tmpIname = "N/A"
		End If
		rsInst.Close
		Set rsInst = Nothing 
		'GET INTERPRETER INFO
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT [last name], [first name] FROM interpreter_T WHERE [index] = " & rsReq("IntrID")
		rsIntr.Open sqlIntr, g_strCONN, 3, 1
		If Not rsIntr.EOF Then
			tmpInName = rsIntr("last name") & ", " & rsIntr("first name")
			tmpInName2 = rsIntr("last name") & """,""" & rsIntr("first name")
		Else
			tmpInName = "N/A"
			tmpInName2 = "N/A"
		End If
		rsIntr.Close
		Set rsIntr = Nothing
		'GET LANGUAGE
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang  = "SELECT [language] FROM language_T WHERE [index] = " & rsReq("LangID")
		rsLang.Open sqlLang , g_strCONN, 3, 1
		If Not rsLang.EOF Then
			tmpSalita = rsLang("language") 
		Else
			tmpSalita = "N/A"
		End If
		rsLang.Close
		Set rsLang = Nothing 
		tmpVer = ""
		If rsReq("SentReq") <> "" Then tmpVer = "checked"
		tmpPaid = ""
		BilledAko = ""
		If rsReq("Processed") <> "" Then 
			tmpPaid = "checked"
			BilledAko = "readonly"
		End If
		Stat = MyStatus(rsReq("Status") )
		myDept =  GetMyDept(rsReq("DeptID"))
		'FOR INST RATE
		Set rsRateko = Server.CreateObject("ADODB.RecordSet")
		sqlRateko = "SELECT Rate FROM Rate_T ORDER BY Rate"
		rsRateko.Open sqlRateko, g_strCONN, 1, 3
		strRate = ""
		Do Until rsRateko.EOF
			RateKo = ""
			If rsReq("InstRate") = rsRateKo("Rate") Then Rateko = "selected"
			strRate = strRate & "<option " & Rateko & " value='" & rsRateko("Rate") & "'>$" & Z_FormatNumber(rsRateko("Rate"), 2) & "</option>" & vbCrLf
			rsRateko.MoveNext
		Loop
		rsRateko.Close
		Set rsRateko = Nothing
		tmpRate = rsReq("InstRate")
		If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then 
			strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
				"<td class='tblgrn2' width='10px'>" & Stat & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsReq("Index") & "'><a class='link2' href='reqconfirm.asp?ID=" & rsReq("Index") & "'><b>" & rsReq("Index") & "</b></a></td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & tmpIname & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & myDept & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("clname") & ", " & rsReq("cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpInName & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & Z_FormatTime(rsReq("appTimeFrom")) & " - " & Z_FormatTime(rsReq("appTimeTo")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' " & BilledAko & " name='txtstime" & x & "' maxlength='5' size='7' value='" & Z_FormatTime(rsReq("astarttime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' " & BilledAko & " name='txtetime" & x & "' maxlength='5' size='7' value='" & Z_FormatTime(rsReq("aendtime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' " & BilledAko & " name='txtBilHrs" & x & "' maxlength='5' size='5' value='" & rsReq("Billable") & "'></td>" & vbCrLf 
			If Cint(Request.Cookies("LBUSERTYPE")) = 1 Then
				strtbl = strtbl & "<td class='tblgrn2' ><select class='seltxt' style='width: 70px;' name='selInstRate" & x & "'><option value='0' >&nbsp;</option>" & strRate & "</select></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' readonly name='txtIntrRate" & x & "' maxlength='7' size='9' value='$" & Z_FormatNumber(rsReq("IntrRate"), 2) & "'></td>" & vbCrLf
			Else
				strtbl = strtbl & "<td class='tblgrn2' >N/A<input type='hidden' name='selInstRate" & x & "' value='" & rsReq("InstRate") & "'></td><td class='tblgrn2' >N/A<input type='hidden' name='txtIntrRate" & x & "' value='" & rsReq("IntrRate") & "'></td>" & vbCrLf
			End If
				strtbl = strtbl & "<td class='tblgrn2' ><input type='checkbox' disabled name='chkver" & x & "' " & tmpVer & "></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='checkbox' disabled name='chkbil" & x & "' " & tmpPaid & "></td>" & vbCrLf & _
				"<td class='tblgrn2' ><textarea class='main' " & BilledAko & " name='txtcom" & x & "'>" & rsReq("comment") & "</textarea></td></tr>" & vbCrLf
			If Cint(Request.Cookies("LBUSERTYPE")) = 1 Then
				csvHEAD = "ID,Institution,Department,Language,Client Last Name, Client First Name, Interpreter First Name, Interpreter Last Name,Date," & _
					"Time (from),Time (to),Actual Time (from), Actual Time (to),Institution Rate,Interpreter Rate,Comment"
				csvTBL = csvTBL & """" & rsReq("Index") & """,""" & tmpIname & """,""" & myDept & """,""" & tmpSalita & """,""" & rsReq("clname") & """,""" & rsReq("cfname") & _
					""",""" & tmpInName2 & """,""" & rsReq("appDate") & """,""" & Z_FormatTime(rsReq("appTimeFrom")) & """,""" & Z_FormatTime(rsReq("appTimeTo")) & _
					""",""" & Z_FormatTime(rsReq("astarttime")) & """,""" & Z_FormatTime(rsReq("aendtime")) & """,""" & tmpRate & """,""" & Z_FormatNumber(rsReq("IntrRate"), 2) & _
					""",""" & rsReq("comment") & """" & vbCrLf
				RepCSV =  "List.csv"
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
			End If
		Else
				strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
				"<td class='tblgrn2' width='10px'>" & Stat & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsReq("r.[index]") & "'><a class='link2' href='reqconfirm.asp?ID=" & rsReq("r.[index]") & "'><b>" & rsReq("r.[index]") & "</b></a></td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & tmpIname & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & myDept & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpInName & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & Z_FormatTime(rsReq("appTimeFrom")) & " - " & Z_FormatTime(rsReq("appTimeTo")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' " & BilledAko & " name='txtstime" & x & "' maxlength='5' size='7' value='" & Z_FormatTime(rsReq("astarttime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' " & BilledAko & " name='txtetime" & x & "' maxlength='5' size='7' value='" & Z_FormatTime(rsReq("aendtime")) & "' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');""></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' " & BilledAko & " name='txtBilHrs" & x & "' maxlength='5' size='5' value='" & rsReq("Billable") & "'></td>" & vbCrLf & _
				"<td class='tblgrn2' ><select class='seltxt' style='width: 70px;' name='selInstRate" & x & "'><option value='0' >&nbsp;</option>" & strRate & "</select></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input class='main2' readonly name='txtIntrRate" & x & "' maxlength='7' size='9' value='$" & Z_FormatNumber(rsReq("IntrRate"), 2) & "'></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='checkbox' disabled name='chkver" & x & "' " & tmpVer & "></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='checkbox' disabled name='chkbil" & x & "' " & tmpPaid & "></td>" & vbCrLf & _
				"<td class='tblgrn2' ><textarea class='main' " & BilledAko & " name='txtcom" & x & "'>" & rsReq("comment") & "</textarea></td></tr>" & vbCrLf
		End If
		x = x + 1
		rsReq.MoveNext
	Loop
Else
	strtbl = "<tr><td colspan='15' align='center'><i>&lt -- No records found. -- &gt</i></td></tr>"
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
'GET DEPT LIST
'Set rsInst = Server.CreateObject("ADODB.RecordSet")
'sqlInst = "SELECT dept, [Index] FROM dept_T ORDER BY [Facility]"
'rsInst.Open sqlInst, g_strCONN, 3, 1
'Do Until rsInst.EOF
'	deptSel = ""
'	If Cint(Request("selDept")) = rsInst("Index") Then deptSel = "selected"
'	deptName = rsInst("dept")
'	strInst = strInst	& "<option value='" & rsInst("Index") & "' " & deptSel & ">" &  deptName & "</option>" & vbCrlf
'	rsInst.MoveNext
'Loop
'rsInst.Close
'Set rsInst = Nothing
'GET INSTITUTION LIST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT Facility, [Index] FROM institution_T ORDER BY [Facility]"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	tmpInst = Request("selInst")
	tmpDept = Request("selDept")
	InstSel = ""
	If Cint(Request("selInst")) = rsInst("Index") Then InstSel = "selected"
	InstName = rsInst("Facility")
	strInst = strInst	& "<option value='" & rsInst("Index") & "' " & InstSel & ">" &  InstName & "</option>" & vbCrlf
	
	InstDept = rsInst("Index")
	'strInstDept = strInstDept & "if (inst == " & InstDept & "){" & vbCrLf
	'Set rsDeptInst = Server.CreateObject("ADODB.RecordSet")
	'sqlDeptInst = "SELECT * FROM dept_T WHERE InstID = " &  InstDept & " ORDER BY Dept"
	'rsDeptInst.Open sqlDeptInst, g_strCONN, 3, 1
	'If Not rsDeptInst.EOF Then
	'	Do Until rsDeptInst.EOF
	'		strInstDept = strInstDept & "if (dept != " & rsDeptInst("index") & ")" & vbCrLf & _
	'			"{var ChoiceInst = document.createElement('option');" & vbCrLf & _
	'			"ChoiceInst.value = " & rsDeptInst("index") & ";" & vbCrLf & _
	'			"ChoiceInst.appendChild(document.createTextNode(""" & rsDeptInst("Dept") & """));" & vbCrLf & _
	'			"document.frmTbl.selDept.appendChild(ChoiceInst);} " & vbCrlf
	'		
	'			deptName = rsDeptInst("dept")
	'			strDept = strDept	& "<option value='" & rsDeptInst("Index") & "' " & deptSel & ">" &  deptName & "</option>" & vbCrlf
	'		
	'		rsDeptInst.MoveNext
	'	Loop
	'End If
	'rsDeptInst.Close
	'Set rsDeptInst = Nothing
	'strInstDept = strInstDept & "}"
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'display dept
If tmpInst = "" Then tmpInst = 0
If tmpDept = "" Then tmpDept = 0
if tmpDept > 0 Then
	Set rsdept = Server.CreateObject("ADODB.RecordSet")
	rsdept.open "select [index] as deptid, dept from dept_T where [index] = " & tmpdept, g_strCONN, 3, 1
	if not rsdept.eof then 
		tmpdepttxt = rsdept("dept")
	end if
	rsdept.close
	set rsdept = Nothing
end if
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
	Case 3 Court = "selected"	
	Case 4 Med = "selected"
	Case 5 Legal = "selected"
	Case 6 Mental = "selected"
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
<!-- #include file="_closeSQL.asp" -->
%>
<html>
	<head>
		<title>Language Bank - Table Request</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function SaveMe()
		{
			var ans = window.confirm("This action will save all entries inside the table to the database. Please double check your enties.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmTbl.action = "action.asp?ctrl=3";
				document.frmTbl.submit();
			}
		}
		function SortMe(sortnum)
		{
			document.frmTbl.action = "reqtable.asp?sort=" + sortnum + "&sType=" + <%=stype%>;
			document.frmTbl.submit();
		}
		function FindMe()
		{
			document.frmTbl.submit();
		}
		function FixSort()
		{
			document.frmTbl.txtFromd8.disabled = true;
			document.frmTbl.txtTod8.disabled = true;
			document.frmTbl.txtFromID.disabled = true;
			document.frmTbl.txtToID.disabled = true;
			document.frmTbl.txtFromHPID.disabled = true;
			document.frmTbl.txtToHPID.disabled = true;
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
			if (document.frmTbl.radioStat[2].checked == true)
			{
				document.frmTbl.txtFromHPID.disabled = false;
				document.frmTbl.txtToHPID.disabled = false;
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
		function opendept(intInst) {
				newwindow = window.open('_deplist.asp?instID=' + intInst,'name','height=500,width=400,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
			}
		function clearme() {
			document.frmTbl.selDept.value = 0;
			document.frmTbl.txtDept.value = '';
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
			<form method='post' name='frmTbl' action='reqtable.asp'>
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
															<option <%=meMisded%> value='3'>Missed</option>
															<option <%=meCanceled%> value='4'>Canceled</option>
															<option <%=meCanceledBill%> value='5'>Canceled (Billable)</option>
															<option <%=meBilled%> value='2'>BILLED</option>
														</select>
														<input class='btntbl' type='button' value='GO' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe();'>
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
															<input type='hidden' name='Hctr' value='<%=x%>'>
															<input class='btntbl' type='button' value='Export Table' style='height: 25px; width: 75px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick="document.location='<%=tmpstring%>';">
															<input class='btntbl' type='button' value='Save Table' style='height: 25px; width: 75px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveMe();'>
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
									<td colspan='10' align='left'>
										<div class='container' style='height: 500px; width:80%; position: relative;'>
											<table class="reqtble" width='100%'>	
												<thead>
													<tr class="noscroll">	
														<td colspan='2' class='tblgrn' onclick='SortMe(1);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Request ID</td>
														<td class='tblgrn' onclick='SortMe(2);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Institution</td>
														<td class='tblgrn' onclick='SortMe(15);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Department</td>
														<td class='tblgrn' onclick='SortMe(3);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Language</td>
														<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
															<td class='tblgrn' onclick='SortMe(4);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Client</td>
														<% End If %>
														<td class='tblgrn' onclick='SortMe(5);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Interpreter</td>
														<td class='tblgrn' onclick='SortMe(6);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Appointment Date</td>
														<td class='tblgrn' onclick='SortMe(7);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Planned Start and End Time</td>
														<td class='tblgrn' onclick='SortMe(8);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual Start Time</td>
														<td class='tblgrn' onclick='SortMe(9);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual End Time</td>
														<td class='tblgrn' onclick='SortMe(10);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Billable Hours</td>
														<td class='tblgrn' onclick='SortMe(11);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Institution Rate</td>
														<td class='tblgrn' onclick='SortMe(14);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Interpreter Rate</td>
														<td class='tblgrn' onclick='SortMe(12);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Confirmed</td>
														<td class='tblgrn' onclick='SortMe(13);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Billed</td>
														<td class='tblgrn'>Comment</td>
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
										* To make the request billable, please set actual time, billable hours, and rates.
									</td>
									<td align='right'>
										<% If x <> 0 Then %>
											<b><u><%=x - 1%></u></b> records &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<% End If %>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='1005px' border='0' style='border: solid 1px;'>
								<tr bgcolor='#FBEEB7'>
									<td align='right' style='border-bottom: solid 1px;'><b>Sort:</b></td>
									<td style='border-right: solid 1px;border-bottom: solid 1px;'>
										<input type='radio' name='radioStat' value='0' <%=radioApp%> onclick='FixSort();'>&nbsp;<b>App. Date Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='10' maxlength='10' name='txtFromd8' value='<%=tmpFromd8%>'>
										&nbsp;-&nbsp;
										<input class='main' size='10' maxlength='10' name='txtTod8' value='<%=tmpTod8%>'>
										<span class='formatsmall'>mm/dd/yyyy</span>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='1' <%=radioID%> onclick='FixSort();'>&nbsp;<b>Request ID Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='7' maxlength='7' name='txtFromID' value='<%=tmpFromID%>'>
										&nbsp;-&nbsp;
										<input class='main' size='7' maxlength='7' name='txtToID' value='<%=tmpToID%>'>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='3' <%=radioHPID%> onclick='FixSort();'>&nbsp;<b>Vendor Site ID Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='7' maxlength='7' name='txtFromHPID' value='<%=tmpFromHPID%>'>
										&nbsp;-&nbsp;
										<input class='main' size='7' maxlength='7' name='txtToHPID' value='<%=tmpToHPID%>'>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='2' <%=radioAll%> onclick='FixSort();'>&nbsp;<b>All</b>
									</td>
									<td align='right' style='border-bottom: solid 1px;'><b>Type:</b></td>
									<td style='border-bottom: solid 1px;'>
										<input type='radio' name='radioAss' value='0' <%=radioAss%> onclick='FixSort();'>&nbsp;<b>Assigned</b>
										&nbsp;&nbsp;
										<input type='radio' name='radioAss' value='1' <%=radioUnAss%> onclick='FixSort();'>&nbsp;<b>Unassigned</b>
										&nbsp;&nbsp;
									</td>
								</tr>
								<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='3'>
										Institution:
										<select class='seltxt' style='width: 150px;' name='selInst' onchange='clearme();'>
											<option value='-1'>&nbsp;</option>
											<%=strInst%>
										</select>
										&nbsp;Department:
										<input type="hidden" class="main" name="selDept" id="selDept" value="<%=tmpDept%>">
										<input type="text" class="main" name="txtDept" id="txtDept" readonly value="<%=tmpdepttxt%>" onclick="opendept(document.frmTbl.selInst.value);">
										<!--<select class='seltxt' style='width: 150px;' name='selDept'>
											<option value='0'>&nbsp;</option>
											<%=strDept%>
										</select>//-->
										&nbsp;Language:
										<select class='seltxt' style='width: 150px;' name='selLang'>
											<option value='-1'>&nbsp;</option>
											<%=strLang%>
										</select>
									</td>
									<td align='center' style='border-left: solid 1px;' rowspan='2'>
										<input class='btntbl' type='button' value='FIND' style='height: 35px; width: 150px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe();'>
									</td>
								</tr>
								<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='3'>
										Interpreter:
										<select class='seltxt' name='selIntr' style='width: 200px;' >
											<option value='-1'>&nbsp;</option>
											<%=strIntr%>
										</select>
										<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
											&nbsp;Client:
											<input class='main' size='20' maxlength='20' name='txtclilname' value="<%=tmpclilname%>">
											&nbsp;,&nbsp;&nbsp;
											<input class='main' size='20' maxlength='20' name='txtclifname' value="<%=tmpclifname%>">
											<span class='formatsmall' >Last name, First name</span>
										<% End If %>
										&nbsp;Classification:
										<select class='seltxt' style='width: 100px;' name='selClass'>
											<option value='-1'>&nbsp;</option>
											<option value='1' <%=SocSer%>>Social Services</option>
											<option value='2' <%=Priv%>>Private</option>
											<option value='3' <%=Court%>>Court</option>
											<option value='4' <%=Med%>>Medical</option>
											<option value='5' <%=Legal%>>Legal</option>
											<option value='6' <%=mental%>>Mental Health</option>
										</select>
									</td>
									
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
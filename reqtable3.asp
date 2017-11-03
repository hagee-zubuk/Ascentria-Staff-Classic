<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
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
x = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST"  Or Request("action") = 3 Then
sqlReq = "SELECT * " & _
	"FROM ( " & _
		"SELECT phoneappt, appTimeFrom, appTimeTo, Emergency, Facility, overmile, payhrs, overpayhrs, r.InstID, IntrID, LangID, InstRate, Processed, ProcessedMedicaid ,Status, DeptID, r.[index], " & _
			"clname, cfname, appDate, InstActTT, InstActMil, astarttime, aendtime, BillInst, Billable, TTRate, MRate, " & _
			"TT_Inst, M_Inst, overMInst, overTTInst, ApproveHRs, vermed, r.medicaid " & _
			", [last name], [first name], [language], class " & _
		"FROM request_T r " & _
			"INNER JOIN institution_T i ON r.InstID = i.[index] " & _
			"INNER JOIN language_T l ON r.LangId = l.[index] " & _
			"INNER JOIN interpreter_T [in] ON r.IntrId = [in].[index] " & _
			"INNER JOIN dept_T [dt] ON r.DeptId = [dt].[index] " & _
		"WHERE r.[instID] <> 479 " & _
		"UNION " & _
		"SELECT phoneappt, appTimeFrom, appTimeTo, Emergency, Facility, overmile, payhrs, overpayhrs, r.InstID, IntrID, LangID, InstRate, Processed, ProcessedMedicaid ,Status, DeptID, r.[index], " & _
			"clname, cfname, appDate, InstActTT, InstActMil, astarttime, aendtime, BillInst, Billable, TTRate, MRate, " & _
			"TT_Inst, M_Inst, overMInst, overTTInst, ApproveHRs, vermed, r.medicaid " & _
			", [last name]='', [first name]='', [language], class " & _
		"FROM request_T r " & _
			"INNER JOIN institution_T i ON r.InstID = i.[index] " & _
			"INNER JOIN language_T l ON r.LangId = l.[index] " & _
			"INNER JOIN dept_T [dt] ON r.DeptId = [dt].[index] " & _
		"WHERE r.[instID] <> 479 AND r.IntrId <= 0 " & _
	") mm " & _
	"WHERE "
'response.write "TEST"
'		sqlReq = "SELECT overmile, payhrs, overpayhrs, request_T.InstID, IntrID, LangID, InstRate, Processed, ProcessedMedicaid ,Status, DeptID, request_T.[index], " & _
'			"clname, cfname, appDate, InstActTT, InstActMil, astarttime, aendtime, BillInst, Billable, TTRate, MRate, " & _
'			"TT_Inst, M_Inst, overMInst, overTTInst, ApproveHRs, vermed, medicaid " & _
'			"FROM request_T, institution_T, language_T, interpreter_T, requester_T, dept_T " & _
'			"WHERE request_T.[instID] <> 479 AND request_T.InstID = institution_T.[index] " & _
'			"AND LangId = language_T.[index] " & _
'			"AND (IntrId = interpreter_T.[index])" & _
'			"AND request_T.DeptId = dept_T.[index] " & _
'			"AND ReqID = requester_T.[index] " 
			'If Request("ctrlX") = 1 Then
				If Request("radioAss") = 0 Then	
					sqlReq = sqlReq & "(status = 0 OR status = 4 OR status = 1) AND ApproveHrs = 0"
					radioAss = "checked"
					radioUnass = ""
					radioUnass2 = ""
					noAppr = ""
				ElseIf Request("radioAss") = 1 Then	
					sqlReq = sqlReq & "(status = 1 OR status = 4 OR status = 0) AND ApproveHrs = 1"
					radioAss = ""
					radioUnass = "checked"
					radioUnass2 = ""
					noAppr = "disabled"
				Else
					radioAss = ""
					radioUnass = ""
					radioUnass2 = "checked"
				End If
		
	
	'FIND
	If Request("radioStat") = 0 Then
		radioApp = "checked"
		radioID = ""
		radioAll = ""
		If Request("txtFromd8") <> "" Then
			If IsDate(Request("txtFromd8")) Then
				sqlReq = sqlReq & " AND appDate >= '" & Request("txtFromd8") & "' "
				tmpFromd8 = Request("txtFromd8") 
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (From)."
				Response.Redirect "reqtable3.asp?ctrlX=1"
			End If
		End If
		If Request("txtTod8") <> "" Then
			If IsDate(Request("txtTod8")) Then
				sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
				tmpTod8 = Request("txtTod8")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (To)."
				Response.Redirect "reqtable3.asp?ctrlX=1"
			End If
		End If
	ElseIf Request("radioStat") = 1 Then
		radioApp = ""
		radioID = "checked"
		radioAll = ""
		If Request("txtFromID") <> "" Then
			If IsNumeric(Request("txtFromID")) Then
				sqlReq = sqlReq & " AND [index] >= " & Request("txtFromID")
				tmpFromID = Request("txtFromID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (From)."
				Response.Redirect "reqtable3.asp?ctrlX=1"
			End If
		End If
		If Request("txtToID") <> "" Then
			If IsNumeric(Request("txtToID")) Then
				sqlReq = sqlReq & " AND [index] <= " & Request("txtToID")
				tmpToID = Request("txtToID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (To)."
				Response.Redirect "reqtable3.asp?ctrlX=1"
			End If
		End If
	Else
		radioApp = ""
		radioID = ""
		radioAll = "checked"
	End If
	'FILTER
	xInst = Cint(Request("selInst"))
	If xInst <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "[instID] = " & xInst
	End If
	xLang = Cint(Request("selLang"))
	If xLang <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "[LangID] = " & xLang
	End If
	If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
			If Trim(Request("txtclilname")) <> "" Then
				sqlReq = sqlReq & " AND Upper(Clname) LIKE '" & CleanMe2(Ucase(Trim(Request("txtclilname")))) & "%'"
			End If
			If Trim(Request("txtclifname")) <> "" Then
				sqlReq = sqlReq & " AND Upper(Cfname) LIKE '" & CleanMe2(Ucase(Trim(Request("txtclifname")))) & "%'"
			End If

	End If
	xIntr = Cint(Request("selIntr"))
	If xIntr <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "IntrID = " & xIntr
	End If
	xClass = Cint(Request("selClass"))
	If xClass <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "Class = " & xClass
	End If
	'If Request("ctrlX") = 1 Then
		sqlReq = sqlReq & " AND (Processed IS NULL AND ProcessedMedicaid IS NULL) AND NOT AStarttime IS NULL AND NOT AEndtime IS NULL"
		If Request("sort") <> "" Then
			If Request("sort") = 1 Then sqlReq = sqlReq & " ORDER BY Request_T.[index]"
			If Request("sort") = 2 Then sqlReq = sqlReq & " ORDER BY Facility"
			If Request("sort") = 3 Then sqlReq = sqlReq & " ORDER BY [Language]"
			If Request("sort") = 4 Then sqlReq = sqlReq & " ORDER BY Clname"
			If Request("sort") = 5 And Request("radioAss") = 0 Then sqlReq = sqlReq & " ORDER BY [Last Name]"
			If Request("sort") = 6 Then sqlReq = sqlReq & " ORDER BY appDate"
			If Request("sort") = 7 Then sqlReq = sqlReq & " ORDER BY InstRate"
			If Request("sort") = 8 Then sqlReq = sqlReq & " ORDER BY AStarttime"
			If Request("sort") = 9 Then sqlReq = sqlReq & " ORDER BY AEndtime"
			If Request("sort") = 10 Then sqlReq = sqlReq & " ORDER BY BillInst"
			If Request("sort") = 11 Then sqlReq = sqlReq & " ORDER BY TTRate"
			If Request("sort") = 12 Then sqlReq = sqlReq & " ORDER BY MRate"
			If Request("sort") = 13 Then sqlReq = sqlReq & " ORDER BY InstActTT"
			If Request("sort") = 14 Then sqlReq = sqlReq & " ORDER BY InstActMil"
			If Request("sort") = 15 Then sqlReq = sqlReq & " ORDER BY TT_Inst"
			If Request("sort") = 16 Then sqlReq = sqlReq & " ORDER BY M_Inst"
			If Request("sort") = 5 And Request("radioAss") = 1 Then
				
			Else
				If Request("stype") = 1 Then sqlReq = sqlReq & " DESC"
				If Request("stype") = 2 Then sqlReq = sqlReq & " ASC"
			End If
			'FIX SORT
			
			If Request("sort") = 4 Then sqlReq = sqlReq & ", Cfname ASC"
			If Request("sort") = 5 And Request("radioAss") = 0 Then sqlReq = sqlReq & ", [First Name] ASC"
		Else
			sqlReq = sqlReq & " ORDER BY appDate, Facility, [last name], [first name]"
		End If	
		
	'Else
	'	sqlReq = sqlReq & " AND (NOT medicaid IS NULL OR medicaid <> '') AND Processed IS NULL AND NOT AStarttime IS NULL AND NOT AEndtime IS NULL ORDER BY appDate, Facility, [last name], [first name]"
	'End If
'End If
'response.write sqlreq
'GET REQUESTS
Set rsReq = Server.CreateObject("ADODB.RecordSet")
rsReq.Open sqlReq, g_strCONN, 3, 1
x = 1
If Not rsReq.EOF Then
	Do Until rsReq.EOF
		kulay = ""
		If Not Z_IsOdd(x) Then kulay = "#FBEEB7"
		If rsReq("phoneappt") Then kulay = "#99ff99" 'Phone Call Appt
		If rsReq("Emergency") Then kulay = "#AFEEEE"
		'GET INSTITUTION
		Set rsInst = Server.CreateObject("ADODB.RecordSet")
		sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & rsReq("InstID")
		rsInst.Open sqlInst, g_strCONN, 3, 1
		If Not rsInst.EOF Then
			tmpIname = rsInst("Facility")  
			'If rsInst("Department") <> "" Then tmpIname = tmpIname & " <br> " & rsInst("Department")
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
		Else
			tmpInName = "N/A"
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
	
		Stat = MyStatus(rsReq("Status") )
		myDept =  GetMyDept(rsReq("DeptID"))
		'If rsReq("BillInst") Then
			TT = Z_FormatNumber(Z_Czero(rsReq("InstActTT")), 2)
			Mile = Z_FormatNumber(Z_Czero(rsReq("InstActMil")), 2)
		'	amtTT = rsReq("TTrate") * rsReq("actTT")
		'	amtMile = rsReq("Mrate") * rsReq("actMil")
		'Else
		'	TT = Z_FormatNumber(rsReq("actTT"), 2)
		'	Mile = 0
		'	amtTT = 0
		'	amtMile = 0
		'End IfZ_FixNull
		'BilHrs = Z_Czero(rsReq("Billable"))
		BilHrs = Z_FixNull(rsReq("Billable"))
		autopop = ""
		If BilHrs = "" Then 
			BilHrs = Z_FormatNumber(InstBillHrs(rsReq("astarttime"), rsReq("aendtime"), rsReq("InstID"), rsReq("DeptID"), rsReq("appdate")), 2)
			autopop = "•"
		End If
		'If BilNullHrs = "" Then autopop = "*"

		FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
		BlnOver2 = ""
		If rsReq("overmile") Then BlnOver2 = "checked"
		'tmpAMT = Z_FormatNumber(rsReq("actMil"), 2)
		tmpBilTInst = rsReq("TT_Inst")
		tmpBilMInst = rsReq("M_Inst")
		OverMInst = ""
		If rsReq("overMInst") = true Then OverMInst = "checked"
		If OverMInst = "" Then
			strjscript = strjscript & "document.frmTbl.txtMile" & x & ".value = Math.round((" & Replace(Mile,",","") & " * document.frmTbl.txtmrate" & x & ".value) * 100)/100;" & vbCrLf
		Else
			strjscript = strjscript & "document.frmTbl.txtMile" & x & ".value = " & tmpBilMInst & ";" & vbCrLf
		End If
		OverTTInst = ""
		If rsReq("overTTInst") = true Then OverTTInst = "checked"
		If OverTTInst = "" Then
			strjscript = strjscript & "document.frmTbl.txtTT" & x & ".value = Math.round((" & TT & " * document.frmTbl.txtTTrate" & x & ".value) * 100)/100;" & vbCrLf
		Else
			strjscript = strjscript & "document.frmTbl.txtTT" & x & ".value = " & tmpBilTInst & ";" & vbCrLf
		End If
		'If Request("ctrlX") = 1 Then
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
		'Else
		'	strRate = strRate & "<option value='45' selected >$45.00</option>" & vbCrLf
		'End If
		tmpRate = rsReq("InstRate")
		BilInst = ""
		BillInstTT = "disabled"
		BillInstMil = "disabled"
		If rsReq("BillInst") Then 
			BilInst = "checked"
			BillInstTT = ""
			BillInstMil = ""
		End If	
		apprHrs = ""
		bilhrsApp = ""
		If rsReq("ApproveHRs") Then 
			apprHrs = "checked disabled"
			BilInst = BilInst & " disabled"
			OverTTInst = OverTTInst & " disabled"
			OverMInst = OverMInst & " disabled"
			bilhrsApp = "readonly"
		End If
		reqHrs = Z_FormatNumber(DateDiff("n", rsReq("appTimeFrom"), rsReq("appTimeTo")) / 60, 2)
		strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
			"<td class='tblgrn2' width='10px'>" & Stat & "</td>" & vbCrLf & _
			"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsReq("Index") & "'><a class='link2' href='reqconfirm.asp?ID=" & rsReq("Index") & "'><b>" & rsReq("Index") & "</b></a></td>" & vbCrLf & _
			"<td class='2' style='font-size: 7pt; text-align: center;'><nobr>" & tmpIname & myDept & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & rsReq("clname") & ", " & rsReq("cfname") & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & tmpInName & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & rsReq("appDate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2' ><select class='seltxt' style='width: 70px;' name='selInstRate" & x & "'><option value='0' >&nbsp;</option>" & strRate & "</select></td>" & vbCrLf & _
			"<td class='tblgrn2' >" & reqHrs & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & Z_FormatTime(rsReq("astarttime")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & Z_FormatTime(rsReq("aendtime")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2' ><input type='checkbox' name='chkBilInst" & x & "' value='1' " & BilInst & " onclick='billme(" & x & ");' ></td>" & vbCrLf & _
			"<td class='tblgrn2' ><nobr>$<input class='main2' name='txtTTrate" & x & "' maxlength='6' size='7' " & BillInstTT & " value='" & Z_CZero(rsReq("TTRate")) & "' onblur='ComputeInstTTM();'>" & vbCrlf & _
			"<td class='tblgrn2' ><nobr>$<input class='main2' name='txtmrate" & x & "' maxlength='6' size='7' " & BillInstMil & " value='" & Z_CZero(rsReq("MRate")) & "' onblur='ComputeInstTTM();'>" & vbCrlf & _
			"<td class='tblgrn2' >" & TT & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & Mile & "</td>" & vbCrLf & _
			"<td class='tblgrn2' ><nobr>$<input class='main2' name='txtTT" & x & "' maxlength='6' size='7' readonly value='" & tmpBilTInst & "'><input type='checkbox' name='chkOverTT" & x & "' value='1' " & OverTTInst & " onclick='overwriteMe(this, document.frmTbl.txtTT" & x & ");'></td>" & vbCrLf & _
			"<td class='tblgrn2' ><nobr>$<input class='main2' name='txtMile" & x & "' maxlength='6' size='7' readonly value='" & tmpBilMInst & "'><input type='checkbox' name='chkOverM" & x & "' value='1' " & OverMInst & " onclick='overwriteMe(this, document.frmTbl.txtMile" & x & ");'></td>" & vbCrLf & _
			"<td class='tblgrn2' ><nobr><input class='main2' name='txtbilHrs" & x & "' maxlength='6' size='7' " & bilhrsApp & " value='" & BilHrs & "'>" & autopop & "</td>" & vbCrLf & _
			"<td class='tblgrn2' ><input type='checkbox' ID='chkM" & x & "' name='chkM" & x & "' " & apprHrs & " value='" & rsReq("Index") & "' ></td></tr>" & vbCrLf
			
			strtbl = strtbl & "</tr>" & vbCrLf
		x = x + 1
		rsReq.MoveNext
	Loop
Else
	strtbl = "<tr><td colspan='20' align='center'><i>&lt -- No records found. -- &gt</i></td></tr>"
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
btndis = "disabled"
if x > 1 then btndis = ""
%>
<html>
	<head>
		<title>Language Bank - Institution Billable Hours</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function overwriteMe(xxx, yyy)
		{
			if (xxx.checked == true)
			{
				yyy.readOnly = false;
			}
			else
			{
				yyy.readOnly = true;
			}
		}
		function billme(xxx)
		{
			var strval = "chkBilInst" + xxx;
			var strTTRate = "txtTTrate" + xxx;
			var strmrate = "txtmrate" + xxx;
			
			if (document.getElementsByName(strval)[0].checked == true)
			{
				document.getElementsByName(strTTRate)[0].disabled = false;
				document.getElementsByName(strmrate)[0].disabled = false;
			}
			else
			{
				document.getElementsByName(strTTRate)[0].value = 0;
				document.getElementsByName(strmrate)[0].value = 0;
				document.getElementsByName(strTTRate)[0].disabled = true;
				document.getElementsByName(strmrate)[0].disabled = true;
			}
		}
		function ComputeInstTTM()
		{
			<%=strjscript%>
		}
		function SaveMe()
		{
			var ans = window.confirm("This action will save all entries inside the table to the database. Please double check your entries.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmTbl.action = "action.asp?ctrl=20&ctrlx=1";
				document.frmTbl.submit();
			}
		}
		function SortMe(sortnum)
		{
			document.frmTbl.action = "reqtable3.asp?sort=" + sortnum + "&sType=" + <%=stype%>;
			document.frmTbl.submit();
		}
		function FindMe(xxx)
		{
			document.frmTbl.action = "reqtable3.asp?action=3&ctrlX=" + xxx;
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
					tmpElem = "chkM" + z;
					document.getElementById(tmpElem).checked = true;
				}	
			}
			else
			{
				for(z = 1; z <= xxx; z ++)
				{
					tmpElem = "chkM" + z;
					document.getElementById(tmpElem).checked = false;
				}	
			}
		}
		

			function ApproveMe()
			{
				var ans = window.confirm("REMINDER: SAVE YOUR DATA FIRST BEFORE APPROVING APPOINTMENTS.\n\nThis action will approve hours in all checked entries inside the table to the database.\nAppointments without Rate or Billable Hours will not be approved.\nPlease double check your enties.\nClick Cancel to stop.");
				if (ans)
				{
					document.frmTbl.action = "action.asp?ctrl=19&ctrlx=1";
					document.frmTbl.submit();
				}
			}
	
		//	function ApproveMe()
		//	{
		//		var ans = window.confirm("This action will approve medicaid in all checked entries inside the table to the database.\nAppointments without Rate or Billable Hours will not be approved.\nPlease double check your enties.\nClick Cancel to stop.");
		//		if (ans)
		//		{
		//			document.frmTbl.action = "action.asp?ctrl=19&ctrlx=2";
		//			document.frmTbl.submit();
		//		}
		//	}
	
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
			<form method='POST' name='frmTbl' action='reqtable3.asp'>
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
												<td align='left' width='1000px' style='vertical-align: bottom;'>
													Legend: <font color='#FF00FF' size='+3'>•</font>&nbsp;-&nbsp;Canceled (billable)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<font color='#336601' size='+3'>•</font>&nbsp;-&nbsp;Computed Billable Hours
												</td>
												<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
													<% If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then %> 
														<td align='right'>
															<input type='hidden' name='Hctr' value='<%=x%>'>
															<input class='btntbl' type='button' <%=btndis%> value='Save Table' style='height: 25px; width: 200px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveMe();'>
														</td>
													<% Else %>
													<td align='left'>
															<input type='hidden' name='ctrlX' value='<%=Request("ctrlX")%>'>
															<input type='hidden' name='Hctr' value='<%=x%>'>
															<%	If Z_CZero(Request("ctrlX")) = 1 Then %>
																<input class='btntbl' type='button' <%=noAppr%> <%=btndis%> value='Save Table' style='height: 25px; width: 150px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='SaveMe();'><br><br>
																<input class='btntbl' type='button' <%=noAppr%> <%=btndis%> value='Approve Hours' style='height: 25px; width: 150px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='ApproveMe();'>
															<% Else %>
																<!--<input class='btntbl' type='button' <%=noAppr%> value='Approve Medicaid' style='height: 25px; width: 150px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='ApproveMe();'>//-->
															<% End If %>
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
														<td colspan='2' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" class='tblgrn' onclick='SortMe(1);'>Request ID</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(2);'>Institution</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(3);'>Language</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(4);'>Client</td>
														<!--<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Medicaid</td>//-->
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(5);'>Interpreter</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(6);'>Appointment Date</td>
													  <td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(7);'>Inst. Rate</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" >Requested Hours</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(8);'>Actual Start Time</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(9);'>Actual End Time</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(10);'>Bill Institution</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(11);'>Travel Time Rate</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(12);'>Mileage Rate</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(13);'>Travel Time (hrs.)</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(14);'>Mileage (m)</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(15);'>Travel Time</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(16);'>Mileage</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" >Billable Hours</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" >
														
															Approve Hours<br>
														
															<input type='checkbox' name='chkall' <%=noAppr%> onclick='checkme(<%=x%>);'>
														</td>
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
											<b><u><%=x - 1%></u></b> records &nbsp;&nbsp;&nbsp;&nbsp;
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
										<!--<input type='radio' name='radioAss' value='2' <%=radioUnAss2%> onclick='FixSort();'>&nbsp;<b>ALL</b>
										&nbsp;&nbsp;//-->
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
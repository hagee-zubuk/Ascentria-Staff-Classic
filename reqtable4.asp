<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_utilsMedicaid.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
Function Z_sqlsinglequote(xxx)
	'CHAR(39)
	Z_sqlsinglequote = xxx
	If Not IsNull(xxx) Or xxx <> "" Then Z_sqlsinglequote = Replace(xxx, "''", "'+CHAR(39)+'") 
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

x = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST"  Or Request("action") = 3 Then
'response.write "TEST"
		sqlReq = "SELECT Happen, Billable, dob, medicaid, meridian, nhhealth, wellsense, vermed, autoacc, wcomp, drg, XID, request_T.InstID, IntrID, LangID, InstRate, ProcessedMedicaid, Status, DeptID, request_T.[index], " & _
			"clname, cfname, appDate " & _
			"FROM request_T, institution_T, language_T, interpreter_T, requester_T, dept_T " & _
			"WHERE request_T.[instID] <> 479 AND request_T.InstID = institution_T.[index] " & _
			"AND LangId = language_T.[index] " & _
			"AND IntrId = interpreter_T.[index] " & _
			"AND request_T.DeptId = dept_T.[index] " & _
			"AND ReqID = requester_T.[index] " & _
			"AND outpatient = 1 AND hasmed = 1 AND autoacc <> 1 AND wcomp <> 1 AND drg = 1 AND (medicaid <> '' OR NOT medicaid IS NULL OR meridian <> '' OR NOT meridian IS NULL OR nhhealth <> '' OR NOT nhhealth IS NULL OR wellsense <> '' OR NOT wellsense IS NULL)"
			'If Request("ctrlX") = 1 Then
				If Request("radioAss") = 0 Then	
					sqlReq = sqlReq & "AND (status = 4 OR status = 1) AND (vermed = 0 OR vermed IS NULL)"
					radioAss = "checked"
					radioUnass = ""
					radioUnass2 = ""
					noAppr = ""
				ElseIf Request("radioAss") = 1 Then	
					sqlReq = sqlReq & "AND (status = 4 OR status = 1) AND vermed = 1"
					radioAss = ""
					radioUnass = "checked"
					radioUnass2 = ""
					noAppr = "disabled"
				ElseIf Request("radioAss") = 2 Then	
					sqlReq = sqlReq & "AND (status = 4 OR status = 1) AND vermed = 2"
					radioAss = ""
					radioUnass = ""
					radioUnass2 = "checked"
					noAppr = "disabled"
				Else
					radioAss = "checked"
					radioUnass = ""
					radioUnass2 = ""
				End If
			'Else
			'	If Request("radioAss") = 0 Then	
			'		sqlReq = sqlReq & "AND (status = 0 OR status = 4) AND vermed = 0"
			'		radioAss = "checked"
			'		radioUnass = ""
			'		radioUnass2 = ""
			'		noAppr = ""
			'	ElseIf Request("radioAss") = 1 Then	
			'		sqlReq = sqlReq & "AND (status = 1 OR status = 4) AND vermed = 1"
			'		radioAss = ""
			'		radioUnass = "checked"
			'		radioUnass2 = ""
			'		noAppr = "disabled"
			'	Else
			'		radioAss = ""
			'		radioUnass = ""
			'		radioUnass2 = "checked"
			'	End If
			'End If
	
	
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
				Response.Redirect "reqtable4.asp"
			End If
		End If
		If Request("txtTod8") <> "" Then
			If IsDate(Request("txtTod8")) Then
				sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
				tmpTod8 = Request("txtTod8")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (To)."
				Response.Redirect "reqtable4.asp"
			End If
		End If
	ElseIf Request("radioStat") = 1 Then
		radioApp = ""
		radioID = "checked"
		radioAll = ""
		If Request("txtFromID") <> "" Then
			If IsNumeric(Request("txtFromID")) Then
				sqlReq = sqlReq & " AND request_T.[index] >= " & Request("txtFromID")
				tmpFromID = Request("txtFromID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (From)."
				Response.Redirect "reqtable4.asp"
			End If
		End If
		If Request("txtToID") <> "" Then
			If IsNumeric(Request("txtToID")) Then
				sqlReq = sqlReq & " AND request_T.[index] <= " & Request("txtToID")
				tmpToID = Request("txtToID")
			Else
				Session("MSG") = "ERROR: Invalid Appointment ID Range (To)."
				Response.Redirect "reqtable4.asp"
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
		sqlReq = sqlReq & "request_T.InstID = " & xInst
	End If
	xLang = Cint(Request("selLang"))
	If xLang <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "LangID = " & xLang
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
	xMCO = Cint(Request("selMCO"))
	selmed = ""
	selmer = ""
 	selnh = ""
 	selwell = ""
	If xMCO > 0 Then
		If xMCO = 1 Then 
			sqlReq = sqlReq & " AND medicaid <> '' AND (meridian = '' OR meridian IS NULL) AND (nhhealth = '' OR  nhhealth IS NULL) AND (wellsense = '' OR wellsense IS NULL) "
			selmed = "SELECTED"
		End If
		If xMCO = 2 Then 
			sqlReq = sqlReq & " AND meridian <> '' "
			selmer= "SELECTED"
		End If
		If xMCO = 3 Then 
			sqlReq = sqlReq & " AND nhhealth <> '' "
			selnh = "SELECTED"
		End If
		If xMCO = 4 Then 
			sqlReq = sqlReq & " AND wellsense <> '' "
			selwell = "SELECTED"
		End If
	End If
	'ADMIN ONLY
	xAdmin = Z_CZero(Request("selAdmin"))
	If xAdmin = 1 Then
		sqlReq = sqlReq & " AND (Status = 1) AND ProcessedMedicaid IS NULL"
		meUnBilled = "selected"
	ElseIf xAdmin = 2 Then
		sqlReq = sqlReq & " AND (Status = 1 OR Status = 4) AND NOT ProcessedMedicaid IS NULL"
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
		sqlReq = sqlReq & " AND ProcessedMedicaid IS NULL "'sqlReq = sqlReq & " AND IsNull(Processed)"
	End If
	'If Request("ctrlX") = 1 Then
		'sqlReq = sqlReq & " AND ProcessedMedicaid IS NULL " 'ORDER BY appDate, Facility, [last name], [first name]"
	'Else
	'	sqlReq = sqlReq & " AND (NOT medicaid IS NULL OR medicaid <> '') AND Processed IS NULL AND NOT AStarttime IS NULL AND NOT AEndtime IS NULL ORDER BY appDate, Facility, [last name], [first name]"
	'End If
'End If
	If Request("sort") <> "" Then
			If Request("sort") = 1 Then sqlReq = sqlReq & " ORDER BY Request_T.[index]"
			If Request("sort") = 2 Then sqlReq = sqlReq & " ORDER BY Facility"
			If Request("sort") = 3 Then sqlReq = sqlReq & " ORDER BY [Language]"
			If Request("sort") = 4 Then sqlReq = sqlReq & " ORDER BY Clname"
			If Request("sort") = 5 Then sqlReq = sqlReq & " ORDER BY dob"
			If Request("sort") = 6 Then sqlReq = sqlReq & " ORDER BY Medicaid"
			If Request("sort") = 7 Then sqlReq = sqlReq & " ORDER BY [last name]"
			If Request("sort") = 8 Then sqlReq = sqlReq & " ORDER BY XID"
			If Request("sort") = 9 Then sqlReq = sqlReq & " ORDER BY appdate"
			If Request("sort") = 10 Then sqlReq = sqlReq & " ORDER BY Billable"
			If Request("sort") = 11 Then sqlReq = sqlReq & " ORDER BY Medicaid, meridian, nhhealth, wellsense"
			If Request("sort") = 12 Then sqlReq = sqlReq & " ORDER BY Happen"	
		

				If Request("stype") = 1 Then sqlReq = sqlReq & " DESC"
				If Request("stype") = 2 Then sqlReq = sqlReq & " ASC"

			'FIX SORT
			
			If Request("sort") = 4 Then sqlReq = sqlReq & ", Cfname ASC"
			If Request("sort") = 7 Then sqlReq = sqlReq & ", [First Name] ASC"
		Else
			sqlReq = sqlReq & " ORDER BY appDate, Facility, [last name], [first name]"
		End If	
'x12 270 Head
dte = FormatDateTime(date, 2)
dteYr = Year(dte)
dteYr2 = Right("0" & Year(dte), 2)
dteMn = Right("0" & Month(dte), 2)
dteDy = Right("0" & Day(dte), 2)
dtetime = FormatDateTime(Now, 4)
tme = Replace(dtetime, ":", "")
If Request("selMCO") = 1 Then
	tradingnumber = "NH100496       "
	idnum = "ZZ*026000618      "
ElseIf Request("selMCO") = 2 Then
ElseIf Request("selMCO") = 3 Then
	tradingnumber = "NH100496       "
	idnum = "ZZ*026000618      "
ElseIf Request("selMCO") = 4 Then
	tradingnumber = "043566243      "
	idnum = "30*043373331      "
End If
strMedHdr = "ISA*00*          *00*          *ZZ*" & tradingnumber & "*" & idnum & "*" & dteYr2 & dteMn & dteDy & "*" & tme & "*^*00501*000000007*0*P*:~" & _ 
	"GS*HS*NH100496*026000618*" & dteYr & dteMn & dteDy & "*" & tme & "*1*X*005010X279A1~"
'GET REQUESTS
Set rsReq = Server.CreateObject("ADODB.RecordSet")
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

		If rsReq("vermed") = 1 Then 
			apprHrs = "checked disabled"
			apprHrs2 = "disabled"
		End If
		If rsReq("vermed") = 2 Then 
			apprHrs = "disabled"
			apprHrs2 = "checked disabled"
		End If
		'medicaid check
		medcheck = ""
		If Request("radioAss") = 0 Then	
			If Not Z_MedicaidCheck(rsReq("index")) Then
				medctr = medctr + 1
				medcheck = "<sup><a href='#' class='question' onclick='medchk(" & rsReq("index") & ");'>[?]</a></sup>"
			End if
		End If
			hmolabel = "" 'Z_FixNull(rsReq("medicaid")) 
			If hmolabel = "" Then hmolabel = Z_FixNull(rsReq("meridian"))
			If hmolabel = "" Then hmolabel = Z_FixNull(rsReq("nhhealth"))
			If hmolabel = "" Then hmolabel = Z_FixNull(rsReq("wellsense"))
			hmo = Z_FixNull(rsReq("medicaid"))
			nm1pr2 = "NH MEDICAID*****PI*026000618" 
			nm11p1 = "OMERBEGOVIC*ALEN****SV*30849597"
			refeo = "REF*EO*820000243~"
			If Z_FixNull(rsReq("meridian")) <> "" Then 
				hmo = Z_FixNull(rsReq("meridian"))
			End If
			If Z_FixNull(rsReq("nhhealth")) <> "" Then 
				hmo = Z_FixNull(rsReq("nhhealth"))
			End If
			If Z_FixNull(rsReq("wellsense")) <> "" Then 
				hmo = Z_FixNull(rsReq("wellsense")) 
				nm1pr2 = "WELL SENSE HEALTH PLAN*****PI*13337"
				nm11p1 = "ASCENTRIA COMMUNITY SERVICES, INC.*****XX*1609133040"
				refeo = ""
			End If
			happen = ""
			If rsReq("happen") = 1 Then happen = "NO"
			If rsReq("happen") = 2 Then happen = "YES"
			strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
				"<td class='tblgrn2' width='10px'>" & Stat & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsReq("Index") & "'><a class='link2' href='reqconfirm.asp?ID=" & rsReq("Index") & "'><b>" & rsReq("Index") & "</b></a></td>" & vbCrLf & _
				"<td class='tblgrn2' ><nobr>" & tmpIname & myDept & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & Z_RemoveDlbQuote(rsReq("clname")) & ", " & Z_RemoveDlbQuote(rsReq("cfname")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("dob") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & Z_FixNull(rsReq("medicaid")) & medcheck & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & hmolabel & medcheck & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & tmpInName & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("XID") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & rsReq("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & Z_FormatNumber(rsReq("Billable"), 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2' >" & happen & "</td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='checkbox' ID='chkM" & x & "' name='chkM" & x & "' " & apprHrs & " value='" & rsReq("Index") & "' onclick='if(this.checked) {document.frmTbl.chkX" & x & ".checked=false;}'></td>" & vbCrLf & _
				"<td class='tblgrn2' ><input type='checkbox' ID='chkX" & x & "' name='chkX" & x & "' " & apprHrs2 & " value='" & rsReq("Index") & "' onclick='if(this.checked) {document.frmTbl.chkM" & x & ".checked=false;}'></td>" & vbCrLf & _
				"</tr>" & vbCrLf
			'x12 270 body
			STnum = Right("0000" & x, 4)
			cleanhmo = Replace(hmo, " ", "")
			strMedBody = strMedBody & "ST*270*" & STnum & "*005010X279A1~" & _
				"BHT*0022*13*10001234*" & dteYr & dteMn & dteDy & "*" & tme & "~" & _
				"HL*1**20*1~" & _
				"NM1*PR*2*" & nm1pr2 & "~" & _
				"HL*2*1*21*1~"
			ptype = 1
			If Z_FixNull(rsReq("wellsense")) <> "" Then ptype = 2
			strMedBody = strMedBody & "NM1*1P*" & ptype & "*" & nm11p1 & "~" & _
				refeo & _
				"HL*3*2*22*0~" & _
				"NM1*IL*1*" & Z_NameMed(rsReq("index")) & "****MI*" & Trim(cleanhmo) & "~" & _
				"DMG*D8*" & Z_DOBMed(rsReq("index")) & "*" & Z_GenderMed(rsReq("index")) & "~" & _
				"DTP*291*RD8*" & Z_DateMed(rsReq("index")) & "-" & Z_DateMed(rsReq("index")) & "~" & _
				"EQ*30~"
			segCount = 13
			If refeo = "" Then segCount = 12
			strMedBody = strMedBody & "SE*" & segCount & "*" & STnum & "~"
		x = x + 1
		rsReq.MoveNext
	Loop
	'x12 270 footer
	strMedFtr = "GE*" & x - 1 & "*1~IEA*1*000000007~"
	strMed = Trim(strMedHdr & strMedBody & strMedFtr)
	'CREATE x12
	tmpdate = replace(date, "/", "") 
	tmpTime = replace(FormatDateTime(time, 3), ":", "")
	tmpTime = replace(tmpTime, " ", "")
	Repx12 =  "x12270-" & tmpdate & tmpTime & ".x12" 
	Set fso = CreateObject("Scripting.FileSystemObject")
	' Response.Write x12Path & Repx12
	Set ofilex12 = fso.CreateTextFile(x12Path & Repx12, 8, True) 
	ofilex12.Write strMed
	Set ofilex12 = Nothing
	' Response.Write "<br /><br /><br />" & x12Path & Repx12 & " to " & x12pathbackup
	fso.CopyFile x12Path & Repx12, x12pathbackup
	Set fso = Nothing
	tmpstring = "x12/" & Repx12
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
		<title>Language Bank - Approve Medicaid</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function medchk(appid) {
			newwindow3 = window.open('medicaidcheck.asp?ReqID=' + appid,'name','height=300,width=800,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow3.focus()}
		}
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
			var ans = window.confirm("This action will save all entries inside the table to the database. Please double check your enties.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmTbl.action = "action.asp?ctrl=20";
				document.frmTbl.submit();
			}
		}
		function SortMe(sortnum)
		{
			document.frmTbl.action = "reqtable4.asp?sort=" + sortnum + "&sType=" + <%=stype%>;
			document.frmTbl.submit();
		}
		function FindMe(xxx) {
			document.frmTbl.action = "reqtable4.asp?action=3";
			document.frmTbl.submit();
		}
		function FixSort() {
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
		function TblFix() {
			/*
			var bodyRect = document.body.getBoundingClientRect();
			var tbl = document.getElementById('tabResults');
			var elemRect = tbl.getBoundingClientRect();
    		var offset  = elemRect.top - bodyRect.top;
    		var y_sz = document.body.clientHeight - offset - 200;
			tbl.style.height = y_sz + "px";	
*/		} 
		function CalendarView(strDate) {
			document.frmTbl.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmTbl.submit();
		}
		function maskMe(str,textbox,loc,delim) {
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
			var tmpElem2;
			var z;
			if (document.frmTbl.chkall.checked == true)
			{
				for(z = 1; z <= xxx; z ++)
				{
					tmpElem = "chkM" + z;
					tmpElem2 = "chkX" + z;
					document.getElementById(tmpElem).checked = true;
					document.getElementById(tmpElem2).checked = false;
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
		function checkme2(xxx)
		{
			var tmpElem;
			var tmpElem2;
			var z;
			if (document.frmTbl.chkall2.checked == true)
			{
				for(z = 1; z <= xxx; z ++)
				{
					tmpElem = "chkX" + z;
					tmpElem2 = "chkM" + z;
					document.getElementById(tmpElem).checked = true;
					document.getElementById(tmpElem2).checked = false;
				}	
			}
			else
			{
				for(z = 1; z <= xxx; z ++)
				{
					tmpElem = "chkX" + z;
					document.getElementById(tmpElem).checked = false;
				}	
			}
		}
		

			function ApproveMe()
			{
				if (document.frmTbl.txtFromd8.value == "" || document.frmTbl.txtTod8.value == "") {
					var ans = window.confirm("You did not select a date range.\nThis will take a considerable amount of time to complete.\nClick Cancel to stop.");
					if (ans)
					{
						var ans = window.confirm("This action will approve/disapprove medicaid in all checked entries inside the table to the database.\nDisaaproved entries will be billed to institution.\nAppointments will only be billable to Medicaid if certain rules are met, even if Medicaid is approved.\nClick Cancel to stop.");
						if (ans)
						{
							document.frmTbl.action = "action.asp?ctrl=22";
							document.frmTbl.submit();
						}
					}
				}
				else {
					var ans = window.confirm("This action will approve/disapprove medicaid in all checked entries inside the table to the database.\nDisaaproved entries will be billed to institution.\nAppointments will only be billable to Medicaid if certain rules are met, even if Medicaid is approved.\nClick Cancel to stop.");
					if (ans)
					{
						document.frmTbl.action = "action.asp?ctrl=22";
						document.frmTbl.submit();
					}
				}
			}
		//function verMed() {
		//	document.frmTbl.action = "action.asp?ctrl=23";
		//	document.frmTbl.submit();
		//}
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
		<body onload='FixSort(); TblFix();'>
			<form method='POST' name='frmTbl' action='reqtable4.asp'>
				<table cellSpacing='0' cellPadding='0' width="100%" border='0' class='bgstyle2'>
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
												<td align='left' width='800px' style='vertical-align: bottom;'>
													Legend: <font color='#FF00FF' size='+3'>•</font>&nbsp;-&nbsp;Canceled (billable)
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													Admin Sort:
														<select class='seltxt' style='width:100px;' name='selAdmin'>
															<option value='0'>&nbsp;</option>
															<option <%=mePending%> value='6'>Pending</option>
															<option <%=meUnBIlled%> value='1'>Completed (Unbilled)</option>
															<option <%=meCanceledBill%> value='5'>Canceled (Billable)</option>
															<option <%=meBilled%> value='2'>BILLED</option>
														</select>
														<input class='btntbl' type='button' value='GO' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe(<%=Request("ctrlX")%>);'>
												</td>
												<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
													<% If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then %> 
														<td align='right'>
															<input type='hidden' name='Hctr' value='<%=x%>'>
																<input class='btntbl' type='button' <%=noAppr%> <%=btndis%> value='Approve/Disapprove Medicaid' style='height: 25px; width: 250px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='ApproveMe();'>
																	<% If Request.Cookies("UID") = 2 Or tmpstring <> "" Then %>
																<input class='btntbl' type="button" value="Download 270 file" style='width: 100px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick="document.location='<%=tmpstring%>';">
															<% End If %>
														</td>
													<% Else %>
													<td align='left'>
															
															<input type='hidden' name='Hctr' value='<%=x%>'>
															
																<% If tmpstring <> "" Then %>
																	<input class='btntbl' type="button" value="Download 270 file" style='height: 25px; width: 275px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick="document.location='<%=tmpstring%>';">
																	<br><br>
																<% End If %>
																<input class='btntbl' type='button' <%=noAppr%> <%=btndis%> value='Approve/Disapprove MCO/Medicaid' style='height: 25px; width: 275px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='ApproveMe();'>
														
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
							</table>
				</table>

<div id="tabResults" style='width:100%; position: relative;'>
	<table class="reqtble" width='100%'>	
		<thead>
			<tr class="noscroll">	
				<td colspan='2' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" class='tblgrn' onclick='SortMe(1);'>Request ID</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(2);'>Institution</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(3);'>Language</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(4);'>Client</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(5);'>DOB</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(6);'>Medicaid</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(11);'>MCO</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(7);'>Interpreter</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(8);'>Xerox ID</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(9);'>Date</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(10);'>Billable Hrs</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'" onclick='SortMe(12);'>Happened</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">
				
					Approve Medicaid<br>
				
					<input type='checkbox' name='chkall' <%=noAppr%> onclick='if(this.checked) {document.frmTbl.chkall2.checked=false;} checkme(<%=x%>);'>
				</td>
				<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">
				
					Disapprove Medicaid<br>
				
					<input type='checkbox' name='chkall2' <%=noAppr%> onclick='if(this.checked) {document.frmTbl.chkall.checked=false;} checkme2(<%=x%>);'>
				</td>
			</tr>
		</thead>
		<tbody style="OVERFLOW: auto;">
			<%=strtbl%>
		</tbody>
	</table>
	<table width='100%'  border='0'>
		<tr><td align='left'>&nbsp;</td>
			<td align='right'>
<% If x <> 0 Then %>
				<b><u><%=x - 1%></u></b> records &nbsp;&nbsp;&nbsp;&nbsp;
<% End If %>
				</td>
			<td>&nbsp;</td>
		</tr>
	</table>
</div>	


		<table cellSpacing='0' cellPadding='0' width='1005px' border='0' style="border: solid 1px; position: relative;">
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
					<input type='radio' name='radioAss' value='0' <%=radioAss%> onclick='FixSort();'>&nbsp;<b>For Review</b>
					&nbsp;&nbsp;
					<input type='radio' name='radioAss' value='1' <%=radioUnAss%> onclick='FixSort();'>&nbsp;<b>Approved</b>
					&nbsp;&nbsp;
					<input type='radio' name='radioAss' value='2' <%=radioUnAss2%> onclick='FixSort();'>&nbsp;<b>Disapprove</b>
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
					&nbsp;Medicaid/MCO:
					<select class='seltxt' style='width: 100px;' name='selMCO'>
						<option value='1' <%=selmed%>>Medicaid</option>
						<option value='2' <%=selmer%>>Meridian Health Plan</option>
						<option value='3' <%=selnh%>>NH Healthy Families</option>
						<option value='4' <%=selwell%>>Well Sense Health Plan</option>
					</select>
				</td>
				<td>&nbsp;</td>
			</tr>
		</table>
<!-- footer! -->
<table cellSpacing='0' cellPadding='0' width="100%" border='0' class='bgstyle2'>
	<tr><td height='50px' valign='bottom'>
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
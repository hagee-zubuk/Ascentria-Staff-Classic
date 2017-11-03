<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
server.scripttimeout = 360000
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function GetLoc(xxx)
	Select Case xxx
		Case 0 
			GetLoc = "Front Door"
		Case 1
			GetLoc = "Cafeteria"
		Case 2
			GetLoc = "Registration Desk"
		Case 3
			GetLoc = "Department"
		Case 4
			GetLoc = "OTHER"
	End Select
End Function
tmpPage = "document.frmMain."
tmpInst = "-1"
tmpIntr = "-1"
'default
selRPEmail = ""
selRPPhone = ""
selRPFax = "checked"
'default
selIntrFax = "checked"
selIntrP2 = ""
selIntrP1 = ""
selIntrEmail = ""
tmpTS = Now
tmpDept = 0
tmpReqP = "-1"
tmpHPID = 0

'RESUPPLY DATA IN ERROR EVENT
If Session("MSG") <> "" And Request("ID") = "" Then
	tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
	tmpNewInst = Split(Z_DoDecrypt(Request.Cookies("LBINST")), "|")
	tmpNewDept = Split(Z_DoDecrypt(Request.Cookies("LBDEPT")), "|")
	tmpNewReq = Split(Z_DoDecrypt(Request.Cookies("LBREQ")), "|")
	tmpNewIntr = Split(Z_DoDecrypt(Request.Cookies("LBINTR")), "|")
	tmpNewIntrBTN = tmpNewIntr(10)
	If tmpNewReq(6) = "BACK" Then
		tmpNewReqLN = tmpNewReq(0)
		tmpNewReqFN = tmpNewReq(1)
		tmpReqExt = tmpNewReq(8)
		tmpNewReqPhone = tmpNewReq(2)
		tmpNewReqEmail = tmpNewReq(3)
		tmpNewReqFax = tmpNewReq(4)
		tmpNewReqPrim = tmpNewReq(7)
		selRPEmail = ""
		selRPPhone = ""
		selRPFax = ""
		if tmpNewReqPrim = 0 Then selRPEmail = "checked"
		if tmpNewReqPrim = 1 Then selRPPhone = "checked"
		if tmpNewReqPrim = 2 Then selRPFax = "checked"
	Else
		tmpReqP = tmpEntry(1)
	End If
	tmplName = tmpEntry(2)
	tmpfName = tmpEntry(3)
	chkClient = ""
	If tmpEntry(20) <> "" Then chkClient = "checked"
	tmpAddr = tmpEntry(4)
	chkUClientadd = ""
	If tmpEntry(29) <> "" Then chkUClientadd = "checked"
	tmpCity = tmpEntry(5)
	tmpState = tmpEntry(6)
	tmpZip = tmpEntry(7)
	tmpCAdrI = tmpEntry(30)
	tmpDir = tmpEntry(8)
	tmpSC = tmpEntry(9)
	tmpDOB = tmpEntry(10)
	tmpLang = tmpEntry(11)
	tmpAppDate = tmpEntry(12)
	tmpAppTFrom = Z_FormatTime(tmpEntry(13))
	tmpAppTTo = Z_FormatTime(tmpEntry(14))
	tmpAppLoc = tmpEntry(15)
	If tmpNewInst(6) = "BACK" Then 
		tmpNewInstchk = tmpNewInst(6)
		tmpNewInstTxt = tmpNewInst(0)
		'tmpNewInstDept = tmpNewInst(1)
		'tmpNewInstAddr = tmpNewInst(2)
		'tmpNewInstCity = tmpNewInst(3)
		'tmpNewInstState = tmpNewInst(4)
		'tmpNewInstZip = tmpNewInst(5)
		'SocSer = ""
		'Priv = ""
		'Legal = ""
		'Med = ""
		'If tmpNewInst(7) = 0 Then SocSer = "selected"
		'If tmpNewInst(7) = 1 Then Priv = "selected"
		'If tmpNewInst(7) = 2 Then Legal = "selected"
		'If tmpNewInst(7) = 3 Then Med = "selected"
		'If tmpNewInst(8) <> "" Then 
		'	chkBillMe = "checked"
		'Else
		'	tmpBilInstAddr = tmpNewInst(9)
		'	tmpBilInstCity = tmpNewInst(10)
		'	tmpBilInstState = tmpNewInst(11)
		'	tmpBilInstZip = tmpNewInst(12)
		'End If	
		'tmpBLname =  tmpNewInst(13)
		'tmpBFname =  tmpNewInst(14)
	Else
		tmpInst = tmpEntry(16)
	End If
	If tmpNewDept(6) = "BACK" Then
		tmpNewInstchk = tmpNewDept(6)
		'tmpNewInstTxt = tmpNewDept(1)
		tmpNewInstDept = tmpNewDept(0)
		SocSer = ""
		Priv = ""
		Legal = ""
		Med = ""
		tmpClass = tmpNewDept(7)
		If tmpNewDept(7) = 1 Then SocSer = "selected"
		If tmpNewDept(7) = 2 Then Priv = "selected"
		If tmpNewDept(7) = 3 Then Court = "selected"
		If tmpNewDept(7) = 4 Then Med = "selected"
		If tmpNewDept(7) = 5 Then Legal = "selected"
		tmpNewInstAddr = tmpNewDept(2)
		tmpNewInstCity = tmpNewDept(3)
		tmpNewInstState = tmpNewDept(4)
		tmpNewInstZip = tmpNewDept(5)
		tmpNewInstAddrI = tmpNewDept(16)
		If tmpNewInst(8) <> "" Then 
			chkBillMe = "checked"
		Else
			tmpBilInstAddr = tmpNewDept(9)
			tmpBilInstCity = tmpNewDept(10)
			tmpBilInstState = tmpNewDept(11)
			tmpBilInstZip = tmpNewDept(12)
		End If	
		tmpBLname =  tmpNewDept(13)
		'tmpBFname =  tmpNewDept(14)
	Else
		tmpDept = tmpEntry(26)
	End If
	tmpRate = tmpEntry(17)
	tmpDoc = tmpEntry(18)
	tmpCRN = tmpEntry(19)
	tmpCFon = tmpEntry(21)
	If Request.Cookies("LBACTION") = 1 Then
		tmpCAFon = tmpEntry(27)
	Else
		tmpCAFon = tmpEntry(38)
	End If
	If tmpNewIntrBTN = "BACK" Then
		tmpIntrLname = tmpNewIntr(0)
		tmpIntrFname = tmpNewIntr(1)
		tmpIntrEmail = tmpNewIntr(2)
		tmpIntrP1 = tmpNewIntr(3)
		tmpIntrExt = tmpNewIntr(13)
		tmpIntrFax = tmpNewIntr(4)
		tmpIntrP2 = tmpNewIntr(5)
		tmpIntrAddr = tmpNewIntr(6)
		tmpIntrCity = tmpNewIntr(7)
		tmpIntrState = tmpNewIntr(8)
		tmpNewIntrAddrI = tmpNewIntr(16)
		tmpIntrZip = tmpNewIntr(9)
		tmpInHouse = ""
		If tmpNewIntr(11) <> "" Then tmpInHouse = "checked"
		tmpIntrPrim = tmpNewIntr(12)
		selIntrFax = ""
		selIntrP2 = ""
		selIntrP1 = ""
		selIntrEmail = ""
		if tmpIntrPrim = 0 Then selIntrEmail = "checked"
		if tmpIntrPrim = 1 Then selIntrP1 = "checked"
		if tmpIntrPrim = 2 Then selIntrP2 = "checked"
		if tmpIntrPrim = 0 Then selIntrFax = "checked"
		tmpIntrRate =  tmpNewIntr(14)
	Else
		tmpIntr = tmpEntry(22)
	End If
	tmpEmer = ""
	If tmpEntry(24) <> "" Then tmpEmer = "checked"
	tmpEmerFee = ""
	If tmpEntry(32) <> "" Then tmpEmerFee = "checked"	
	tmpCombil = tmpEntry(33) 
	If Request.Cookies("LBACTION") = 1 Then	
		tmpCom =  tmpEntry(25)
		tmpComintr = tmpEntry(31)
		'tmpHPID = tmpEntry(30)
	Else
		tmpCom =  tmpEntry(30)
		'tmpHPID = tmpEntry(44)
	End If
End If
'CLONE REQUEST
If Request("Clone") <> "" Then
	Set rsClone = Server.CreateObject("ADODB.RecordSet")
	sqlClone = "SELECT * FROM request_T WHERE index = " & Request("Clone")
	rsClone.Open sqlClone, g_strCONN, 3, 1
	If Not rsCLone.EOF Then
		tmpReqP = rsClone("ReqID") 
		tmplName = rsClone("clname") 
		tmpfName = rsClone("cfname")	
		chkClient = ""
		If rsClone("Client") = True Then chkClient = "checked"
		chkUClientadd = ""
		If  rsClone("CliAdd")  = True Then chkUClientadd = "checked"
		tmpAddr = rsClone("caddress") 
		tmpCity = rsClone("cCity") 
		tmpState = rsClone("cstate") 
		tmpZip = rsClone("czip")
		tmpCAdrI = rsClone("CliAdrI")
		tmpCFon = rsClone("Cphone")
		tmpCAFon = rsClone("CAphone")
		tmpDir = rsClone("directions")
		tmpSC = rsClone("spec_cir")
		tmpDOB = rsClone("DOB")
		tmpLang = rsClone("langID")
		tmpAppDate = rsClone("appDate")
		tmpAppTFrom = Z_FormatTime(rsClone("appTimeFrom"))
		tmpAppTTo = Z_FormatTime(rsClone("appTimeTo"))
		tmpAppLoc = rsClone("appLoc")
		tmpInst = rsClone("instID")
		tmpDept = rsClone("deptID")
		tmpInstRate = rsClone("InstRate")
		tmpDoc = rsClone("docNum")
		tmpCRN = rsClone("CrtRumNum")
		tmpIntr = rsClone("IntrID")
		tmpIntrRate = rsClone("IntrRate")
		tmpEmer = ""
		If rsClone("Emergency") =True Then tmpEmer = "checked"
		tmpEmerFee = ""
		If rsClone("EmerFee") =True Then tmpEmerFEE = "checked"
		tmpGender	= Z_CZero(rsClone("Gender"))
		tmpMale = ""
		tmpFemale = ""
		If tmpGender = 0 Then 
			tmpMale = "SELECTED"
		Else
			tmpFemale = "SELECTED"
		End If
		chkMinor = ""
		If rsClone("Child") Then chkMinor = "CHECKED"
		'tmpHPID = rsClone("HPID")
		Session("MSG") = "NOTE: Entries cloned from Request: " & Request("Clone")
	End If
	rsClone.CLose
	Set rsClone = Nothing
End If
'GET INTERPRETER LIST
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE Active = True ORDER BY [last name], [first name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	IntrSel = ""
	If CInt(tmpIntr) = rsIntr("index") Then IntrSel = "selected"
	strIntr = strIntr	& "<option " & IntrSel & " value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbCrlf
	tmpIntrName = CleanMe(rsIntr("last name")) & ", " & CleanMe(rsIntr("first name"))
	strIntr2 = strIntr2 & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsIntr("index") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
			"document.frmMain.selIntr.appendChild(ChoiceReq);}" & vbCrLf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
'GET INTERPRETER INFO
Set rsIntr2 = Server.CreateObject("ADODB.RecordSet")
sqlInst2 = "SELECT * FROM interpreter_T WHERE Active = True"
rsIntr2.Open sqlInst2, g_strCONN, 3, 1
Do Until rsIntr2.EOF
	CtrLang = 0
	If rsIntr2("Language1") <> "" Then CtrLang =  CtrLang + 1 
	If rsIntr2("Language2") <> "" Then CtrLang =  CtrLang + 1
	If rsIntr2("Language3") <> "" Then CtrLang =  CtrLang + 1
	If rsIntr2("Language4") <> "" Then CtrLang =  CtrLang + 1
	If rsIntr2("Language5") <> "" Then CtrLang =  CtrLang + 1
	strJScript2 = strJScript2 & "if (Intr == " & rsIntr2("Index") & ") " & vbCrLf & _
		"{document.frmMain.selIntr.value = """ & rsIntr2("Index") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrEmail.value = """ & rsIntr2("E-mail") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrP1.value = """ & rsIntr2("phone1") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrExt.value = """ & rsIntr2("P1Ext") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrP2.value = """ & rsIntr2("phone2") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrFax.value = """ & rsIntr2("fax") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrAddr.value = """ & rsIntr2("address1") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrCity.value = """ & rsIntr2("City") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrState.value = """ & rsIntr2("State") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrZip.value = """ & rsIntr2("Zip Code") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrAddrI.value = """ & rsIntr2("IntrAdrI") &"""; " & vbCrLf & _
		"document.frmMain.txtIntrRate.value = """ & rsIntr2("Rate") &"""; " & vbCrLf & _
		"document.frmMain.LangCtr.value = " & CtrLang &"; " & vbCrLf & _
		"document.frmMain.Lang1.value = GetLangID(""" & Trim(rsIntr2("Language1")) & """); " & vbCrLf & _
		"document.frmMain.Lang2.value = GetLangID(""" & Trim(rsIntr2("Language2")) &"""); " & vbCrLf & _
		"document.frmMain.Lang3.value = GetLangID(""" & Trim(rsIntr2("Language3")) &"""); " & vbCrLf & _
		"document.frmMain.Lang4.value = GetLangID(""" & Trim(rsIntr2("Language4")) &"""); " & vbCrLf & _
		"document.frmMain.Lang5.value = GetLangID(""" & Trim(rsIntr2("Language5")) &"""); " & vbCrLf 
		If rsIntr2("InHouse") = True Then 
			strJScript2 = strJScript2 & "document.frmMain.chkInHouse.checked = true; " & vbCrLf 
		Else
			strJScript2 = strJScript2 & "document.frmMain.chkInHouse.checked = false; " & vbCrLf 
		End If
		If rsIntr2("prime") = 0 Then
			strJScript2 = strJScript2 & "document.frmMain.radioPrim2[0].checked = true;" & vbCrLf 
		ElseIf rsIntr2("prime") = 1 Then
			strJScript2 = strJScript2 & "document.frmMain.radioPrim2[1].checked = true;" & vbCrLf 
		ElseIf rsIntr2("prime") = 2 Then
			strJScript2 = strJScript2 & "document.frmMain.radioPrim2[3].checked = true;" & vbCrLf 
		ElseIf rsIntr2("prime") = 3 Then
			strJScript2 = strJScript2 & "document.frmMain.radioPrim2[2].checked = true;" & vbCrLf 
		End If
		strJScript2 = strJScript2 & "}"
		rsIntr2.MoveNext
Loop
rsIntr2.Close
Set rsIntr2 = Nothing
'GET INTERPRETER INFO IF INACTIVE BUT ASSIGNED
If (tmpIntr <> "" Or tmpIntr <> "-1") And Request("ID") <> "" Then
	'CHECK IF INACTIVE
	If IsActive(tmpIntr) = False Then
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM interpreter_T WHERE index = " & tmpIntr
		rsIntr.Open sqlIntr, g_strCONN, 3, 1
		If Not rsIntr.EOF Then
			strIntr = strIntr	& "<option selected value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & "(INACTIVE)</option>" & vbCrlf
			tmpIntrName = CleanMe(rsIntr("last name")) & ", " & CleanMe(rsIntr("first name"))
			strIntr2 = strIntr2 & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
				"ChoiceReq.value = " & rsIntr("index") & ";" & vbCrLf & _
				"ChoiceReq.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
				"document.frmMain.selIntr.appendChild(ChoiceReq);}" & vbCrLf
			CtrLang = 0
			If rsIntr("Language1") <> "" Then CtrLang =  CtrLang + 1 
			If rsIntr("Language2") <> "" Then CtrLang =  CtrLang + 1
			If rsIntr("Language3") <> "" Then CtrLang =  CtrLang + 1
			If rsIntr("Language4") <> "" Then CtrLang =  CtrLang + 1
			If rsIntr("Language5") <> "" Then CtrLang =  CtrLang + 1
			strJScript2 = strJScript2 & "if (Intr == " & rsIntr("Index") & ") " & vbCrLf & _
				"{document.frmMain.selIntr.value = """ & rsIntr("Index") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrEmail.value = """ & rsIntr("E-mail") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrP1.value = """ & rsIntr("phone1") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrExt.value = """ & rsIntr("P1Ext") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrP2.value = """ & rsIntr("phone2") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrFax.value = """ & rsIntr("fax") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrAddr.value = """ & rsIntr("address1") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrCity.value = """ & rsIntr("City") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrState.value = """ & rsIntr("State") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrZip.value = """ & rsIntr("Zip Code") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrAddrI.value = """ & rsIntr("IntrAdrI") &"""; " & vbCrLf & _
				"document.frmMain.txtIntrRate.value = """ & rsIntr("Rate") &"""; " & vbCrLf & _
				"document.frmMain.LangCtr.value = " & CtrLang &"; " & vbCrLf & _
				"document.frmMain.Lang1.value = GetLangID(""" & Trim(rsIntr("Language1")) & """); " & vbCrLf & _
				"document.frmMain.Lang2.value = GetLangID(""" & Trim(rsIntr("Language2")) &"""); " & vbCrLf & _
				"document.frmMain.Lang3.value = GetLangID(""" & Trim(rsIntr("Language3")) &"""); " & vbCrLf & _
				"document.frmMain.Lang4.value = GetLangID(""" & Trim(rsIntr("Language4")) &"""); " & vbCrLf & _
				"document.frmMain.Lang5.value = GetLangID(""" & Trim(rsIntr("Language5")) &"""); " & vbCrLf 
			If rsIntr("InHouse") = True Then 
				strJScript2 = strJScript2 & "document.frmMain.chkInHouse.checked = true; " & vbCrLf 
			Else
				strJScript2 = strJScript2 & "document.frmMain.chkInHouse.checked = false; " & vbCrLf 
			End If
			If rsIntr("prime") = 0 Then
				strJScript2 = strJScript2 & "document.frmMain.radioPrim2[0].checked = true;" & vbCrLf 
			ElseIf rsIntr("prime") = 1 Then
				strJScript2 = strJScript2 & "document.frmMain.radioPrim2[1].checked = true;" & vbCrLf 
			ElseIf rsIntr("prime") = 2 Then
				strJScript2 = strJScript2 & "document.frmMain.radioPrim2[3].checked = true;" & vbCrLf 
			ElseIf rsIntr("prime") = 3 Then
				strJScript2 = strJScript2 & "document.frmMain.radioPrim2[2].checked = true;" & vbCrLf 
			End If
			strJScript2 = strJScript2 & "}"
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	End If
End If
'GET INSTITUTION LIST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T ORDER BY Facility"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	tmpDO = ""
	If Cint(tmpInst) = rsInst("index") Then tmpDO = "selected"
	InstName = rsInst("Facility")
	strInst = strInst	& "<option " & tmpDO & " value='" & rsInst("Index") & "'>" &  InstName & "</option>" & vbCrlf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'GET DEPT INFO
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T ORDER BY dept"
rsDept.Open sqlDept, g_strCONN, 3, 1
Do Until rsDept.EOF
	tmpOLDAddr = rsDept("address") & "|" & rsDept("city") & "|" & rsDept("state") & "|" & rsDept("zip")
	strDept = strDept & "if (dept == " & rsDept("Index") & ") " & vbCrLf & _
		"{document.frmMain.txtInstAddr.value = """ & rsDept("address") &"""; " & vbCrLf & _
		"document.frmMain.selDept.value = " & rsDept("Index") & "; " & vbCrLf & _
		"document.frmMain.txtInstCity.value = """ & rsDept("city") &"""; " & vbCrLf & _
		"document.frmMain.txtInstState.value = """ & rsDept("state") &"""; " & vbCrLf & _
		"document.frmMain.txtInstZip.value = """ & rsDept("zip") &"""; " & vbCrLf & _
		"document.frmMain.txtInstAddrI.value = """ & rsDept("InstAdrI") &"""; " & vbCrLf & _
		"document.frmMain.txtBlname.value = """ & rsDept("BLname") &"""; " & vbCrLf & _
		"document.frmMain.txtBilAddr.value = """ & rsDept("Baddress") &"""; " & vbCrLf & _
		"document.frmMain.txtBilCity.value = """ & rsDept("Bcity") &"""; " & vbCrLf & _
		"document.frmMain.txtBilState.value = """ & rsDept("Bstate") &"""; " & vbCrLf & _
		"document.frmMain.txtBilZip.value = """ & rsDept("Bzip") &"""; " & vbCrLf & _
		"document.frmMain.OldAddr.value = """ & tmpOLDAddr &"""; " & vbCrLf & _
		"document.frmMain.selClass.value = """ & rsDept("Class") &"""; }" & vbCrLf 
	rsDept.MoveNext
Loop
rsDept.Close
Set rsDept = Nothing
'GET AVAILABLE DEPARTMENTS
Set rsInstDept = Server.CreateObject("ADODB.RecordSet")
sqlInstDept = "SELECT * FROM institution_T ORDER BY Facility"
rsInstDept.Open sqlInstDept, g_strCONN, 3, 1
Do Until rsInstDept.EOF
	InstDept = rsInstDept("Index")
	strInstDept = strInstDept & "if (inst == " & InstDept & "){" & vbCrLf
	Set rsDeptInst = Server.CreateObject("ADODB.RecordSet")
	sqlDeptInst = "SELECT * FROM dept_T WHERE InstID = " &  InstDept & " ORDER BY Dept"
	rsDeptInst.Open sqlDeptInst, g_strCONN, 3, 1
	If Not rsDeptInst.EOF Then
		Do Until rsDeptInst.EOF
			strInstDept = strInstDept & "if (dept != " & rsDeptInst("index") & ")" & vbCrLf & _
				"{var ChoiceInst = document.createElement('option');" & vbCrLf & _
				"ChoiceInst.value = " & rsDeptInst("index") & ";" & vbCrLf & _
				"ChoiceInst.appendChild(document.createTextNode(""" & rsDeptInst("Dept") & """));" & vbCrLf & _
				"document.frmMain.selDept.appendChild(ChoiceInst);} " & vbCrlf
			rsDeptInst.MoveNext
		Loop
	End If
	rsDeptInst.Close
	Set rsDeptInst = Nothing
	rsInstDept.MoveNext
	strInstDept = strInstDept & "}"
Loop
rsInstDept.Close
Set rsInstDept = Nothing
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT * FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpL = ""
	If tmpLang = "" Then tmpLang = -1
	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
	strLangChk = strLangChk & "if (xxx == """ & Trim(rsLang("Language")) & """){ " & vbCrLf & _
		"return " & rsLang("index") & ";}"
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'GET REQUESTING PERSON LIST
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T ORDER BY Lname, Fname"
rsReq.Open sqlReq, g_strCONN, 3, 1
Do Until rsReq.EOF
	ReqSel = ""
	If tmpReqP = "" Then tmpReqP = -1
	If CInt(tmpReqP) = rsReq("index") Then ReqSel = "selected"
	tmpReqName = CleanMe(rsReq("lname")) & ", " & CleanMe(rsReq("fname"))
	strReq2 = strReq2 & "<option " & ReqSel & " value='" & rsReq("Index") & "'>" & rsReq("Lname") & ", " & rsReq("Fname") & "</option>" & vbCrLf
	strReq = strReq & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsReq("index") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpReqName & """));" & vbCrLf & _
			"document.frmMain.selReq.appendChild(ChoiceReq);}" & vbCrLf
	rsReq.MoveNext
Loop
rsReq.Close
Set rsReq = Nothing
'GET REQUESTING PERSON INFO
Set rsReqI = Server.CreateObject("ADODB.RecordSet")
sqlReqI = "SELECT * FROM requester_T ORDER BY Lname, Fname"
rsReqI.Open sqlReqI, g_strCONN, 3, 1
Do Until rsReqI.EOF
	strJScript3 = strJScript3 & "if (Req == " & rsReqI("Index") & ") " & vbCrLf & _
		"{document.frmMain.txtphone.value = """ & rsReqI("Phone") &"""; " & vbCrLf & _
		"document.frmMain.selReq.value = " & rsReqI("Index") & "; " & vbCrLf & _
		"document.frmMain.txtReqExt.value = """ & rsReqI("pExt") &"""; " & vbCrLf & _
		"document.frmMain.txtfax.value = """ & rsReqI("Fax") &"""; " & vbCrLf & _
		"document.frmMain.txtemail.value = """ & rsReqI("Email") &"""; " & vbCrLf
		If rsReqI("prime") = 0 Then
			strJScript3 = strJScript3 & "document.frmMain.radioPrim1[2].checked = true;" & vbCrLf 
		ElseIf rsReqI("prime") = 1 Then
			strJScript3 = strJScript3 & "document.frmMain.radioPrim1[0].checked = true;" & vbCrLf 
		ElseIf rsReqI("prime") = 2 Then
			strJScript3 = strJScript3 & "document.frmMain.radioPrim1[1].checked = true;" & vbCrLf 
		End If
		strJScript3 = strJScript3 & "}"
	rsReqI.MoveNext
Loop
rsReqI.Close
Set rsReqI = Nothing
'REQUESTING PERSON CHECKER
Set rsReqCHK = Server.CreateObject("ADODB.RecordSet")
sqlReqCHK = "SELECT * FROM requester_T"
rsReqCHK.Open sqlReqCHK, g_strCONN, 3, 1
Do Until rsReqCHK.EOF
	strReqCHK = strReqCHK & "if (document.frmMain.txtReqLname.value == """ & Trim(rsReqCHK("lname")) & """ && document.frmMain.txtReqFname.value == """ & Trim(rsReqCHK("Fname")) & """) " & vbCrLf & _
		"{var ans = window.confirm(""Requester's name already exists. Click on Cancel to rename. Click on OK to continue.""); " & vbCrLf & _
		"{if (ans){ " & vbCrLf & _
		"pnt = 1; " & vbCrLf & _
		"} " & vbCrLf & _
		"else " & vbCrLf & _
		"{ " & vbCrLf & _
		"return; " & vbCrLf & _
		"} " & vbCrLf & _
		"} " & vbCrLf & _
		"} " & vbCrLf & _
		"else " & vbCrLf & _
		"{pnt = 1; " & vbCrLf & _
		"} " & vbCrLf
	rsReqCHK.MoveNext 
Loop
rsReqCHK.Close
Set rsReqCHK = Nothing
'INTERPRETER CHECKER ACTIVE
Set rsIntrCHK = Server.CreateObject("ADODB.RecordSet")
sqlIntrCHK = "SELECT * FROM interpreter_T WHERE Active = true"
rsIntrCHK.Open sqlIntrCHK, g_strCONN, 3, 1
Do Until rsIntrCHK.EOF
	strIntrCHK = strIntrCHK & "if (document.frmMain.txtIntrLname.value == """ & Trim(rsIntrCHK("last name")) & """ && document.frmMain.txtIntrFname.value == """ & Trim(rsIntrCHK("First name")) & """) " & vbCrLf & _
		"{var ans = window.confirm(""Interpreter's name already exists. Click on Cancel to rename. Click on OK to continue.""); " & vbCrLf & _
		"{if (ans){ " & vbCrLf & _
		"pnt = 1; " & vbCrLf & _
		"} " & vbCrLf & _
		"else " & vbCrLf & _
		"{ " & vbCrLf & _
		"return; " & vbCrLf & _
		"} " & vbCrLf & _
		"} " & vbCrLf & _
		"} " & vbCrLf & _
		"else " & vbCrLf & _
		"{pnt = 1; " & vbCrLf & _
		"} " & vbCrLf
	rsIntrCHK.MoveNext 
Loop
rsIntrCHK.Close
Set rsIntrCHK = Nothing
'INTERPRETER CHECKER INACTIVE
Set rsIntrCHK = Server.CreateObject("ADODB.RecordSet")
sqlIntrCHK = "SELECT * FROM interpreter_T WHERE Active = false"
rsIntrCHK.Open sqlIntrCHK, g_strCONN, 3, 1
Do Until rsIntrCHK.EOF
	strIntrCHK = strIntrCHK & "if (document.frmMain.txtIntrLname.value == """ & Trim(rsIntrCHK("last name")) & """ && document.frmMain.txtIntrFname.value == """ & Trim(rsIntrCHK("First name")) & """) " & vbCrLf & _
		"{alert(""Interpreter's name already exists but is currently inactive.\nPlease contact your system adminitrator.""); " & vbCrLf & _
		"return; " & vbCrLf & _
		"} " & vbCrLf & _
		"else " & vbCrLf & _
		"{pnt = 1; " & vbCrLf & _
		"} " & vbCrLf
	rsIntrCHK.MoveNext 
Loop
rsIntrCHK.Close
Set rsIntrCHK = Nothing

'GET INSTITUTION RATES
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM rate_T ORDER BY Rate"
rsRate.Open sqlRate, g_strCONN, 3, 1
Do Until rsRate.EOF
	RateKo = ""
	If tmpInstRate = rsRate("Rate") Then Rateko = "selected"
	strRate1 = strRate1 & "<option " & Rateko & " value='" & rsRate("Rate") & "'>$" & Z_FormatNumber(rsRate("Rate"), 2) & "</option>" & vbCrLf
	rsRate.MoveNext
Loop
rsRate.Close
Set rsRate = Nothing
'GET INTERPRETER RATES
Set rsRate2 = Server.CreateObject("ADODB.RecordSet")
sqlRate2 = "SELECT * FROM rate2_T ORDER BY Rate2"
rsRate2.Open sqlRate2, g_strCONN, 3, 1
Do Until rsRate2.EOF
	RateKo2 = ""
	strRate2 = strRate2 & "<option value='" & rsRate2("Rate2") & "'>$" & Z_FormatNumber(rsRate2("Rate2"), 2) & "</option>" & vbCrLf
	rsRate2.MoveNext
Loop
rsRate2.Close
Set rsRate2 = Nothing
'GET ALLOWED INTERPRETER
Set rsLangIntr = Server.CreateObject("ADODB.RecordSet")
sqlLangIntr = "SELECT * FROM language_T ORDER BY [Language]"
rsLangIntr.Open sqlLangIntr, g_strCONN, 3, 1
Do Until rsLangIntr.EOF
	IntrLang = UCase(rsLangIntr("Language"))
	strIntrLang = strIntrLang & "if (dialect == " & rsLangIntr("index") & "){" & vbCrLf
	Set rsIntrLang = Server.CreateObject("ADODB.RecordSet")
	sqlIntrLang = "SELECT * FROM interpreter_T WHERE (UCase(Language1) = '" & IntrLang & "' OR UCase(Language2) = '" & IntrLang & "' OR UCase(Language3) = '" & IntrLang & _
		"' OR UCase(Language4) = '" & IntrLang & "' OR UCase(Language5) = '" & IntrLang & "') AND ACtive = true ORDER BY [Last Name], [First Name]" 
	rsIntrLang.Open sqlIntrLang, g_strCONN, 3, 1
	Do Until rsIntrLang.EOF
		tmpIntrName = CleanMe(rsIntrLang("last name")) & ", " & CleanMe(rsIntrLang("first name"))
		strIntrLang = strIntrLang	& "if(intr != "& rsIntrLang("index") & ")" & vbCrLf & _
			"{var ChoiceIntr = document.createElement('option');" & vbCrLf & _
			"ChoiceIntr.value = " & rsIntrLang("index") & ";" & vbCrLf & _
			"ChoiceIntr.appendChild(document.createTextNode(""" & tmpIntrName & """));" & vbCrLf & _
			"document.frmMain.selIntr.appendChild(ChoiceIntr);}" & vbCrLf
		rsIntrLang.MoveNext
	Loop
	rsIntrLang.Close
	Set rsIntrLang = Nothing
	rsLangIntr.MoveNext
	strIntrLang = strIntrLang & "}"
Loop
rsLangIntr.Close
Set rsLangIntr = Nothing
'GET AVAILABLE REQUESTING PERSON PER DEPARTMENT
Set rsInstReq = Server.CreateObject("ADODB.RecordSet")
sqlInstReq = "SELECT * FROM dept_T ORDER BY dept"
rsInstReq.Open sqlInstReq, g_strCONN, 3, 1
Do Until rsInstReq.EOF
	InstReq = rsInstReq("index")
	strInstReqDept = strInstReqDept & "if (dept == " & InstReq & "){" & vbCrLf
	Set rsReqInst = Server.CreateObject("ADODB.RecordSet")
	sqlReqInst = "SELECT * FROM requester_T, reqdept_T WHERE  ReqID = requester_T.index AND DeptID = " & InstReq & " ORDER BY lname, fname"
	rsReqInst.Open sqlReqInst, g_strCONN, 3, 1
	Do Until rsReqInst.EOF
		tmpReqName = CleanMe(rsReqInst("lname")) & ", " & CleanMe(rsReqInst("fname"))
		strInstReqDept = strInstReqDept	& "if(req != "& rsReqInst("requester_T.index") & ")" & vbCrLf & _
			"{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsReqInst("requester_T.index") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpReqName & """));" & vbCrLf & _
			"document.frmMain.selReq.appendChild(ChoiceReq);}" & vbCrLf
		rsReqInst.MoveNext
	Loop
	rsReqInst.Close
	Set rsReqInst = Nothing
	rsInstReq.MoveNext
	strInstReqDept = strInstReqDept & "}"
Loop
rsInstReq.Close
Set rsLangIntr = Nothing
'GET DEPARTMENTS
Set rsDept2 = Server.CreateObject("ADODB.RecordSet")
sqlDept2 = "SELECT * FROM dept_T ORDER BY Dept"
rsDept2.Open sqlDept2, g_strCONN, 3, 1
Do Until rsDept2.EOF
	tmpDpt = ""
	If Cint(tmpDept) = rsDept2("index") Then tmpDpt = "selected"
	DeptName = rsDept2("Dept")
	'If rsInst("Department") <> "" Then InstName = rsInst("Facility") & " - " & rsInst("Department")
	strDept2 = strDept2	& "<option " & tmpDpt & " value='" & rsDept2("Index") & "'>" &  DeptName & "</option>" & vbCrlf
	rsDept2.MoveNext
Loop
rsDept2.Close
Set rsDept2 = Nothing
%>
<html>
	<head>
		<title>Language Bank - Interpreter Request Form</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
		<!--
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
		}
		function chkPrim()
		{
			if (document.frmMain.radioPrim1[0].checked == true)
			{
				document.frmMain.txtPRim1.value = "Phone";
			}
			if (document.frmMain.radioPrim1[1].checked == true)
			{
				document.frmMain.txtPRim1.value = "Fax";
			}
			if (document.frmMain.radioPrim1[2].checked == true)
			{
				document.frmMain.txtPRim1.value = "E-Mail";
			}
		}
		function chkPrim2()
		{
			if (document.frmMain.radioPrim2[0].checked == true)
			{
				document.frmMain.txtPRim2.value = "E-Mail";
			}
			if (document.frmMain.radioPrim2[1].checked == true)
			{
				document.frmMain.txtPRim2.value = "Home Phone";
			}
			if (document.frmMain.radioPrim2[2].checked == true)
			{
				document.frmMain.txtPRim2.value = "Fax";
			}
			if (document.frmMain.radioPrim2[3].checked == true)
			{
				document.frmMain.txtPRim2.value = "Mobile Phone";
			}
		}
		function ReqChkMe()
		{
			var pnt = 0;
			
			//CHECK IF ADDRESS HAS BEEN CHANGED
			var strNewAddr = document.frmMain.txtInstAddr.value + "|" + document.frmMain.txtInstCity.value + "|" + document.frmMain.txtInstState.value + "|" + document.frmMain.txtInstZip.value;
			if (strNewAddr != document.frmMain.OldAddr.value)
			{
				var ans = window.confirm("WARNING: Changing of institution address will be effective for all instances of that institution. Click Cancel to stop.");
				if (!ans)
				{
					return;
				}
			}
			if (document.frmMain.radioPrim1[2].checked == true && document.frmMain.txtemail.value == "")
			{
				alert("ERROR: Please supply an E-mail address to requesting person."); 
				document.frmMain.txtemail.focus();
				return;
			}
			if (document.frmMain.radioPrim1[0].checked == true && document.frmMain.txtphone.value == "")
			{
				alert("ERROR: Please supply a Phone Number to requesting person."); 
				document.frmMain.txtphone.focus();
				return;
			}
			if (document.frmMain.radioPrim1[1].checked == true && document.frmMain.txtfax.value == "")
			{
				alert("ERROR: Please supply a Fax Number to requesting person."); 
				document.frmMain.txtfax.focus();
				return;
			}
			//CHECK VALID FAX
			if (document.frmMain.radioPrim1[1].checked == true && document.frmMain.txtfax.value != "")
			{
				var tmpFax =  document.frmMain.txtfax.value
				tmpFax = tmpFax.replace("-", "")
				if (tmpFax.length < 10) 
				{
					alert("ERROR: Please include area code in Fax Number to requesting person."); 
					document.frmMain.txtfax.focus();
					return;
				}
			}
			if (document.frmMain.chkClientAdd.checked == true)
			{
				if (document.frmMain.txtCliAdd.value == "" || document.frmMain.txtCliCity.value == "" || document.frmMain.txtCliState.value == "" || document.frmMain.txtCliZip.value == "")
				{
					alert("ERROR: Please include full address of client."); 
					document.frmMain.txtCliAdd.focus();
					return;
				}
			}
			if (document.frmMain.selIntr.value !== "-1")
			{
				if (document.frmMain.radioPrim2[0].checked == true && document.frmMain.txtIntrEmail.value == "")
				{
					alert("ERROR: Please supply an E-mail address to interpreter."); 
					document.frmMain.txtIntrEmail.focus();
					return;
				}
				if (document.frmMain.radioPrim2[1].checked == true && document.frmMain.txtIntrP1.value == "")
				{
					alert("ERROR: Please supply a Home Number to interpreter."); 
					document.frmMain.txtIntrP1.focus();
					return;
				}
				if (document.frmMain.radioPrim2[3].checked == true && document.frmMain.txtIntrP2.value == "")
				{
					alert("ERROR: Please supply a Mobile Number to interpreter."); 
					document.frmMain.txtIntrP2.focus();
					return;
				}
				if (document.frmMain.radioPrim2[2].checked == true && document.frmMain.txtIntrFax.value == "")
				{
					alert("ERROR: Please supply a Fax Number to interpreter."); 
					document.frmMain.txtIntrFax.focus();
					return;
				}
				//CHECK VALID FAX
				if (document.frmMain.radioPrim2[2].checked == true && document.frmMain.txtIntrFax.value != "")
				{
					var tmpFax =  document.frmMain.txtIntrFax.value
					tmpFax = tmpFax.replace("-", "")
					if (tmpFax.length < 10) 
					{
						alert("ERROR: Please include area code in Fax Number to interpreter."); 
						document.frmMain.txtIntrFax.focus();
						return;
					}
				}
			}
			<%=strReqCHK%>
			<%=strIntrCHK%>
			//CHECK IF INTERPRETER ALLOWED
			if (document.frmMain.chkAll2.checked == true)
			{
				var LangOK = 0;
				if(document.frmMain.selLang.value == document.frmMain.Lang1.value){LangOK = 1;}
				if(document.frmMain.selLang.value == document.frmMain.Lang2.value){LangOK = 1;}
				if(document.frmMain.selLang.value == document.frmMain.Lang3.value){LangOK = 1;}
				if(document.frmMain.selLang.value == document.frmMain.Lang4.value){LangOK = 1;}
				if(document.frmMain.selLang.value == document.frmMain.Lang5.value){LangOK = 1;}
				if (LangOK != 1)
				{
					if (document.frmMain.LangCtr.value < 5)
					{
						pnt = 1;
					}
					else
					{
						alert("ERROR: Cannot add this language to interpreter.");
						return;
					}
				}
				else
				{
					pnt = 1;
				}
			}
			if (pnt = 1)
			{
				SavReq(<%=Request("ID")%>);
			} 
		}
		function SavReq(zzz)
		{
			var zzz;
			if (zzz == undefined)
				{	
					document.frmMain.action = "action.asp?ctrl=1";
					document.frmMain.submit();
				}
			else
				{
					document.frmMain.action = "action.asp?ctrl=2";
					document.frmMain.submit();
				}
		}
		function printMe(xxx)
		{
			document.frmMain.action = "action.asp?ctrl=2&Print='Yes'&PID=" + xxx;
			document.frmMain.submit();
		}
		function textboxchangeReq() 
		{
			if (document.frmMain.btnNewReq.value == 'NEW')
			{
				alert("To save a new Requesting Person, complete the form and click 'Save Request' button.");
				document.frmMain.btnNewReq.value = 'BACK';
				document.frmMain.selReq.disabled = true;
				document.frmMain.txtReqLname.style.visibility = 'visible';
				document.frmMain.txtReqFname.style.visibility = 'visible';
				document.frmMain.txtcoma2.style.visibility = 'visible';
				document.frmMain.txtformat2.style.visibility = 'visible';
				document.frmMain.txtemail.value = "";
				document.frmMain.txtphone.value = "";
				document.frmMain.txtReqExt.value = "";
				document.frmMain.txtfax.value = "";
				document.frmMain.txtReqLname.focus();
				document.frmMain.HnewReq.value = 'BACK';
			}
			else
			{
				document.frmMain.btnNewReq.value = 'NEW';
				document.frmMain.selReq.disabled = false;
				document.frmMain.txtReqLname.value = "";
				document.frmMain.txtReqFname.value = "";
				document.frmMain.txtReqLname.style.visibility = 'hidden';
				document.frmMain.txtReqFname.style.visibility = 'hidden';
				document.frmMain.txtcoma2.style.visibility = 'hidden';
				document.frmMain.txtformat2.style.visibility = 'hidden';
				ReqInfo(document.frmMain.selReq.value);
				document.frmMain.HnewReq.value = 'NEW';
			}
		}
		function hideNewReq() 
		{
			if (document.frmMain.txtReqLname.value == "" && document.frmMain.txtReqFname.value == "")
			{	
				document.frmMain.txtReqLname.style.visibility = 'hidden';
				document.frmMain.txtReqFname.style.visibility = 'hidden';
				document.frmMain.txtcoma2.style.visibility = 'hidden';
				document.frmMain.txtformat2.style.visibility = 'hidden';
				document.frmMain.btnNewReq.value = 'NEW';
				document.frmMain.txtReqLname.value = "";
				document.frmMain.txtReqFname.value = "";
				document.frmMain.HnewReq.value = 'NEW';
			}
			else
			{
				document.frmMain.txtReqLname.style.visibility = 'visible';
				document.frmMain.txtReqFname.style.visibility = 'visible';
				document.frmMain.txtcoma2.style.visibility = 'visible';
				document.frmMain.txtformat2.style.visibility = 'visible';
				document.frmMain.btnNewReq.value = 'BACK';
				document.frmMain.selReq.disabled = true;
				document.frmMain.txtReqLname.value = '<%=tmpNewReqLN%>';
				document.frmMain.txtReqFname.value = '<%=tmpNewReqFN%>';
				document.frmMain.txtemail.value = '<%=tmpNewReqeMail%>';
				document.frmMain.txtReqExt.value = '<%=tmpReqExt%>';
				document.frmMain.txtphone.value = '<%=tmpNewReqPhone%>';
				document.frmMain.txtfax.value = '<%=tmpNewReqFax%>';
				document.frmMain.HnewReq.value = 'BACK';
			}
		}
		function ReqInfo(Req)
		{
			if (Req == " -1")
			{
				if  (document.frmMain.txtReqLname.value == "" || document.frmMain.txtReqFname.value == "")
					{
						hideNewReq();
					}
					else
					{document.frmMain.txtphone.value = ""; 
					document.frmMain.txtReqExt.value = ""; 
					document.frmMain.radioPrim1[1].checked = true;
					document.frmMain.txtfax.value = ""; 
					document.frmMain.txtemail.value = ""; }
			}
			<%=strJScript3%>
			chkPrim();
		}
		function textboxchangeInst() 
		{
			if (document.frmMain.btnNew.value == 'NEW')
			{
				alert("To save a new Institution, complete the form and click 'Save Request' button.");
				document.frmMain.btnNew.value = 'BACK';
				document.frmMain.selInst.disabled = true;
				document.frmMain.txtNewInst.style.visibility = 'visible';
				document.frmMain.txtInstDept.value = "";
				document.frmMain.txtInstAddr.value = "";
				document.frmMain.txtInstCity.value = "";
				document.frmMain.txtInstState.value = "";
				document.frmMain.txtInstZip.value = "";
				document.frmMain.txtInstAddrI.value = "";
				document.frmMain.txtBlname.value = ""; 
				document.frmMain.txtBilAddr.value = ""; 
				document.frmMain.txtBilCity.value = ""; 
				document.frmMain.txtBilState.value = ""; 
				document.frmMain.txtBilZip.value = ""; 
				document.frmMain.selClass.value = "1";
				document.frmMain.txtNewInst.focus();
				document.frmMain.HnewInt.value = 'BACK';
				DeptChoice();
			}
			else
			{
				document.frmMain.btnNew.value = 'NEW';
				document.frmMain.selInst.disabled = false;
				document.frmMain.txtNewInst.value = "";
				document.frmMain.txtNewInst.style.visibility = 'hidden';
				document.frmMain.HnewInt.value = 'NEW';
				DeptChoice(document.frmMain.selInst.value);
			}
		}
		function hideNewInts() 
		{
			if (document.frmMain.txtNewInst.value == "")
			{	
				document.frmMain.txtNewInst.style.visibility = 'hidden';
				document.frmMain.btnNew.value = 'NEW';
				document.frmMain.txtNewInst.value = "";
				document.frmMain.HnewInt.value = 'NEW';
			}
			else
			{
				document.frmMain.txtNewInst.style.visibility = 'visible';
				document.frmMain.btnNew.value = 'BACK';
				document.frmMain.selInst.disabled = true;
				document.frmMain.txtInstDept.value = '<%=tmpNewInstDept%>';
				document.frmMain.txtInstAddr.value = '<%=tmpNewInstAddr%>';
				document.frmMain.txtInstCity.value = '<%=tmpNewInstCity%>';
				document.frmMain.txtInstState.value = '<%=tmpNewInstState%>';
				document.frmMain.txtInstZip.value = '<%=tmpNewInstZip%>';
				document.frmMain.txtInstAddrI.value = '<%=tmpNewInstAddrI%>';
				document.frmMain.HnewInt.value = 'BACK';
			}
		}
		function InstInfo(Inst)
		{
			<%=strJScript%>
			if (Inst == -1)
			{
				document.frmMain.selInst.value = -1; 
				document.frmMain.txtInstDept.value = ""; 
				document.frmMain.txtInstAddr.value = ""; 
				document.frmMain.txtInstCity.value = ""; 
				document.frmMain.txtInstState.value = ""; 
				document.frmMain.txtInstZip.value = ""; 
				document.frmMain.txtInstAddrI.value = "";
				document.frmMain.txtBlname.value = ""; 
				document.frmMain.txtBilAddr.value = ""; 
				document.frmMain.txtBilCity.value = ""; 
				document.frmMain.txtBilState.value = ""; 
				document.frmMain.txtBilZip.value = ""; 
				document.frmMain.selClass.value = "1";
			}
		}
	
		function textboxchangeIntr() 
		{
			if (document.frmMain.btnNewIntr.value == 'NEW')
			{
				alert("To save a new Interpreter, complete the form and click 'Save Request' button.");
				document.frmMain.btnNewIntr.value = 'BACK';
				document.frmMain.selIntr.disabled = true;
				document.frmMain.txtIntrLname.style.visibility = 'visible';
				document.frmMain.txtIntrFname.style.visibility = 'visible';
				document.frmMain.txtcoma.style.visibility = 'visible';
				document.frmMain.txtformat.style.visibility = 'visible';
				document.frmMain.selIntrRate.style.visibility = 'visible';
				document.frmMain.txtIntrRate.value = "";
				document.frmMain.txtIntrEmail.value = "";
				document.frmMain.txtIntrP1.value = "";
				document.frmMain.txtIntrExt.value = "";
				document.frmMain.txtIntrFax.value = "";
				document.frmMain.txtIntrP2.value = "";
				document.frmMain.txtIntrAddr.value = "";
				document.frmMain.txtIntrCity.value = "";
				document.frmMain.txtIntrState.value = "";
				document.frmMain.txtIntrZip.value = "";
				document.frmMain.txtIntrAddrI.value = "";
				document.frmMain.LangCtr.value = 0;
				document.frmMain.Lang1.value = "";
				document.frmMain.Lang2.value = "";
				document.frmMain.Lang3.value = "";
				document.frmMain.Lang4.value = "";
				document.frmMain.Lang5.value = "";
				document.frmMain.txtIntrLname.focus();
				document.frmMain.HnewIntr.value = 'BACK';
			}
			else
			{
				document.frmMain.btnNewIntr.value = 'NEW';
				document.frmMain.selIntr.disabled = false;
				document.frmMain.txtNewInst.value = "";
				document.frmMain.selIntrRate.value = 0;
				document.frmMain.txtIntrLname.style.visibility = 'hidden';
				document.frmMain.txtIntrFname.style.visibility = 'hidden';
				document.frmMain.txtcoma.style.visibility = 'hidden';
				document.frmMain.txtformat.style.visibility = 'hidden';
				document.frmMain.selIntrRate.style.visibility = 'hidden';
				IntrInfo(document.frmMain.selIntr.value);
				document.frmMain.HnewIntr.value = 'NEW';
			}
		}
		function hideNewIntr() 
		{
			if (document.frmMain.txtIntrLname.value == "" && document.frmMain.txtIntrFname.value == "")
			{	
				document.frmMain.txtIntrLname.style.visibility = 'hidden';
				document.frmMain.txtIntrFname.style.visibility = 'hidden';
				document.frmMain.txtcoma.style.visibility = 'hidden';
				document.frmMain.txtformat.style.visibility = 'hidden';
				document.frmMain.btnNewIntr.value = 'NEW';
				document.frmMain.txtIntrLname.value = "";
				document.frmMain.txtIntrFname.value = "";
				document.frmMain.selIntrRate.style.visibility = 'hidden';
				document.frmMain.HnewIntr.value = 'NEW';
			}
			else
			{
				document.frmMain.txtIntrLname.style.visibility = 'visible';
				document.frmMain.txtIntrFname.style.visibility = 'visible';
				document.frmMain.txtcoma.style.visibility = 'visible';
				document.frmMain.txtformat.style.visibility = 'visible';
				document.frmMain.selIntrRate.style.visibility = 'visible';
				document.frmMain.btnNewIntr.value = 'BACK';
				document.frmMain.selIntr.disabled = true;
				document.frmMain.txtIntrLname.value = '<%=tmpIntrLname%>';
				document.frmMain.txtIntrFname.value = '<%=tmpIntrFname%>';
				document.frmMain.txtIntrEmail.value = '<%=tmpIntrEmail%>';
				document.frmMain.txtIntrP1.value = '<%=tmpIntrP1%>';
				document.frmMain.txtIntrExt.value = '<%=tmpIntrExt%>';
				document.frmMain.txtIntrFax.value = '<%=tmpIntrFax%>';
				document.frmMain.txtIntrP2.value = '<%=tmpIntrP2%>';
				document.frmMain.txtIntrAddr.value = '<%=tmpIntrAddr%>';
				document.frmMain.txtIntrCity.value = '<%=tmpIntrCity%>';
				document.frmMain.txtIntrState.value = '<%=tmpIntrState%>';
				document.frmMain.txtIntrZip.value = '<%=tmpIntrZip%>';
				document.frmMain.txtIntrAddrI.value = '<%=tmpIntrAddrI%>';
				document.frmMain.selIntrRate.value = '<%=tmpIntrRate%>';
				document.frmMain.LangCtr.value = 0;
				document.frmMain.Lang1.value = "";
				document.frmMain.Lang2.value = "";
				document.frmMain.Lang3.value = "";
				document.frmMain.Lang4.value = "";
				document.frmMain.Lang5.value = "";
				document.frmMain.HnewIntr.value = 'BACK';
			}
		}
		function IntrInfo(Intr)
		{	
			if (Intr == -1)
			{
				document.frmMain.selIntr.value = -1;
				document.frmMain.txtIntrEmail.value = ""; 
				document.frmMain.txtIntrP1.value = ""; 
				document.frmMain.txtIntrExt.value = ""; 
				document.frmMain.txtIntrP2.value = ""; 
				document.frmMain.txtIntrFax.value = ""; 
				document.frmMain.txtIntrAddr.value = ""; 
				document.frmMain.txtIntrCity.value = ""; 
				document.frmMain.txtIntrState.value = ""; 
				document.frmMain.txtIntrZip.value = ""; 
				document.frmMain.txtIntrAddrI.value = ""; 
				document.frmMain.chkInHouse.checked = false; 
				document.frmMain.txtIntrRate.value = 0;
				document.frmMain.radioPrim2[2].checked = true;
				document.frmMain.LangCtr.value = 0;
				document.frmMain.Lang1.value = "";
				document.frmMain.Lang2.value = "";
				document.frmMain.Lang3.value = "";
				document.frmMain.Lang4.value = "";
				document.frmMain.Lang5.value = "";
			}
			<%=strJScript2%>
			chkPrim2();
		}
		function KillMe(xxx)
		{
			var ans = window.confirm("Delete Request? Click Cancel to stop.");
			if (ans){
				document.frmMain.action = "action.asp?ctrl=9&ReqID=" + xxx;
				document.frmMain.submit();
			}
		}
		function CancelMe()
		{
			document.frmMain.selCancel.value = 0;
			document.frmMain.selCancel.disabled = true;
			if (document.frmMain.radioStat[3].checked == true || document.frmMain.radioStat[4].checked == true)
			{
				document.frmMain.selCancel.disabled = false;
			}
		}
		function CancelReason(xxx)
		{
			if (xxx !== 0)
			{
				document.frmMain.selCancel.value = xxx;
			}
		}
		function MissedMe()
		{
			document.frmMain.selMissed.value = 0;
			document.frmMain.selMissed.disabled = true;
			if (document.frmMain.radioStat[2].checked == true)
			{
				document.frmMain.selMissed.disabled = false;
			}
		}
		function MissedReason(xxx)
		{
			if (xxx !== 0)
			{
				document.frmMain.selMissed.value = xxx;
			}
		}
		function CompleteMe()
		{
			if (document.frmMain.radioStat[1].checked == true || document.frmMain.radioStat[4].checked == true)
			{
				document.frmMain.radioStat[0].disabled = true;
				document.frmMain.radioStat[2].disabled = true;
				document.frmMain.radioStat[3].disabled = true;
			}
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
		function PopMe()
		{
			newwindow = window.open('find.asp','name','height=150,width=400,scrollbars=0,directories=0,status=0,toolbar=0,resizable=0');
			if (window.focus) {newwindow.focus()}
		}
		function PopMe2(tmpDate, tmpTime, tmpLang)
		{
			if (tmpDate == "" || tmpTime == "" || tmpLang == "")
			{
				alert("Error: Appointment Date/Time/Language is Required.")
				return;
			}
			else
			{
				var zDate = tmpDate; 
				var zTime = tmpTime;
				var zLang = tmpLang
				var zDT = zDate + "|" + zTime + "|" + zLang
				//alert(zipus);
				newwindow2 = window.open('IntrAvail.asp?AppInfo=' + zDT,'name','width=400,scrollbars=0,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow2.focus()}
			}
		}
		function LangChoice(dialect, intr)
		{
			 var i;
			for(i=document.frmMain.selIntr.options.length-1;i>=1;i--)
			{
				if (intr != "undefined")
				{
					if (document.frmMain.selIntr.options[i].value != intr)
					{
						document.frmMain.selIntr.remove(i);
					}
				}
				else
				{
					document.frmMain.selIntr.remove(i);
				}
			}
			<%=strIntrLang%>
		}
		function DeptChoice(inst, dept)
		{
			var i;
			for(i=document.frmMain.selDept.options.length-1;i>=1;i--)
			{
				if (dept != "undefined")
				{
					if (document.frmMain.selDept.options[i].value != dept)
					{
						document.frmMain.selDept.remove(i);
					}
				}
				else
				{
					document.frmMain.selReq.remove(i);
				}
			}
			<%=strInstDept%>
		}
		function DeptInfo(dept)
		{
			if (dept == 0 && document.frmMain.txtInstDept.value == "" )
			{
				document.frmMain.selDept.value =0;
				document.frmMain.txtInstDept.value = "";
				document.frmMain.txtInstAddr.value = "";
				document.frmMain.txtInstCity.value = "";
				document.frmMain.txtInstState.value = "";
				document.frmMain.txtInstZip.value = "";
				document.frmMain.txtInstAddrI.value = "";
				document.frmMain.txtBlname.value = "";
				document.frmMain.txtBilAddr.value = "";
				document.frmMain.txtBilCity.value = "";
				document.frmMain.txtBilState.value = "";
				document.frmMain.txtBilZip.value = "";
				document.frmMain.OldAddr.value = "";
			}
			else
			{
				hideNewDept();
			}
			<%=strDept%>
		}
		function textboxchangeDept()
		{
			if (document.frmMain.btnNewDept.value == 'NEW')
			{
				alert("To save a new Department, complete the form and click 'Save Request' button.");
				document.frmMain.btnNewDept.value = 'BACK';
				document.frmMain.selDept.disabled = true;
				document.frmMain.txtInstDept.style.visibility = 'visible';
				document.frmMain.txtInstDept.value = "";
				document.frmMain.selClass.value = 1;
				document.frmMain.txtInstAddr.value = "";
				document.frmMain.txtInstCity.value = "";
				document.frmMain.txtInstState.value = "";
				document.frmMain.txtInstZip.value = "";
				document.frmMain.txtInstAddrI.value = "";
				document.frmMain.txtBlname.value = "";
				document.frmMain.txtBilAddr.value = "";
				document.frmMain.txtBilCity.value = "";
				document.frmMain.txtBilState.value = "";
				document.frmMain.txtBilZip.value = "";
				document.frmMain.txtInstDept.focus();
				document.frmMain.HnewDept.value = 'BACK';
			}
			else
			{
				document.frmMain.btnNewDept.value = 'NEW';
				document.frmMain.selDept.disabled = false;
				document.frmMain.txtInstDept.value = "";
				document.frmMain.txtInstDept.style.visibility = 'hidden';
				DeptInfo(document.frmMain.selDept.value);
				document.frmMain.HnewDept.value = 'NEW';
			}
		}
		function hideNewDept() 
		{
			if (document.frmMain.txtInstDept.value == "")
			{	
				document.frmMain.txtInstDept.style.visibility = 'hidden';
				document.frmMain.btnNewDept.value = 'NEW';
				document.frmMain.txtInstDept.value = "";
				document.frmMain.HnewDept.value = 'NEW';
			}
			else
			{
				document.frmMain.txtInstDept.style.visibility = 'visible';
				document.frmMain.btnNewDept.value = 'BACK';
				document.frmMain.selDept.disabled = true;
				document.frmMain.selClass.value = '<%=tmpClass%>';
				document.frmMain.txtInstAddr.value = '<%=tmpNewInstAddr%>';
				document.frmMain.txtInstCity.value = '<%=tmpNewInstCity%>';
				document.frmMain.txtInstState.value = '<%=tmpNewInstState%>';
				document.frmMain.txtInstZip.value = '<%=tmpNewInstZip%>';
				document.frmMain.txtInstAddrI.value = '<%=tmpNewInstAddrI%>';
				document.frmMain.txtBlname.value = '<%=tmpBLname%>';
				if (document.frmMain.chkBill.checked != true)
				{
					document.frmMain.chkBill.checked = false;
					document.frmMain.txtBilAddr.value = '<%=tmpBilInstAddr%>';
					document.frmMain.txtBilCity.value = '<%=tmpBilInstCity%>';
					document.frmMain.txtBilState.value = '<%=tmpBilInstState%>';
					document.frmMain.txtBilZip.value = '<%=tmpBilInstZip%>';
				}
				else
				{
					document.frmMain.chkBill.checked = true;
					document.frmMain.txtBilAddr.value = "";
					document.frmMain.txtBilCity.value = "";
					document.frmMain.txtBilState.value = "";
					document.frmMain.txtBilZip.value = "";
				}
				document.frmMain.HnewDept.value = 'BACK';
			}
		}
		function ReqChoice(dept, req)
		{
			 var i;
			for(i=document.frmMain.selReq.options.length-1;i>=1;i--)
			{
				if (req != "undefined")
				{
					if (document.frmMain.selReq.options[i].value != req)
					{
						document.frmMain.selReq.remove(i);
					}
				}
				else
				{
					document.frmMain.selReq.remove(i);
				}
			}
			<%=strInstReqDept%>
		}
		function ReqShowMe()
		{
			if (document.frmMain.chkAll.checked == true) 
			{
				for(i=document.frmMain.selReq.options.length-1;i>=1;i--)
				{
					document.frmMain.selReq.remove(i);
				}
				<%=strReq%>
			}
			else
			{
				ReqChoice(document.frmMain.selDept.value);
			}
		}
		function IntrShowMe()
		{
			if (document.frmMain.chkAll2.checked == true) 
			{
				for(i=document.frmMain.selIntr.options.length-1;i>=1;i--)
				{
					document.frmMain.selIntr.remove(i);
				}
				<%=strIntr2%>
			}
			else
			{
				LangChoice(document.frmMain.selLang.value);
			}
		}
		function CalendarView(strDate)
		{
			document.frmMain.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmMain.submit();
		}
		function GetLangID(xxx)
		{
			if (xxx == "" )
			{
				return 0;
			}
			<%=strLangChk%>
		}
		
		function IsNumeric(sText)
		{
			var ValidChars = "0123456789.";
		   	var IsNumber=true;
		   	var Char;
			newText = sText.replace("-", "")
			if (newText.length = 10)
			{
			 	for (i = 0; i < newText.length && IsNumber == true; i++) 
			     { 
			     	Char = newText.charAt(i); 
			      	if (ValidChars.indexOf(Char) == -1) 
			         {
			         		IsNumber = false;
			         }
			      }
			   	return IsNumber;
			 }
			 
		  }
		function myfee()
		{
			if (document.frmMain.chkEmer.checked == true)
			{
				document.frmMain.chkEmerFee.disabled = false;
			}
			else
			{
				document.frmMain.chkEmerFee.disabled = true;
				document.frmMain.chkEmerFee.checked = false;
			}
		}
		 //-->
		</script>
		</head>
		<body onload='InstInfo(<%=tmpInst%>); hideNewInts(); hideNewReq(); IntrInfo(<%=tmpIntr%>); hideNewIntr(); chkPrim(); chkPrim2(); hideNewDept(); myfee();
		
			 DeptInfo(<%=tmpDept%>); DeptChoice(<%=tmpInst%>, <%=tmpDept%>); ReqChoice(<%=tmpDept%>, <%=tmpReqP%>); ReqInfo(<%=tmpReqP%>); LangChoice(<%=tmpLang%>, <%=tmpIntr%>); 
			'>
			<form method='post' name='frmMain' action='main.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<form name='frmService' method='post' action=''>
								<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
									<tr>
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form <%=RemEdit%></td>
									</tr>
									<tr>
										<td align='center' colspan='10'><nobr>(*) required</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td  align='left'>
											<div name="dErr" style="width:100%; height:55px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr>
										<td class='header' colspan='10'><nobr>Contact Information</td>
									</tr>
									
									<tr>
										<td align='right'>Timestamp:</td>
										<td width='300px'><input class='main' size='23' readonly name='txttstamp' value='<%=tmpTS%>'></td>
										<td align='right'>Emergency:</td>
										<td width='300px'><input type='checkbox' name='chkEmer' value='1' <%=tmpEmer%> onclick='myfee();'></td>
									</tr>
									<tr><td>&nbsp;</td>
										<td>&nbsp;</td>
										<td align='right'>Apply Emergency Fee:</td>
										<td><input type='checkbox' name='chkEmerFee' value='1' <%=tmpEmerFee%>></td>
									</tr>
									<tr>
										<td align='right'>*Institution:</td>
										<td width='350px'>
											<select class='seltxt' name='selInst'  style='width:250px;' onfocus='DeptChoice(document.frmMain.selInst.value);DeptInfo(document.frmMain.selDept.value); ' onchange='DeptChoice(document.frmMain.selInst.value);DeptInfo(document.frmMain.selDept.value); '>
												<option value='-1'>&nbsp;</option>
												<%=strInst%>
											</select>
											<input type='button'  value="FIND" <%=HpLock%>  name="findReq" onclick='PopMe();' title='Search instiution' class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<input class='btnLnk' type='button' name='btnNew' value='NEW'  <%=HpLock%> onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeInst();'>
											<input type='hidden' name='hideInst' value='<%=tmpInst%>'>
										</td>
										<td>
											<input type='hidden' name='HnewInt'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td><input size='50' class='main' maxlength='50' name='txtNewInst' value='<%=tmpNewInstTxt%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right' width='15%'>*Department:</td>
										<td>	
											<select class='seltxt' name='selDept'  style='width:250px;' onfocus='DeptInfo(document.frmMain.selDept.value); ReqChoice(document.frmMain.selDept.value); '  onchange='DeptInfo(document.frmMain.selDept.value); ReqChoice(document.frmMain.selDept.value); '>
												<option value='0'>&nbsp;</option>
												<%=strDept2%>
											</select>
											<input class='btnLnk' type='button' name='btnNewDept' <%=HpLock%> value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeDept();'>
											<input type='hidden' name='HnewDept'>
											<input type='hidden' name='hideDept' value='<%=tmpDept%>'>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td><input class='main' size='50' maxlength='50' name='txtInstDept' value='<%=tmpNewInstDept%>' onkeyup='bawal(this);'></td>
										<td align='right' width='15%'>Rate:</td>
										<td>
											<input class='main' size='5' maxlength='5'  readonly name='txtInstRate' value='<%=tmpInstRate%>'>
											<select class='seltxt' style='width: 70px;' name='selInstRate'>
												<option value='0' >&nbsp;</option>
												<%=strRate1%>
											</select>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">
												*Rate varies per request
											</span>
										</td>
									</tr>
									<tr>
										<td align='right'>Classification:</td>
										<td>
											<select class='seltxt' name='selClass'>
												<option value='1' <%=SocSer%>>Social Services</option>
												<option value='2' <%=Priv%>>Private</option>
												<option value='3' <%=court%>>Court</option>
												<option value='4' <%=Med%>>Medical</option>
												<option value='5' <%=Med%>>Legal</option>
											</select>
										</td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtInstAddrI' value='<%=tmpNewInstAddrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtInstAddr' value='<%=tmpNewInstAddr%>' onkeyup='bawal(this);'>
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtInstCity' value='<%=tmpNewInstCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtInstState' value='<%=tmpNewInstState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtInstZip' value='<%=tmpNewInstZip%>' onkeyup='bawal(this);'>
											<input type='hidden' name='OldAddr'>
										</td>
									</tr>
									<tr>
										<td align='right'>Billed To:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtBlname' value='<%=tmpBLname%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Billing Address:</td>
										<td>
											<input type='checkbox' name='chkBill' <%=chkBillMe%>>
											(same as appointment address)
										</td>
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtBilAddr' value='<%=tmpBilInstAddr%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtBilCity' value='<%=tmpBilInstCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtBilState' value='<%=tmpBilInstState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtBilZip' value='<%=tmpBilInstZip%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>*Requesting Person:</td>
										<td width='200px'>
											<nobr>
											<select id='selReq' class='seltxt' name='selReq'  style='width:250px;' onfocus='JavaScript:ReqInfo(document.frmMain.selReq.value);' onchange='JavaScript:ReqInfo(document.frmMain.selReq.value);'>
												<option value='-1'>&nbsp;</option>
												<%=strReq2%>
											</select>
											<input class='btnLnk' type='button' name='btnNewReq' <%=HpLock%> value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeReq();'>
											<input type='checkbox' <%=HpLock%> name='chkAll' onclick='ReqShowMe(); ReqInfo(document.frmMain.selReq.value);'>
											Show All
											<input type='hidden' name='HnewReq'>
											<input type='hidden' name='hideReq' value='<%=tmpReqP%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td align='left'>
											<input class='main' size='20' maxlength='20' name='txtReqLname' value='<%=tmpNewReqLN%>' onkeyup='bawal(this);'>
											<input class='trans' style='width: 5px;' name='txtcoma2' readonly value=', '>
											<input class='main' size='20' maxlength='20' name='txtReqFname' value='<%=tmpNewReqFN%>' onkeyup='bawal(this);'>
											<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" size='22' name='txtformat2' readonly value='last name, first name'>
										</td>
									</tr>
									<tr>
										<td align='right'><b>*Contact Numbers:</b></td>
										<td align='left'><b>(any of the following)</b></td>
									</tr>
									<tr>
										<td align='right'>Primary:</td>
										<td>
											<input class='main2'  name='txtPRim1'  readonly size='6'>
										</td>
									</tr>
									<tr>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='1' <%=selRPPhone%> onclick='chkPrim();'>
											Phone:
										</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtphone' value='<%=tmpNewReqPhone%>' onkeyup='bawal(this);'>
											&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtReqExt' value='<%=tmpReqExt%>' onkeyup='bawal(this);'>
										</td>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='2'  <%=selRPFax%> onclick='chkPrim();'>
											Fax:
										</td>
										<td width='300px'><input class='main' size='12' maxlength='12' name='txtfax' value='<%=tmpNewReqFax%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='0' <%=selRPEmail%> onclick='chkPrim();'>
											E-Mail:
										</td>
										<td><input class='main' size='50' maxlength='50' name='txtemail' value='<%=tmpNewReqeMail%>' onkeyup='bawal(this);'></td>
										<td>&nbsp;</td>
										<td>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Please include area code on fax number</span>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' class='header'><nobr>Appointment Information</td>
									</tr>
									<tr>
										<td align='right'>*Client Last Name:</td>
										<td>
											<input class='main' size='20' maxlength='20' name='txtClilname' value='<%=tmplname%>' onkeyup='bawal(this);'>&nbsp;First Name:
											<input class='main' size='20' maxlength='20' name='txtClifname' value='<%=tmpfname%>' onkeyup='bawal(this);'>
										</td>
										<td align='right'>LSS Client:</td>
											<td><input type='checkbox' name='chkClient' value='1' <%=chkClient%>></td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtCliAddrI' value='<%=tmpCAdrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Client Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtCliAdd' value='<%=tmpAddr%>' onkeyup='bawal(this);'>
											<input type='checkbox' name='chkClientAdd' value='1' <%=chkUClientadd%>>Use Client Address
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
										<td align='right'>Client Phone:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCliFon' value='<%=tmpCFon%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td>
											<input class='main' size='25' maxlength='25' name='txtCliCity' value='<%=tmpCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtCliState' value='<%=tmpState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtCliZip' value='<%=tmpZip%>' onkeyup='bawal(this);'>
										</td>
										<td align='right'>Alter. Phone:</td>
										<td align='left' rowspan='2'>
											<textarea name='txtAlter' class='main' onkeyup='bawal(this);' ><%=tmpCAFon%></textarea>
										</td>
									</tr>
									<tr>
										<td align='right'>Gender:</td>
										<td>
											<select class='seltxt' name='selGender' style='width: 75px;'>
												<option value='0' <%=tmpMale%>>Male</option>
												<option value='1' <%=tmpfeMale%>>Female</option>
											</select>
											&nbsp;&nbsp;
											Minor:
											<input type='checkbox' name='chkMinor' value='1' <%=chkMinor%>>
										</td>
									</tr>
									<tr>
										<td align='right'>Directions / Landmarks:</td>
										<td><input class='main' size='50' maxlength='50' name='txtCliDir' value='<%=tmpDir%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Special Circumstances:</td>
										<td><input class='main' size='50' maxlength='50' name='txtCliCir' value='<%=tmpSC%>' onkeyup='bawal(this);'></td>
										<td>&nbsp;</td>
									
									</tr>
									<tr>
										<td align='right'>DOB:</td>
										<td>
											<input class='main' size='11' maxlength='10' name='txtDOB' value='<%=tmpDOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										</td>
										
									</tr>
									<tr>
										<td align='right'>*Language:</td>
										<td>
											<select class='seltxt' name='selLang'  style='width:100px;' onchange='LangChoice(document.frmMain.selLang.value);'>
												<option value='-1'>&nbsp;</option>
												<%=strLang%>
											</select>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Date:</td>
										<td>
											<input class='main' size='10' maxlength='10' name='txtAppDate'  readonly value='<%=tmpAppDate%>'>
											<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
											onclick="showCalendarControl(document.frmMain.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Time:</td>
										<td>
											&nbsp;From:<input class='main' size='5' maxlength='5' name='txtAppTFrom' value='<%=tmpAppTFrom%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
											&nbsp;To:<input class='main' size='5' maxlength='5' name='txtAppTTo' value='<%=tmpAppTTo%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
										</td>
									</tr>
									<tr>
										<td align='right'>Appointment Location:</td>
										<td><input class='main' size='50' maxlength='50' name='txtAppLoc' value='<%=tmpAppLoc%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'><b>For legal appointments:</b></td>
										<td><b>(also fill in)</b></td>
									</tr>
									<tr>
										<td align='right'>Docket Number:</td>
										<td><input class='main' size='50' maxlength='50' name='txtDocNum' value='<%=tmpDoc%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Court Room No:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCrtNum' value='<%=tmpCRN%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
									<td align='right' valign='top'>Appointment Comment:</td>
										<td>
											<textarea name='txtcom' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpCom%></textarea>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' class='header'><nobr>Interpreter Information</td>
									</tr>
									<tr>
										<td align='right'>Interpreter:</td>
										<td>
											<select class='seltxt' name='selIntr' style='width: 200px;' onchange='JavaScript:IntrInfo(document.frmMain.selIntr.value);'>
												<option value='-1'>&nbsp;</option>
												<%=strIntr%>
											</select>
											<input type="button" value="..." name="btnchkavail"
											onclick='PopMe2(document.frmMain.txtAppDate.value, document.frmMain.txtAppTFrom.value, document.frmMain.selLang.value);' title='Check Availability' class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<input class='btnLnk' type='button' name='btnNewIntr' value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeIntr();'>
											<input type='checkbox' name='chkAll2' onclick='IntrShowMe(); IntrInfo(document.frmMain.selIntr.value);'>
											Show All
											<input type='hidden' name='HnewIntr'>
											<input type='hidden' name='Lang1'>
											<input type='hidden' name='Lang2'>
											<input type='hidden' name='Lang3'>
											<input type='hidden' name='Lang4'>
											<input type='hidden' name='Lang5'>
											<input type='hidden' name='LangCtr'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td>
											<input class='main' size='20' maxlength='20' name='txtIntrLname' value='<%=tmpIntrLname%>' onkeyup='bawal(this);'>
											<input class='trans' size='1' style='width: 5px;' name='txtcoma' readonly value=', '>
											<input class='main' size='20' maxlength='20' name='txtIntrFname' value='<%=tmpIntrFname%>' onkeyup='bawal(this);'>
											<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" size='22' name='txtformat' readonly value='last name, first name'>
										</td>
									</tr>
									<tr>
										<td align='right'>Primary:</td>
										<td>
											<input class='main2'  name='txtPRim2'  readonly size='12'>
										</td>
									</tr>
									<tr>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='0'  <%=selIntrEmail%> onclick='chkPrim2();'>
											E-Mail:
										</td>
										<td><input class='main' size='50' maxlength='50' name='txtIntrEmail' value='<%=tmpIntrEmail%>' onkeyup='bawal(this);'></td>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='1' <%=selIntrP1%> onclick='chkPrim2();'>&nbsp;
											Home Phone:
										</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtIntrP1' value='<%=tmpIntrP1%>' onkeyup='bawal(this);'>
											&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtIntrExt' value='<%=tmpIntrExt%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='3' <%=selIntrFax%> onclick='chkPrim2();'>
											&nbsp;&nbsp;&nbsp;Fax:
											</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtIntrFax' value='<%=tmpIntrFax%>' onkeyup='bawal(this);'>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Please include area code on fax number</span>
										</td>
										<td align='right' width='15%'>
											<input type='radio' name='radioPrim2' value='2' <%=selIntrP2%> onclick='chkPrim2();'>
											Mobile Phone:
											</td>
										<td><input class='main' size='12' maxlength='12' name='txtIntrP2' value='<%=tmpIntrP2%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtIntrAddrI' value='<%=tmpNewIntrAddrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtIntrAddr' value='<%=tmpIntrAddr%>' onkeyup='bawal(this);'>
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td colspan='5'>
											<input class='main' size='25' maxlength='25' name='txtIntrCity' value='<%=tmpIntrCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtIntrState' value='<%=tmpIntrState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtIntrZip' value='<%=tmpIntrZip%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>In-House Interpreter:</td>
										<td><input type='checkbox' name='chkInHouse' value='1' <%=tmpInHouse%>></td>
									</tr>
									<tr>
										<td align='right' width='15%'>Default Rate:</td>
										<td>
											<input class='main' size='5' maxlength='5'  readonly  name='txtIntrRate' value='<%=tmpIntrRate%>'>
											<select class='seltxt' style='width: 70px;' name='selIntrRate'>
												<option value='0' >&nbsp;</option>
												<%=strRate2%>
											</select>
										</td>
									<tr>
								
									<tr><td>&nbsp;</td></tr>
									<tr>	
										<td align='right' valign='top'>Interpreter Comment:</td>
										<td colspan='3' >
											<textarea name='txtcomintr' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpComintr%></textarea>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>	
										<td align='right' valign='top'>Billing Comment:</td>
										<td colspan='3' >
											<textarea name='txtcombil' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpCombil%></textarea>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' class='header'><nobr>Language Bank Notes</td>
									</tr>
									<tr>	
										<td align='right' valign='top'>Notes:</td>
										<td colspan='3' >
											<textarea name='txtLBcom' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpLBcom%></textarea>
										</td>
									</tr>
										
										<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
										
												<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ReqChkMe();'>
												<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
								
										</td>
									</tr>
									<tr>
										<td valign="top"><div id="directions" style="display: none;"></div></td>
									</tr>
									<tr>
										<td valign="top"><div id="map_canvas" style="display: none;"></div></td>
									</tr>
								</table>
							</form>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
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

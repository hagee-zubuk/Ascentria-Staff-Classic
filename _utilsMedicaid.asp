<%
Function Z_getmedicaid(reqid)
	Z_getmedicaid = ""
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT medicaid, meridian, nhhealth, wellsense FROM request_T WHERE [index] = " & reqid, g_strCONN, 3, 1
	If Not rsMed.EOF Then
		hmo = Trim(Ucase(Z_FixNull(rsMed("medicaid")))) 
		If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("meridian"))))
		If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("nhhealth"))))
		If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("wellsense")))) 
		Z_getmedicaid = Trim(Ucase(hmo))
	End If
End Function
Function Z_MedicaidCheck(appid)
	Z_MedicaidCheck = True
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT clname, cfname, dob, medicaid, meridian, nhhealth, wellsense FROM [request_T] WHERE [index] = " & appid, g_strCONN, 3, 1
	If Not rsMed.EOF Then
		lname = Trim(Ucase(rsMed("clname")))
		fname = Trim(Ucase(rsMed("cfname")))
		dob = rsMed("dob")
		'hmo = Trim(Ucase(Z_FixNull(rsMed("medicaid")))) 
		'If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("meridian"))))
		'If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("nhhealth"))))
		'If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("wellsense")))) 
		hmo = Z_FixNull(Ucase(Trim(rsMed("medicaid")))) 
		If Z_FixNull(rsMed("meridian")) <> "" Then 
			hmo = Z_FixNull(Ucase(Trim(rsMed("meridian"))))
		End If
		If Z_FixNull(rsMed("nhhealth")) <> "" Then 
			hmo = Z_FixNull(Ucase(Trim(rsMed("nhhealth"))))
		End If
		If Z_FixNull(rsMed("wellsense")) <> "" Then 
			hmo = Z_FixNull(Ucase(Trim(rsMed("wellsense")))) 
		End If
	End If
	rsMed.close
	set rsMed = nothing
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	rsTBL.Open "SELECT * FROM medapprove_T WHERE medicaid = '" & hmo & "'", g_strCONN, 3, 1
	if Not rstbl.EOF Then
		Z_MedicaidCheck = false
		if Trim(Ucase(rsTBL("lname"))) = lname And Trim(Ucase(rsTBL("fname"))) = fname And rsTBL("dob") = dob Then Z_MedicaidCheck = True
	end if
End Function
Function Z_IntrInfo(intrID)
	Z_IntrInfo = "NOT*AVAILABLE****SV*00000000"
	If Z_Czero(intrID) = 0 Then Exit Function
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT [first name], [last name], PID FROM interpreter_T WHERE [index] = " & intrID, g_strCONN, 3, 1
	If Not rsMed.EOF Then
		pid = Z_fixNull(rsMed("PID"))
		If pid= "" Then pid = "00000000"
		cleanIntrLName = Ucase(Trim(Replace(rsMed("last name"), "*", "")))
		cleanIntrFName = Ucase(Trim(Replace(rsMed("first name"), "*", "")))
		Z_IntrInfo = cleanIntrLName & "*" & cleanIntrFName & "****SV*" & pid
	End If
	rsMed.Close
	Set rsMed = Nothing
End Function
Function Z_NameMed(ReqID)
	Z_NameMed = "NOT*AVAILABLE"
	If Z_Czero(ReqID) = 0 Then Exit Function
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT clname, cfname FROM Request_T WHERE [index] = " & ReqID, g_strCONN, 3, 1
	If Not rsMed.EOF Then
		Z_NameMed = Z_RemoveDlbQuote(Ucase(Trim(rsMed("clname")))) & "*" & Z_RemoveDlbQuote(Ucase(Trim(rsMed("cfname"))))
	End If
	rsMed.Close
	Set rsMed = Nothing
End Function
Function Z_DOBMed(ReqID)
	Z_DOBMed = ""
	If Z_Czero(ReqID) = 0 Then Exit Function
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT DOB FROM Request_T WHERE [index] = " & ReqID, g_strCONN, 3, 1
	If Not rsMed.EOF Then
		if not isnull(rsMed("DOB")) Then
		dte = FormatDateTime(rsMed("DOB"), 2)
		dteYr = Year(dte)
		dteMn = Right("0" & Month(dte), 2)
		dteDy = Right("0" & Day(dte), 2)
		Z_DOBMed = dteYr & dteMn & dteDy
	end if
	End If
	rsMed.Close
	Set rsMed = Nothing
End Function
Function Z_GenderMed(ReqID)
	If Z_Czero(ReqID) = 0 Then Exit Function
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT Gender FROM Request_T WHERE [index] = " & ReqID, g_strCONN, 3, 1
	If Not rsMed.EOF Then
		If IsNull(rsMed("gender")) Then
			Z_GenderMed = "U"
		Else
			If rsMed("gender") = 1 Then
				Z_GenderMed = "F"
			Else
				Z_GenderMed = "M"
			End If
		End If
	Else
		Z_GenderMed = "U"
	End If
	rsMed.Close
	Set rsMed = Nothing
End Function
Function Z_DateMed(ReqID)
	Z_DateMed = ""
	If Z_Czero(ReqID) = 0 Then Exit Function
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT appdate FROM Request_T WHERE [index] = " & ReqID, g_strCONN, 3, 1
	If Not rsMed.EOF Then
		dte = FormatDateTime(rsMed("appdate"), 2)
		dteYr = Year(dte)
		dteMn = Right("0" & Month(dte), 2)
		dteDy = Right("0" & Day(dte), 2)
		Z_DateMed = dteYr & dteMn & dteDy
	End If
	rsMed.Close
	Set rsMed = Nothing
End Function
%>
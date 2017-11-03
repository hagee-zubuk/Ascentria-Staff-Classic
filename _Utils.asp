<%
'list of functions used
Function Z_Mod(numerator, denominator)
	Dim x
	x = (numerator / denominator)
	x = x - Int(x)
	Z_Mod = x * denominator
End Function
Function Z_RemoveDlbQuote(xxx)
	' clean string
	Z_RemoveDlbQuote = xxx
	If Not IsNull(xxx) Or xxx <> "" Then Z_RemoveDlbQuote = Replace(xxx, "''", "'")
End Function
Function CleanMe2(xxx)
	' clean string
	CleanMe2 = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe2 = Replace(xxx, "'", "''")
End Function
Function Z_GetDec(xxx)
	Dim i, j
	j = Len(xxx)
	i = InStrRev(xxx, ".")
	If i>0 Then Z_GetDec = UCase(Right(xxx, j-i)) Else Z_GetDec = 0
	Z_GetDec = Z_GetDec
End Function
Function GUIDExists(xxx)
	GUIDExists = False
	If xxx = "" Then Exit Function
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(F604AStr & xxx & ".PDF") Then GUIDExists = True
	Set fso = Nothing
End Function
Function GUIDExists271(xxx, yyy)
	GUIDExists271 = False
	If xxx = "" Then Exit Function
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(f271Str & xxx & "." & yyy) Then GUIDExists271 = True
	Set fso = Nothing
End Function
Function FileUpload(xxx)
	FileUpload = False
	If xxx = "" Then Exit Function
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(F604AStr & xxx & ".PDF") Then FileUpload = True
	Set fso = Nothing
End Function
Function Z_GenerateGUID()
	Dim objGUID
	Set objGUID = Server.CreateObject("Z_MkGUID.ZGUID")
	Z_GenerateGUID = objGUID.GetGUID()
	Set objGUID = Nothing
End Function
Function InstBillHrs(xxx, yyy, zzz, aaa, appdate)
	defHrs = 2
	mintime = 120
	If Z_CDate(appdate) < Z_CDate("8/1/2014") Then  
		defHrs = 1.5
		mintime = 90
	End If
	'get billable hours
	InstBillHrs = 0
	If xxx <> "12:00:00 AM" Then
		If Z_DateNull(xxx) = empty Then
			InstBillHrs = 0
			Exit Function
		End If
	ElseIf Z_DateNull(xxx) = empty Then
			InstBillHrs = 0
			Exit Function
	End If
	If yyy <> "12:00:00 AM" Then
		If Z_DateNull(yyy) = empty Then
			InstBillHrs = 0
			Exit Function
		End If
	ElseIf Z_DateNull(yyy) = empty Then
			InstBillHrs = 0
			Exit Function
	End If
	If zzz = 27 Then
		If Z_dates(yyy) = Empty Then 
			InstBillHrs = defHrs '2 '1.5
			Exit Function
		End If
		If DateDiff("n", xxx, yyy) <= mintime And DateDiff("n", xxx, yyy) >= 0 Then 
			InstBillHrs = defHrs '2 '1.5
			Exit Function
		Else
			tmpBillMin = DateDiff("n", xxx, yyy)
			If tmpBillMin < 0 Then tmpBillMin = 1440 - Mid(tmpBillMin, 2)
			tmpBillHrs = tmpBillMin / 60
			tmpBillMHrs = Int(tmpBillHrs)
			tmpLen = Len(tmpBillHrs)
			tmpPosDec = Instr(tmpBillHrs, ".")
			tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
			If ClassInt(aaa) = 1 Or ClassInt(aaa) = 4 Then
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.25 Then
					InstBillHrs = tmpBillMHrs + 0.25
				ElseIf  Cdbl(tmpBillMMin) > 0.25 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.75 Then
					InstBillHrs = tmpBillMHrs + 0.75
				ElseIf  Cdbl(tmpBillMMin) > 0.75 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs = tmpBillMHrs + 1
				Else
					InstBillHrs = tmpBillMHrs
				End If
			Else
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs = tmpBillMHrs + 1
				Else
					InstBillHrs = tmpBillMHrs
				End If
			End If
		End If
	Else
		If Z_dates(yyy) = Empty Then 
			InstBillHrs = 2
			Exit Function
		End If	
		If DateDiff("n", xxx, yyy) <= 120 And DateDiff("n", xxx, yyy) >= 0 Then 
			InstBillHrs = 2
			Exit Function
		Else
			tmpBillMin = DateDiff("n", xxx, yyy)
			If tmpBillMin < 0 Then tmpBillMin = 1440 - Mid(tmpBillMin, 2)
			tmpBillHrs = tmpBillMin / 60
			tmpBillMHrs = Int(tmpBillHrs)
			tmpLen = Len(tmpBillHrs)
			tmpPosDec = Instr(tmpBillHrs, ".")
			tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
			If ClassInt(aaa) = 1 Or ClassInt(aaa) = 4 Then
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.25 Then
					InstBillHrs = tmpBillMHrs + 0.25
				ElseIf  Cdbl(tmpBillMMin) > 0.25 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.75 Then
					InstBillHrs = tmpBillMHrs + 0.75
				ElseIf  Cdbl(tmpBillMMin) > 0.75 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs = tmpBillMHrs + 1
				Else
					InstBillHrs = tmpBillMHrs
				End If
			Else
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs = tmpBillMHrs + 1
				Else
					InstBillHrs = tmpBillMHrs
				End If
			End If
		End If
	End If
End Function
Function InstBillHrs2(xxx, yyy, zzz, aaa, appdate)
	defHrs = 2
	mintime = 120
	If Z_CDate(appdate) < Z_CDate("8/1/2014") Then  
		defHrs = 1.5
		mintime = 90
	End If
	'get billable hours
	InstBillHrs2 = 0
	If zzz = 27 Then
		If Z_dates(yyy) = Empty Then 
			InstBillHrs2 = defHrs'2 '1.5
			Exit Function
		End If
		If DateDiff("n", xxx, yyy) <= mintime And DateDiff("n", xxx, yyy) >= 0 Then 
			InstBillHrs2 = defHrs'2 '1.5
			Exit Function
		Else
			tmpBillMin = DateDiff("n", xxx, yyy)
			If tmpBillMin < 0 Then tmpBillMin = 1440 - Mid(tmpBillMin, 2)
			tmpBillHrs = tmpBillMin / 60
			tmpBillMHrs = Int(tmpBillHrs)
			tmpLen = Len(tmpBillHrs)
			tmpPosDec = Instr(tmpBillHrs, ".")
			tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
			If ClassInt(aaa) = 1 Or ClassInt(aaa) = 4 Then
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.25 Then
					InstBillHrs2 = tmpBillMHrs + 0.25
				ElseIf  Cdbl(tmpBillMMin) > 0.25 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs2 = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.75 Then
					InstBillHrs2 = tmpBillMHrs + 0.75
				ElseIf  Cdbl(tmpBillMMin) > 0.75 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs2 = tmpBillMHrs + 1
				Else
					InstBillHrs2 = tmpBillMHrs
				End If
			Else
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs2 = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs2 = tmpBillMHrs + 1
				Else
					InstBillHrs2 = tmpBillMHrs
				End If
			End If
		End If
	Else
		If Z_dates(yyy) = Empty Then 
			InstBillHrs2 = 2
			Exit Function
		End If	
		If DateDiff("n", xxx, yyy) <= 120 And DateDiff("n", xxx, yyy) >= 0 Then 
			InstBillHrs2 = 2
			Exit Function
		Else
			tmpBillMin = DateDiff("n", xxx, yyy)
			If tmpBillMin < 0 Then tmpBillMin = 1440 - Mid(tmpBillMin, 2)
			tmpBillHrs = tmpBillMin / 60
			tmpBillMHrs = Int(tmpBillHrs)
			tmpLen = Len(tmpBillHrs)
			tmpPosDec = Instr(tmpBillHrs, ".")
			tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
			If ClassInt(aaa) = 1 Or ClassInt(aaa) = 4 Then
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.25 Then
					InstBillHrs2 = tmpBillMHrs + 0.25
				ElseIf  Cdbl(tmpBillMMin) > 0.25 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs2 = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.75 Then
					InstBillHrs2 = tmpBillMHrs + 0.75
				ElseIf  Cdbl(tmpBillMMin) > 0.75 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs2 = tmpBillMHrs + 1
				Else
					InstBillHrs2 = tmpBillMHrs
				End If
			Else
				If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.5 Then
					InstBillHrs2 = tmpBillMHrs + 0.5
				ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.99 Then
					InstBillHrs2 = tmpBillMHrs + 1
				Else
					InstBillHrs2 = tmpBillMHrs
				End If
			End If
		End If
	End If
End Function
Function CTime(tmptime)
	if z_fixnull(tmptime) <> "" Then 
		myTime = Right(tmptime, 11)
		If instr(myTime, "/") > 0 Then
			Ctime = ""
		Else
			Ctime = cdate(myTime)
		End If
	Else
		Ctime = ""
	End If
End Function
Function Z_dates(xxx)
	Z_dates = xxx
	If xxx = "12:00:00 AM" Or xxx = "24:00" Then
		Z_dates = "12:00:01 AM"
		Exit Function
	Else
		If IsDate(xxx) Then
			Z_dates = Cdate(xxx)
		Else
			Z_dates = Empty
		End If
	End If
End Function
Function RoundDown(Value)
	If InStr(Value, ".") Then
		RoundDown = Left(Value, InStr(Value, ".") - 1)
	Else
		RoundDown = Value
	End If
End Function
Function MakeTime(xxx)
	If xxx < 60 Then
		If Len(xxx) = 1 Then
			MakeTime = 0 & ":0" & xxx
		Else
			MakeTime = 0 & ":" & xxx
		End If
		Exit Function
	End If
	DecTime = xxx / 60
	myHrs = RoundDown(DecTime)
	DecTime = DecTime - RoundDown(DecTime)
	myMins = Z_FormatNumber(DecTime * 60, 0)
	If Len(myMins) = 1 Then myMins = "0" & myMins
	MakeTime = myHrs & ":" & myMins
End Function
Function IntrBillHrs(xxx, yyy)
	'get billable hours
	IntrBillHrs = 0
	If xxx <> "12:00:00 AM" Then
		If Z_DateNull(xxx) = empty Then
			IntrBillHrs = 0
			Exit Function
		End If
	ElseIf Z_DateNull(xxx) = empty Then
			IntrBillHrs = 0
			Exit Function
	End If
	If yyy <> "12:00:00 AM" Then
		If Z_DateNull(yyy) = empty Then
			IntrBillHrs = 0
			Exit Function
		End If
	ElseIf Z_DateNull(yyy) = empty Then
			IntrBillHrs = 0
			Exit Function
	End If
	If DateDiff("n", xxx, yyy) <= 120 And DateDiff("n", xxx, yyy) >= 0 Then 
		IntrBillHrs = 2
		Exit Function
	Else
		tmpBillMin = DateDiff("n", xxx, yyy)
		If tmpBillMin < 0 Then tmpBillMin = 1440 - Mid(tmpBillMin, 2)
		tmpBillHrs = tmpBillMin / 60
		tmpBillMHrs = Int(tmpBillHrs)
		tmpLen = Len(tmpBillHrs)
		tmpPosDec = Instr(tmpBillHrs, ".")
		tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
		If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.25 Then
			IntrBillHrs = tmpBillMHrs + 0.25
		ElseIf  Cdbl(tmpBillMMin) > 0.25 And  Cdbl(tmpBillMMin) <= 0.50 Then
			IntrBillHrs = tmpBillMHrs + 0.5
		ElseIf  Cdbl(tmpBillMMin) > 0.50 And  Cdbl(tmpBillMMin) <= 0.75 Then
			IntrBillHrs = tmpBillMHrs + 0.75
		ElseIf  Cdbl(tmpBillMMin) > 0.75 And  Cdbl(tmpBillMMin) <= 0.99 Then
			IntrBillHrs = tmpBillMHrs + 1
		Else
			IntrBillHrs = tmpBillMHrs
		End If
	End If
End Function
Function GetSun(xxx)
	If WeekDay(xxx) = 1 Then 
		GetSun = xxx
	Else
		tmpWkDay = WeekDay(xxx)
		ctr = 1
		Do Until tmpWkDay = 1
			tmpWkDay = tmpWkDay - 1
			ctr = ctr + 1
		Loop
		GetSun = DateAdd("d", -(ctr - 1), xxx)
	End If
End Function
Function GetSat(xxx)
	If WeekDay(xxx) = 7 Then 
		GetSat = xxx
	Else
		tmpWkDay = WeekDay(xxx)
		ctr = 1
		Do Until tmpWkDay = 7
			tmpWkDay = tmpWkDay + 1
			ctr = ctr + 1
		Loop
		GetSat = DateAdd("d", (ctr - 1), xxx)
	End If
End Function
Function Z_Replace(var, del, rpl)
	If IsNull(var) Then 
		Z_Replace = 0
		Exit Function
	End If
	If Instr(var, del) <> 0 Then 
		Z_Replace = Replace(var, del, rpl)
	Else
		Z_Replace  = var
	End If
End Function
Function Z_DateNull(var)
	Dim dblTmp
    'Z_DateNull = False
    If IsNull(var) Then 
    	Z_DateNull = Empty
    ElseIf var = "" Then 
    	Z_DateNull = Empty
    ElseIf Not IsDate(var) Then
    	Z_DateNull = Empty
    Else
    	Z_DateNull = cdate(var)
    End If
End Function
Function Z_IsOdd2(var)
	Dim dblTmp
    Z_IsOdd2 = False
    If IsNull(var) Then Exit Function
    If var = "" Then Exit Function
    If Not IsNumeric(var) Then Exit Function
    Z_IsOdd2 = CBool((var Mod 2) = 1) Or CBool((var Mod 2) = -1)
End Function

Function Z_CZero(var)
	If IsNull(var) Then 
		Z_CZero = Cdbl(0)
	ElseIf var = "" Then 
		Z_CZero = Cdbl(0)	
	ElseIf Not IsNumeric(var) Then
		Z_CZero = Cdbl(0)
	Else
		Z_CZero = Cdbl(var)
	End If
End Function
Function Z_ZeroToNull(xxx)
	Z_ZeroToNull = xxx
	If xxx = 0 Then Z_ZeroToNull = ""
End Function
Function Z_CEmpty(var)
	If IsNull(var) Then 
		Z_CEmpty = ""
	ElseIf var = "" Then 
		Z_CEmpty = ""	
	Else
		Z_CEmpty = var
	End If
End Function

Function Z_CDate(var)
	If IsNull(var) Then Z_CDate = Empty
	If var = "" Then Z_CDate = Empty
	If IsDate(var) Then Z_CDate = CDate(var)
End Function

Function Z_IsOdd(var)
DIM dblTmp
	Z_IsOdd = False
	If IsNull(var) Then Exit Function
	If var = "" Then Exit Function
	If Not IsNumeric(var) Then Exit Function
	Z_IsOdd = CBool( (var Mod 2) = 1 )
End Function

Function Z_FixNull(vntZ)
	If IsNull(vntZ) Then
		Z_FixNull = ""
	ElseIf IsEmpty(vntZ) Then
		Z_FixNull = ""
	ElseIf Trim(vntZ) = "" Then
		Z_FixNull = ""
	Else
		Z_FixNull = vntZ
	End If
End Function

Function Z_NullFix(vntZ)
	If IsNull(vntZ) Then
		Z_NullFix = Null
	ElseIf Trim(vntZ) = "" Then
		Z_NullFix = Null
	Else
		Z_NullFix = vntZ
	End If
End Function

Function Z_Blank(vntZ)
	Z_Blank = False
	If IsNull(vntZ) Then
		Z_Blank = True
		Exit Function
	ElseIf Trim(vntZ) = "" Then
		Z_Blank = True
	End If
End Function

Function Z_MDYDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = Z_CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_MDYDate = "0"
	Z_MDYDate = Z_MDYDate & lngTmp & "/"
	lngTmp = Z_CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_MDYDate = Z_MDYDate & "0"
	Z_MDYDate = Z_MDYDate & lngTmp & "/"
	strTmp = DatePart("yyyy", dtDate)
	Z_MDYDate = Z_MDYDate & Right(strTmp,2)
End Function

Function Z_SFDate(dtDate)
' semiflowery date
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = DatePart("m", dtDate)
	strTmp = MonthName(lngTmp, False) & " "
	strTmp = strTmp & DatePart("d", dtDate)
	lngTmp = Z_CLng(Right(strTmp,1))
	If lngTmp = 1 Then
		strTmp = strTmp & "st"
	ElseIf lngTmp = 2 Then
		strTmp = strTmp & "nd"
	ElseIf lngTmp = 3 Then
		strTmp = strTmp & "rd"
	Else
		strTmp = strTmp & "th"
	End If
	strTmp = strTmp & ", " & DatePart("yyyy", dtDate)
	Z_SFDate = strTmp
End Function

Function Z_DateAdd(dtDate, lngPd)
' returns a Date: lngPd business days from dtDate
DIM	lngAdded, dtTmp, lngDy, lngTmp, lngYr
	If Not IsDate(dtDate) Then Exit Function
	lngPd = Z_CLng(lngPd)
	If lngPd = 0 Then
		Z_DateAdd = Z_SFDate(dtDate)
		Exit Function
	ElseIf lngPd > 0 Then
		lngDy = 1
	ElseIf lngPd < 0 Then
		lngDy = -1
	End If
	dtTmp = dtDate
	lngAdded = 0
	Do While lngAdded < lngPd
		dtTmp = DateAdd("d", lngDy, dtTmp)
		lngTmp = DatePart("w", dtTmp, vbSunday)
		lngYr = DatePart("yyyy", dtTmp)
		'Response.Write "<!-- " & dtTmp & ": " & lngTmp & " -->" & vbCrLf
		If lngTmp > 1 And lngTmp < 7 Then
			lngAdded = lngAdded + 1
			' holiday check
			If dtTmp = CDate("12/25/" & lngYr) Or dtTmp = CDate("1/1/" & lngYr) Or _
					dtTmp = CDate("1/2/" & lngYr) Or dtTmp = CDate("1/19/" & lngYr) Or _
					dtTmp = CDate("2/16/" & lngYr) Or dtTmp = CDate("5/31/" & lngYr) Or _
					dtTmp = CDate("7/5/" & lngYr) Or dtTmp = CDate("9/6/" & lngYr) Or _
					dtTmp = CDate("10/11/" & lngYr) Or dtTmp = CDate("11/25/" & lngYr) Or _
					dtTmp = CDate("11/26/" & lngYr) Or dtTmp = CDate("12/24/" & lngYr) Then
				lngAdded = lngAdded - 1
			End If
		End If
	Loop
	Z_DateAdd = Z_SFDate(dtTmp)
End Function


Function Z_MDYDateAdd(dtDate, lngPd)
' returns a Date: lngPd business days from dtDate
DIM	lngAdded, dtTmp, lngDy, lngTmp, lngYr
	If Not IsDate(dtDate) Then Exit Function
	lngPd = Z_CLng(lngPd)
	If lngPd = 0 Then
		Z_MDYDateAdd = Z_MDYDate(dtDate)
		Exit Function
	ElseIf lngPd > 0 Then
		lngDy = 1
	ElseIf lngPd < 0 Then
		lngDy = -1
	End If
	dtTmp = dtDate
	lngAdded = 0
	Do While lngAdded < lngPd
		dtTmp = DateAdd("d", lngDy, dtTmp)
		lngTmp = DatePart("w", dtTmp, vbSunday)
		lngYr = DatePart("yyyy", dtTmp)
		'Response.Write "<!-- " & dtTmp & ": " & lngTmp & " -->" & vbCrLf
		If lngTmp > 1 And lngTmp < 7 Then
			lngAdded = lngAdded + 1
			' holiday check
			If dtTmp = CDate("12/25/" & lngYr) Or dtTmp = CDate("1/1/" & lngYr) Or _
					dtTmp = CDate("1/2/" & lngYr) Or dtTmp = CDate("1/19/" & lngYr) Or _
					dtTmp = CDate("2/16/" & lngYr) Or dtTmp = CDate("5/31/" & lngYr) Or _
					dtTmp = CDate("7/5/" & lngYr) Or dtTmp = CDate("9/6/" & lngYr) Or _
					dtTmp = CDate("10/11/" & lngYr) Or dtTmp = CDate("11/25/" & lngYr) Or _
					dtTmp = CDate("11/26/" & lngYr) Or dtTmp = CDate("12/24/" & lngYr) Then
				lngAdded = lngAdded - 1
			End If
		End If
	Loop
	Z_MDYDateAdd = Z_MDYDate(dtTmp)
End Function


Function Z_MDYCalDateAdd(dtDate, lngPd)
' returns a Date: lngPd business days from dtDate
DIM	lngAdded, dtTmp, lngDy, lngTmp, lngYr
	If Not IsDate(dtDate) Then Exit Function
	lngPd = Z_CLng(lngPd)
	If lngPd = 0 Then
		Z_MDYCalDateAdd = Z_MDYDate(dtDate)
		Exit Function
	ElseIf lngPd > 0 Then
		lngDy = 1
	ElseIf lngPd < 0 Then
		lngDy = -1
	End If
	dtTmp = dtDate
	lngAdded = 0
	dtTmp = DateAdd("d", lngPd, dtTmp)
	Do While True
		lngTmp = DatePart("w", dtTmp, vbSunday)
		If lngTmp > 1 And lngTmp < 7 Then
			lngYr = DatePart("yyyy", dtTmp)
			' holiday check
			If dtTmp = CDate("12/25/" & lngYr) Or dtTmp = CDate("1/1/" & lngYr) Or _
					dtTmp = CDate("1/2/" & lngYr) Or dtTmp = CDate("1/19/" & lngYr) Or _
					dtTmp = CDate("2/16/" & lngYr) Or dtTmp = CDate("5/31/" & lngYr) Or _
					dtTmp = CDate("7/5/" & lngYr) Or dtTmp = CDate("9/6/" & lngYr) Or _
					dtTmp = CDate("10/11/" & lngYr) Or dtTmp = CDate("11/25/" & lngYr) Or _
					dtTmp = CDate("11/26/" & lngYr) Or dtTmp = CDate("12/24/" & lngYr) Then
				dtTmp = DateAdd("d", 1, dtTmp)
			Else
				Exit Do
			End If
		Else
			dtTmp = DateAdd("d", 1, dtTmp)
		End If
	Loop
	Z_MDYCalDateAdd = Z_MDYDate(dtTmp)
End Function

Function Z_FixPath(path)
	If Right(path,1)<>"\" Then
		Z_FixPath = path & "\"
	Else
		Z_FixPath = path
	End If
End Function

Function Z_FixVRoot(strWD, strBase)
	Dim strRes, i, strArry
	i = (Len(strWD)-Len(g_FilesPath))
	If i > 0 Then 
		strRes = Right(strWD, i)
		strArry = Split(strRes,"\")
		strRes = ""
		For i = 0 to UBound(strArry)
			if strArry(i)<>"" Then strRes= strRes & strArry(i) & "/"
		Next
		Z_FixVRoot = strRes
	End If
End Function

Function Z_CleanExt(name)
	Dim i
	i = InStrRev(name, ".")
	If i>0 Then Z_CleanExt = Left(name, i-1) Else Z_CleanExt = name
End Function

Function Z_GetExt(name)
	Dim i, j
	j = Len(name)
	i = InStrRev(name, ".")
	If i>0 Then Z_GetExt = UCase(Right(name, j-i)) Else Z_GetExt = ""
	Z_GetExt = UCase(Z_GetExt)
End Function

Function Z_GetPath(name)
	Dim i, j
	If Right(name, 1) = "\" Then name = Left(name, Len(name)-1)
	i = InStrRev(name, "\")
	If i>0 Then Z_GetPath = LCase(Left(name, i)) Else Z_GetPath = LCase(name)
End Function

Function Z_GetFilename(name)
	Dim i, j
	j = Len(name)
	i = InStrRev(name, "\")
	If i > 0 Then Z_GetFilename = Right(name, j-i) Else Z_GetFilename = name
End Function

Function Z_FormatNumber(strN, Decimals)
	Dim strTmp
	Z_FormatNumber = ""
	If IsNull(Decimals) Then Exit Function
	If Not IsNumeric(Decimals) Then Exit Function
	If Not IsNull(strN) Then
		strN = Trim(strN)
		If Trim(strN) <> "" Then
			If IsNumeric(strN) Then
				Z_FormatNumber = FormatNumber(strN, Decimals, -1, -1, -1)
			Else
				Z_FormatNumber = strN
			End If
		End If
	End If
End Function

Function Z_FormatNumberNC(strN, Decimals)
	Dim strTmp
	Z_FormatNumberNC = ""
	If IsNull(Decimals) Then Exit Function
	If Not IsNumeric(Decimals) Then Exit Function
	If Not IsNull(strN) Then
		strN = Trim(strN)
		If Trim(strN) <> "" Then
			If IsNumeric(strN) Then
				Z_FormatNumberNC = FormatNumber(strN, Decimals, -1, -1, 0)
			Else
				Z_FormatNumberNC = strN
			End If
		End If
	End If
End Function


Function Z_MapMime(strM)
	strM = UCase(strM)
	Select Case strM
		Case "PDF"
			Z_MapMime = "application/PDF"
		Case "DOC"
		Case "DOT"
			Z_MapMime = "application/msword"
		Case "XLS"
			Z_MapMime = "application/vnd.ms-excel"
		Case "TAR"
			Z_MapMime = "application/x-tar"
		Case "ZIP"
			Z_MapMime = "application/x-zip-compressed"
		Case "TXT"
			Z_MapMime = "text/plain"
		Case else
			Z_MapMime = "application/x-octetstream"
	End Select
End Function

Function Z_CDbl(var)
	If Not IsNull(var) Then
		If IsNumeric(var) Then
			var = Replace(var," ","")
			Z_CDbl = var
			If Len(var)<=10 then Z_CDbl = CDbl(Replace(var,",",""))
		Else
			Z_CDbl = 0.0
		End If
	Else
		Z_CDbl = 0.0
	End If
End Function

Function Z_CLng(var)
DIM lngI, lngZ, blnLeading, strTmp
	If Not IsNull(var) Then
		If var = "" Then
			Z_CLng = 0
			Exit Function
		End If
		If IsNumeric(var) Then
			var = Replace(var, " ", "")
			If Len(var)<=5 Then
				Z_CLng = CLng(Replace(var, ",", ""))
			Else
				Z_CLng = ""
				blnLeading = True
				lngZ = Len(var)
				For lngI = 1 to lngZ
					strTmp = Mid(var,lngI,1)
					If IsNumeric(strTmp) Then
						If Not blnLeading And strTmp = "0" Then
							Z_CLng = Z_CLng & strTmp
						ElseIf blnLeading and strTmp <> "0" Then
							Z_CLng = Z_CLng & strTmp
							blnLeading = False
						ElseIf Not(blnLeading) Or strTmp <> "0" Then
							Z_CLng = Z_CLng & strTmp
						End If
					Else
						Exit For
					End If
				Next
			End If
		Else
			Z_CLng = 0
		End If
	Else
		Z_CLng = 0
	End If
End Function

Function Replca(totest, backup)
	If IsNull(totest) Then
		Replca = Z_FixNull(backup)
	Else
		If Trim(totest)="" Then
			Replca = Z_FixNull(backup)
		Else
			Replca = Trim(totest)
		End If
	End If
End Function

Function Z_DoEncrypt(strZZ)
	DIM objEncrypt
	If Trim(strZZ) <> "" Then
		Set objEncrypt = Server.CreateObject("ZEnc.ZBlowfish")
		Z_DoEncrypt = objEncrypt.Encrypt3(strZZ)
		Set objEncrypt = Nothing
	Else
		Z_DoEncrypt = ""
	End If
End Function

Function Z_DoDecrypt(strZZ)
	DIM objEncrypt
	If Trim(strZZ) <> "" Then
		Set objEncrypt = Server.CreateObject("ZEnc.ZBlowfish")
		Z_DoDecrypt = objEncrypt.Decrypt3(strZZ)
		Set objEncrypt = Nothing
	Else
		Z_DoDecrypt = ""
	End If
End Function

Function Z_SQLCBool(var)
	If IsNull(var) Then
		Z_SQLCBool = 0
	ElseIf var = "" Then
		Z_SQLCBool = 0
	Else
		If CBool(var) Then
			Z_SQLCBool = 1
		Else
			Z_SQLCBool = 0
		End If
	End If
End Function

Function Z_CleanName(vntN)
DIM	strTmp
	strTmp = Replace(vntN,"""","")
	'strTmp = Replace(strTmp,"'","")
	strTmp = Replace(strTmp,"+","")
	strTmp = Replace(strTmp,"=","")
	strTmp = Replace(strTmp,"\","")
	strTmp = Replace(strTmp,"/","")
	strTmp = Replace(strTmp,"[","")
	strTmp = Replace(strTmp,"]","")
	strTmp = Replace(strTmp,";","")
	strTmp = Replace(strTmp,":","")
	strTmp = Replace(strTmp,"<","")
	strTmp = Replace(strTmp,">","")
	strTmp = Replace(strTmp,"?","")
	strTmp = Replace(strTmp,"|","")
	Z_CleanName = strTmp 
End Function
%>
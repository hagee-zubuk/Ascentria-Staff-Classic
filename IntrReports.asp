<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
DIM ts0, ts1
ts0 = Now
DIM tmpIntr(), tmpHours()
Function GetUname(xxx)
	GetUname = "N/A"
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	sqlUser = "SELECT lname, fname FROM User_T WHERE [index] = " & xxx
	rsUser.Open sqlUser, g_strCONN, 3, 1
	If Not rsUser.EOF Then
		GetUname = rsUser("lname") & ", " & rsUser("fname")
	End If
	rsUser.Close
	Set rsUser = Nothing
End Function
Function GetTrain(xxx)
	GetTrain = ""
	Set rsTrain = Server.CreateOBject("ADODB.RecordSet")
	sqlTrain = "SELECT * FROM  Training_T WHERE [index] = " & xxx
	rsTrain.Open sqlTrain, g_strCONN, 1, 3
	If Not rsTrain.EOF Then
		GetTrain = rsTrain("training")
	End If
	rsTrain.Close
	Set rsTrain = Nothing
End Function
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
Server.ScriptTimeout = 360000

DIM tmpUser

tmpDate2 = date
tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpUser = Request("selUser")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")
If Request("ctrl") = 1 Then
	RepCSV =  "ExpireDocs" & tmpdate & ".csv" 
	strMSG = "Expiring Interpreter Documents report."
	strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Document</td>" & vbCrlf & _
		"<td class='tblgrn'>Expiration Date</td>" & vbCrlf 
	CSVHead = "Last Name, First Name, Document, Expiration Date"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
	rsRep.Open sqlRep, g_strCONN,1 ,3
	y = 0
	Do Until rsRep.EOF
		tmpName = rsRep("last name") & ", " & rsRep("first name")
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		If Not IsNull(rsRep("passexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("passexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Passport</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("passexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Passport," & rsRep("passexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("driveexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("driveexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Driver's License</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("driveexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Driver's License," & rsRep("driveexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("greenexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("greenexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Green Card</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("greenexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Green Card," & rsRep("greenexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("employexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("employexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Employment Authorization</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("employexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Employment Authorization," & rsRep("employexp") & vbCrLf
				y = y + 1
			End If
		End If
		If Not IsNull(rsRep("carexp")) Then 
			If DateDiff("d", tmpDate2, rsRep("carexp")) < 15 Then 
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>Car Insurance</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("carexp") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & ",Car Insurance," & rsRep("carexp") & vbCrLf
				y = y + 1
			End If
		End If
		rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf Request("ctrl")= 2 Then
	If Request("selRep") = 1 Then 'training
		RepCSV =  "IntrTrain" & tmpdate & ".csv" 
		strMSG = "Interpreter Training"
		If Request("txtyear") <> 0 Then
			strMSG = strMSG & " for the year " & Request("txtyear")
			tmpDate1 = cdate("1/1/" & Request("txtyear"))
		End IF
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Hours</td>" & vbCrlf & _
			"<td class='tblgrn'>Training</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Date, Hours, Training"
		If IsDate(tmpDate2) Then
			tmpYear = Year(tmpDate1)
			Set rsRep = Server.CreateObject("ADODB.RecordSet")
			If Request("txtyear") <> 0 Then
				sqlRep = "SELECT * FROM IntrTraining_T, interpreter_T WHERE Year(date) = " & tmpYear & _
					" AND intrID = interpreter_T.[index] AND Active = 1 ORDER BY [last name], [first name], date"
			Else
				sqlRep = "SELECT * FROM IntrTraining_T, interpreter_T WHERE Active = 1 AND intrID = interpreter_T.[index] ORDER BY [last name], [first name], date"
			End If
			rsRep.Open sqlRep, g_strCONN, 1, 3
			Do Until rsRep.EOF
				tmpName = rsRep("last name") & ", " & rsRep("first name")
				kulay = "#FFFFFF"
				If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("date") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("Hours") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetTrain(rsRep("Type")) & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("date") & "," & rsRep("Hours") & "," & GetTrain(rsRep("Type")) & vbCrLf
				rsRep.MoveNext
			Loop
			rsRep.Close
			Set rsRep = Nothing 
		End If
	ElseIf Request("selRep") = 2 Then'eval/feed 
		RepCSV =  "IntrEval" & tmpdate & ".csv" 
		strMSG = "Interpreter Evaluation/Feedback"
		If Request("txtyear") <> 0 Then
			strMSG = strMSG & " for the year " & Request("txtyear")
			tmpDate1 = cdate("1/1/" & Request("txtyear"))
		End IF
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>ID</td>" & vbCrlf & _
			"<td class='tblgrn'>User</td>" & vbCrlf & _
			"<td class='tblgrn'>Evaluation/Feedback</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Date, ID, User, Evaluation/Feedback"

		tmpYear = Year(tmpDate1)
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM InterpreterEval_T WHERE NOT comment IS NULL "
		If Request("txtRepFrom") <> "" Then
			sqlRep = sqlRep & "AND date >= '" & Request("txtRepFrom") & "' "
			strMSG = strMSG & " from " & Request("txtRepFrom")
		End If
		If Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & "AND date <= '" & Request("txtRepTo") & "' "
			strMSG = strMSG & " to " & Request("txtRepTo")
		End If
		If Request("selIntr") > 0 Then
			sqlRep = sqlRep & "AND IntrID = " & Request("selIntr") & " "
			strMSG = strMSG & " for " & GetIntr(Request("selIntr")) & "."
		End If
		sqlRep = sqlRep & "ORDER BY appID, Date"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			If IsActive(rsRep("intrID")) Then
				If rsRep("comment") <> "" Then
					tmpName = GetIntr(rsRep("IntrID"))
					kulay = "#FFFFFF"
					If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
					strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("date") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("appid") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & GetUname(rsRep("UID")) & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("comment") & "</td></tr>" & vbCrLf
					CSVBody = CSVBody & """" & tmpName & """,""" & rsRep("date") & """,""" & _
						rsRep("appid") & """,""" & GetUname(rsRep("UID")) & """,""" & rsRep("comment") & """" & vbCrLf
					y = y + 1
				End If
			End If
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing 

	ElseIf Request("selRep") = 3 Then'docs
		RepCSV =  "IntrDocs" & tmpdate & ".csv" 
		strMSG = "Interpreter Documents" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Document</td>" & vbCrlf & _
			"<td class='tblgrn'>Number</td>" & vbCrlf & _
			"<td class='tblgrn'>Expiration Date</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Document, Number, Expiration Date"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'>" & tmpName & "</td></tr>"
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & vbCrLf
			If rsRep("ssnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Social Security</td>" & _
					"<td class='tblgrn2'>" & rsRep("ssnum") & "</td></tr>"
				CSVBody = CSVBody & "Social Security" & "," & rsRep("ssnum") & vbCrLf
			End If
			If rsRep("Passnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Passport</td>" & _
					"<td class='tblgrn2'>" & rsRep("Passnum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("Passexp") & "</td></tr>"
				CSVBody = CSVBody & "Passport" & "," & rsRep("Passnum") & "," & rsRep("Passexp") & vbCrLf
			End If
			If rsRep("drivenum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Driver's License</td>" & _
					"<td class='tblgrn2'>" & rsRep("drivenum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("driveexp") & "</td></tr>"
				CSVBody = CSVBody & "Driver's License" & "," & rsRep("Drivenum") & "," & rsRep("Driveexp") & vbCrLf
			End If
			If rsRep("employnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Employment Authorization</td>" & _
					"<td class='tblgrn2'>" & rsRep("employnum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("employexp") & "</td></tr>"
				CSVBody = CSVBody & "Employment Authorization" & "," & rsRep("employnum") & "," & rsRep("employexp") & vbCrLf
			End If
			If rsRep("greennum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Green Card</td>" & _
					"<td class='tblgrn2'>" & rsRep("greennum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("greenexp") & "</td></tr>"
				CSVBody = CSVBody & "Green Card" & "," & rsRep("greennum") & "," & rsRep("greenexp") & vbCrLf
			End If
			If rsRep("carnum") <> "" Then
				strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'>Car Insurance</td>" & _
					"<td class='tblgrn2'>" & rsRep("carnum") & "</td>" & _
					"<td class='tblgrn2'>" & rsRep("carexp") & "</td></tr>"
				CSVBody = CSVBody & "Car Insurance" & "," & rsRep("carnum") & "," & rsRep("carexp") & vbCrLf
			End If
			'strBody = strBody & "</tr>" & vbCrlf
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing
	ElseIf Request("selRep") = 4 Then'hire 
		RepCSV =  "IntrHiredDate" & tmpdate & ".csv" 
		strMSG = "Interpreter Date of Hire" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date Of Hire</td>" & vbCrlf & vbCrlf 
		CSVHead = "Last Name, First Name, Date of Hire"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1" 
		If Request("txtRepFrom") <> "" And Request("txtRepTo") = "" Then
			sqlRep = sqlRep & " AND DateHired >= '" & Request("txtRepFrom") & "'"
		End If
		If Request("txtRepFrom") = "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND DateHired <= '" & Request("txtRepTo") & "'"
		End If
		If Request("txtRepFrom") <> "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND DateHired >= '" & Request("txtRepFrom") & "' AND DateHired <= '" & Request("txtRepTo") & "'"
		End If
		sqlRep = sqlRep & " ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("DateHired") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("datehired") & vbCrLf
			y = y + 1
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing		
	ElseIf Request("selRep") = 5 Then'driver and crime
		RepCSV =  "IntrDriveCrime" & tmpdate & ".csv" 
		strMSG = "Interpreter Driver and Criminal Check" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Driver Record</td>" & vbCrlf & _
			"<td class='tblgrn'>Criminal Record</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Document, Number, Expiration Date"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("DriveDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("CrimeDate") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("drivedate") & "," & rsRep("crimedate") & vbCrLf
			y = y + 1
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing	
	ElseIf Request("selRep") = 6 Then'term 
		RepCSV =  "IntrTermDate" & tmpdate & ".csv" 
		strMSG = "Interpreter Date of Termination" 
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Date Of Termination</td>" & vbCrlf & vbCrlf 
		CSVHead = "Last Name, First Name, Date of Termination"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Interpreter_T WHERE Active = 0 " 
		If Request("txtRepFrom") <> "" And Request("txtRepTo") = "" Then
			sqlRep = sqlRep & " AND dateTerm >= '" & Request("txtRepFrom") & "'"
		End If
		If Request("txtRepFrom") = "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND dateTerm <= '" & Request("txtRepTo") & "'"
		End If
		If Request("txtRepFrom") <> "" And Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND dateTerm >= '" & Request("txtRepFrom") & "' AND dateTerm <= '" & Request("txtRepTo") & "'"
		End If
		sqlRep = sqlRep & " ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 1, 3
		Do Until rsRep.EOF
			tmpName = rsRep("last name") & ", " & rsRep("first name")
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("dateTerm") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & rsRep("dateterm") & vbCrLf
			y = y + 1
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing		
	ElseIf Request("selRep") = 7 Then'user record
		RepCSV =  "UserRecord" & tmpdate & ".csv" 
		strMSG = "User Record from " & Request("txtRepFrom") & " to " & Request("txtRepTo")
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Number of Appointments</td>" & vbCrlf & vbCrlf 
		CSVHead = "Last Name, First Name, Number of Appointments"
		
		'GET USERS
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM User_T WHERE type <> 2 ORDER BY lname, fname"
		rsUser.Open sqlUser, g_strCONN, 3, 1
		Do Until rsUser.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			tmpCreator = rsUser("fname") & " " & rsUser("lname")
			username = UCase(Trim(rsUser("fname") & " " & rsUser("lname")))
			Set rsrep = Server.CreateObject("ADODB.RecordSet")
			sqlrep = "SELECT Count(creator) AS appcount FROM History_T WHERE Upper(creator) = '" & username & "' AND dateTS >= '" & Request("txtRepFrom") & _
				"' AND dateTS <= '" & Request("txtRepTo") & "'"
			rsrep.Open sqlrep, g_strCONNHist, 3, 1
			If Not rsrep.EOF Then
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpCreator & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("appcount") & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & rsUser("lname") & "," & rsUser("fname") & "," & rsRep("appcount") & vbCrLf
				apptot = apptot + rsRep("appcount")
			End If
			rsrep.Close
			Set rsrep = Nothing
			y = y + 1
			rsUser.MoveNext
		Loop
		rsUser.Close
		Set rsUser = Nothing
		strBody = strBody & "<tr><td class='tblgrn4'>TOTAL</td><td class='tblgrn4'>" & apptot & "</td>" & vbCrLf
		CSVBody = CSVBody & ",TOTAL," & apptot & vbCrLf
	ElseIf Request("selRep") = 8 Then'PD
		RepCSV =  "PublicDefender" & tmpdate & ".csv"
		strMSG = "Public Defender Record"
		strHead = "<td class='tblgrn'>Docket Number</td>" & vbCrlf & _
			"<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Amount</td>" & vbCrlf
		CSVHead = "Docket Number, Date, Amount"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT * FROM Request_T, Institution_T WHERE InstID = Institution_T.[index] AND PD = 1 ORDER BY DocNum, appdate"
		rsRep.Open sqlRep, g_strCONN, 3, 1
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("DocNum") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td><td class='tblgrn2'><nobr>" & rsRep("PDamount") & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & rsRep("DocNum") & "," & rsRep("appdate") & "," & rsRep("PDamount") & vbCrLf
			y = y + 1
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing
	ElseIf Request("selRep") = 9 Then'client import
		ts = now
		RepCSV =  "ClientImport" & tmpdate & ".csv"
		strMSG = "Client Import Record"
		strHead = "<td class='tblgrn'>Medicaid/MCO</td>" & vbCrlf & _
			"<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>DOB</td>" & vbCrlf & _
			"<td class='tblgrn'>Gender</td>" & vbCrlf
		'CSVHead = "Company,Client ID	Division	Customer Number	Program Code	Last Name	First Name	Program Description	Address1	City	State	Zip	Sex	DOB	Client Identifier	Diag1	Diag2	Diag3	Diag4	SSN

		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		rsRep.Open "SELECT * FROM clientList_T", g_strCONN, 3, 1
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			cliname = rsRep("lname") & ", " & rsRep("fname")
			gender = "U"
			strGender = " AND [gender] IS NULL"
			gender2 = -1
			If rsRep("gender") = 0 Then 
				gender = "M"
				gender2 = 0
				strGender = " AND [gender]=0"
			ElseIf rsRep("gender") = 1 Then 
				gender = "F"
				gender2 = 1
				strGender = " AND [gender]=1"
			End If
			'save in uploaded client
			medicaid = Ucase(Trim(rsRep("medicaid")))
			'If medicaid = "" Then medicaid = Ucase(Trim(rsRep("meridian")))
			'If medicaid = "" Then medicaid = Ucase(Trim(rsRep("nhhealth")))
			'If medicaid = "" Then medicaid = Ucase(Trim(rsRep("wellsense")))
			CleanLname = Replace(Ucase(Trim(rsRep("lname"))), "'", "''")
			CleanFname = Replace(Ucase(Trim(rsRep("fname"))), "'", "''")
			Set rsCli = Server.CreateObject("ADODB.RecordSet")
			sqlCli = "SELECT * FROM clientuploaded_T WHERE lname = '" & CleanLname & "' AND fname = '" & CleanFname & _
				"' AND medicaid = '" & medicaid & "' AND dob = '" & rsRep("dob") & "' " ' & strGender
			'" AND gender = " & gender2
			rsCli.Open sqlCli, g_strCONN, 1, 3
			If rsCli.EOF Then
				rsCli.AddNew
				rsCli("lname") = Ucase(Trim(rsRep("lname")))
				rsCli("fname") = Ucase(Trim(rsRep("fname")))
				rsCli("medicaid") = medicaid
				rsCli("dob") = rsRep("DOB")
				If gender2 < 0 Then
					rsCli("gender") = vbNull
				Else
					rsCli("gender") = rsRep("gender")
				End If
				rsCli("timestamp") = ts
				rsCli.Update
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & medicaid & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & cliname & "</td><td class='tblgrn2'><nobr>" & rsRep("dob") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & gender & "</td></tr>" & vbCrLf
				CSVBody = CSVBody & """" & "LSS" & """,""" & medicaid & """,""" & "" & """,""" & medicaid & """,""" & "LB" & _
					""",""" & rsRep("lname") & """,""" & rsRep("fname") & """,""" & "" & """,""" & "261 Sheep Davis Road Suite A-1" & """,""" & _
					"Concord" & """,""" & "NH" & """,""" & "033015750" & """,""" & gender & """,""" & FixDateFormat(rsRep("DOB")) & """,""" & medicaid & _
					""",""" & "" & """,""" & "" & """,""" & "" & """,""" & "" & """,""" & "" & """" & vbCrLf
				y = y + 1
			End If
			rsCli.Close
			Set rsCli = Nothing
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing
		'DELETE client list
		Set rsDel = Server.CreateObject("ADODB.RecordSet")
		rsDel.Open "DELETE FROM clientList_T", g_strCONN, 1, 3
		Set rsDel = Nothing
	ElseIf Request("selRep") = 10 Then'Yes interpreters
		RepCSV =  "YesNoNAInterpreters" & tmpdate & ".csv"
		strMSG = "Interpreter Answers"
		strHead = "<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
			"<td class='tblgrn'>ANSWER</td>" & vbCrlf & _
			"<td class='tblgrn'>Timestamp</td>" & vbCrlf & _
			"<td class='tblgrn'>ID</td>" & vbCrlf & _
			"<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Institution</td>" & vbCrlf & _
			"<td class='tblgrn'>Department</td>" & vbCrlf
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT DISTINCT(appt_t.[IntrID]) AS myIntrID FROM appt_T, request_T WHERE request_T.[index] = appt_T.appid"
		If Request("txtRepFrom") <> "" Then
			sqlRep = sqlRep & " AND appdate >= '" & Request("txtRepFrom") & "'"
			strMSG = strMSG & " from " & Request("txtRepFrom")
		End If
		If Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND appdate <= '" & Request("txtRepTo") & "'"
			strMSG = strMSG & " to " & Request("txtRepTo")
		End If
		If Request("selIntr") > 0 Then
			sqlRep = sqlRep & "AND appt_T.IntrID = " & Request("selIntr") & " "
			strMSG = strMSG & " for " & GetIntr(Request("selIntr")) & "."
		End If
		'response.write sqlRep & "<br>"
		rsRep.Open sqlRep, g_strCONN, 3, 1
		y = 0
		Do Until rsRep.EOF
			'YES
			YesIntr = 0
			Set rsYes = Server.CreateObject("ADODB.RecordSet")
			sqlYes = "SELECT appID, appdate, facility, dept, ansTS FROM appt_T, request_T, Institution_T, dept_T WHERE request_T.[InstID] = Institution_T.[index] AND " & _
				"request_T.[index] = appt_T.appid AND DeptID = dept_T.[index] AND accept = 1 AND appt_T.IntrID = " & _
				rsRep("myIntrID")
			If Request("txtRepFrom") <> "" Then
				sqlYes = sqlYes & " AND appdate >= '" & Request("txtRepFrom") & "'"
			End If
			If Request("txtRepTo") <> "" Then
				sqlYes = sqlYes & " AND appdate <= '" & Request("txtRepTo") & "'"
			End If
			'response.write sqlYes & "<br>"
			x = 0
			yesIntr = 0
			rsYes.Open sqlYes, g_strCONN, 3, 1
			If Not rsYes.EOF Then
				'YesIntr = rsYes("myCount")
				Do Until rsYes.EOF
					kulay = "#FFFFFF"
					If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
					strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(Request("selIntr")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>YES</td>" & _
					"<td class='tblgrn2'><nobr>" & rsYes("ansTS") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsYes("appID") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsYes("appDate") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsYes("Facility") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsYes("dept") & "</td>" & _
					"</tr>"
					x = x + 1
					yesIntr = yesIntr + 1
					rsYes.MoveNext
				Loop
			End If
			rsYes.Close
			Set rsYes = Nothing
			'NO
			NoIntr = 0
			Set rsNo = Server.CreateObject("ADODB.RecordSet")
			sqlNo = "SELECT appID, appdate, facility, dept, ansTS FROM appt_T, request_T, Institution_T, dept_T WHERE request_T.[InstID] = Institution_T.[index] AND " & _
				"request_T.[index] = appt_T.appid AND DeptID = dept_T.[index] AND accept = 2 AND appt_T.IntrID = " & _
				rsRep("myIntrID")
			If Request("txtRepFrom") <> "" Then
				sqlNo = sqlNo & " AND appdate >= '" & Request("txtRepFrom") & "'"
			End If
			If Request("txtRepTo") <> "" Then
				sqlNo = sqlNo & " AND appdate <= '" & Request("txtRepTo") & "'"
			End If
			rsNo.Open sqlNo, g_strCONN, 3, 1
			If Not rsNo.EOF Then
				Do Until rsNo.EOF
					kulay = "#FFFFFF"
					If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
					strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(Request("selIntr")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>NO</td>" & _
					"<td class='tblgrn2'><nobr>" & rsNo("ansTS") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsNo("appID") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsNo("appDate") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsNo("Facility") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsNo("dept") & "</td>" & _
					"</tr>"
					x = x + 1
					NoIntr = NoIntr + 1
					rsNo.MoveNext
				Loop
			End If
			rsNo.Close
			Set rsNo = Nothing
			'NA
			NAIntr = 0
			Set rsNA = Server.CreateObject("ADODB.RecordSet")
			sqlNA = "SELECT COUNT(UID) AS myCount FROM appt_T, request_T WHERE request_T.[index] = appt_T.appid AND accept = 0 AND appt_T.IntrID = " & _
				rsRep("myIntrID")
			If Request("txtRepFrom") <> "" Then
				sqlNA = sqlNA & " AND appdate >= '" & Request("txtRepFrom") & "'"
			End If
			If Request("txtRepTo") <> "" Then
				sqlNA = sqlNA & " AND appdate <= '" & Request("txtRepTo") & "'"
			End If
			rsNA.Open sqlNA, g_strCONN, 3, 1
			If Not rsNA.EOF Then
				NAIntr = rsNA("myCount")
				kulay = "#FFFFFF"
				If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(Request("selIntr")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>COUNT</td>" & _
				"<td class='tblgrn2' colspan='5' style='text-align: left;'><nobr><i>YES: " & YesIntr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;NO: " & NoIntr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;NA: " & NAIntr & "</i></td>" & _
				"</tr>"
			End If
			rsNA.Close
			Set rsNA = Nothing
			'kulay = "#FFFFFF"
			'If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			'tmpName = GetIntr(rsRep("myIntrID"))
			'strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpName & "</td>" & vbCrLf & _
			'		"<td class='tblgrn2'><nobr>" & YesIntr & "</td>" & _
			'		"<td class='tblgrn2'><nobr>" & NoIntr & "</td>" & _
			'		"<td class='tblgrn2'><nobr>" & NAIntr & "</td>" & _
			'		"</tr>"
			rsRep.MoveNext
			y = y + 1
		Loop
		rsRep.Close
		Set rsRep = Nothing
	ElseIf Request("selRep") = 11 Then'Yes interpreters per appt
		RepCSV =  "YesNoNAInterpretersAppt" & tmpdate & ".csv"
		strMSG = "Interpreter Answers per Appointment"
		strHead = "<td class='tblgrn'>ID</td>" & vbCrlf & _
			"<td class='tblgrn'>Institution - Department</td>" & vbCrlf & _
			"<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Yes</td>" & vbCrlf & _
			"<td class='tblgrn'>No</td>" & vbCrlf & _
			"<td class='tblgrn'>No Answer</td>" & vbCrlf & _
			"<td class='tblgrn'>Assigned</td>" & vbCrlf & _
			"<td class='tblgrn'>Assigned by</td>" & vbCrlf
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT DISTINCT(appID), appdate FROM appt_T, request_T WHERE request_T.[index] = appt_T.appid AND [Status] <> 3 "
		If Request("txtRepFrom") <> "" Then
			sqlRep = sqlRep & " AND appdate >= '" & Request("txtRepFrom") & "'"
			strMSG = strMSG & " from " & Request("txtRepFrom")
		End If
		If Request("txtRepTo") <> "" Then
			sqlRep = sqlRep & " AND appdate <= '" & Request("txtRepTo") & "'"
			strMSG = strMSG & " to " & Request("txtRepTo")
		End If
		sqlRep = sqlRep & " ORDER BY appDate"
		rsRep.Open sqlRep, g_strCONN, 3, 1
		y = 0
		Do Until rsRep.EOF
			'YES
			YesIntr = ""
			Set rsYes = Server.CreateObject("ADODB.RecordSet")
			sqlYes = "SELECT appt_T.intrID AS myIntrID FROM appt_T, request_T WHERE request_T.[index] = appt_T.appid AND accept = 1 AND appt_T.appID = " & _
				rsRep("appID")
			If Request("txtRepFrom") <> "" Then
				sqlYes = sqlYes & " AND appdate >= '" & Request("txtRepFrom") & "'"
			End If
			If Request("txtRepTo") <> "" Then
				sqlYes = sqlYes & " AND appdate <= '" & Request("txtRepTo") & "'"
			End If
			'response.write sqlYes & "<br>"
			rsYes.Open sqlYes, g_strCONN, 3, 1
			Do Until rsYes.EOF 
				YesIntr = YesIntr & GetIntr(rsYes("myIntrID")) & "<br>"
				rsYes.MoveNext
			Loop
			rsYes.Close
			Set rsYes = Nothing
			'NO
			NoIntr = ""
			Set rsNo = Server.CreateObject("ADODB.RecordSet")
			sqlNo = "SELECT appt_T.intrID AS myIntrID FROM appt_T, request_T WHERE request_T.[index] = appt_T.appid AND accept = 2 AND appt_T.appID = " & _
				rsRep("appID")
			If Request("txtRepFrom") <> "" Then
				sqlNo = sqlNo & " AND appdate >= '" & Request("txtRepFrom") & "'"
			End If
			If Request("txtRepTo") <> "" Then
				sqlNo = sqlNo & " AND appdate <= '" & Request("txtRepTo") & "'"
			End If
			rsNo.Open sqlNo, g_strCONN, 3, 1
			Do Until rsNo.EOF 
				NoIntr = NoIntr & GetIntr(rsNo("myIntrID")) & "<br>"
				rsNo.MoveNext
			Loop
			rsNo.Close
			Set rsNo = Nothing
			'NA
			NAIntr = ""
			Set rsNA = Server.CreateObject("ADODB.RecordSet")
			sqlNA = "SELECT appt_T.intrID AS myIntrID FROM appt_T, request_T WHERE request_T.[index] = appt_T.appid AND accept = 0 AND appt_T.appID = " & _
				rsRep("appID")
			If Request("txtRepFrom") <> "" Then
				sqlNA = sqlNA & " AND appdate >= '" & Request("txtRepFrom") & "'"
			End If
			If Request("txtRepTo") <> "" Then
				sqlNA = sqlNA & " AND appdate <= '" & Request("txtRepTo") & "'"
			End If
			rsNA.Open sqlNA, g_strCONN, 3, 1
			Do Until rsNA.EOF 
				NAIntr = NAIntr & GetIntr(rsNA("myIntrID")) & "<br>"
				rsNa.MoveNext
			Loop
			rsNA.Close
			Set rsNA = Nothing
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			InstDept = GetInst(Z_GetInfoFROMAppID(rsRep("appID"), "instID")) & " - " & GetDept(Z_GetInfoFROMAppID(rsRep("appID"), "deptID"))
			IntrName = GetIntr(Z_GetInfoFROMAppID(rsRep("appID"), "IntrID"))
			Username = GetUsername(Z_GetInfoFROMAppID(rsRep("appID"), "assignedby"))
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appID") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & InstDept & "</td>" & _
					"<td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td>" & _
					"<td class='tblgrn2'><nobr>" & YesIntr & "</td>" & _
					"<td class='tblgrn2'><nobr>" & NoIntr & "</td>" & _
					"<td class='tblgrn2'><nobr>" & NAIntr & "</td>" & _
					"<td class='tblgrn2'><nobr>" & IntrName & "</td>" & _
					"<td class='tblgrn2'><nobr>" & Username & "</td>" & _
					"</tr>"
			rsRep.MoveNext
			y = y + 1
		Loop
		rsRep.Close
		Set rsRep = Nothing
	ElseIf Request("selRep") = 12 Then'DHHS survey
		g_strCONNDBDHHS = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=DHHSSurvey;Integrated Security=SSPI;"
		Set g_strCONNDHHS = Server.CreateObject("ADODB.Connection")
		g_strCONNDHHS.Open g_strCONNDBDHHS
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM survey_T WHERE timestamp >= '" & Request("txtRepFrom") & "' AND timestamp <= '" & Request("txtRepTo") & "'"
		rsUser.Open sqlUser, g_strCONNDHHS, 3, 1
		Do Until rsUser.EOF
			If rsUser("q1") = 1 Then q11 = q11 + 1
			If rsUser("q1") = 2 Then q12 = q12 + 1
			If rsUser("q1") = 3 Then q13 = q13 + 1
		
			If rsUser("q2") = 1 Then q21 = q21 + 1
			If rsUser("q2") = 2 Then q22 = q22 + 1
			If rsUser("q2") = 3 Then q23 = q23 + 1
			If rsUser("q2") = 4 Then q24 = q24 + 1
			
			If rsUser("q3") = 1 Then q31 = q31 + 1
			If rsUser("q3") = 2 Then q32 = q32 + 1
			If rsUser("q3") = 3 Then q33 = q33 + 1
			If rsUser("q3") = 4 Then q34 = q34 + 1
			
			If rsUser("q4") = 1 Then q41 = q41 + 1
			If rsUser("q4") = 2 Then q42 = q42 + 1
			If rsUser("q4") = 3 Then q43 = q43 + 1
			If rsUser("q4") = 4 Then q44 = q44 + 1
			
			If rsUser("q5") = 1 Then q51 = q51 + 1
			If rsUser("q5") = 2 Then q52 = q52 + 1
			If rsUser("q5") = 3 Then q53 = q53 + 1
			If rsUser("q5") = 4 Then q54 = q54 + 1
			
			If rsUser("q6") = 1 Then q61 = q61 + 1
			If rsUser("q6") = 2 Then q62 = q62 + 1
			If rsUser("q6") = 3 Then q63 = q63 + 1
			If rsUser("q6") = 4 Then q64 = q64 + 1
			
			If rsUser("q7") = 1 Then q71 = q71 + 1
			If rsUser("q7") = 2 Then q72 = q72 + 1
			If rsUser("q7") = 3 Then q73 = q73 + 1
			If rsUser("q7") = 4 Then q74 = q74 + 1
				
			If rsUser("q8") = 1 Then q81 = q81 + 1
			If rsUser("q8") = 2 Then q82 = q82 + 1
			If rsUser("q8") = 3 Then q83 = q83 + 1
			If rsUser("q8") = 4 Then q84 = q84 + 1
				
			If Trim(rsUser("comment")) <> "" Then	
				strID = rsUser("LBID2") 'Left(rsUser("LBID"), Len(rsUser("LBID") - 1))
				strCOM = strCOM & "* " & strID & " -- " & Trim(rsUser("comment")) & "<br>"
			End If
			rsUser.MoveNext
		Loop
		rsUser.Close
		Set rsUser = Nothing
		strBody = "<tr><td><p align='left'>" & _
			"How many encounters with an interpreter did you have today?<br>" & _
			"1: " & q11 & "<br>" & _
			"2-5: " & q12 & "<br>" & _
			"More than 5: " & q13 & "<br>" & _
			"<br><br>" & _
			"How would you rate the Language Bank interpreters today?<br>" & _
			"<br>" & _
			"Appear to be fluent/competent in facilitating dialog<br>" & _
			"Very satisfied: " & q21 & "<br>" & _
			"Satisfied: " & q22 & "<br>" & _
			"Dissatisfied: " & q23 & "<br>" & _
			"Very dissatisfied: " & q24 & "<br>" & _
			"<br>" & _
			"Arrive early/on time to the appointments<br>" & _
			"Very satisfied: " & q31 & "<br>" & _
			"Satisfied: " & q32 & "<br>" & _
			"Dissatisfied: " & q33 & "<br>" & _
			"Very dissatisfied: " & q34 & "<br>" & _
			"<br>" & _
			"Avoid adding personal opinions or asking personal questions<br>" & _
			"Very satisfied: " & q41 & "<br>" & _
			"Satisfied: " & q42 & "<br>" & _
			"Dissatisfied: " & q43 & "<br>" & _
			"Very dissatisfied: " & q44 & "<br>" & _
			"<br>" & _
			"Remain impartial<br>" & _
			"Very satisfied: " & q51 & "<br>" & _
			"Satisfied: " & q52 & "<br>" & _
			"Dissatisfied: " & q53 & "<br>" & _
			"Very dissatisfied: " & q54 & "<br>" & _
			"<br>" & _
			"Are professional/courteous<br>" & _
			"Very satisfied: " & q61 & "<br>" & _
			"Satisfied: " & q62 & "<br>" & _
			"Dissatisfied: " & q63 & "<br>" & _
			"Very dissatisfied: " & q64 & "<br>" & _
			"<br>" & _
			"Assist with cultural issues when appropriate<br>" & _
			"Very satisfied: " & q71 & "<br>" & _
			"Satisfied: " & q72 & "<br>" & _
			"Dissatisfied: " & q73 & "<br>" & _
			"Very dissatisfied: " & q74 & "<br>" & _
			"<br><br>" & _
			"What is your OVERALL level of satisfaction with the Language Bank interpreter(s) today?<br>" & _
			"Very satisfied: " & q81 & "<br>" & _
			"Satisfied: " & q82 & "<br>" & _
			"Dissatisfied: " & q83 & "<br>" & _
			"Very dissatisfied: " & q84 & "<br>" & _
			"<br><br>" & _
			"Comments:<br>" & _
			strCOM & _
		"</p></td></tr>"
		g_strCONNDHHS.Close
		Set g_strCONNDHHS = Nothing
	ElseIf Request("selRep") = 13 Then'Interpreter feedback 2
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM survey_T WHERE timestamp >= '" & Request("txtRepFrom") & "' AND timestamp <= '" & Request("txtRepTo") & "'"
		rsUser.Open sqlUser, g_strCONNHP, 3, 1
		Do Until rsUser.EOF
			If rsUser("q1") = 1 Then q11 = q11 + 1
			If rsUser("q1") = 2 Then q12 = q12 + 1
			If rsUser("q1") = 3 Then q13 = q13 + 1
			If rsUser("q1") = 4 Then q14 = q14 + 1
			If rsUser("q1") = 5 Then q15 = q15 + 1
			
			If rsUser("q2") = 1 Then q21 = q21 + 1
			If rsUser("q2") = 2 Then q22 = q22 + 1
			If rsUser("q2") = 3 Then q23 = q23 + 1
			If rsUser("q2") = 4 Then q24 = q24 + 1
			If rsUser("q2") = 5 Then q25 = q25 + 1
		
			If rsUser("q3") = 1 Then q31 = q31 + 1
			If rsUser("q3") = 2 Then q32 = q32 + 1
			If rsUser("q3") = 3 Then q33 = q33 + 1
			If rsUser("q3") = 4 Then q34 = q34 + 1
			If rsUser("q3") = 5 Then q35 = q35 + 1
				
			If rsUser("q4") = 1 Then q41 = q41 + 1
			If rsUser("q4") = 2 Then q42 = q42 + 1
			If rsUser("q4") = 3 Then q43 = q43 + 1
			If rsUser("q4") = 4 Then q44 = q44 + 1
			If rsUser("q4") = 5 Then q45 = q45 + 1
				
			If rsUser("q5") = 1 Then q51 = q51 + 1
			If rsUser("q5") = 2 Then q52 = q52 + 1
			If rsUser("q5") = 3 Then q53 = q53 + 1
			If rsUser("q5") = 4 Then q54 = q54 + 1
			If rsUser("q5") = 5 Then q55 = q55 + 1
			
			If rsUser("q6") = 1 Then q61 = q61 + 1
			If rsUser("q6") = 2 Then q62 = q62 + 1
			If rsUser("q6") = 3 Then q63 = q63 + 1
			If rsUser("q6") = 4 Then q64 = q64 + 1
			If rsUser("q6") = 5 Then q65 = q65 + 1
				
			If rsUser("q7") = 1 Then q71 = q71 + 1
			If rsUser("q7") = 2 Then q72 = q72 + 1
			If rsUser("q7") = 3 Then q73 = q73 + 1
			If rsUser("q7") = 4 Then q74 = q74 + 1
			If rsUser("q7") = 5 Then q75 = q75 + 1
				
			If rsUser("q8") = 0 Then q81 = q81 + 1
			If rsUser("q8") = 1 Then q82 = q82 + 1
			
			If rsUser("q9") = 0 Then q91 = q91 + 1
			If rsUser("q9") = 1 Then q92 = q92 + 1
			
			If rsUser("q10") = 0 Then q101 = q101 + 1
			If rsUser("q10") = 1 Then q102 = q102 + 1
				
			If Trim(rsUser("qcom1")) <> "" Then qcom1 = qcom1 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom1")) & "<br>" 			
			If Trim(rsUser("qcom2")) <> "" Then qcom2 = qcom2 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom2")) & "<br>"
			If Trim(rsUser("qcom3")) <> "" Then qcom3 = qcom3 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom3")) & "<br>"
			If Trim(rsUser("qcom4")) <> "" Then qcom4 = qcom4 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom4")) & "<br>"
			If Trim(rsUser("qcom5")) <> "" Then qcom5 = qcom5 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom5")) & "<br>"
			If Trim(rsUser("qcom6")) <> "" Then qcom6 = qcom6 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom6")) & "<br>"
			If Trim(rsUser("qcom7")) <> "" Then qcom7 = qcom7 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom7")) & "<br>"
			If Trim(rsUser("qcom8")) <> "" Then qcom8 = qcom8 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("qcom8")) & "<br>"
			If Trim(rsUser("q11")) <> "" Then qcom11 = qcom11 & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- " & Trim(rsUser("q11")) & "<br>"	
			
			If Trim(rsUser("phone")) <> "" Or Trim(rsUser("email")) <> "" Then
				contact = contact & "[" & rsUser("appID") & "]" & "[" & GetIntr(rsUser("IntrID")) & "] -- n " & Trim(rsUser("fname")) & " " & Trim(rsUser("lname")) & " | p " & Trim(rsUser("phone")) & " | e " & Trim(rsUser("email")) & "<br>"
			End If
			rsUser.MoveNext
		Loop
		rsUser.Close
		Set rsUser = Nothing
		strBody = "<tr><td><p align='left'>" & _
			"1) Introduced himself/herself and role of the interpreter (Hold Pre-session):<br>" & _
			"1: " & q11 & "<br>" & _
			"2: " & q12 & "<br>" & _
			"3: " & q13 & "<br>" & _
			"4: " & q14 & "<br>" & _
			"5: " & q15 & "<br>" & _
			"Comment:<br>" & _
			qcom1 & _
			"<br><br>" & _
			"2) Interpreted everything it was said (all the conversations) during the appointment:<br>" & _
			"1: " & q21 & "<br>" & _
			"2: " & q22 & "<br>" & _
			"3: " & q23 & "<br>" & _
			"4: " & q24 & "<br>" & _
			"5: " & q25 & "<br>" & _
			"Comment:<br>" & _
			qcom2 & _
			"<br><br>" & _
			"3) Able to keep up with the pace of communication:<br>" & _
			"1: " & q31 & "<br>" & _
			"2: " & q32 & "<br>" & _
			"3: " & q33 & "<br>" & _
			"4: " & q34 & "<br>" & _
			"5: " & q35 & "<br>" & _
			"Comment:<br>" & _
			qcom3 & _
			"<br><br>" & _
			"4) Maintained transparency by keeping either party ( provider or LEP client / patient ) in the loop when communicating with the other for clarification:<br>" & _
			"1: " & q41 & "<br>" & _
			"2: " & q42 & "<br>" & _
			"3: " & q43 & "<br>" & _
			"4: " & q44 & "<br>" & _
			"5: " & q45 & "<br>" & _
			"Comment:<br>" & _
			qcom4 & _
			"<br><br>" & _
			"5) Used the first person while interpreting:<br>" & _
			"1: " & q51 & "<br>" & _
			"2: " & q52 & "<br>" & _
			"3: " & q53 & "<br>" & _
			"4: " & q54 & "<br>" & _
			"5: " & q55 & "<br>" & _
			"Comment:<br>" & _
			qcom5 & _
			"<br><br>" & _
			"6) Impartiality and boundaries –did not stay alone in room with patient at any time, keeps personal opinions/feelings/believes out of the triadic setting:<br>" & _
			"1: " & q61 & "<br>" & _
			"2: " & q62 & "<br>" & _
			"3: " & q63 & "<br>" & _
			"4: " & q64 & "<br>" & _
			"5: " & q65 & "<br>" & _
			"Comment:<br>" & _
			qcom6 & _
			"<br><br>" & _
			"7) Professionalism –communicated with provider and others with respect:<br>" & _
			"1: " & q61 & "<br>" & _
			"2: " & q62 & "<br>" & _
			"3: " & q63 & "<br>" & _
			"4: " & q64 & "<br>" & _
			"5: " & q65 & "<br>" & _
			"Comment:<br>" & _
			qcom7 & _
			"<br><br>" & _
			"8) Was interpreter dressed professionally:<br>" & _
			"NO: " & q81 & "<br>" & _
			"YES: " & q82 & "<br>" & _
			"Comment:<br>" & _
			qcom8 & _
			"<br><br>" & _
			"9) Did interpreter arrive on time:<br>" & _
			"NO: " & q91 & "<br>" & _
			"YES: " & q92 & "<br>" & _
			"<br><br>" & _
			"10) Was interpreting wearing LB badge:<br>" & _
			"NO: " & q101 & "<br>" & _
			"YES: " & q102 & "<br>" & _
			"<br><br>" & _
			"11) Please feel free to provide additional comments in regards to this interpreter:<br>" & _
			qcom11 & _
			"<br><br>" & _
			"12) Please provide your contact information if you would like a follow up to your response:<br>" & _
			contact & _
			"</p></td></tr>"
	ElseIf Request("selRep") = 14 Then'Intr avg hrs 
		RepCSV =  "IntrAvgHrs" & tmpdate & ".csv"
		strMSG = "Interpreter Average Hours"
		strHead = "<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Avg. Hours</td>" & vbCrlf 
		CSVHead = "Last Name, First Name, Avg. Hours"
		Set rsRep = Server.CreateObject("ADODB.RecordSet")
		sqlRep = "SELECT overpayhrs, IntrID, payhrs, AStarttime, AEndtime, appdate FROM Request_T, Interpreter_T WHERE IntrID = Interpreter_T.[index] AND [status] = 1 "
		If Request("txtRepTo") <> "" Then
			lastday = Request("txtRepTo")
			firstday = DateAdd("d", -83, Request("txtRepTo"))
			sqlRep = sqlRep & "AND appdate <= '" & lastday & "' "
			sqlRep = sqlRep & "AND appdate >= '" & firstday & "' "
			strMSG = strMSG & " from " & firstday
			strMSG = strMSG & " to " & lastday
			'get week range
			wk1start = firstday
			wk1end = DateAdd("d", 6, wk1start)
			wk2start = DateAdd("d", 1, wk1end)
			wk2end = DateAdd("d", 6, wk2start)
			wk3start = DateAdd("d", 1, wk2end)
			wk3end = DateAdd("d", 6, wk3start)
			wk4start = DateAdd("d", 1, wk3end)
			wk4ends = DateAdd("d", 6, wk4start)
			wk5start = DateAdd("d", 1, wk4ends)
			wk5end = DateAdd("d", 6, wk5start)
			wk6start = DateAdd("d", 1, wk5end)
			wk6end = DateAdd("d", 6, wk6start)
			wk7start = DateAdd("d", 1, wk6end)
			wk7end = DateAdd("d", 6, wk7start)
			wk8start = DateAdd("d", 1, wk7end)
			wk8end = DateAdd("d", 6, wk8start)
			wk9start = DateAdd("d", 1, wk8end)
			wk9end = DateAdd("d", 6, wk9start)
			wk10start = DateAdd("d", 1, wk9end)
			wk10end = DateAdd("d", 6, wk10start)
			wk11start = DateAdd("d", 1, wk10end)
			wk11end = DateAdd("d", 6, wk11start)
			wk12start = DateAdd("d", 1, wk11end)
			wk12end = DateAdd("d", 6, wk12start)
		End If
		If Request("selIntr") > 0 Then
			sqlRep = sqlRep & "AND IntrID = " & Request("selIntr") & " "
			strMSG = strMSG & " for " & GetIntr(Request("selIntr")) & "."
		End If
		sqlRep = sqlRep & "ORDER BY [last name], [first name]"
		rsRep.Open sqlRep, g_strCONN, 3, 1
		x = 0
		Do Until rsRep.EOF
			strIntr = rsRep("IntrID")
			If rsRep("overpayhrs") Then 
				PHrs = rsRep("payhrs")
			Else
				PHrs = IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime"))
			End If
			lngIDx = SearchArraysHours(strIntr, tmpIntr)
			If lngIdx < 0 Then
				ReDim Preserve tmpIntr(x)
				ReDim Preserve tmpHours(x)
				ReDim Preserve tmpHours2(x)
				ReDim Preserve tmpHours3(x)
				ReDim Preserve tmpHours4(x)
				ReDim Preserve tmpHours5(x)
				ReDim Preserve tmpHours6(x)
				ReDim Preserve tmpHours7(x)
				ReDim Preserve tmpHours8(x)
				ReDim Preserve tmpHours9(x)
				ReDim Preserve tmpHours10(x)
				ReDim Preserve tmpHours11(x)
				ReDim Preserve tmpHours12(x)
				'ReDim Preserve tmpCount(x)
				
				tmpIntr(x) = strIntr
				tmpHours(x) = 0
				tmpHours2(x) = 0
				tmpHours3(x) = 0
				tmpHours4(x) = 0
				tmpHours5(x) = 0
				tmpHours6(x) = 0
				tmpHours7(x) = 0
				tmpHours8(x) = 0
				tmpHours9(x) = 0
				tmpHours10(x) = 0
				tmpHours11(x) = 0
				tmpHours12(x) = 0
				If rsRep("appDate") >= CDate(wk1start) And rsRep("appDate") <= CDate(wk1end) Then
					tmpHours(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk2start) And rsRep("appDate") <= CDate(wk2end) Then
					tmpHours2(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk3start) And rsRep("appDate") <= CDate(wk3end) Then
					tmpHours3(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk4start) And rsRep("appDate") <= CDate(wk4ends) Then
					tmpHours4(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk5start) And rsRep("appDate") <= CDate(wk5end) Then
					tmpHours5(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk6start) And rsRep("appDate") <= CDate(wk6end) Then
					tmpHours6(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk7start) And rsRep("appDate") <= CDate(wk7end) Then
					tmpHours7(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk8start) And rsRep("appDate") <= CDate(wk8end) Then
					tmpHours8(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk9start) And rsRep("appDate") <= CDate(wk9end) Then
					tmpHours9(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk10start) And rsRep("appDate") <= CDate(wk10end) Then
					tmpHours10(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk11start) And rsRep("appDate") <= CDate(wk11end) Then
					tmpHours11(x) = Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk12start) And rsRep("appDate") <= CDate(wk12end) Then
					tmpHours12(x) = Cdbl(PHrs)
				End If
				'tmpCount(x) = 1
				x = x + 1
			Else
				If rsRep("appDate") >= CDate(wk1start) And rsRep("appDate") <= CDate(wk1end) Then
					tmpHours(lngIdx) = tmpHours(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk2start) And rsRep("appDate") <= CDate(wk2end) Then
					tmpHours2(lngIdx) = tmpHours2(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk3start) And rsRep("appDate") <= CDate(wk3end) Then
					tmpHours3(lngIdx) = tmpHours3(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk4start) And rsRep("appDate") <= CDate(wk4ends) Then
					tmpHours4(lngIdx) = tmpHours4(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk5start) And rsRep("appDate") <= CDate(wk5end) Then
					tmpHours5(lngIdx) = tmpHours5(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk6start) And rsRep("appDate") <= CDate(wk6end) Then
					tmpHours6(lngIdx) = tmpHours6(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk7start) And rsRep("appDate") <= CDate(wk7end) Then
					tmpHours7(lngIdx) = tmpHours7(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk8start) And rsRep("appDate") <= CDate(wk8end) Then
					tmpHours8(lngIdx) = tmpHours8(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk9start) And rsRep("appDate") <= CDate(wk9end) Then
					tmpHours9(lngIdx) = tmpHours9(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk10start) And rsRep("appDate") <= CDate(wk10end) Then
					tmpHours10(lngIdx) = tmpHours10(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk11start) And rsRep("appDate") <= CDate(wk11end) Then
					tmpHours11(lngIdx) = tmpHours11(lngIdx) + Cdbl(PHrs)
				ElseIf rsRep("appDate") >= CDate(wk12start) And rsRep("appDate") <= CDate(wk12end) Then
					tmpHours12(lngIdx) = tmpHours12(lngIdx) + Cdbl(PHrs)
				End If	
				'tmpHours(lngIdx) = tmpHours(lngIdx) + Cdbl(PHrs)
				'tmpCount(lngIdx) = tmpCount(lngIdx) + 1
			End If
			rsRep.MoveNext
		Loop
		rsRep.Close
		Set rsRep = Nothing
		y = 0
		Do Until y = x
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
				'if tmpIntr(y)  = 318 Then
				'	response.write "1: " & tmpHours(y) & "<br>"
				'	response.write "2: " & tmpHours2(y) & "<br>"
				'	response.write "3: " & tmpHours3(y) & "<br>"
				'	response.write "4: " & tmpHours4(y) & "<br>"
				'	response.write "5: " & tmpHours5(y) & "<br>"
				'	response.write "6: " & tmpHours6(y) & "<br>"
				'	response.write "7: " & tmpHours7(y) & "<br>"
				'	response.write "8: " & tmpHours8(y) & "<br>"
				'	response.write "9: " & tmpHours9(y) & "<br>"
				'	response.write "10: " & tmpHours10(y) & "<br>"
				'	response.write "11: " & tmpHours11(y) & "<br>"
				'	response.write "12: " & tmpHours12(y) & "<br>"
				'End If
			'response.write y & ": " & tmpHours(y) & " / " & tmpCount(y) & "<br>"
			'avgHrs = tmpHours(y) / 7
			'avgHrs2 = tmpHours2(y) / 7
			'avgHrs3 = tmpHours3(y) / 7
			'avgHrs4 = tmpHours4(y) / 7
			'avgHrs5 = tmpHours5(y) / 7
			'avgHrs6 = tmpHours6(y) / 7
			'avgHrs7 = tmpHours7(y) / 7
			'avgHrs8 = tmpHours8(y) / 7
			'avgHrs9 = tmpHours9(y) / 7
			'avgHrs10 = tmpHours10(y) / 7
			'avgHrs11 = tmpHours11(y) / 7
			'avgHrs12 = tmpHours12(y) / 7
			'FinalAvgHrs = (avgHrs + avgHrs2 + avgHrs3 + avgHrs4 + avgHrs5 + avgHrs6 + avgHrs7 + avgHrs8 + avgHrs9 + avgHrs10 + avgHrs11 + avgHrs12) / 12 
			FinalAvgHrs = (tmpHours(y) + tmpHours2(y) + tmpHours3(y) + tmpHours4(y) + tmpHours5(y) + tmpHours6(y) + tmpHours7(y) + tmpHours8(y) + tmpHours9(y) + tmpHours10(y) + tmpHours11(y) + tmpHours12(y)) / 12 
			fontred = ""
			If Z_FormatNumber(FinalAvgHrs, 2) >= 30 Then fontred = "style='color: #ff0000;'"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2' " & fontred & " ><nobr>" & GetIntr(tmpIntr(y)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2' " & fontred & " ><nobr>" & Z_FormatNumber(FinalAvgHrs, 2) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & GetIntr(tmpIntr(y)) & "," & Z_FormatNumber(FinalAvgHrs, 2) & vbCrLf
			y = y + 1
		Loop
	ElseIf Request("selRep") = 15 Then'Intr comment
		RepCSV =  "IntrComment" & tmpdate & ".csv" 
		strMSG = "Interpreter Comment "
		strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
			"<td class='tblgrn'>Appointment ID</td>" & vbCrlf & _
			"<td class='tblgrn'>Comment</td>" & vbCrlf
		CSVHead = "Last Name, First Name, Appointment ID, Comment"
		
		'GET comment
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "select [index] AS myID, IntrComment, intrID FROM request_T where IntrComment <> '' "
		If Request("selIntr") > 0 Then
			sqlUser = sqlUser & "AND IntrID = " & Request("selIntr") & " "
			strMSG = strMSG & " for " & GetIntr(Request("selIntr")) & "."
		End If
		If Request("txtRepFrom") <> "" Then
			sqlUser = sqlUser & "AND appdate >= '" & Request("txtRepFrom") & "' "
			strMSG = strMSG & " from " & Request("txtRepFrom")
		End If
		If Request("txtRepTo") <> "" Then
			sqlUser = sqlUser & "AND appdate <= '" & Request("txtRepTo") & "' "
			strMSG = strMSG & " to " & Request("txtRepTo")
		End If
		rsUser.Open sqlUser, g_strCONN, 3, 1
		Do Until rsUser.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(rsUser("intrID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsUser("myID") & "</td><td class='tblgrn2'>" & trim(rsUser("intrComment")) & "</td></tr>" & vbCrLf
			CSVBody = CSVBody & GetIntr(rsUser("intrID")) & "," & rsUser("myID") & "," & trim(rsUser("intrComment")) & vbCrLf
			rsUser.MoveNext
		Loop
		rsUser.Close
		Set rsUser = Nothing
	ElseIf Request("selRep") = 16 Then'user assign
		RepCSV =  "UserAssign" & tmpdate & ".csv" 
		tmpUser = Z_CLng(Request("selUser"))
		strMSG = "User Assign from " & Request("txtRepFrom") & " to " & Request("txtRepTo")
		
		'GET USERS
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM User_T WHERE type <> 2 ORDER BY lname, fname"
		If tmpUser > 0 Then 
			sqlUser = "SELECT * FROM User_T WHERE [index]= " & tmpUser & " AND [type]<> 2 ORDER BY lname, fname"
		End If
		rsUser.Open sqlUser, g_strCONN, 3, 1
		Do Until rsUser.EOF
			CSVHead = ""
			y = 0
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			tmpCreator = rsUser("fname") & " " & rsUser("lname")
			username = UCase(Trim(rsUser("fname") & " " & rsUser("lname")))
			Set rsRep = Server.CreateObject("ADODB.RecordSet")
			If tmpUser > 0 Then 
				apptot = 0
				strHead = "<td class='tblgrn'>Encoded</td>" & _
						"<td class='tblgrn'>Assigned</td>" & _
						"<td class='tblgrn'>Req ID</td>" & _
						"<td class='tblgrn'>Appt Date</td>" & _
						"<td class='tblgrn'>Interpreter</td>" & _
						"<td class='tblgrn'>Language</td>" & vbCrlf & vbCrlf 
				CSVHead = "Assignment report for " & tmpCreator & vbCrLf & "Req timestamp,Assigned,Req_ID,Appt Date,Interpreter,Language"
				tmpDT = Z_YMDDate(DateAdd("d", 1, Z_CDate(Request("txtRepTo"))))
				sqlRep = "SELECT rr.[timestamp], hh.[interTS], hh.[reqID], rr.[appDate], it.[First Name], it.[Last Name] " & _
						", ln.[Language] " & _
						"FROM [HistLangBank].dbo.[History_T] AS hh " & _
						"LEFT JOIN [Langbank].dbo.[request_T] AS rr ON hh.[reqID] = rr.[index] " & _
						"INNER JOIN [Langbank].dbo.[language_T] AS ln ON rr.[langID] = ln.[index] " & _
						"INNER JOIN [Langbank].dbo.[interpreter_T] AS it ON rr.[IntrID]=it.[index] " & _
						"WHERE (hh.[interU]) LIKE '" & tmpCreator  & "' AND interTS >= '" & Z_YMDDate(Request("txtRepFrom")) & _
						"' AND interTS <= '" & tmpDT & "' ORDER BY rr.[timestamp] ASC"
				'Response.Write sqlRep
				rsRep.Open sqlrep, g_strCONN, 3, 1
				If Not rsRep.EOF Then
					Do Until rsRep.EOF
						If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
						strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
								"<td class='tblgrn2'><nobr>" & rsRep("timestamp") & "</td>" & vbCrLf & _
								"<td class='tblgrn2'><nobr>" & rsRep("interTS") & "</td>" & vbCrLf & _
								"<td class='tblgrn2'><nobr>" & rsRep("reqID") & "</td>" & vbCrLf & _
								"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
								"<td class='tblgrn2'><nobr>" & rsRep("first name") & " " & rsRep("last name") & "</nobr></td>" & vbCrLf & _
								"<td class='tblgrn2'><nobr>" & rsRep("language") & "</td></tr>" & vbCrLf
						CSVBody = CSVBody & rsRep("timestamp") & "," & _
								rsRep("interTS") & "," & rsRep("reqID") & "," & rsRep("appDate") & "," & _
								"""" & rsRep("first name") & " " & rsRep("last name") & """,""" & rsRep("language") & """" & vbCrLf
						'apptot = apptot + 1
						'Response.Write "-->" & rsRep("timestamp") & "<br />" & vbCrLf
						rsRep.MoveNext
						y = y + 1
					Loop
				End If
				rsRep.Close
				apptot = y
			Else
				strHead = "<td class='tblgrn'>Name</td>" & vbCrlf & _
						"<td class='tblgrn'>Number of Appointments</td>" & vbCrlf & vbCrlf 
				CSVHead = "Last Name, First Name, Number of Appointments"
				sqlrep = "SELECT Count(interU) AS appcount FROM History_T " & _
						"WHERE Upper(interU) = '" & username & "' AND interTS >= '" & Request("txtRepFrom") & _
						"' AND interTS <= '" & Request("txtRepTo") & "'"
				rsRep.Open sqlrep, g_strCONNHist, 3, 1
				If Not rsRep.EOF Then
					strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmpCreator & "</td>" & vbCrLf & _
							"<td class='tblgrn2'><nobr>" & rsRep("appcount") & "</td></tr>" & vbCrLf
					CSVBody = CSVBody & rsUser("lname") & "," & rsUser("fname") & "," & rsRep("appcount") & vbCrLf
					apptot = apptot + rsRep("appcount")
				End If
				rsRep.Close
			End If
			Set rsRep = Nothing
			y = y + 1
			rsUser.MoveNext
		Loop
		rsUser.Close
		Set rsUser = Nothing
		strBody = strBody & "<tr><td class='tblgrn4'>TOTAL</td><td class='tblgrn4'>" & apptot & "</td>" & vbCrLf
		CSVBody = CSVBody & ",TOTAL," & apptot & vbCrLf
	End If
	
End If
If Request("csv") <> 1 And Request("selrep") <> 12  And Request("selrep") <> 13 Then
	'CONVERT TO CSV
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
	If Request("selRep") <> 9 Then
		Prt.WriteLine "LANGUAGE BANK - REPORT"
		Prt.WriteLine strMSG
		Prt.WriteLine CSVHead
		Prt.WriteLine CSVBody
	Else
		Prt.Write CSVBody
	End If
	Prt.Close	
	Set Prt = Nothing
	
	'COPY FILE TO BACKUP
	
	fso.CopyFile RepPath & RepCSV, BackupStr
	
	Set fso = Nothing
	
	tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)	
Else
	
End If
ts1 = Now
'DateAdd
%>
<html>
	<head>
		<title>Language Bank - Admin Report</title>
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
								340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br>
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
								<td colspan='10' align='center'>
									
										<b><%=strMSG%></b>
									
								</td>
							</tr>
							<tr>
								
								<%=strHead%>
							
							</tr>
							
								<%=strBody%>
							
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td colspan='10' align='center' height='100px' valign='bottom'>
									<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
									<% If Request("selrep") <> 12 And Request("selrep") <> 13 Then %>
										<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
									<% End If %>
								</td>
							</tr>
								<td colspan='5' align='center' height='100px' valign='bottom'>
									* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
								</td>
							<tr>
							</tr>
						</table>	
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>

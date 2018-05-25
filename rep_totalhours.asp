<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
DIM tmpIntr(), tmpTown(), tmpIntrName(), tmpLang(), tmpClass(), tmpBill(), tmpAhrs(), tmpApp(), tmpInst(), tmpDept(), tmpAmt(), tmpFac(), tmpMonthYr(), tmpCtr(), tmpMonthYr2(), tmpMonthYr3()
DIM tmpMonthYr4(), tmpHrs(), tmpHHrs(), tmpMile(), tmpToll(), arrTS(), arrAuthor(), arrPage(), tmpTrain(), tmpIHTrain(), tmpbhrs(), arrBody(), tmpHrs2(), tmpHrs3(), tmpHrs4() , tmpHrs5(), tmpZip()
DIM tmpHrsHP(), tmpHrsHP2()
server.scripttimeout = 360000
%>
<%
tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")
ctr = 13

RepCSV =  "TotalHours" & tmpdate & ".csv"
strMSG = "Total Hours report"
strHead = "<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>File Number</td>" & vbCrlf & _
		"<td class='tblgrn'>Regular Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Holiday Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Over Time Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Back Hours</td>" & vbCrlf 
CSVHead = "Co Code,Batch ID,Last Name,First Name,File #,temp dept,temp rate,reg hours,o/t hours,hours 3 code" & _
		",hours 3 amount,hours 4 code,hours 4 amount,earnings 3 code,earnings 3 amount,earnings 4 code" & _
		",earnings 4 amount,earnings 5 code,earnings 5 amount,memo code,memo amount"
		' removed -- "Last Name,First Name,File Number,Regular Hours,Holiday Hours,Over Time Hours"
Set rsRep = Server.CreateObject("ADODB.RecordSet")
sqlRep = "SELECT * FROM [request_T] AS r " & _
		"INNER JOIN [interpreter_T] AS i ON r.[IntrID] = i.[index] " & _
		"WHERE (IntrID <> 0 OR intrID = -1) " & _
				"AND [Status] <> 2 " & _
				"AND [Status] <> 3 " & _
				"AND [ShowIntr] = 1 "
If tmpReport(1) <> "" Then
	sqlRep = sqlRep & "AND [appDate] >= '" & tmpReport(1) & "' "
	strMSG = strMSG & " from " & tmpReport(1)
End If
If tmpReport(2) <> "" Then
	sqlRep = sqlRep & "AND [appDate] <= '" & tmpReport(2) & "' "
	strMSG = strMSG & " to " & tmpReport(2)
End If
sqlRep = sqlRep & "ORDER BY [last name], [first name], [AppDate]"
rsRep.Open sqlRep, g_strCONN, 3, 1

x = 0
If rsRep.EOF Then
	strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
	CSVBody = "< --- No records found --- >"
End If
Do While Not rsRep.EOF
	strIntr = rsRep("IntrID")
	TT = Z_FormatNumber(rsRep("actTT"), 2)
	' i'm guessing -- PHrs = paid hours; actual
	If rsRep("overpayhrs") Then 
		PHrs = Z_FormatNumber(rsRep("payhrs"), 2)
	Else
		PHrs = Z_FormatNumber(IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime")), 2)
	End If

	If rsRep("deptID") = 1876 Then 'back hours
		FPHHrs = 0
		FPHrs = 0
		thours = 0
		ihthours = 0
		FPHrsHP = 0
		bhrs = Z_Czero(PHrs) + Z_Czero(TT)
	Else
		bhrs = 0
		If rsRep("training") = 0 Then
			ihthours = 0
			thours = 0
			If IsHoliday(rsRep("appdate")) Then
				FPHHrs = Z_Czero(PHrs) + Z_Czero(TT)
				FPHrs = 0
			Else
				''''
				If Z_EligibleHigherPay(rsRep("LangID")) Then
					FPHHrs = 0
					FPHrs = 0
					FPHrsHP = Z_Czero(PHrs) + Z_Czero(TT)
				Else
					FPHHrs = 0
					FPHrsHP = 0
					FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
				End If
			End If
		ElseIf rsRep("training") = 1 Then
			FPHHrs = 0
			FPHrs = 0
			ihthours = 0
			FPHrsHP = 0
			thours = Z_Czero(PHrs) + Z_Czero(TT)
		ElseIf rsRep("training") = 2 Then 		' In house Training
			FPHHrs = 0
			FPHrs = 0
			thours = 0
			FPHrsHP = 0 
			ihthours = Z_Czero(PHrs) + Z_Czero(TT)
		ElseIf rsRep("training") = 3 Then		' --- Interpreter Training Hours --- added 2017-12-07
			FPHHrs = 0
			FPHrs = 0
			thours = 0
			FPHrsHP = 0 
			ihthours = Z_Czero(PHrs) + Z_Czero(TT)
		End If
	End If
	lngIDx = SearchArraysHours(strIntr, tmpIntr)
	If lngIdx < 0 Then
		' NOT FOUND
		ReDim Preserve tmpIntr(x)
		ReDim Preserve tmpHrs(x)
		ReDim Preserve tmpHrs2(x)
		ReDim Preserve tmpHHrs(x)
		ReDim Preserve tmpTrain(x)
		ReDim Preserve tmpIHTrain(x)
		ReDim Preserve tmpbhrs(x)
		ReDim Preserve tmpHrsHP(x)
		ReDim Preserve tmpHrsHP2(x)
				
		tmpIntr(x) = strIntr
		If rsRep("appDate") >= Z_DateNull(tmpReport(1)) And rsRep("appDate") <= DateAdd("d", 6, tmpReport(1)) Then
			tmpHrs(x) = FPHrs
			tmpHrs2(x) = 0
			tmpHrsHP(x) = FPHrsHP
			tmpHrsHP2(x) = 0
		ElseIf rsRep("appDate") <= Z_DateNull(tmpReport(2)) And rsRep("appDate") >= DateAdd("d", -6, tmpReport(2)) Then
			tmpHrs(x) = 0
			tmpHrs2(x) = FPHrs
			tmpHrsHP(x) = 0
			tmpHrsHP2(x) = FPHrsHP
		End If
		tmpHHrs(x) = FPHHrs
		tmpTrain(x) = thours
		tmpIHTrain(x) = ihthours
		tmpbhrs(x) = bhrs
		x = x + 1
	Else	
		' interpreter already in the array!
		If rsRep("appDate") >= Z_DateNull(tmpReport(1)) And rsRep("appDate") <= DateAdd("d", 6, tmpReport(1)) Then
			' appDate is after FROM date but less than a week from it
			tmpHrs(lngIdx) = tmpHrs(lngIdx) + FPHrs
			tmpHrsHP(lngIdx) = tmpHrsHP(lngIdx) + FPHrsHP
		ElseIf rsRep("appDate") <= Z_DateNull(tmpReport(2)) And rsRep("appDate") >= DateAdd("d", -6, tmpReport(2)) Then
			' appDate is on/before TO date but less than a week from it
			tmpHrs2(lngIdx) = tmpHrs2(lngIdx) + FPHrs
			tmpHrsHP2(lngIdx) = tmpHrsHP2(lngIdx) + FPHrsHP
		End If
		tmpHHrs(lngIdx) = tmpHHrs(lngIdx) + FPHHrs
		tmpTrain(lngIdx) = tmpTrain(lngIdx) + thours
		tmpIHTrain(lngIdx) = tmpIHTrain(lngIdx) + ihthours
		tmpbhrs(lngIdx) = tmpbhrs(lngIdx) + bhrs
	End If
	rsRep.MoveNext
Loop

y = 0
Do Until y = x
	kulay = "#FFFFFF"
	If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
	TotHours = tmpHrs(y) + tmpHHrs(y) + tmpTrain(y) + tmpHrs2(y) + tmpIHTrain(y) + tmpHrsHP(y) + tmpHrsHP2(y)
	myTrain = Z_Czero(tmpTrain(y))
	myIHTrain = Z_Czero(tmpIHTrain(y))
	myBhrs = Z_Czero(tmpbhrs(y))
	myHHrs = Z_Czero(tmpHHrs(y))
	myhrs1 = Z_Czero(tmpHrs(y))
	myhrsHP1 = Z_Czero(tmpHrsHP(y))
	myhrsHP2 = Z_Czero(tmpHrsHP2(y))
	myOTHrs1 = 0
	If tmpHrs(y) > 40 Then 
		myOTHrs1 = tmpHrs(y) - 40
		myhrs1 = tmpHrs(y) - myOTHrs1
	End If
	myhrs2 = Z_Czero(tmpHrs2(y))
	myOTHrs2 = 0
	If tmpHrs2(y) > 40 Then 
		myOTHrs2 = tmpHrs2(y) - 40
		myhrs2 = tmpHrs2(y) - myOTHrs2
	End If
	myHrs = myhrs1 + myhrs2
	myOTHrs = myOTHrs1 + myOTHrs2
	myhrsHP = myhrsHP1 + myhrsHP2
	If TotHours <> 0 Or myBhrs <> 0 Then
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(tmpIntr(y)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetFileNum(tmpIntr(y)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myHrs,2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myHHrs,2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myOTHrs,2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myBhrs,2) & "</td></tr>" & vbCrLf
									
		CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",,," & _
				Z_FormatNumber(myHrs,2) & "," & Z_FormatNumber(myOTHrs,2) & vbCrLf
		If myhrsHP > 0 Then
			CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",," & _
					Z_GetHigherPay(0, tmpIntr(y)) & "," & Z_FormatNumber(myhrsHP,2) & vbCrLf
		End If
		If 	myHHrs > 0 Then
			IntrRate = Z_GetDefRate(tmpIntr(y)) * 1.5
			CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",," & _
					Z_FormatNumber(IntrRate,2) & ",0,0" & ",HWK," & Z_FormatNumber(myHHrs,2) & vbCrLf
		End If
		If myTrain > 0 Then
			CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",," & _
					Z_MinRate() & "," & Z_FormatNumber(myTrain,2) & "," & Z_FormatNumber(myOTHrs,2) & vbCrLf
		End If
		If myIHTrain > 0 Then
			CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & ",," & _
					Z_InHouseRate() & "," & Z_FormatNumber(myIHTrain,2) & "," & Z_FormatNumber(myOTHrs,2) & vbCrLf
		End If
		If myBhrs > 0 Then
			CSVBody = CSVBody & "F7M,LB," & GetIntr(tmpIntr(y)) & "," & GetFileNum(tmpIntr(y)) & _
					",,,0,0" & ",BCK," & Z_FormatNumber(myBhrs,2) & vbCrLf
		End If
	End If
	y = y + 1
Loop

rsRep.Close
Set rsRep = Nothing

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
	
	'COPY FILE TO BACKUP
	fso.CopyFile RepPath & RepCSV, BackupStr
	
	Set fso = Nothing
	'EXPORT CSV
	'If Request("bill") <> 1 Then
	tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV) & "&NF=" & Z_DoEncrypt("TotalHours_Report")
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
				<tr><td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr bgcolor='#f58426'>
									<td colspan='<%=ctr + 7%>' align='center'>
										<% If Request("bill") <> 1 Then %>
											<b><%=strMSG%></b>
										<% Else %>
											<b><%=strMSG2%></b>
										<% End If%>
									</td>
								</tr>
<tr><%=strHead%></tr>
<%=strBody%>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
									</td>
								</tr>
								<tr><td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
										* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
									</td>
								</tr>
							</table>	
						</td>
					</tr>
				</table>
		</form>
	</body>
</html>

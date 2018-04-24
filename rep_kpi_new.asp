<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
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
Function Z_InHouseRate()
	Z_InHouseRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT inHouse FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		Z_InHouseRate = rsRate("inHouse")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function


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


RepCSV =  "TotalHoursNEW" & tmpdate & ".csv"
strMSG = "Total Hours report (NEW)"
strHead = "<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>File Number</td>" & vbCrlf & _
		"<td class='tblgrn'>Regular Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Holiday Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Over Time Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Back Hours</td>" & vbCrlf 
CSVHead = "Interpreter Last Name, Interpreter First Name,Company Code,Charge Date" & _
		",File #,Temp Dept,Temp Rate,Regular Hours,OT Hours,Regular Backup Pay Code" & _
		",Regular Backup Pay Hours,OT Back Pay Code,OT Back Pay Hours,Earnings Code,Amount"
Set rsRep = Server.CreateObject("ADODB.RecordSet")
strIDT = ""
If tmpReport(1) <> "" Then
	strIDT = strIDT & "AND [appDate] >= '" & tmpReport(1) & "' "
	strMSG = strMSG & " from " & tmpReport(1)
End If
If tmpReport(2) <> "" Then
	strIDT = strIDT & "AND [appDate] <= '" & tmpReport(2) & "' "
	strMSG = strMSG & " to " & tmpReport(2)
End If

sqlRep = "SELECT [last name], [first name], [appdate], [langid], [deptid] " & _
			", [IntrID], [actTT], [overpayhrs], [payhrs] " & _
			", [AStarttime], [AEndtime], [training] " & _
			"FROM [request_T] AS r " & _
			"INNER JOIN [interpreter_T] AS i ON r.[IntrId] = i.[Index] " & _
			"WHERE ([IntrID] <> 0 OR [IntrID] = -1) " & _
					"AND STATUS <> 2 " & _
					"AND STATUS <> 3 " & _
					"AND showintr = 1 " & _
					strIDT & _
					"ORDER BY [last name], [first name], [appdate] "
rsRep.Open sqlRep, g_strCONN, 3, 1
If rsRep.EOF Then 
	strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
	CSVBody = "< --- No records found --- >"
End If

x = 0
Do While Not rsRep.EOF
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
		ElseIf rsRep("training") = 2 Then
			FPHHrs = 0
			FPHrs = 0
			thours = 0
			FPHrsHP = 0 
			ihthours = Z_Czero(PHrs) + Z_Czero(TT)
		ElseIf rsRep("training") = 3 Then
			FPHHrs = 0
			FPHrs = 0
			thours = 0
			FPHrsHP = 0 
			ihthours = Z_Czero(PHrs) + Z_Czero(TT)
		End If
	End If
	lngIDx = SearchArraysHours(strIntr, tmpIntr)
	If lngIdx < 0 Then
		' not FOUND!
		ReDim Preserve tmpIntr(x)
		ReDim Preserve tmpHrs(x)
		ReDim Preserve tmpHrs2(x)
		ReDim Preserve tmpHHrs(x)
		ReDim Preserve tmpTrain(x)
		ReDim Preserve tmpIHTrain(x)
		ReDim Preserve tmpbhrs(x)
		ReDim Preserve tmpHrsHP(x)
				
		tmpIntr(x) = strIntr
		tmpHrs(x) = FPHrs
		tmpHrsHP(x) = FPHrsHP
		tmpHHrs(x) = FPHHrs
		tmpTrain(x) = thours
		tmpIHTrain(x) = ihthours
		tmpbhrs(x) = bhrs
		x = x + 1
	Else	
		tmpHrs(lngIdx) = tmpHrs(lngIdx) + FPHrs
		tmpHrsHP(lngIdx) = tmpHrsHP(lngIdx) + FPHrsHP
		tmpHHrs(lngIdx) = tmpHHrs(lngIdx) + FPHHrs
		tmpTrain(lngIdx) = tmpTrain(lngIdx) + thours
		tmpIHTrain(lngIdx) = tmpIHTrain(lngIdx) + ihthours
		tmpbhrs(lngIdx) = tmpbhrs(lngIdx) + bhrs
	End If
	rsRep.MoveNext
Loop
rsRep.Close
Set rsRep = Nothing

y = 0
Do Until y = x
	kulay = "#FFFFFF"
	If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
	TotHours = tmpHrs(y) + tmpHHrs(y) + tmpTrain(y) + tmpIHTrain(y) + tmpHrsHP(y)
	myTrain = Z_Czero(tmpTrain(y))
	myIHTrain = Z_Czero(tmpIHTrain(y))
	myBhrs = Z_Czero(tmpbhrs(y))
	myHHrs = Z_Czero(tmpHHrs(y))
	myhrs1 = Z_Czero(tmpHrs(y))
	myhrsHP1 = Z_Czero(tmpHrsHP(y))
	myOTHrs1 = 0
	If tmpHrs(y) > 40 Then 
		myOTHrs1 = tmpHrs(y) - 40
		myhrs1 = tmpHrs(y) - myOTHrs1
	End If
	myHrs = myhrs1 
	myOTHrs = myOTHrs1 
	myhrsHP = myhrsHP1
	If TotHours > 0 Or myBhrs > 0 Then
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(tmpIntr(y)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetFileNum(tmpIntr(y)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myHrs + myhrsHP, 2) & "</td>" & vbCrLf & _ 
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myHHrs, 2) & "</td>" & vbCrLf & _ 		
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myOTHrs, 2) & "</td>" & vbCrLf & _ 		
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myBhrs, 2) & "</td></tr>" & vbCrLf 		
		If myHrs > 0 Or myOTHrs > 0 Or myBhrs > 0 Then 
			cleanOTHrs = myOTHrs
			If myOTHrs = 0 Then cleanOTHrs = ""
			If myHrs = 0 Then myHrs = "" 	
			CSVBody = CSVBody & GetIntr(tmpIntr(y)) & ",ACS," & tmpReport(2) & "," & GetFileNum(tmpIntr(y)) & ",,," & Z_FormatNumber(myHrs,2) & "," & Z_FormatNumber(cleanOTHrs,2)
			If myBhrs > 0 Then
				CSVBody = CSVBody & ",RBCK," & Z_FormatNumber(myBhrs,2)
			End If
			CSVBody = CSVBody & vbCrLf
		End If ' myHrs > 0 Or myOTHrs > 0 Or myBhrs > 0 Then 
		If myhrsHP > 0 Then
			CSVBody = CSVBody & GetIntr(tmpIntr(y)) & ",ACS," & tmpReport(2) & "," & _
					GetFileNum(tmpIntr(y)) & ",," & Z_GetHigherPay(0, tmpIntr(y)) & "," & _
					Z_FormatNumber(myhrsHP,2) & vbCrLf
		End If
		If 	myHHrs > 0 Then
			IntrRate = Z_GetDefRate(tmpIntr(y)) * 1.5
			CSVBody = CSVBody & GetIntr(tmpIntr(y)) & ",ACS," & tmpReport(2) & "," & _
					GetFileNum(tmpIntr(y)) & ",," & Z_FormatNumber(IntrRate,2) & ",0," & _
					",HWK," & Z_FormatNumber(myHHrs,2) & vbCrLf
		End If
		If myTrain > 0 Then
			cleanOTHrs = myOTHrs
			If myOTHrs = 0 Then cleanOTHrs = "" 	
			CSVBody = CSVBody & GetIntr(tmpIntr(y)) & ",ACS," & tmpReport(2) & "," & GetFileNum(tmpIntr(y)) & ",," & Z_MinRate() & "," & Z_FormatNumber(myTrain,2) & "," & Z_FormatNumber(cleanOTHrs,2) & vbCrLf
		End If
		If myIHTrain > 0 Then
			cleanOTHrs = myOTHrs
			If myOTHrs = 0 Then cleanOTHrs = "" 	
			CSVBody = CSVBody & GetIntr(tmpIntr(y)) & ",ACS," & tmpReport(2) & "," & GetFileNum(tmpIntr(y)) & ",," & Z_InHouseRate() & "," & Z_FormatNumber(myIHTrain,2) & "," & Z_FormatNumber(cleanOTHrs,2) & vbCrLf
		End If
	End If
	y = y + 1
Loop

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
	tmpstring = "CSV/" & repCSV 'add for RepCSVBill
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
	<title>Language Bank - Report Result</title>
	<link href='style.css' type='text/css' rel='stylesheet'>
	<script language='JavaScript'>
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
	</script>
</head>
<body>
<div style="width: 300px; text-align: center; margin: 0px auto 20px;">
	<img src='images/LBISLOGO.jpg' align='center' style="width: 287px; height: 67px;" />
	340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br />
	Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
</div>
		<form method='post' name='frmResult'>
			<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
				<tr><td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr bgcolor='#f58426'>
									<td colspan='<%=ctr + 7%>' align='center'>
<b><%=strMSG%></b>
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

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
Function Z_GetRequestCount(sqlDt, sqlFilt)
	Z_GetRequestCount = 0
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([appDate]) AS CTR FROM [request_T] INNER JOIN [dept_T] ON [request_T].[DeptID]=[dept_T].[index] WHERE [request_T].[instID] <> 479 " & _
				sqlDt & sqlFilt
On Error Resume Next				
	rsRef.Open sqlRef, g_strCONN, 1, 3
	Z_GetRequestCount = rsRef("CTR")
	rsRef.Close
On Error Goto 0	
	Set rsRef = Nothing
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


RepCSV =  "KPI" & tmpdate & ".csv"
strMSG = "KPI report (rev. 2018-01-17)"
CSVHead = "Interpreter Last Name, Interpreter First Name,Company Code,Charge Date" & _
		",File #,Temp Dept,Temp Rate,Regular Hours,OT Hours,Regular Backup Pay Code" & _
		",Regular Backup Pay Hours,OT Back Pay Code,OT Back Pay Hours,Earnings Code,Amount"
Set rsRep = Server.CreateObject("ADODB.RecordSet")
strIDT = ""

If tmpReport(1) <> "" Then strMSG = strMSG & " from " & tmpReport(1)
If tmpReport(2) <> "" Then strMSG = strMSG & " to " & tmpReport(2)

strHead = "<td class='tblgrn'>Classification</td>" & vbCrlf & _
		"<td class='tblgrn'>Status</td>" & vbCrlf & _
		"<td class='tblgrn'>" & tmpReport(1) & " - " & tmpReport(2) & "</td>" & vbCrlf
CSVHead = "Classification,Status," & tmpReport(1) & " - " & tmpReport(2)
tmpRef = 0
tmpCan = 0
tmpCanB = 0
tmpMis = 0
tmpMis2 = 0
tmpPen = 0
tmpCom = 0
tmpEmer = 0

DIM strClasses(3), strSeq(3)
strSeq(0) = " AND Class = 3 "
strClasses(0) = "Court"
strSeq(1) = " AND Class = 5 "
strClasses(1) = "Legal"
strSeq(2) = " AND Class = 4 "
strClasses(2) = "Medical"
strSeq(3) = " AND (Class = 1 OR Class = 2) "
strClasses(3) = "Other"
strBody = ""
CSVBody = ""

' date clause
sqlDT = " " 'AND [request_T].[DeptID]=[dept_T].[index] "
If tmpReport(1) <> "" Then sqlDT = sqlDT & " AND appDate >= '" & tmpReport(1) & "'"
If tmpReport(2) <> "" Then sqlDT = sqlDT & " AND appDate <= '" & tmpReport(2) & "'"
sqlDT = sqlDT & " "
DIM lngReqs
For lngI = 0 To 3
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>" & strClasses(lngI) & "</nobr></td>" & vbCrLf
	CSVBody = CSVBody & strClasses(lngI) & ","
	'REFERRALS
	strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td>" & vbCrLf
	CSVBody = CSVBody & "# of Referrals,"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI))
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & lngReqs  & "," & vbCrLf
	tmpRef = tmpRef + lngReqs

	'CANCELLED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments,"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=3 ")
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & lngReqs  & "," & vbCrLf
	tmpCan = tmpCan + lngReqs

	'CANCELLED BILLABLE
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments (Billable),"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=4 ")
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & lngReqs & "," & vbCrLf
	tmpCanB = tmpCanB + lngReqs

	'MISSED
	strBody = strBody & "<tr><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments Missed by Interpreters,"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=2 AND [missed]<>1 ")
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  lngReqs & "," & vbCrLf
	tmpMis = tmpMis + lngReqs
	
	'MISSED 2
	strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter,"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=2 AND [missed]=1 ")
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & lngReqs & "," & vbCrLf
	tmpMis2 = tmpMis2 + lngReqs
	
	'PENDING
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Pending Appointments,"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=0 ")
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & lngReqs & "," & vbCrLf
	tmpPen = tmpPen + lngReqs
	
	'EMERGENCY
	strBody = strBody & "<tr bgcolor='#F5F5F5'><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Emergency Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Emergency Appointments,"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [Emergency]=1 ")
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody & lngReqs & "," & vbCrLf
	tmpEmer = tmpEmer + lngReqs

	'COMLPETED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Completed Appointments,"
	lngReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=1 ")
	strBody = strBody & "<td class='tblgrn4'>" & lngReqs & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  lngReqs & "," & vbCrLf
	tmpCom = tmpCom + lngReqs
	
	' NEW ROW! 171204: Facilities Clients requesting appointments'
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Facilities Clients</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Facilities Clients,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(DISTINCT([request_T].[InstID])) AS instcnt FROM [request_T] INNER JOIN [dept_T] ON [request_T].[DeptID]=[dept_T].[index] " & _
			" WHERE [request_T].[instID] <> 479 " & _
			sqlDT & _
			strSeq(lngI)
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4'>" & rsRef("instcnt") & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("instcnt") & "," & vbCrLf
	rsRef.Close
	Set rsRef = Nothing

	' NEW ROW! 180117: How many distinct interpreters took appointments at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Interpreters Involved</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Interpreters Involved,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(	DISTINCT([request_T].[IntrID])	) AS intrcnt FROM [request_T] INNER JOIN [dept_T] ON [request_T].[DeptID]=[dept_T].[index] " & _
			"WHERE [request_T].[instID] <> 479 AND [IntrID] > 0 " & sqlDT & strSeq(lngI)
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4'>" & rsRef("intrcnt") & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("intrcnt") & "," & vbCrLf
	rsRef.Close
	Set rsRef = Nothing

	' NEW ROW! 180117: How many distinct languages provided at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Languages Requested</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Languages Requested,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(	DISTINCT([request_T].[LangID])	) AS langcnt FROM [request_T] INNER JOIN [dept_T] ON [request_T].[DeptID]=[dept_T].[index] " & _
			"WHERE [request_T].[instID] <> 479 AND [IntrID] > 0 " & sqlDT & strSeq(lngI)
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4'>" & rsRef("langcnt") & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("langcnt") & "," & vbCrLf
	rsRef.Close
	Set rsRef = Nothing

	strBody = strBody & "<tr><td>&nbsp;</td></tr>"
	CSVBody = CSVBody &  vbCrLf
Next

'''''''''''TOTALS'''''''''''''''
strBody = strBody & "<tr  bgcolor='#F5F5F5'><td class='tblgrn2'><nobr>TOTALS</td>" & vbCrLf
CSVBody = CSVBody &  "TOTALS,"
'REFERRALS
strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td><td class='tblgrn4'>" & tmpRef & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  "# of Referrals," & tmpRef & vbCrLf
'CANCELLED
strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td><td class='tblgrn4'>" & tmpCan & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Canceled Appointments," & tmpCan & vbCrLf
'CANCELLED BILLABLE
strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td><td class='tblgrn4'>" & tmpCanB & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Canceled Appointments (Billable)," & tmpCanB & vbCrLf
'MISSED
strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td><td class='tblgrn4'>" & tmpMis & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Appointments Missed by Interpreter," & tmpMis & vbCrLf
'MISSED 2
strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td><td class='tblgrn4'>" & tmpMis2 & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter," & tmpMis2 & vbCrLf
'PENDING
strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td><td class='tblgrn4'>" & tmpPen & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Pending Appointments," & tmpPen & vbCrLf
'PENDING
strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Emergency Appointments</td><td class='tblgrn4'>" & tmpEmer & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Emergency Appointments," & tmpEmer & vbCrLf
'COMLPETED
strBody = strBody & "<tr bgcolor='#F5F5F5'><td>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td><td class='tblgrn4'>" & tmpCom & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Completed Appointments," & tmpCom & vbCrLf
' NEW ROW! 171204: Facilities Clients requesting appointments'
strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Facilities Clients</td>" & vbCrLf
CSVBody = CSVBody &  ",# of Facilities Clients,"
Set rsRef = Server.CreateObject("ADODB.RecordSet")
sqlRef = "SELECT COUNT(DISTINCT([request_T].[InstID])) AS instcnt FROM [request_T] INNER JOIN [dept_T] ON [request_T].[DeptID]=[dept_T].[index] " & _
		" WHERE [request_T].[instID] <> 479 " & _
		sqlDT
rsRef.Open sqlRef, g_strCONN, 1, 3
strBody = strBody & "<td class='tblgrn4'>" & rsRef("instcnt") & "</td></tr>" & vbCrLf
CSVBody = CSVBody &  rsRef("instcnt") & "," & vbCrLf
tmpCom = tmpCom + rsRef("instcnt")
rsRef.Close
Set rsRef = Nothing
' NEW ROW! 180117: How many distinct interpreters took appointments at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Interpreters Involved</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Interpreters Involved,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(	DISTINCT([request_T].[IntrID])	) AS intrcnt FROM [request_T] INNER JOIN [dept_T] ON [request_T].[DeptID]=[dept_T].[index] " & _
			"WHERE [request_T].[instID] <> 479 AND [IntrID] > 0 " & sqlDT
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4'>" & rsRef("intrcnt") & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("intrcnt") & "," & vbCrLf
	rsRef.Close
	Set rsRef = Nothing
' NEW ROW! 180117: How many distinct languages provided at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Languages Requested</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Languages Requested,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT(	DISTINCT([request_T].[LangID])	) AS langcnt FROM [request_T] INNER JOIN [dept_T] ON [request_T].[DeptID]=[dept_T].[index] " & _
			"WHERE [request_T].[instID] <> 479 AND [IntrID] > 0 " & sqlDT
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4'>" & rsRef("langcnt") & "</td></tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("langcnt") & "," & vbCrLf
	rsRef.Close
	Set rsRef = Nothing
strBody = strBody & "<tr><td>&nbsp;</td></tr>"
CSVBody = CSVBody &  vbCrLf
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
	tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV) & "&NF=" & Z_DoEncrypt("KPI_Report.csv")
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

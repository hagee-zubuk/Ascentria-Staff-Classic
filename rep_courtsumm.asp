<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
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
Function Z_Time24(dtDate)
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = DatePart("h", dtDate)
	If lngTmp < 10 Then Z_Time24 = Z_Time24 & "0"
	Z_Time24 = Z_Time24 & lngTmp & ":"
	lngTmp = DatePart("n", dtDate)
	If lngTmp < 10 Then Z_Time24 = Z_Time24 & "0"
	Z_Time24 = Z_Time24 & lngTmp
End Function
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
function CloseMe() {
	window.close();
}
	</script>
</head>
<body>
	<div style="width: 300px; text-align: center; margin: 0px auto 20px;">
		<img src='images/LBISLOGO.jpg' align='center' style="width: 287px; height: 67px;" />
		340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br />
		Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
	</div>
<%

DIM tmpIntr(), tmpTown(), tmpIntrName(), tmpLang(), tmpClass(), tmpBill(), tmpAhrs(), tmpApp(), tmpInst(), tmpDept(), tmpAmt(), tmpFac(), tmpMonthYr(), tmpCtr(), tmpMonthYr2(), tmpMonthYr3()
DIM tmpMonthYr4(), tmpHrs(), tmpHHrs(), tmpMile(), tmpToll(), arrTS(), arrAuthor(), arrPage(), tmpTrain(), tmpIHTrain(), tmpbhrs(), arrBody(), tmpHrs2(), tmpHrs3(), tmpHrs4() , tmpHrs5(), tmpZip()
DIM tmpHrsHP(), tmpHrsHP2()
Server.scripttimeout = 3600000	' 1 hour!

tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")

tmpDate = "." & Replace(Z_YMDDate(date), "-", "")
tmpTime = "." & Replace(Replace(FormatDateTime(time, 4), ":", ""), " ", "")
tmpOver =Z_CLng(Request("override"))
ctr = 10

Set rsRep = Server.CreateObject("ADODB.RecordSet")

strMSG = "Court Appointment Summary"
strSta = Z_YMDDate(tmpReport(1))
strEnd = Z_YMDDate(tmpReport(2))
strPd = Z_MDYDate(tmpReport(1)) & " to " & Z_MDYDate(tmpReport(2))


strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Requests</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Charged</td>" & vbCrlf & _
		"</tr>"

strSQL = "SELECT req.[index], ins.[Facility] AS [Institution]" & _
		", dep.[dept] AS [Department]" & _
		", lng.[Language]" & _
		", req.[Billable]" & _
		", CASE WHEN req.[emerFee] = 1 THEN " & _
		"CASE WHEN dep.[Class] = 3 OR dep.[Class] = 5 THEN req.[Billable] * emf.[FeeLegal] + req.[TT_Inst] + req.[M_Inst] " & _
		"WHEN dep.[Class] = 1 OR dep.[Class] = 2 OR dep.[Class] = 4 THEN " & _
		"req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] + emf.[FeeOther] " & _
		"ELSE req.[InstRate] " & _
		"END " & _
		"ELSE req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] " & _
		"END AS [TotalCharged] " & _
		"FROM [request_T] AS req " & _
		"INNER JOIN [institution_T]		AS ins ON req.[InstID] =  ins.[index] " & _
		"INNER JOIN [dept_T]			AS dep ON req.deptID =  dep.[index] " & _
		"INNER JOIN [language_T]		AS lng ON req.[LangID] = lng.[index] " & _
		"INNER JOIN [EmergencyFee_T]	AS emf ON req.[index] > 0 " & _
		"WHERE dep.[class]=3 " & _
		"AND (req.[processed] IS NOT NULL OR req.[processedmedicaid] IS NOT NULL) " & _
		"AND req.[appDate] >= '" & strSta & "' " & _
		"AND req.[appDate] <= '" & strEnd & "' " & _
		"ORDER BY ins.[facility], dep.[dept], req.[appDate]"		

' Costs of foreign language and ASL interpreter services, by court location; this needs to include travel time and mileage in the cost
strRp1 = "SELECT ins.[Facility] AS [Institution]" & _
		", dep.[dept] AS [Department]" & _
		", COUNT( req.[index]) AS [svcs] " & _
		", SUM( CASE WHEN req.[emerFee] = 1 THEN " & _
		"CASE WHEN dep.[Class] = 3 OR dep.[Class] = 5 THEN req.[Billable] * emf.[FeeLegal] + req.[TT_Inst] + req.[M_Inst] " & _
		"WHEN dep.[Class] = 1 OR dep.[Class] = 2 OR dep.[Class] = 4 THEN " & _
		"req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] + emf.[FeeOther] " & _
		"ELSE req.[InstRate] " & _
		"END " & _
		"ELSE req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] " & _
		"END ) AS [TotalCharged] " & _		
		"FROM request_T AS req " & _
		"INNER JOIN [institution_T]		AS ins ON req.InstID =  ins.[index] " & _
		"INNER JOIN [dept_T]			AS dep ON req.deptID =  dep.[index] " & _
		"INNER JOIN [EmergencyFee_T]	AS emf ON req.[index] > 0 " & _
		"WHERE dep.[class]=3 " & _
		"AND (req.[processed] IS NOT NULL OR req.[processedmedicaid] IS NOT NULL) " & _
		"AND req.[appDate] >= '" & strSta & "' " & _
		"AND req.[appDate] <= '" & strEnd & "' " & _
		"GROUP BY ins.[facility], dep.[dept] " & _
		"ORDER BY ins.[facility], dep.[dept]"

' Costs of foreign language and ASL interpreter services, by language; this needs to include travel time and mileage in the cost  
strRp2 = "SELECT lng.[Language]" & _
		", COUNT( req.[index]) AS [svcs] " & _
		", SUM( CASE WHEN req.[emerFee] = 1 THEN " & _
		"CASE WHEN dep.[Class] = 3 OR dep.[Class] = 5 THEN req.[Billable] * emf.[FeeLegal] + req.[TT_Inst] + req.[M_Inst] " & _
		"WHEN dep.[Class] = 1 OR dep.[Class] = 2 OR dep.[Class] = 4 THEN " & _
		"req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] + emf.[FeeOther] " & _
		"ELSE req.[InstRate] " & _
		"END " & _
		"ELSE req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] " & _
		"END ) AS [TotalCharged] " & _
		"FROM request_T AS req " & _
		"INNER JOIN [institution_T]		AS ins ON req.InstID =  ins.[index] " & _
		"INNER JOIN [dept_T]			AS dep ON req.deptID =  dep.[index] " & _
		"INNER JOIN [language_T]		AS lng ON req.[LangID] = lng.[index] " & _
		"INNER JOIN [EmergencyFee_T]	AS emf ON req.[index] > 0 " & _
		"WHERE dep.[class]=3 " & _
		"AND (req.[processed] IS NOT NULL OR req.[processedmedicaid] IS NOT NULL) " & _
		"AND req.[appDate] >= '" & strSta & "' " & _
		"AND req.[appDate] <= '" & strEnd & "' " & _
		"GROUP BY lng.[language] " & _
		"ORDER BY lng.[language]"

' Number of services provided in foreign language and ASL interpreter services by court location and by language;
strRp3 = "SELECT ins.[Facility] AS [Institution]" & _
		", dep.[dept] AS [Department]" & _
		", lng.[Language]" & _
		", COUNT( req.[index]) AS [svcs] " & _
		", SUM( CASE WHEN req.[emerFee] = 1 THEN " & _
		"CASE WHEN dep.[Class] = 3 OR dep.[Class] = 5 THEN req.[Billable] * emf.[FeeLegal] + req.[TT_Inst] + req.[M_Inst] " & _
		"WHEN dep.[Class] = 1 OR dep.[Class] = 2 OR dep.[Class] = 4 THEN " & _
		"req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] + emf.[FeeOther] " & _
		"ELSE req.[InstRate] " & _
		"END " & _
		"ELSE req.[Billable] * req.[InstRate] + req.[TT_Inst] + req.[M_Inst] " & _
		"END ) AS [TotalCharged] " & _		
		"FROM request_T AS req " & _
		"INNER JOIN [institution_T]		AS ins ON req.InstID =  ins.[index] " & _
		"INNER JOIN [dept_T]			AS dep ON req.deptID =  dep.[index] " & _
		"INNER JOIN [language_T]		AS lng ON req.[LangID] = lng.[index] " & _
		"INNER JOIN [EmergencyFee_T]	AS emf ON req.[index] > 0 " & _
		"WHERE dep.[class]=3 " & _
		"AND (req.[processed] IS NOT NULL OR req.[processedmedicaid] IS NOT NULL) " & _
		"AND req.[appDate] >= '" & strSta & "' " & _
		"AND req.[appDate] <= '" & strEnd & "' " & _
		"GROUP BY ins.[facility], dep.[dept], lng.[Language] " & _
		"ORDER BY ins.[facility], dep.[dept], lng.[Language]"

strHTML = ""
'CONVERT TO CSV
Set fso = CreateObject("Scripting.FileSystemObject")
RepCSVZ =  "CourtRows" & tmpDate & tmpTime & ".csv" 
Set Prt = fso.CreateTextFile(RepPath &  RepCSVZ, True)
Prt.WriteLine "LANGUAGE BANK - COURT APPOINTMENT LIST REPORT: " & strPd
Prt.WriteLine """ID"", ""Institution"", ""Department"", ""Language"", ""Billable"", ""TotalCharged""" & vbCrLf
'ins.[Facility] AS [Institution]" & _
''		", dep.[dept] AS [Department]" & _
''		", lng.[Language]" & _
''		", COUNT(req.[index] ) as [Requests]" & _
''		", SUM( CASE WHEN req.[emerFee] = 1 THEN " & _
''		"END ) AS [TotalCharged]'

Set rsRep = Server.CreateObject("ADODB.RecordSet")
rsRep.Open strSQL, g_strCONN, 3, 1
Do Until rsRep.EOF
	Prt.WriteLine rsRep("index")  & ", """ & rsRep("institution") & """, """ & rsRep("department") & """, """ & rsRep("language") & _
			"""," & rsRep("billable") & "," & rsRep("TotalCharged") & vbCrLf
	rsRep.MoveNext
Loop
rsRep.Close
Set rsRep = Nothing
Prt.Close
Set Prt = Nothing

RepCSV1 =  "CourtSummByLoc" & tmpDate & tmpTime & ".csv" 
Set Prt = fso.CreateTextFile(RepPath &  RepCSV1, True)
Prt.WriteLine "LANGUAGE BANK - SUMMARY OF COURT APPOINTMENTS BY LOCATION: " & strPd
Prt.WriteLine """Institution"", ""Department"", ""Requests"", ""TotalCharged""" & vbCrLf
'ins.[Facility] AS [Institution]" & _
''		", dep.[dept] AS [Department]" & _
''		", COUNT( req.[index]) AS [svcs] " & _
''		", SUM( CASE WHEN req.[emerFee] = 1 THEN " & _
''		"END ) AS [TotalCharged] " &
Set rsRep = Server.CreateObject("ADODB.RecordSet")
rsRep.Open strRp1, g_strCONN, 3, 1
dblGrdCharged = 0
lngGrdReqests = 0
x = 0
Do Until rsRep.EOF
	kulay = "#FFFFc5"
	If Not Z_IsOdd(x) Then kulay = "#c5c5c5"
	strHTML = strHTML & "<tr><td style=""background-color: " & kulay & ";"">" & rsRep("institution") & "</td>" & _
			"<td style=""background-color: " & kulay & ";"">" & rsRep("department") & "</td>" & _
			"<td style=""background-color: " & kulay & "; text-align: center; "">" & rsRep("svcs") & "</td>" & _
			"<td style=""background-color: " & kulay & "; text-align: right; padding-right: 5px;"">" & _
			Z_FormatNumber(rsRep("TotalCharged"), 2) & "</td></tr>" & vbCrLf
	Prt.WriteLine """" & rsRep("institution") & """, """ & rsRep("department") & """," & _
			rsRep("svcs") & "," & rsRep("TotalCharged") & vbCrLf
	dblGrdCharged = dblGrdCharged + Z_CDbl( rsRep("TotalCharged") )
	lngGrdReqests = lngGrdReqests + Z_CLng( rsRep("svcs") )
	x = x + 1
	rsRep.MoveNext
Loop
rsRep.Close
Set rsRep = Nothing
Prt.Close
Set Prt = Nothing

RepCSV2 =  "CourtSummByLang" & tmpDate & tmpTime & ".csv" 
Set Prt = fso.CreateTextFile(RepPath &  RepCSV2, True)
Prt.WriteLine "LANGUAGE BANK - SUMMARY OF COURT APPOINTMENTS BY LANGUAGE: " & strPd
Prt.WriteLine """Language"", ""Requests"", ""TotalCharged""" & vbCrLf
'lng.[Language]" & _
''		", COUNT( req.[index]) AS [svcs] " & _
''		", SUM( CASE WHEN req.[emerFee] = 1 THEN " & _
''		"END ) AS [TotalCharged]
Set rsRep = Server.CreateObject("ADODB.RecordSet")
rsRep.Open strRp2, g_strCONN, 3, 1
Do Until rsRep.EOF
	Prt.WriteLine """" & rsRep("language") & """," & rsRep("svcs") & "," & rsRep("TotalCharged") & vbCrLf
	rsRep.MoveNext
Loop
rsRep.Close
Set rsRep = Nothing
Prt.Close
Set Prt = Nothing

RepCSV3 =  "CourtSummByLang" & tmpDate & tmpTime & ".csv" 
Set Prt = fso.CreateTextFile(RepPath &  RepCSV3, True)
Prt.WriteLine "LANGUAGE BANK - SUMMARY OF COURT APPOINTMENTS BY LOCATION & LANGUAGE: " & strPd
Prt.WriteLine """Institution"", ""Department"", ""Language"", ""Requests"", ""TotalCharged""" & vbCrLf
'ins.[Facility] AS [Institution]" & _
''		", dep.[dept] AS [Department]" & _
''		", lng.[Language]" & _
''		", COUNT( req.[index]) AS [svcs] " & _
'		"END ) AS [TotalCharged]
Set rsRep = Server.CreateObject("ADODB.RecordSet")
rsRep.Open strRp3, g_strCONN, 3, 1
Do Until rsRep.EOF
	Prt.WriteLine """" & rsRep("institution") & """, """ & rsRep("department") & """, """ & rsRep("language") & _
			"""," & rsRep("svcs") & "," & rsRep("TotalCharged") & vbCrLf
	rsRep.MoveNext
Loop
rsRep.Close
Set rsRep = Nothing
Prt.Close
Set Prt = Nothing

%>
		<form method='post' name='frmResult'>
			<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
				<tr><td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr bgcolor='#f58426'>
									<td colspan="4" align="center">
<b><%=strMSG%></b>
									</td></tr>
<tr><th>Institution</th><th>Department</th><th>Appts</th><th>Charged</th></tr>
<%=strHTML%>
							</table>
<%



Set fso = Nothing
'EXPORT CSV
'If Request("bill") <> 1 Then
tmpD = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSVZ)
tmp1 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSV1)
tmp2 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSV2)
tmp3 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSV3)

%>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<input class="btn" type="button" value="Appt Rows CSV"
												onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'"
												onclick="document.location='<%=tmpD%>';">
										<input class="btn" type="button" value="Summ by Loc CSV"
												onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'"
												onclick="document.location='<%=tmp1%>';">
										<input class="btn" type="button" value="Summ by Lang CSV"
												onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'"
												onclick="document.location='<%=tmp2%>';" />
										<input class="btn" type="button" value="Loc/Lang CSV"
												onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'"
												onclick="document.location='<%=tmp3%>';" />
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

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
tmpdate = Replace(Date, "/", "") 
tmpTime = Replace(FormatDateTime(Time, 3), ":", "")
tmpOver =Z_CLng(Request("override"))
ctr = 10

RepCSV =  "Timesheet" & tmpdate & ".csv"
tmpMonthYear = MonthName(Month(tmpReport(1))) & " - " & Year(tmpReport(1))
mysundate = GetSun(tmpReport(1))
mysatdate = GetSat(tmpReport(1))
strMSG = "Timsheet report "'for the week of " & mysundate & " - " & mysatdate
strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
	"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
	"<td class='tblgrn'>Rate</td>" & vbCrlf & _
	"<td class='tblgrn'>Language</td>" & vbCrlf & _
	"<td class='tblgrn'>Activity</td>" & vbCrlf & _
	"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
	"<td class='tblgrn'>Appt. Start Time</td>" & vbCrlf & _
	"<td class='tblgrn'>Appt. End Time</td>" & vbCrlf & _
	"<td class='tblgrn'>Total Hours</td>" & vbCrlf & _
	"<td class='tblgrn'>Payable Hours</td>" & vbCrlf & _
	"<td class='tblgrn'>Final Payable Hours</td>" & vbCrlf 
CSVHead = """Date"",""Last Name"",""First Name"",""Rate"",""Language"",""Activity"",""Travel Time""" & _
		",""Appt. Start Time"",""Appt. End Time"",""Total Hours"",""Payable Hours"",""Final Payable Hours"""
Set rsRep = Server.CreateObject("ADODB.RecordSet")
strPD = ""
sqlRep = "SELECT lan.[Language], ins.[Facility], itr.[Last Name], itr.[First Name], itr.[Rate]" & _
		", req.[AStarttime], req.[AEndtime], req.[appDate], req.[Cfname], req.[totalhrs], req.[actTT], req.[overpayhrs]" & _
		", req.[payhrs], itr.[index] as myintrID, req.InstID " & _
		"FROM [Request_T] AS req " & _
		"INNER JOIN [Interpreter_T] AS itr ON req.[IntrID]=itr.[index] " & _
		"INNER JOIN [institution_T] AS ins ON req.[InstID]=ins.[Index] " & _
		"INNER JOIN [language_T] AS lan ON req.[langID]=lan.[index]  " & _
		"WHERE req.[showintr] = 1  " & _
		"AND req.[LBconfirm] = 1 "
If tmpReport(1) <> "" Then
	sqlRep = sqlRep & " AND req.[appDate] >= '" & tmpReport(1) & "'"
	strMSG = strMSG & " from " & tmpReport(1)
End If
If tmpReport(2) <> "" Then
	sqlRep = sqlRep & " AND req.[appDate] <= '" & tmpReport(2) & "'"
	strMSG = strMSG & " to " & tmpReport(2)
End If
If Z_CZero(tmpReport(4)) > 0 Then
	sqlRep = sqlRep & " AND req.[IntrID] = " & tmpReport(4) & " "
	strMSG = strMSG & " for " & GetIntr(tmpReport(4)) & "."
End If
sqlRep = sqlRep & " ORDER BY itr.[last name], itr.[first name], req.[appDate]"

'CONVERT TO CSV
Set fso = CreateObject("Scripting.FileSystemObject")
Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
Prt.WriteLine "LANGUAGE BANK - REPORT"
Prt.WriteLine strMSG
Prt.WriteLine CSVHead
%>
		<form method="post" name="frmResult" id="frmResult">
			<table cellSpacing="0" cellPadding="0" width="100%" bgColor='white' border="0">
				<tr><td valign="top" >
							<table bgColor="white" border="0" cellSpacing="2" cellPadding="0" align="center">
								<tr bgcolor="#f58426">
									<td colspan="11" align="center">
<b><%=strMSG%></b>
									</td>
								</tr>
<tr><%=strHead%></tr>
<%
rsRep.Open sqlRep, g_strCONN, 3, 1

y = 0
IntrID2 = ""
totHrs = 0
Do Until rsRep.EOF
	kulay = "#FFFFFF"
	If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
	IntrName = rsRep("Last Name") & ", " & rsRep("First Name")
	CliName = rsRep("Facility") '& " - " & rsRep("Cfname")
	tmpAMTs = rsRep("totalhrs")
	TT = Z_FormatNumber(rsRep("actTT"), 2)
	If rsRep("overpayhrs") Then 
		PHrs = Z_FormatNumber(rsRep("payhrs"), 2)
		OvrHrs = "*"
	Else
		PHrs = Z_FormatNumber(IntrBillHrs(rsRep("AStarttime"), rsRep("AEndtime")), 2)
		OvrHrs = ""
	End If
	FPHrs = Z_Czero(PHrs) + Z_Czero(TT)
	If Z_CZero(tmpReport(4)) > 0 Then
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & IntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(rsRep("Rate"), 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Language") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & TT & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AEndtime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpAMTs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(PHrs, 2) & OvrHrs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(FPHrs, 2) & "</td></tr>" & vbCrLf 
	Else
		IntrID = rsRep("myintrID")
			
		If IntrID <> IntrID2 And IntrID2 <> "" Then
			strBody = strBody & "<tr bgcolor='#FFFFCE'><td colspan='10' class='tblgrn2'>&nbsp;</td><td class='tblgrn2'>" & _
					Z_FormatNumber(totHrs,2) & "</td></tr>"
			If IntrID2 <> "" Then strBody = strBody & "<P CLASS='pagebreakhere'>"
			totHrs = 0
		End If
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & IntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(rsRep("Rate"), 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Language") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & TT & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CTime(rsRep("AEndtime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpAMTs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(PHrs, 2) & OvrHrs & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(FPHrs, 2) & "</td></tr>" & vbCrLf 
		IntrID2 = IntrID
	End If
	totHrs = totHrs + Z_CZero(FPHrs)
	CSVBody = CSVBody & """" & rsRep("appDate") & """,""" & rsRep("Last Name") & """,""" & rsRep("First Name") & """," & _
			Z_FormatNumber(rsRep("Rate"), 2) & ",""" & rsRep("Language") & """,""" & _
			CliName & """,""" & TT & """,""" & CTime(rsRep("AStarttime")) & """,""" & CTime(rsRep("AEndtime")) & _
			""",""" & tmpAMTs & """,""" & Z_FormatNumber(PHrs, 2) & OvrHrs & """,""" & Z_FormatNumber(FPHrs, 2) & """" & vbCrLf
	y = y + 1
	rsRep.MoveNext

	Response.Write strBody
	Response.Flush
	strBody = ""
	x = x + 1
	If Not Response.IsClientConnected Then 
		' Response.End
		Exit Do
	End If
	Prt.Write CSVBody
	'Prt.Flush
	CSVBody = ""
Loop
rsRep.Close
Set rsRep = Nothing
strBody = strBody & "<tr bgcolor='#FFFFCE'><td colspan='10' class='tblgrn2'>&nbsp;</td><td class='tblgrn2'>" & Z_FormatNumber(totHrs,2) & "</td></tr>"

Prt.WriteLine CSVBody
Prt.Close	
Set Prt = Nothing

'COPY FILE TO BACKUP
fso.CopyFile RepPath & RepCSV, BackupStr

Set fso = Nothing
'EXPORT CSV
'If Request("bill") <> 1 Then
tmpstring = "CSV/" & repCSV 'add for RepCSVBill
tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)

%>
<%=strBody%>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan="11" align="center" height="100px" valign="bottom">
										<input class="btn" type="button" value="Print" onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"
												onclick="print({bShrinkToFit: true});">
										<input class="btn" type="button" value="CSV Export" onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
									</td>
								</tr>
								<tr><td colspan="11" align="center" height="100px" valign="bottom">
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

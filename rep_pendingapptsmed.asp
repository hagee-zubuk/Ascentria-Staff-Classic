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
Server.scripttimeout = 360000

tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")

ctr = 10
RepCSV =  "PendingApptsMedicaid" & tmpdate & ".csv" 
Set rsRep = Server.CreateObject("ADODB.RecordSet")
sqlRep = "SELECT req.[index] AS [Request ID], ins.[Facility], dep.[dept] AS [Department]" & _
		", lan.[Language], req.[Clname], req.[Cfname], itr.[Last Name] , itr.[First Name]" & _
		", req.[appDate], req.[appTimeFrom], req.[appTimeTo], req.[InstRate]" & _
		", COALESCE(req.[Medicaid], '') AS [Medicaid]" & _
		", REPLACE(REPLACE(RTRIM(SUBSTRING(req.[Comment], 0, 200)), CHAR(13), ''), CHAR(10), '') AS [Comments] " & _
		"FROM [request_T] AS req " & _
		"INNER JOIN [Interpreter_T] AS itr ON req.[IntrID]=itr.[index] " & _
		"INNER JOIN [institution_T] AS ins ON req.[InstID]=ins.[index] " & _
		"INNER JOIN [Dept_T] AS dep ON req.[DeptID]=dep.[index] " & _
		"LEFT JOIN [Language_T] AS lan ON req.[LangID]=lan.[index] " & _
		"WHERE req.[instID] <> 479 " & _
		"AND req.[Medicaid] IS NOT NULL " & _
		"AND req.[Medicaid] <> '' " & _
		"AND req.[Status] = 0 "
strMSG = "Pending appointment report"
strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Medicaid</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
CSVHead = "Request ID, Institution, Department,Language, Client Last Name, Client First Name, Interpreter Last Name" & _
		", Interpreter First Name, Appointment Date, Appointment Start Time, Appointment End Time, Rate, Medicaid" & _
		", Comments"
If tmpReport(1) <> "" Then
	sqlRep = sqlRep & " AND req.[appDate] >= '" & tmpReport(1) & "' "
	strMSG = strMSG & " from " & tmpReport(1)
End If
If tmpReport(2) <> "" Then
	sqlRep = sqlRep & " AND req.[appDate] <= '" & tmpReport(2) & "' "
	strMSG = strMSG & " to " & tmpReport(2)
End If
sqlRep = sqlRep & " ORDER BY ins.[Facility], req.[appDate], req.[Clname], req.[Cfname]"

rsRep.Open sqlRep, g_strCONN, 1, 3	
If Not rsRep.EOF Then
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
		fltRate = Z_FormatNumber(rsRep("InstRate"), 2)
		strFrom = ctime(rsRep("appTimeFrom"))
		strTo = ctime(rsRep("appTimeTo"))
		strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("Request ID") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("Request ID") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""text-align: left;""><nobr>" & rsRep("Facility") & " - " & rsRep("Department") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Language") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""text-align: left;""><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""text-align: left;""><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strFrom & " - " & strTo &"</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""text-align: right;""><nobr>" & fltRate & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Medicaid") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""text-align: left;""><nobr>" & rsRep("Comments") & "</td>" & _
				"</tr>" & vbCrLf
		CSVBody = CSVBody & rsRep("Request ID") & "," & rsRep("Facility") & "," &  Replace(rsRep("Department"), " - ", "") & "," & _
				rsRep("Language") & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & rsRep("Last Name") & "," & _
				rsRep("First Name") & ","  & rsRep("appDate") & "," & strFrom & "," & strTo & "," & _
				fltRate & ",""" & rsRep("Medicaid") & """,""" & rsRep("Comments") & """" &  vbCrLf
		rsRep.MoveNext
		x = x + 1
	Loop
Else
	strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
	CSVBody = "< --- No records found --- >"
End If
rsRep.Close
Set rsRep = Nothing

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
	tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)
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

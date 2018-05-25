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

DIM tmpReport, tmpDate, tmpTime, RepCSV, strMSG, CSVHead, strHead, rsInst, strIDT, sqlDT, sqlInst, x, y, z, kulay
Server.ScriptTimeout = 360000


tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")


RepCSV =  "BillablesMA" & tmpdate & ".csv"
strMSG = "Billable Appointments Report (rev. 2018-01-23)"
CSVHead = "Interpreter Last Name, Interpreter First Name,Company Code,Charge Date" & _
		",File #,Temp Dept,Temp Rate,Regular Hours,OT Hours,Regular Backup Pay Code" & _
		",Regular Backup Pay Hours,OT Back Pay Code,OT Back Pay Hours,Earnings Code,Amount"

If tmpReport(1) <> "" Then strMSG = strMSG & " from " & tmpReport(1)
If tmpReport(2) <> "" Then strMSG = strMSG & " to " & tmpReport(2)
sqlDT = " " 'AND [request_T].[DeptID]=[dept_T].[index] "
If tmpReport(1) <> "" Then sqlDT = sqlDT & " AND appDate >= '" & tmpReport(1) & "'"
If tmpReport(2) <> "" Then sqlDT = sqlDT & " AND appDate <= '" & tmpReport(2) & "'"
sqlDT = sqlDT & " "

strMSG = strMSG & " for the state of Massachusetts."

strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf & _
		"<td class='tblgrn'>Amount</td>" & vbCrlf
CSVHead = "Institution,Count,Amount"
Set rsInst = CreateObject("ADODB.RecordSet")
sqlInst = "SELECT (r.[instID]) AS myINST, [Facility], COUNT(r.[Index]) AS tot FROM request_T AS r " & _
		"INNER JOIN [institution_T] AS i ON  r.[InstID] = i.[index] " & _
		"INNER JOIN [dept_T] AS d ON r.[DeptID]=d.[Index] " & _
		"WHERE UPPER(BState)='MA' AND [status]<>2 AND [status] <> 3 " & _
		sqlDT & " AND r.[instID] <> 479 GROUP BY r.[InstID], [Facility]"
sqlInst = "EXEC [dbo].[spRepBillable_MA] '" & Z_YMDDate(tmpReport(1)) & "', '" & Z_YMDDate(tmpReport(2)) & "'"
'Response.Write "<!-- " & sqlInst & " -->" & vbCrLf
rsInst.Open sqlInst, g_strCONN, 3, 1
x = 0
y = 0
z = 0
If Not rsInst.EOF Then
	Do Until rsInst.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"

		strBody = strBody & "<tr bgcolor='" & kulay & "'>" & _
				"<td class='tblgrn2' style=""text-align: right;""><nobr>" & rsInst("Facility") & "</td>" & _
				"<td class='tblgrn2'><nobr>" & rsInst("tot") & "</td>" & _
				"<td class='tblgrn2' style=""text-align: right;""><nobr>" & Z_FormatNumber(rsInst("tot_billable"), 2) & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & """" & rsInst("Facility") & """," & rsInst("tot") & ",""" & Z_FormatNumber(rsInst("tot_billable"), 2) & """" & vbCrLf
		x = x + 1
		y = y + Z_CLng(rsInst("tot"))
		z = z + CDbl(rsInst("tot_billable"))
		rsInst.MoveNext
	Loop
	strBody = strBody & "<tr><td align='center'><i>" & x & " records found</i></td>" & _
			"<td align=""center"">" & y & "</td><td align=""right"">" & Z_FormatNumber(z, 2) & "</td></tr>" & vbCrLf
Else
	strBody = "<tr><td colspan='3' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
	CSVBody = "< --- No records found --- >"
End If
rsInst.Close
Set rsInst = Nothing


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

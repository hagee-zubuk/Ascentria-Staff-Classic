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

RepCSV =  "NoTotalHours" & tmpdate & ".csv"
strMSG = "No Total Hours report"
strHead = "<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>File Number</td>" & vbCrlf & _
		"<td class='tblgrn'>Regular Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Holiday Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Over Time Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Back Hours</td>" & vbCrlf 
CSVHead = "Co Code,Batch ID,Last Name,First Name,File #,temp dept,temp rate,reg hours,o/t hours" & _
		",hours 3 code,hours 3 amount,hours 4 code,hours 4 amount,earnings 3 code,earnings 3 amount" & _
		",earnings 4 code,earnings 4 amount,earnings 5 code,earnings 5 amount,memo code,memo amount"
		' removed: "Last Name,First Name,File Number,Regular Hours,Holiday Hours,Over Time Hours"
Set rsIntrL = Server.CreateObject("ADODB.RecordSet")
Set rsRep = Server.CreateObject("ADODB.RecordSet")
sqlDT = ""
If tmpReport(1) <> "" Then
	sqlDT = sqlDT & "AND appDate >= '" & tmpReport(1) & "' "
End If
If tmpReport(2) <> "" Then
	sqlDT = sqlDT & "AND appDate <= '" & tmpReport(2) & "' "
End If

strSQL = "SELECT "
rsIntrL.Open "SELECT [index] AS intrID FROM interpreter_T ORDER BY [last name], [first name]", g_strCONN, 3, 1


Do Until rsIntrL.EOF
	sqlRep = "SELECT intrID FROM request_T WHERE STATUS <> 2 AND STATUS <> 3 AND showintr = 1 AND intrid = " & rsIntrL("intrID") & " " & sqlDT
	rsRep.Open sqlRep, g_strCONN, 3, 1
	If rsRep.EOF Then 
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetIntr(rsIntrL("intrID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetFileNum(rsIntrL("intrID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>0.00</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>0.00</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>0.00</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>0.00</td></tr>" & vbCrLf
		CSVBody = CSVBody & "F7M,LB," & GetIntr(rsIntrL("intrID")) & "," & GetFileNum(rsIntrL("intrID")) & ",,," & "0.00" & "," & "0.00" & vbCrLf
	End If
	rsRep.Close
	rsIntrL.MoveNext
Loop
rsIntrL.Close
Set rsIntrL = Nothing
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

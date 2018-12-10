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
tmpdate = Replace(Date, "/", "") 
tmpTime = Replace(FormatDateTime(Time, 3), ":", "")
tmpOver =Z_CLng(Request("override"))
ctr = 10
RepCSV =  "IntrFreqs" & tmpdate & tmpTime & ".csv" 
tmpDate = Z_YMDDate(Date)
Set rsRep = Server.CreateObject("ADODB.RecordSet")

strMSG = "Interpreter Utilization Frequency Report"
strHead = "<td class='tblgrn' colspan=""2"">Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Completed</td>" & vbCrlf & _
		"<td class='tblgrn'>Pending</td>" & vbCrlf & _
		"<td class='tblgrn'>Missed</td>" & vbCrlf & _
		"<td class='tblgrn'>Cancelled</td>" & vbCrlf & _
		"<td class='tblgrn'>Cancelled/Billable</td>" & vbCrlf
CSVHead = """Interpreter Last Name"",""Interpreter First Name"",""Total"",""Completed"",""Pending"",""Cancelled"",""Cancelled/Billable"""
strPD = ""
If tmpReport(1) <> "" Then
	strPD = strPD & " '" & Z_YMDDate(tmpReport(1)) & "' "
	strMSG = strMSG & " from " & tmpReport(1)
Else
	strPD = strPD & " '" & Z_YMDDate(Date) & "' "
	strMSG = strMSG & " from " & Z_MDYDate(Date)
End If
If tmpReport(2) <> "" Then
	strPD = strPD & ", '" & Z_YMDDate(tmpReport(2)) & "' "
	strMSG = strMSG & " to " & tmpReport(2)
Else
	strPD = strPD & ", '" & Z_YMDDate(Date) & "' "
	strMSG = strMSG & " to " & Z_MDYDate(Date)
End If

'CONVERT TO CSV
Set fso = CreateObject("Scripting.FileSystemObject")
Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
Prt.WriteLine "LANGUAGE BANK - REPORT"
Prt.WriteLine strMSG
Prt.WriteLine CSVHead
%>
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
<%
sqlRep = "EXEC [dbo].[RepCountIntrFreq] " & strPD ' '2018-08-15', '2018-08-17'
rsRep.Open sqlRep, g_strCONN, 3, 1
x = 1
If Not rsRep.EOF Then
	strOldID = ""
	Do Until rsRep.EOF
		kulay = "#FFFFc5"
		If Not Z_IsOdd(x) Then kulay = "#c5c5c5"
		strBody = strBody & "<tr style=""background-color: " & kulay & ";"" >" & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("last name")  & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("first name")  & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("total_appts") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("completed") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("pending") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("missed") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("cancelled") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";"">" & rsRep("canc_billable") & "</td>" & vbCrLf & _
				"</tr>" & vbCrLf
		CSVBody = CSVBody & rsRep("last name") & "," & rsRep("first name") & "," & _
				rsRep("total_appts") & "," & rsRep("completed") & "," & _
				rsRep("pending") & "," & rsRep("missed") & "," & _
				rsRep("cancelled") & "," & rsRep("canc_billable") & vbCrLf
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
Else
	strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
	CSVBody = "< --- No records found --- >"
End If
rsRep.Close
Set rsRep = Nothing

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

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
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")
tmpOver =Z_CLng(Request("override"))
ctr = 10
RepCSV =  "IntrApptResps" & tmpdate & tmpTime & ".csv" 
tmpDate = Z_YMDDate(date)
Set rsRep = Server.CreateObject("ADODB.RecordSet")

strMSG = "Interpreter Appointment Response report"
strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Status</td>" & vbCrlf & _
		"<td class='tblgrn' colspan=""2"">Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Answer</td>" & vbCrlf & _
		"<td class='tblgrn'>Answer Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Assigend?</td>" & vbCrlf
CSVHead = "Request ID,Status,Department,Language,Date,Appointment Start,Appointment End" & _
		",Interpreter Last Name,Interpreter First Name,Answer,Answer Date,Assigned"
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
If tmpOver < 1 Then
	sqlRep = "EXEC [dbo].[CountRepIntrResponses] " & strPD
	rsRep.Open sqlRep, g_strCONN, 3, 1
	lngCount = 0
	If Not rsRep.EOF Then
		lngCount = Z_CLng(rsRep("num"))
	End If
	If lngCount > 2000 Then
	%>
		<div id="uhoh" style="margin-left: 50px;">
		<h1>Response count produces: <%=lngCount%> records</h1>
		<p style="font-size: 150%;">Select a smaller report period instead.</p>
		<button type="button" class="button button-secondary" onclick="CloseMe()">Close</button>
			<br /><br /><br /><br />
		<button type="button" class="button button-primary" onclick="Continue()">Continue Anyway</button>
		* warning -- execution time might take long!
		</div>
		<div id="wait" style="display: none; width: 70px; margin-left: auto; margin-right: auto;">
			<img src="images/ajax-loader.gif" style="width: 66px; height: 66px;" />
		</div>
		</body>
	</html>
	<script>
		function Continue() {
			var zzz = document.getElementById("wait");
			zzz.style.display = "block";
			var aaa = document.getElementById("uhoh");
			aaa.style.display = "none";
			document.location = "rep_intrresp.asp?override=1";
		}
	</script>

	<%
		Response.Flush
		Response.End
	End if
	rsRep.Close
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
sqlRep = "EXEC [dbo].[RepIntrResponses] " & strPD ' '2018-08-15', '2018-08-17'
rsRep.Open sqlRep, g_strCONN, 3, 1
x = 1
If Not rsRep.EOF Then
	strOldID = ""
	Do Until rsRep.EOF
		kulay = "#FFFFc5"
		If Not Z_IsOdd(x) Then kulay = "#c5c5c5"
		If rsRep("appID") <> strOldID Then
			strFrom = Z_Time24(rsRep("from"))
			strTo = Z_Time24(rsRep("to"))
			strDt = Z_MDYDate(rsRep("date"))
			strBody = strBody & "<tr style=""background-color: " & kulay & ";"" onclick='PassMe(" & rsRep("appID") & ")'>" & _
					"<td class='tblgrn2' style=""background-color: " & kulay & ";""><nobr>" & rsRep("appID")  & "</td>" & vbCrLf & _
					"<td class='tblgrn2' style=""background-color: " & kulay & ";""><nobr>" & rsRep("status") & "</td>" & vbCrLf & _
					"<td class='tblgrn2' style=""background-color: " & kulay & ";""><nobr>" & strDt & "</td>" & vbCrLf & _
					"<td class='tblgrn2' style=""background-color: " & kulay & ";""><nobr>" & strFrom & " - " & strTo &"</td>" & vbCrLf & _
					"<td class='tblgrn2' style=""background-color: " & kulay & "; text-align: left;""><nobr>" & rsRep("Dept") & "</td>" & vbCrLf & _
					"<td class='tblgrn2' style=""background-color: " & kulay & "; text-align: left;""><nobr>" & rsRep("Language") & "</td>" & vbCrLf
			strOldID = rsRep("appID")
		Else
			strBody = strBody & "<tr style=""background-color: " & kulay & ";""><td colspan=""6""></td>" & vbCrLf
		End If
		strAD = ""
		If Z_FixNull(rsRep("ans_dt")) <> "" Then
			strAD = Z_MDYDate(rsRep("ans_dt")) & " " & FormatDateTime(rsRep("ans_dt"), 4)
		End If
		strBody = strBody & "<td class='tblgrn2' style=""background-color: " & kulay & "; text-align: left;""><nobr>" & _
				rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";""><nobr>" & rsRep("answer") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";""><nobr>" &  strAD & "</td>" & vbCrLf & _
				"<td class='tblgrn2' style=""background-color: " & kulay & ";""><nobr>" &  rsRep("assigned") & "</td></tr>" & vbCrLf 
		CSVBody = CSVBody & rsRep("appID") & "," & rsRep("Status") & "," &  Replace(rsRep("Dept"), " - ", "") & "," & _
				rsRep("Language") & "," & strDt & "," & strFrom &  ","  & strTo & "," & _
				rsRep("Last Name") & ","  & rsRep("First Name") & ",""" & _
				rsRep("answer") & """,""" & strAD & """,""" &  rsRep("assigned") & """" & vbCrLf
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

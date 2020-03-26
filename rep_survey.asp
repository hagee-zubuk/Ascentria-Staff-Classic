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
Function Z_YesNo(blnZZ) 
	Z_YesNo = "<td>&mdash;</td>"
On Error Resume Next	
	If CBool(blnZZ) Then
		Z_YesNo = "<td class=""yiz"">YES</td>"
	ElseIf Not CBool(blnZZ) Then
		Z_YesNo = "<td>no</td>"
	End If
End Function
Function Z_YesNoPlain(blnZZ) 
	Z_YesNoPlain = """-"""
On Error Resume Next	
	If CBool(blnZZ) Then
		Z_YesNoPlain = """YES"""
	ElseIf Not CBool(blnZZ) Then
		Z_YesNoPlain = """no"""
	End If
End Function
Function Z_FmtTS(dtDate)
	Z_FmtTS = "<td></td>"
	If Not IsDate(dtDate) Then Exit Function
	If Abs(DateDiff("d", dtDate, Date)) > 1 Then
		Z_FmtTS = "<td class=""ts late"">"
	Else
		Z_FmtTS = "<td class=""ts"">"
	End If
DIM lngTmp, strDay, strTmp
	lngTmp = Z_CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_FmtTS = Z_FmtTS & "0"
	Z_FmtTS = Z_FmtTS & lngTmp & "/"
	lngTmp = Z_CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_FmtTS = Z_FmtTS & "0"
	Z_FmtTS = Z_FmtTS & lngTmp & "/"
	strTmp = DatePart("yyyy", dtDate)
	Z_FmtTS = Z_FmtTS & Right(strTmp,2) & " "
	lngTmp = DatePart("h", dtDate)
	If lngTmp < 10 Then Z_FmtTS = Z_FmtTS & "0"
	Z_FmtTS = Z_FmtTS & lngTmp & ":"
	lngTmp = DatePart("n", dtDate)
	If lngTmp < 10 Then Z_FmtTS = Z_FmtTS & "0"
	Z_FmtTS = Z_FmtTS & lngTmp & "</td>"
End Function
Function Z_MakeUniqueFileName()
	tmpdate = replace(date, "/", "") 
	tmpTime = replace(FormatDateTime(time, 3), ":", "")
	tmpTime = replace(tmpTime, " ", "")
	Z_MakeUniqueFileName = tmpdate & tmpTime
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
<style>
table.zzz { margin-bottom: 20px; min-width: 660px; }
.zzz tr:nth-child(even) { background: #CCC }
.zzz tr:nth-child(odd)  { background: #FFF }
.zzz th { font-size: 80%; background-color: #f9e79f; }
.zzz td {text-align: center; padding: 1px 3px;}
.zzz td:first-child{ text-align: left; } 
td.ts { font-size: 80%; }
td.late { background-color: #FFC300; }
td.yiz  { background-color: #FE5C5C; font-weight: bold; font-size: 110%; }
</style>	
</head>
<body>
	<div style="width: 300px; text-align: center; margin: 0px auto 20px;">
		<img src='images/LBISLOGO.jpg' align='center' style="width: 287px; height: 67px;" />
		340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br />
		Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
	</div>
<%
Server.scripttimeout = 3600000	' 1 hour!


strSQL = "SELECT [ts], [First Name], [Last Name]" & _
			", isu.[username] AS [username]" & _
			", [Q1], [Q2], [Q3], [sig], yyy.[cnt] " & _
		"FROM  [interpreter_T] AS itr " & _
			"INNER JOIN [user_t] AS isu ON itr.[index]=isu.[intrID] " & _
			"LEFT JOIN ( " & _
		"SELECT svy.[IntrID], svy.[ts], svy.[sig], svy.[Q1], svy.[Q2], svy.[Q3], zzz.[latest], zzz.[cnt] " & _
		"FROM [2020Survey] AS svy " & _
			"INNER JOIN ( " & _
				"SELECT [intrID]" & _
					", MAX([ts]) AS [latest]" & _
					", COUNT([id]) AS [cnt] " & _
				"FROM [2020Survey] " & _
				"GROUP BY [intrID] " & _
			") AS zzz ON svy.[ts]=zzz.[latest] AND svy.[IntrID]=zzz.[IntrID]" & _
		") AS yyy ON itr.[index]=yyy.[IntrID] " & _
		"WHERE itr.[active]=1 AND itr.[index]<>770 " & _
		"ORDER BY itr.[Last Name] ASC"


Set rsReq = CreateObject("ADODB.Recordset")
rsReq.Open strSQL, g_strCONN, 3, 1
strCSV = """Last Name"",""First Name"",""username"",""timestamp"",""Q1: travel"",""Q2: contact"",""Q3: symptoms"",""initials"",""# surveys""" & vbCrLf
Do Until rsReq.EOF
	strBody = strBody & "<tr><td>" & rsReq("last name") & ", " & rsReq("first name") & "</td>"
	strCSV = strCSV & """" & rsReq("last name") & """,""" & rsReq("first name") & """,""" & LCase( rsReq("username") ) & ""","
	If (Z_CDate(rsReq("ts")) < CDate("2020-01-01")) Then
		strBody = strBody & "<td colspan=""6"" style=""text-align: center;""><i>no response</i></td>"
		strCSV = strCSV & """no response"""
	Else 
		strBody = strBody & _
				Z_FmtTS(rsReq("ts")) & _
				Z_YesNo(rsReq("Q1")) & _
				Z_YesNo(rsReq("Q2")) & _
				Z_YesNo(rsReq("Q3")) & _
				"<td>" & rsReq("sig") & "</td>" & _
				"<td>" & rsReq("cnt") & "</td>"
		strCSV = strCSV & """" & rsReq("ts") & ""","
		strCSV = strCSV & Z_YesNoPlain(rsReq("Q1")) & "," & _
				Z_YesNoPlain(rsReq("Q2")) & "," & _
				Z_YesNoPlain(rsReq("Q3")) & ","
		strCSV = strCSV & """" & rsReq("sig") & ""","
		strCSV = strCSV & Z_CLng(rsReq("cnt") )
	End If
	strBody = strBody & "</tr>" & vbCrLf 
	strCSV = strCSV & vbCrLf
	rsReq.MoveNext
Loop
rsReq.Close
Set rsReq = Nothing

Set fso = Server.CreateObject("Scripting.FileSystemObject")
RepCSV =  "NCOV_survey" & Z_MakeUniqueFileName() & ".csv" 
Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
Prt.WriteLine "LANGUAGE BANK - SUMMARY COVID-19 Survey Responses " & Z_MDYDate(Date)
Prt.WriteLine strCSV
Prt.Close
Set Prt = Nothing
tmp3 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSV)
%>
<form method='post' name='frmResult'>
	<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
				<tr><td valign='top' >
<table class="zzz" bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
	<tr bgcolor='#f58426'>
		<td colspan="7" align="center">Interpreter Survey</td></tr>
	<tr><th style="min-width: 200px;">Interpreter</th>
		<th style="width: 100px;">timestamp</th>
		<th style="width: 50px;">Q1<br />(travel)</th>
		<th style="width: 50px;">Q2<br />(contact)</th>
		<th style="width: 50px;">Q3<br />(syptoms)</th>
		<th style="width: 120px;">initials</th>
		<th style="width: 40px;"># of sur-<br />veys taken</th>
		</tr>
<%= strBody %>
</table>
				</td></tr>
	<tr><td style="text-align: center;">
		<button id="btnCSV" name="btnCSV" type="button" onclick="document.location='<%= tmp3%>';">
			CSV Export</button>
		</td></tr>
<!--								<tr><td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
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
									</td></tr>
-->									
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
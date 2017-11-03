<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function GetInst(zzz)
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM institution_T WHERE index = " & zzz
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInst	= rsInst("Facility")
	Else
		GetInst = "N/A"
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetIntr(zzz)
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM interpreter_T WHERE index = " & zzz
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		GetIntr	= rsIntr("Last Name") & ", " & rsIntr("First Name")
	Else
		GetIntr = "N/A"
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	tmp1Day = Request("selMonth") & "/01/" & Request("txtyear")
	tmpMonth = MonthName(Request("selMonth")) & " - " & Request("txtyear")
End If
'SET CALENDAR
If Request("selMonth") <> "" And Request("txtyear") <> "" Then
	tmp1Day = Request("selMonth") & "/01/" & Request("txtyear")
	tmpMonth = MonthName(Request("selMonth")) & " - " & Request("txtyear")
End If
If tmp1Day = "" Then 
	tmp1Day = Month(Date) & "/01/" & Year(Date)
	tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
End If
If Not IsDate(tmp1Day) Then 
	tmp1day = Month(Date) & "/01/" & Year(Date)
	tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
	Session("MSG") = "ERROR: Year inputted is not valid. Set to current month and year."
End If
CorrectMonth = True
tmpToday = tmp1Day
Do While CorrectMonth = True 
	strCal = strCAL & "<tr>"
	If WeekdayName(Weekday(tmpToday), True) <> "Sun" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM request_T WHERE appDate = #" & tmpToday & "# ORDER BY astarttime, apptimefrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			Do Until rsReq.EOF
				tmpInst = GetInst(rsReq("InstID"))
				tmpIntr = GetIntr(rsReq("IntrID"))
				tmpStr = tmpInst & " - " & tmpIntr
				If Len(tmpStr) > 25 Then tmpStr = Left(tmpStr, 25) & "...."
				tmptime = Z_DateNull(rsReq("AStarttime"))
				If tmptime = Empty Then tmptime = rsReq("appTimeFrom")
				strApp = strApp & "<a href='main.asp?ID=" & rsReq("index") & "' class='callink' title='" &  tmpInst & " - " & tmpIntr & _
				" (" & tmptime & ")" & "'><nobr>" & tmpStr & "</a><br>"
				rsReq.MoveNext
			Loop
		End If
		rsReq.Close
		Set rsReq = Nothing
		strCal = strCAL & "<td class='caltbl' valign='top'><span class='calheader'>" & Day(tmpToday) & "</span><br>" & strApp & "</td>"
		strApp = ""
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Mon" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
			Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM request_T WHERE appDate = #" & tmpToday & "# ORDER BY astarttime, apptimefrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			Do Until rsReq.EOF
				tmpInst = GetInst(rsReq("InstID"))
				tmpIntr = GetIntr(rsReq("IntrID"))
				tmpStr = tmpInst & " - " & tmpIntr
				If Len(tmpStr) > 25 Then tmpStr = Left(tmpStr, 25) & " ..."
				tmptime = Z_DateNull(rsReq("AStarttime"))
				If tmptime = Empty Then tmptime = rsReq("appTimeFrom")
				strApp = strApp & "<a href='main.asp?ID=" & rsReq("index") & "' class='callink' title='" &  tmpInst & " - " & tmpIntr & _
				" (" & tmptime & ")" & "'><nobr>" & tmpStr & "</a><br>"
				rsReq.MoveNext
			Loop
		End If
		rsReq.Close
		Set rsReq = Nothing
		strCal = strCAL & "<td class='caltbl' valign='top'><span class='calheader'>" & Day(tmpToday) & "</span><br>" & strApp & "</td>"
		strApp = ""
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Tue" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM request_T WHERE appDate = #" & tmpToday & "# ORDER BY astarttime, apptimefrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			Do Until rsReq.EOF
				tmpInst = GetInst(rsReq("InstID"))
				tmpIntr = GetIntr(rsReq("IntrID"))
				tmpStr = tmpInst & " - " & tmpIntr
				If Len(tmpStr) > 25 Then tmpStr = Left(tmpStr, 25) & " ..."
				tmptime = Z_DateNull(rsReq("AStarttime"))
				If tmptime = Empty Then tmptime = rsReq("appTimeFrom")
				strApp = strApp & "<a href='main.asp?ID=" & rsReq("index") & "' class='callink' title='" &  tmpInst & " - " & tmpIntr & _
				" (" & tmptime & ")" & "'><nobr>" & tmpStr & "</a><br>"
				rsReq.MoveNext
			Loop
		End If
		rsReq.Close
		Set rsReq = Nothing
		strCal = strCAL & "<td class='caltbl' valign='top'><span class='calheader'>" & Day(tmpToday) & "</span><br>" & strApp & "</td>"
		strApp = ""
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Wed" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM request_T WHERE appDate = #" & tmpToday & "# ORDER BY astarttime, apptimefrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			Do Until rsReq.EOF
				tmpInst = GetInst(rsReq("InstID"))
				tmpIntr = GetIntr(rsReq("IntrID"))
				tmpStr = tmpInst & " - " & tmpIntr
				If Len(tmpStr) > 25 Then tmpStr = Left(tmpStr, 25) & " ..."
				tmptime = Z_DateNull(rsReq("AStarttime"))
				If tmptime = Empty Then tmptime = rsReq("appTimeFrom")
				strApp = strApp & "<a href='main.asp?ID=" & rsReq("index") & "' class='callink' title='" &  tmpInst & " - " & tmpIntr & _
				" (" & tmptime & ")" & "'><nobr>" & tmpStr & "</a><br>"
				rsReq.MoveNext
			Loop
		End If
		rsReq.Close
		Set rsReq = Nothing
		strCal = strCAL & "<td class='caltbl' valign='top'><span class='calheader'>" & Day(tmpToday) & "</span><br>" & strApp & "</td>"
		strApp = ""
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Thu" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM request_T WHERE appDate = #" & tmpToday & "# ORDER BY astarttime, apptimefrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			Do Until rsReq.EOF
				tmpInst = GetInst(rsReq("InstID"))
				tmpIntr = GetIntr(rsReq("IntrID"))
				tmpStr = tmpInst & " - " & tmpIntr
				If Len(tmpStr) > 25 Then tmpStr = Left(tmpStr, 25) & " ..."
				tmptime = Z_DateNull(rsReq("AStarttime"))
				If tmptime = Empty Then tmptime = rsReq("appTimeFrom")
				strApp = strApp & "<a href='main.asp?ID=" & rsReq("index") & "' class='callink' title='" &  tmpInst & " - " & tmpIntr & _
				" (" & tmptime & ")" & "'><nobr>" & tmpStr & "</a><br>"
				rsReq.MoveNext
			Loop
		End If
		rsReq.Close
		Set rsReq = Nothing
		strCal = strCAL & "<td class='caltbl' valign='top'><span class='calheader'>" & Day(tmpToday) & "</span><br>" & strApp & "</td>"
		strApp = ""
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Fri" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM request_T WHERE appDate = #" & tmpToday & "# ORDER BY astarttime, apptimefrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			Do Until rsReq.EOF
				tmpInst = GetInst(rsReq("InstID"))
				tmpIntr = GetIntr(rsReq("IntrID"))
				tmpStr = tmpInst & " - " & tmpIntr
				If Len(tmpStr) > 25 Then tmpStr = Left(tmpStr, 25) & " ..."
				tmptime = Z_DateNull(rsReq("AStarttime"))
				If tmptime = Empty Then tmptime = rsReq("appTimeFrom")
				strApp = strApp & "<a href='main.asp?ID=" & rsReq("index") & "' class='callink' title='" &  tmpInst & " - " & tmpIntr & _
				" (" & tmptime & ")" & "'><nobr>" & tmpStr & "</a><br>"
				rsReq.MoveNext
			Loop
		End If
		rsReq.Close
		Set rsReq = Nothing
		strCal = strCAL & "<td class='caltbl' valign='top'><span class='calheader'>" & Day(tmpToday) & "</span><br>" & strApp & "</td>"
		strApp = ""
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Sat" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		Set rsReq = Server.CreateObject("ADODB.RecordSet")
		sqlReq = "SELECT * FROM request_T WHERE appDate = #" & tmpToday & "# ORDER BY astarttime, apptimefrom"
		rsReq.Open sqlReq, g_strCONN, 3, 1
		If Not rsReq.EOF Then
			Do Until rsReq.EOF
				tmpInst = GetInst(rsReq("InstID"))
				tmpIntr = GetIntr(rsReq("IntrID"))
				tmpStr = tmpInst & " - " & tmpIntr
				If Len(tmpStr) > 25 Then tmpStr = Left(tmpStr, 25) & " ..."
				tmptime = Z_DateNull(rsReq("AStarttime"))
				If tmptime = Empty Then tmptime = rsReq("appTimeFrom")
				strApp = strApp & "<a href='main.asp?ID=" & rsReq("index") & "' class='callink' title='" &  tmpInst & " - " & tmpIntr & _
				" (" & tmptime & ")" & "'><nobr>" & tmpStr & "</a><br>"
				rsReq.MoveNext
			Loop
		End If
		rsReq.Close
		Set rsReq = Nothing
		strCal = strCAL & "<td class='caltbl' valign='top'><span class='calheader'>" & Day(tmpToday) & "</span><br>" & strApp & "</td>"
		strApp = ""
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	strCal = strCAL & "</tr>"
	If Month(tmp1Day) <> Month(tmpToday) Then CorrectMonth = False
Loop
%>
<html>
	<head>
		<title>Language Bank - Calendar View</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function ChangeMonth(xxx)
		{
			document.frmCal.action = "action.asp?ctrl=4&dir=" + xxx;
			document.frmCal.submit();
		}
		function SearchMonth()
		{
			if (document.frmCal.txtyear.value == "")
			{
				alert("ERROR: Year is required.");
			}
			else
			{
				document.frmCal.action = "calendarview.asp"
				document.frmCal.submit();
			}
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmCal'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td height='100px' valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td>
						<table cellSpacing='2' cellPadding='2' style='width: 100%;' border='0' align='center'>
							<tr>
								<td>
									<input class='btn' type='button' value='<<' title='Previous Month' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(0);'>
								</td>
								<td align='center' colspan='5'>
									Month:
									<select class='seltxt' style='width: 50px;' name='selMonth'>
										<option value='01'>Jan</option>
										<option value='02'>Feb</option>
										<option value='03'>Mar</option>
										<option value='04'>Apr</option>
										<option value='05'>May</option>
										<option value='06'>Jun</option>
										<option value='07'>Jul</option>
										<option value='08'>Aug</option>
										<option value='09'>Sep</option>
										<option value='10'>Oct</option>
										<option value='11'>Nov</option>
										<option value='12'>Dec</option>
									</select>
									Year:
									<input class='main' name='txtyear' maxlength='4' size='5'>
									<input class='btn' type='button' value='GO' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SearchMonth();'>
								</td>
								<td align='right'>
									<input class='btn' type='button' value='>>' title='Next Month' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(1);'>
								</td>
							</tr>
							<tr>
								<td colspan='7'>
									<table cellSpacing='0' cellPadding='0' bgcolor='#FFFFFF' align='center' style='width: 100%;'>
										<tr>
											<td colspan='7' align='center' class='calheader'><%=tmpMonth%></td>
										</tr>
										<tr>
											<td class='calweekday'>Sun</td>
											<td class='calweekday'>Mon</td>
											<td class='calweekday'>Tue</td>
											<td class='calweekday'>Wed</td>
											<td class='calweekday'>Thu</td>
											<td class='calweekday'>Fri</td>
											<td class='calweekday'>Sat</td>
										</tr>
										<%=strCal%>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<input type='hidden' name='Hmonth' value='<%=tmpMonth%>'>
				<tr>
					<td height='50px' valign='bottom'>
						<!-- #include file="_footer.asp" -->
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
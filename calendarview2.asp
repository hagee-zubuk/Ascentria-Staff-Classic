<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request.Cookies("LBUSERTYPE") = 2 Then
	If Session("UIntr") = "" Then
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.Redirect "default.asp"
	End If
End If
Function MyStatus(xxx)
	'get status of request
	Select Case xxx
		Case 1
			MyStatus = "<font color='#000000' size='+3'>•</font>" 'complete
		Case 2
			MyStatus = "<font color='#0000FF' size='+3'>•</font>" 'missed
		Case 3
			MyStatus = "<font color='#FF0000' size='+3'>•</font>" 'canceled
		Case 4
			MyStatus = "<font color='#FF00FF' size='+3'>•</font>" 'canceled bill
		Case Else
			MyStatus = ""
	End Select
End Function
CalendarPage = True
tmpReqMonth = Request("selMonth")
tmpReqYear = Request("txtyear")
'institution
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM Institution_T WHERE Active = 1 ORDER BY Facility"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF 
	strInst = strInst & "<option value='" & rsInst("index") & "'>" & rsInst("Facility") & "</option>" & vbCrLf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
If Request("appdate") <> "" Then
		tmpAppDate = Split(Request("appdate"), "/")
		tmpReqMonth = tmpAppDate(0)
		tmpReqYear = tmpAppDate(2)
End If
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	tmp1Day = tmpReqMonth & "/01/" & tmpReqYear
	tmpMonth = MonthName(tmpReqMonth) & " - " & tmpReqYear
End If
'SET CALENDAR
If tmpReqMonth <> "" And tmpReqYear <> "" Then
	tmp1Day = tmpReqMonth & "/01/" & tmpReqYear
	tmpMonth = MonthName(tmpReqMonth) & " - " & tmpReqYear
	myMonth = tmpReqMonth
	myYear = tmpReqYear
End If
If tmp1Day = "" Then 
	tmp1Day = Month(Date) & "/01/" & Year(Date)
	tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
	myMonth = Month(Date)
	myYear = Year(Date)
End If
If Not IsDate(tmp1Day) Then 
	tmp1day = Month(Date) & "/01/" & Year(Date)
	tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
	myMonth = Month(Date)
	myYear = Year(Date)
	Session("MSG") = "ERROR: Year inputted is not valid. Set to current month and year."
End If
CorrectMonth = True
tmpToday = tmp1Day
Do While CorrectMonth = True 
	'set calendar
	strCal = strCAL & "<tr><td>&nbsp;</td>"
	If WeekdayName(Weekday(tmpToday), True) <> "Sun" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Mon" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Tue" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Wed" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Thu" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Fri" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Sat" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	strCal = strCAL & "</tr>"
	If Month(tmp1Day) <> Month(tmpToday) Then CorrectMonth = False
Loop
'SET ORGANIZER
If Request("txtday") <> "" Then
	tmpDate = tmpReqMonth & "/" & Request("txtday") & "/" & tmpReqYear
Else
	tmpDate = Date
	If tmpReqMonth <> "" And tmpReqYear <> "" Then tmpDate = tmp1Day
	If Request("appdate") <> "" Then tmpDate = Request("appdate")
End If
myHol = ""
If IsHoliday(tmpDate) Then myHol = "&nbsp;<i>(Holiday)</i>"
'SET TIMESHEET DATE and MILEAGE
mysundate = GetSun(tmpDate)
mysatdate = GetSat(tmpDate)
mytsheet = mysundate & " - " & mysatdate
mymileage = tmpMonth

sqlDup = "SELECT COUNT(qq) AS dup FROM (" & _
		"SELECT COUNT([index]) AS qq, [Clname], [Cfname] " & _
		"FROM [request_T] " & _
		"WHERE [appDate]='" & tmpDate & "' AND ([Status]<2 OR [Status]>3) " & _
		"GROUP BY [Clname], [Cfname]) AS z WHERE qq > 1"
Set rsDup = Server.CreateObject("ADODB.RecordSet")		
rsDup.Open sqlDup, g_strConn, 3, 1
blnDup = False
If Not rsDup.EOF Then
	blnDup = CBool(rsDup("dup") > 0)
End If
rsDup.Close
Set rsDup = Nothing
Set rsReq = Server.CreateObject("ADODB.RecordSet")
If Request.Cookies("LBUSERTYPE") <> 2 Then
	sqlReq = "SELECT * FROM request_T WHERE appDate = '" & tmpDate & "' ORDER BY appTimeFrom"
Else
	sqlReq = "SELECT [index], DeptID, InstID, IntrID, LangID, Status, Cphone, Clname, Cfname, appTimeFrom, appTimeTo FROM request_T WHERE appDate = '" & tmpDate & "' AND IntrID = " & Session("UIntr") & " " & _
		"AND showintr = 1 AND NOT(STATUS = 2 OR STATUS = 3) ORDER BY appTimeFrom"
End If
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	Do Until rsReq.EOF
		myDept =  GetMyDept(rsReq("DeptID"))
		tmpInst = Split(GetInst(rsReq("InstID")), "|")
		tmpIntr = GetIntr(rsReq("IntrID"))
		assigned = ""
		If tmpIntr <> "N/A, N/A" Then assigned = "disabled"
		tmpLang = GetLang(rsReq("LangID"))
		tmpStat = MyStatus(rsReq("Status") )
		tmpTime = CTime(rsReq("appTimeFrom"))
		If Z_fixnull(rsReq("appTimeTo")) <> "" Then 
			strtmpTime = tmpTime & " - " & CTime(rsReq("appTimeTo"))
		Else
			strtmpTime = tmpTime
		End If
		tmpPhone = "N/A"
		If rsReq("Cphone") <> "" Then tmpPhone = rsReq("Cphone")
		tmp12 = cdate("12:00 AM")
		tmp1259 = cdate("1:00 AM")
		tmpID = rsReq("Index")
		cbk = "#F5F5F5"
		tmpHPID = Z_CZero(rsReq("HPID"))
		If rsReq("InstID") = 273 And rsReq("LangID") = 25 And (Weekday(tmpDate) = 2 Or Weekday(tmpDate) = 3 Or Weekday(tmpDate) = 4 _
			Or Weekday(tmpDate) = 5) Then 
				If tmpHPID <> 0 Then
					If IsBlock(tmpHPID) Then cbk = "#FFFFCE"
				End If
				'cbk = "#FFFFCE"
		End If	
		If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then
			If  Cint(Request.Cookies("LBUSERTYPE")) <> 2 Then
				tmpstr = "<tr bgcolor='" & cbk & "' onclick=''>" & _
					"<td align='center'><input class='btn' " & assigned & " type='button' value='Email' style='width: 40px;' onmouseover=""this.className='hovbtn'"" onmouseout=""this.className='btn'"" onclick='AssignMe(" & tmpID & ");'>"  & _
					"<td align='center' onclick='PassMe(" & tmpID & ");'><nobr>" & strtmpTime & "</td>" & _
					"<td align='center' onclick='PassMe(" & tmpID & ");'>" &  rsReq("Clname") & ", " & rsReq("Cfname") & "</td>" & _
					"<td align='center' onclick='PassMe(" & tmpID & ");'>" & tmpInst(0) & myDept & "</td><td align='center' onclick='PassMe(" & tmpID & ");'>" & tmpLang & "</td>" & _
					"<td align='center' onclick='PassMe(" & tmpID & ");'>" & tmpIntr & "</td><td align='center'><nobr>" &  tmpPhone & "</td>" & _
					"<td align='center' onclick='PassMe(" & tmpID & ");'>" & tmpStat & "</td></tr>" & vbCrLf
			Else
				tmpCli = left(rsReq("Cfname"), 1) & ". " & left(rsReq("Clname"), 1) & ". "
				tmpstr = "<tr bgcolor='" & cbk & "' onclick=''>" & _
					"<td align='center'>&nbsp;</td>"  & _
					"<td align='center'><nobr>" & strtmpTime & _
					"</td><td align='center'>" &  tmpCli & "</td>" & _
					"<td align='center'>" & tmpInst(0) & "</td><td align='center' onclick=''>" & tmpLang & "</td>" & _
					"<td align='center'>" & tmpIntr & "</td><td align='center'><nobr>N/A</td>" & _
					"<td align='center'>" & tmpStat & "</td></tr>" & vbCrLf
			End If
		Else
			tmpstr = "<tr bgcolor='" & cbk & "' onclick=''>" & _
				"<td align='center'><input class='btn' type='button' value='View' style='width: 40px;' onmouseover=""this.className='hovbtn'"" onmouseout=""this.className='btn'"" onclick='PassMe(" & tmpID & ");'>" & _
				"<td align='center'><nobr>" & strtmpTime & "</td>" & _
				"<td align='center'>" & tmpInst(0) & myDept & "</td><td align='center'>" & tmpLang & "</td>" & _
				"<td align='center'>" & tmpIntr & "</td><td align='center'><nobr>" &  tmpPhone & "</td>" & _
				"<td align='center'>" & tmpStat & "</td></tr>" & vbCrLf
		End If
		If tmpTime >= tmp12 And tmpTime < tmp1259 Then
			str12a = str12a & tmpstr
			x12a = x12a + 1
		ElseIf tmpTime >= DateAdd("H", 1, tmp12) And tmpTime < DateAdd("H", 1, tmp1259) Then
			str1a = str1a & tmpstr
			x1a = x1a + 1
		ElseIf tmpTime >= DateAdd("H", 2, tmp12) And tmpTime < DateAdd("H", 2, tmp1259) Then
			str2a = str2a & tmpstr
			x2a = x2a + 1
		ElseIf tmpTime >= DateAdd("H", 3, tmp12) And tmpTime < DateAdd("H", 3, tmp1259) Then
			str3a = str3a & tmpstr
			x3a = x3a + 1
		ElseIf tmpTime >= DateAdd("H", 4, tmp12) And tmpTime < DateAdd("H", 4, tmp1259) Then
			str4a = str4a & tmpstr
			x4a = x4a + 1
		ElseIf tmpTime >= DateAdd("H", 5, tmp12) And tmpTime < DateAdd("H", 5, tmp1259) Then
			str5a = str5a & tmpstr
			x5a =  x5a + 1
		ElseIf tmpTime >= DateAdd("H", 6, tmp12) And tmpTime < DateAdd("H", 6, tmp1259) Then
			str6a = str6a & tmpstr
			x6a = x6a + 1
		ElseIf tmpTime >= DateAdd("H", 7, tmp12) And tmpTime < DateAdd("H", 7, tmp1259) Then
			str7a = str7a & tmpstr
			x7a = x7a + 1
		ElseIf tmpTime >= DateAdd("H", 8, tmp12) And tmpTime < DateAdd("H", 8, tmp1259) Then
			str8a = str8a & tmpstr
			x8a = x8a + 1
		ElseIf tmpTime >= DateAdd("H", 9, tmp12) And tmpTime < DateAdd("H", 9, tmp1259) Then
			str9a = str9a & tmpstr
			x9a = x9a + 1
		ElseIf tmpTime >= DateAdd("H", 10, tmp12) And tmpTime < DateAdd("H", 10, tmp1259) Then
			str10a = str10a & tmpstr
			x10a = x10a + 1
		ElseIf tmpTime >= DateAdd("H", 11, tmp12) And tmpTime < DateAdd("H", 11, tmp1259) Then
			str11a = str11a & tmpstr
			x11a = x11a + 1
		ElseIf tmpTime >= DateAdd("H", 12, tmp12) And tmpTime < DateAdd("H", 12, tmp1259) Then
			str12p = str12p & tmpstr
			x12p = x12p + 1
		ElseIf tmpTime >= DateAdd("H", 13, tmp12) And tmpTime < DateAdd("H", 13, tmp1259) Then
			str1p = str1p & tmpstr
			x1p = x1p + 1
		ElseIf tmpTime >= DateAdd("H", 14, tmp12) And tmpTime < DateAdd("H", 14, tmp1259) Then
			str2p = str2p & tmpstr
			x2p = x2p + 1
		ElseIf tmpTime >= DateAdd("H", 15, tmp12) And tmpTime < DateAdd("H", 15, tmp1259) Then
			str3p = str3p & tmpstr
			x3p = x3p + 1
		ElseIf tmpTime >= DateAdd("H", 16, tmp12) And tmpTime < DateAdd("H", 16, tmp1259) Then
			str4p = str4p & tmpstr
			x4p = x4p + 1
		ElseIf tmpTime >= DateAdd("H", 17, tmp12) And tmpTime < DateAdd("H", 17, tmp1259) Then
			str5p = str5p & tmpstr
			x5p = x5p + 1
		ElseIf tmpTime >= DateAdd("H", 18, tmp12) And tmpTime < DateAdd("H", 18, tmp1259) Then
			str6p = str6p & tmpstr
			x6p = x6p + 1
		ElseIf tmpTime >= DateAdd("H", 19, tmp12) And tmpTime < DateAdd("H", 19, tmp1259) Then
			str7p = str7p & tmpstr
			x7p = x7p + 1
		ElseIf tmpTime >= DateAdd("H", 20, tmp12) And tmpTime < DateAdd("H", 20, tmp1259) Then
			str8p = str8p & tmpstr
			x8p = x8p + 1
		ElseIf tmpTime >= DateAdd("H", 21, tmp12) And tmpTime < DateAdd("H", 21, tmp1259) Then
			str9p = str9p & tmpstr
			x9p = x9p + 1
		ElseIf tmpTime >= DateAdd("H", 22, tmp12) And tmpTime < DateAdd("H", 22, tmp1259) Then
			str10p = str10p & tmpstr
			x10p = x1a + 1
		ElseIf tmpTime >= DateAdd("H", 23, tmp12) And tmpTime < DateAdd("H", 23, tmp1259) Then
			str11p = str11p & tmpstr
			x11p = x11p + 1
		End If
		rsReq.MoveNext
	Loop
End If
rsReq.Close
Set rsReq = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Calendar View</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function AssignMe(xxx)
		{
			newwindow = window.open('emailIntr.asp?ID=' + xxx,'','height=250,width=500,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
		}
		function PassMe(xxx)
		{
			document.frmCal.action = "reqconfirm.asp?ID=" + xxx;
			document.frmCal.submit();
		}
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
				document.frmCal.action = "calendarview2.asp";
				document.frmCal.submit();
			}
		}
		function GoToday(xm, xd, xy)
		{
			document.frmCal.action = "calendarview2.asp?selMonth=" + xm + "&txtday=" + xd + "&txtyear=" + xy;
			document.frmCal.submit();
		}
		function PublishMe()
		{
			document.frmCal.action = "action.asp?ctrl=8";
			document.frmCal.submit();	
		}
		function PublishMe2()
		{
			document.frmCal.action = "action.asp?ctrl=15&tmpDate=" + '<%=tmpDate%>';
			document.frmCal.submit();	
		}
<%	If blnDup Then %>		
		function chkDuplicates() {
			dupwin = window.open("calDuplicates.asp?dt=<%=tmpDate%>","WinDup",
							"channelmode=0,directories=0,fullscreen=0,height=400," + 
							"left=100,location=0,menubar=0,resizable=1,scrollbars=1" +
							"status=0,titlebar=0,toolbar=0,top=100,width=600");
		}
<%	End If %>		
		function PublishMe3()
		{
			document.frmCal.action = "action.asp?ctrl=26&tmpDate=" + '<%=tmpDate%>';
			document.frmCal.submit();	
		}
		<%If Request("rep") <> "" Then%>
			<%If Request("rep") = 25 Then%>
				function PopMe(zzz)
				{
					if (zzz !== undefined)
						{
						<% If Request.Cookies("LBUSERTYPE") = 2 Then %>	
							newwindow = window.open("printreport.asp?publish=1&Hdate='" + <%=Request("tmpdate")%> + "'&selRP=" + <%=Request("tmpRP")%>,"name","height=800,width=750,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1");
						<% Else %>
							newwindow = window.open("printreport.asp?publish=2&Hdate='" + <%=Request("tmpdate")%> + "' ","name","height=800,width=750,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1");
						<% End If%>
						if (window.focus) {newwindow.focus()}
						}
				}
			<%End If%>
			<%If Request("rep") = 26 Then%>
				function PopMe(zzz)
				{
					if (zzz !== undefined)
						{
						<% If Request.Cookies("LBUSERTYPE") = 2 Then %>	
							newwindow = window.open("printreport.asp?publish=1&Hdate='" + <%=Request("tmpdate")%> + "'&selRP=" + <%=Request("tmpRP")%> + "&mytype=" + <%=Request("mytype")%>,"name","height=800,width=750,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1");
						<% Else %>
							newwindow = window.open("printreport.asp?publish=2&Hdate='" + <%=Request("tmpdate")%> + "'&mytype=" + <%=Request("mytype")%>,"name","height=800,width=750,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1");
						<% End If%>
						if (window.focus) {newwindow.focus()}
						}
				}
				<%End If%>
			<%If Request("rep") = 1 Then%>
				function PopMe(zzz)
			{
				if (zzz !== undefined)
					{
					<% If Request.Cookies("LBUSERTYPE") = 2 Then %>	
						newwindow = window.open("printreport.asp?publish=1&Hmonth='" + <%=Request("tmpM")%> + "'&selRP=" + <%=Request("tmpRP")%>,"name","height=800,width=750,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1");
					<% Else %>
						newwindow = window.open("printreport.asp?publish=1&Hmonth='" + <%=Request("tmpM")%> + "' ","name","height=800,width=750,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1");
					<% End If%>
					if (window.focus) {newwindow.focus()}
					}
			}
			<%End If%>
		<%End If%>
		function findSNHMC(xxx, yyy)
		{
			newwindow = window.open("calSNHMC.asp?InstID=" + xxx + "&appdate=" + yyy,"name","height=800,width=850,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1");
			if (window.focus) {newwindow.focus()}
		}
		function tsheets()
		{
			document.frmCal.action = "tsheet.asp?tmpDate=" + '<%=tmpDate%>';
			document.frmCal.submit();	
		}
		function mileage()
		{
			document.frmCal.action = "mileage.asp?tmpMonth=" + '<%=myMonth%>' + '&tmpYear=' + '<%=myYear%>';
			document.frmCal.submit();	
		}
		function disFind() {
			if (document.frmCal.selInst.value == 0) {
				document.frmCal.btnFIND.disabled = true;
			}
			else {
				document.frmCal.btnFIND.disabled = false;
			}
		}
		//-->
		</script>
		<style type="text/css">
	 	.container
	      {
	          border: solid 1px black;
	          overflow: auto;
	      }
	      .noscroll
	      {
	          position: relative;
	          background-color: white;
	          top: expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
	</head>
	<body onload="disFind();
		<%If Request("rep") <> "" Then%>
			PopMe(<%=Request("rep")%>);
		<%End If%>
		"
		>
		<form method='post' name='frmCal'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td height='100px' valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table cellSpacing='2' cellPadding='2' border='0' align='center'>
							<!-- #include file="_greetme.asp" -->
							<tr>
								<td style='width: 75%;' valign='top'>
									<table cellSpacing='2' cellPadding='2' style='width: 100%;' border='0' align='center'>
										<tr>
											<td colspan='2' align='center' class='timeheader'>
												<%=FormatDateTime(Cdate(tmpDate), 1)%><%=myHol%>
											</td>
										</tr>
										<tr>
											<td>
												<div class='container' style='height: 440px; width:100%; position: relative;'>
													<table cellSpacing='2' cellPadding='0' height='100%' width='100%' border='0' align='left' bgcolor='#FFFFFF'>
														<thead>
															<tr bgcolor='#D4D0C8' class="noscroll">
																<td align='center'  class='time2'>&nbsp;</td>
																<td align='center'  class='time2'>&nbsp;</td>
																<td align='center'  class='time2'>Time</td>
																<% If Cint(Request.Cookies("LBUSERTYPE")) <> 4 Then %>
																	<td align='center'  class='time2'>Client</td>
																<% End If %>
																<td align='center'  class='time2'>Institution</td>
																<td align='center'  class='time2'>Language</td>
																<td align='center'  class='time2'>Interpreter</td>
																<td align='center'  class='time2'>Phone</td>
																<td align='center'  class='time2'>Status</td>	
															</tr>
														</thead>
														<tbody style="OVERFLOW: auto;">
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x12a + 1%>' >
																	12&nbsp;AM
																</td>
																<% If str12a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str12a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x1a + 1%>'>
																	1&nbsp;AM
																</td>
																<% If str1a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str1a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x2a + 1%>'>
																	2&nbsp;AM
																</td>
																<% If str2a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str2a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x3a + 1%>'>
																	3&nbsp;AM
																</td>
																<% If str2a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str3a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x4a + 1%>'>
																	4&nbsp;AM
																</td>
																<% If str4a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str4a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x5a + 1%>'>
																	5&nbsp;AM
																</td>
																<% If str5a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str5a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x6a + 1%>'>
																	6&nbsp;AM
																<% If str2a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str6a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x7a + 1%>'>
																	7&nbsp;AM
																</td>
																<% If str7a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str7a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x8a + 1%>'>
																	8&nbsp;AM
																</td>
																<% If str8a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
																	<%=str8a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x9a + 1%>'>
																	9&nbsp;AM
																</td>
																<% If str9a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str9a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x10a + 1%>'>
																	10&nbsp;AM
																</td>
																<% If str10a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str10a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x11a + 1%>'>
																	11&nbsp;AM
																</td>
																<% If str11a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str11a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x12p + 1%>'>
																	12&nbsp;PM
																</td>
																<% If str12p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str12p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x1p + 1%>'>
																	1&nbsp;PM
																</td>
																<% If str1p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str1p%>
															</tr>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x2p + 1%>'>
																	2&nbsp;PM
																</td>
																<% If str2p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str2p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x3p + 1%>'>
																	3&nbsp;PM
																</td>
																<% If str3p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str3p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x4p + 1%>'>
																	4&nbsp;PM
																</td>
																<% If str4p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str4p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x5p + 1%>'>
																	5&nbsp;PM
																</td>
																<% If str5p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str5p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x6p + 1%>'>
																	6&nbsp;PM
																</td>
																<% If str6p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str6p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x7p + 1%>'>
																	7&nbsp;PM
																</td>
																<% If str7p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str7p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x8p + 1%>'>
																	8&nbsp;PM
																</td>
																<% If str8p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str8p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x9p + 1%>'>
																	9&nbsp;PM
																</td>
																<% If str9p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str9p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x10p + 1%>'>
																	10&nbsp;PM
																</td>
																<% If str10p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str10p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x11p + 1%>'>
																	11&nbsp;PM
																</td>
																<% If str11p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str11p%>
														</tbody>
													</table>
												</div>
											</td>
										</tr>
										<tr>
											<td align='right'>
												Legend: <font color='#000000' size='2'>•</font>&nbsp;-&nbsp;Completed&nbsp;<font color='#0000FF' size='2'>•</font>&nbsp;-&nbsp;Missed&nbsp;<font color='#FF0000 ' size='2'>•</font>&nbsp;-&nbsp;Canceled&nbsp;
													<font color='#FF00FF' size='2'>•</font>&nbsp;-&nbsp;Canceled (billable)
											</td>
										</tr>
									</table>
								</td>
								<td valign='top' style='width: 25%;'>	
									<table cellSpacing='2' cellPadding='2' style='width: 95%;' border='0' align='center'>
										<tr>
											<td colspan='7'>
												<table cellSpacing='0' cellPadding='0' align='center' style='width: 100%;' border='0'>
													<tr>
														<td align='left'>
															<input class='btn' type='button' value='&lt&lt' title='Previous Month' style='width: 25px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(0);'>
														</td>
														<td colspan='7' align='center' class='calheader'>
															<%=tmpMonth%>
														</td>
														<td align='right'>
															<input class='btn' type='button' value='&gt&gt' title='Next Month' style='width: 25px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(1);'>
														</td>
													</tr>
													<tr>
														<td>&nbsp;</td>
														<td class='calweekday'>Sun</td>
														<td class='calweekday'>Mon</td>
														<td class='calweekday'>Tue</td>
														<td class='calweekday'>Wed</td>
														<td class='calweekday'>Thu</td>
														<td class='calweekday'>Fri</td>
														<td class='calweekday'>Sat</td>
														<td>&nbsp;</td>
													</tr>
													<%=strCal%>
													<tr><td>&nbsp;</td></tr>
													<tr>
														<td colspan='9' align='center'>
															<table cellSpacing='0' cellPadding='0' style='width: 100%; height: 100%;' border='0' align='center'>
																<tr>
																	<td align='center'>
																		<nobr>Month:
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
																		<input class='btn' type='button' value='GO' style='width: 25px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SearchMonth();'>
																	</td>
																</tr>
																<tr><td>&nbsp;</td></tr>
																<tr><td><hr align='center' width='75%'></td></tr>
																<tr><td>&nbsp;</td></tr>
																<% If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then %>
																	<tr>
																		<td align='center'>	
																			<table cellspacing='1' cellpadding='1' border='0'>
																				<tr>
																					<td align='left'>Timesheet:</td>
																				</tr>
																				<tr>
																					<td>	
																						<input class='btn' type='button' value='<%=mytsheet%>' style='width: 150px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='tsheets();'>
																					</td>
																				</tr>
																				<tr>
																					<td align='left'>Mileage:</td>
																				</tr>
																				<tr>
																					<td>	
																						<input class='btn' type='button' value='<%=mymileage%>' style='width: 150px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='mileage();'>
																					</td>
																				</tr>
																			</table>
																		</td>
																	</tr>
																<% Else %>
																	<tr>
																		<td align='center'>		
																			<input class='btn' type='button' value='Print Today' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PublishMe2();'>
																			<input class='btn' type='button' value='Assigned Appts.' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PublishMe3();'>
<% If blnDup Then %>																			
																			<br />
																			<input class='btn' type='button' value='Review Duplicates' style='width: 180px; margin-top: 20px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='chkDuplicates();'>
<% End If %>																			
																		</td>
																	</tr>
																<% End If %>
																<% If Cint(Request.Cookies("LBUSERTYPE")) = 0 Or Cint(Request.Cookies("LBUSERTYPE")) = 1 Or Cint(Request.Cookies("LBUSERTYPE")) = 3 Then %>
																	<tr>
																		<td valign='bottom' align='center' style='height: 250px;'>
																			<select class="seltxt" name="selInst" style='width: 175px;' onchange="disFind();">
																				<option value="0">&nbsp;</option>
																				<%=strInst%>
																			</select>
																			<input class='btn' type='button' value='FIND' name="btnFIND" style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="findSNHMC(document.frmCal.selInst.value, '<%=tmpDate%>');">
																			<!--<input class='btn' type='button' value='SNHMC' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='findSNHMC();'>//-->
																		</td>
																	</tr>
																<% End If %>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
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
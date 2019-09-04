<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
Function Z_FriendlyTime(dtZ)
	If Not IsDate(dtZ) Then Exit Function
	strAPI = "am"
	lngThr = Z_CLng(DatePart("h", dtZ))
	If lngThr > 12 Then
		strAPI = "pm"
		lngThr = lngThr - 12
	End If
	lngTmn = Z_CLng(DatePart("n", dtZ))
	If lngTmn < 10 Then
		strAPI = "0" & lngTmn & strAPI
	Else
		strAPI = lngTmn & strAPI
	End If
	Z_FriendlyTime = lngThr & ":" & strAPI
End Function
Function Avail(myID, myTime)
	Avail = False
	Set rsAvail = Server.CreateObject("ADODB.RecordSet")
	sqlAvail = "SELECT [index] FROM Avail_T WHERE intrID = " & myID & " AND Avail = '" & myTime & "'"
	rsAvail.Open sqlAvail, g_strCONN, 3, 1
	If Not rsAvail.EOF Then Avail = True
	rsAvail.Close
	set rsAvail = Nothing
	If Avail Then Exit Function
	Set rsAvail2 = Server.CreateObject("ADODB.RecordSet")
	sqlAvail2 = "SELECT [index] FROM Avail_T WHERE IntrID = " & myID
	rsAvail2.Open sqlAvail2, g_strCONN, 3, 1
	If rsAvail2.EOF Then Avail = True
	rsAvail2.Close
	set rsAvail2 = Nothing
End Function

Set rsReq = Server.CreateObject("ADODB.REcordSet")
sqlReq = "SELECT [appDate], [apptimefrom], [apptimeto], [ccity], [CliAdd]" & _
		", r.[InstID], [DeptID], d.[city], COALESCE(l.[language], '') AS [language] " & _
		"FROM [request_T] AS r " & _
		"INNER JOIN [dept_T] AS d ON r.[DeptID]=d.[index] " & _
		"INNER JOIN [language_T] AS l ON r.[langid]=l.[index] " & _
		"WHERE r.[index]=" & Request("ID")
rsReq.Open sqlReq, g_strCONN, 1, 3
If Not rsReq.EOF Then
	appDate = Z_MDYDate(rsReq("appdate"))
	'appTime = rsReq("apptimefrom") & " - " & rsReq("apptimeto")
	appTime = Z_FriendlyTime(rsReq("apptimefrom")) & " to " & Z_FriendlyTime(rsReq("apptimeto"))
	tmpAppTFrom = rsReq("apptimefrom")
	tmpAppTTo = rsReq("apptimeto")
	appCity = rsReq("city")
	If rsReq("cliadd") Then appCity = rsReq("Ccity")
	strLang = rsReq("language")
	IntrLang = Ucase(rsReq("language"))
	tmpAvail = Weekday(appdate) & "," & Hour(tmpAppTFrom)
	tmpInst = rsReq("instID")
	tmpDept = rsReq("deptID")
End If
rsReq.Close
Set rsReq = Nothing


If Request("mail") = 1 Then 'Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	'SEND THE MESSAGE
	' -- get email info
	strTo = zGetInterpreterEmailByID(Request("selIntr"))
	strBcc = "sysdump1@ascentria.org"
	strSubject = strLang & " Appointment on " & appDate & " " & appTime & " in " & appCity
	strMSG = "<p>Are you available to do this appointment?</p>"
	If (Z_Czero(Request("txtMile")) > 0) Then
		strMSG = strMSG & "<p>If you accept this appointment this is the amount you will be reimbursed " & _
				"for mileage and travel time.</p>Payable travel time is: " & Z_Czero(Request("txtTravel")) & " hrs.<br /> " & _
				"Payable mileage is: " & Z_Czero(Request("txtMile")) & "  miles.<br /><br />"
	End If
	strMSG = strMSG  & "<p>Please reply to this email or contact " & Request.Cookies("LBUsrName") & _
			" of LanguageBank.</p><p>Thank you.</p>"

	' -- try to send the message out
	lngMesgErr = zSendMessage(strTo, strBCC, strSubject, strMSG)
	
	IntrName = GetIntr2(Request("selIntr"))
	If lngMesgErr<>0 Then
		' -- got an error
		strErrMsg = "[" & Now & "]: Availability Email sent to " & IntrName & "<" & strTo & "> with result=" & lngMesgErr & ". Message may have to be resent."
		Session("MSG") = "Email process returned: " & lngMesgErr & ". Contact Tech Support."
	Else
		strErrMsg = "[" & Now & "]: Availability Email sent to " & IntrName
		Session("MSG") = "Email Sent."
	End If

	' Save lngMesgErr to notes
	Set rsNotes = Server.CreateObject("ADODB.RecordSet")
	sqlNotes = "SELECT [LBComment] FROM request_T WHERE [index] = " & Request("ReqID")
	rsNotes.Open sqlNotes, g_StrCONN, 1, 3
	If Not rsNotes.EOF Then
		rsNotes("LBComment") = rsNotes("LBComment") & vbCrlF & strErrMsg
		rsNotes.Update
	End If
	rsNotes.CLose
	Set rsNotes = Nothing
	If SaveHist(Request("ReqID"), "emailIntr.asp") Then
		'Session("MSG") = "HIST SAVED"
	End If
	 ' Session("MSG") = "Email Sent."
End If
	

'PREPARE EMAIL	
strAvail = 0
'strSubj = "Appointment on " & appDate & " at " & apptime & " in " & appCity
strSubj = strLang & " Appointment on " & appDate & " " & appTime & " in " & appCity
strMSG = "Are you available to do this appointment?" & vbCrlf & vbCrlf
If (Z_Czero(Request("txtMile")) > 0) Then
	strMSG = strMSG  & "If you accept this appointment this is the amount you will be reimbursed for mileage " & _
			"and travel time." & vbCrlf & vbCrlf & "Payable travel time is " & Z_Czero(Request("txtTravel")) & _
			" hrs. and payable mileage is " & Z_Czero(Request("txtMile")) & " miles." & vbCrlf & vbCrlf
End If
strMSG = strMSG  & 	"Please reply to this email or contact " & Request.Cookies("LBUsrName") & " of LanguageBank." & vbCrlf & vbCrlf & _
		"Thank you."
'INTERPRETER LIST
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE  Active = 1 " & _
		"AND [e-mail] <> '' " & _
		"AND (Upper(Language1) = '" & IntrLang & "' OR Upper(Language2) = '" & IntrLang & _
				"' OR Upper(Language3) = '" & IntrLang & "' OR Upper(Language4) = '" & IntrLang & _
				"' OR Upper(Language5) = '" & IntrLang & "') " & _
		"ORDER BY [Last Name], [First Name]"
rsIntr.Open sqlIntr, g_strCONN, 1, 3
Do Until rsIntr.EOF
	'include vacation
	myIntr = ""
	mark = 0
	If cint(Request("selIntr")) = rsIntr("index") Then myIntr = "selected"
	tmpIntr = cint(Request("selIntr"))
	tmpIntrName = rsIntr("Last Name") & ", " & rsIntr("First Name")
	'If OnVacation(rsIntr("index"), appdate) = False Then 'If isNull(rsIntrLang("vacfrom")) Then
	If tmpIntr = rsIntr("index") Or (Avail(rsIntr("index"), tmpAvail) And NotRestrict(rsIntr("index"), tmpInst, tmpDept)) Then
		mark = 1
		rest = ""
		If NotRestrict(rsIntr("index"), tmpInst, tmpDept) = false Then rest = " (restricted)"
		strIntr = strIntr	& "<option " & myIntr & " value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & rest & "</option>" & vbCrlf
		tmpIntrName = CleanMe(rsIntr("last name")) & ", " & CleanMe(rsIntr("first name"))
		strAvail = SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo)
	End If
	'End If
	If mark = 0 And NotRestrict(rsIntr("index"), tmpInst, tmpDept) Then 'And OnVacation(rsIntr("index"), appdate) = False Then
		strIntr2 = strIntr2 & "<option value='" & rsIntr("index") & "' " & IntrSel & ">" & tmpIntrName & "</option>" & vbCrlf
		strAvail = SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo)
	End If

	strJScript2 = strJScript2 & "if (Intr == " & rsIntr("Index") & ") {" & vbCrLf & _
		"document.frmEmail.chkAvail.value = " & SkedCheck(rsIntr("Index"), Request("ID"), appdate, tmpAppTFrom, tmpAppTTo) &"}; " & vbCrLf & _
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
If strIntr2 <> "" Then strIntr = strIntr & "<option value='0'>-----</option>" & vbCrlf & strIntr2
'End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Email Interpreter</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
			function sendMe()
			{
				if (document.frmEmail.selIntr.value == 0)
				{
					alert("ERROR: Please select an interpreter.")
					return;
				}
				if (document.frmEmail.chkAvail.value == 1) {
					var ans = window.confirm("WARNING: Interpreter already has an appointment for this date and time range.\nPlease check the calendar. \nClick OK to override. \nClick Cancel to stop."); 
					if (ans) {
						document.frmEmail.mail.value = 1;
						document.frmEmail.action = "emailIntr.asp";
						document.frmEmail.submit();	
					}
					else {
						return;
					}
				}
				else {
					document.frmEmail.mail.value = 1;
					document.frmEmail.action = "emailIntr.asp";
					document.frmEmail.submit();	
				}
			}
			function GetMile(xxx)
			{
				if (document.getElementById("btnSend").value == 'Send') {
					document.getElementById("btnSend").value = "Please Wait..."
					document.getElementById("btnSend").disabled = true
					document.frmEmail.action = "Intrmile.asp?selIntr=" + xxx + "&RID=" + <%=Request("ID")%>;
					document.frmEmail.submit();	
				}
			}
			function IntrInfo(Intr) {
				if (Intr == -1 || Intr == 0) {
					document.frmEmail.chkAvail.value = 0;
				}
				<%=strJScript2%>
			}
		</script>
	</head>
	<body onload='IntrInfo(<%=tmpIntr%>);'>
		<form name='frmEmail' method='post'>
			<table cellpadding='1' cellspacing='0' border='0'>
				<tr><td>&nbsp;
					<input name="mail" value="" type="hidden" />
					</td></tr>
				<tr>
					<td align='right'>&nbsp;Interpreter:</td>
					<td align='left'>&nbsp;
						<select class='seltxt' name='selIntr' style='width: 150px;' onchange='GetMile(this.value); IntrInfo(this.value);' onfocus='IntrInfo(this.value);'>
							<option value='0'>&nbsp;</option>
							<%=strIntr%>
						</select>
					</td>
				</tr>
				<tr>
					<td align='right'>Subject:</td>
					<td align='left'>&nbsp;<b><%=strSubj%></b></td>
				</tr>
				<tr>
					<td align='right' valign='top'>Message:</td>
					<td align='left'>&nbsp;
<textarea name='txtMSG' readonly cols='56' rows='15' style="border: 2px solid #aaa;"><%=strMSG%>
						</textarea>
					</td>
				</tr>
				<tr>
					<td align='right' colspan='2'>
						<input class='btn' type='button' name="btnSend" id="btnSend" value='Send' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='sendMe();'>
						<input class='btn' type='button' value='Close' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='window.close();'>
						<input type='hidden' name='ReqID' value='<%=Request("ID")%>'>
						<input type="hidden" name="id" value='<%=Request("ID")%>' />
						<input type="hidden" name="txtTravel" value='<%=Request("txtTravel")%>' />
						<input type="hidden" name="txtMile" value='<%=Request("txtMile")%>' />
						<input type='hidden' name='adr1'  value='<%=Request("adr1")%>'>
						<input type='hidden' name='adr2'  value='<%=Request("adr2")%>'>
						<input type='hidden' name='zip1'  value='<%=Request("zip1")%>'>
						<input type='hidden' name='zip2'  value='<%=Request("zip2")%>'>
						<input type='hidden' name='chkAvail'>	
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
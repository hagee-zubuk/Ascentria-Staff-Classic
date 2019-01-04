<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
tmpPage = "document.frmTbl."
frmdte = Request("frmdte")
todte = Request("todte")
myLang = Request("selLang")
Dim selClass(6)
For lngI = 0 To 6
	selClass(lngI) = ""
Next
Server.Scripttimeout = 360000
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT [Index], [language] FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	LangSel = ""
	If Z_CZero(myLang) = rsLang("Index") Then LangSel = "selected"
	strLang = strLang	& "<option value='" & rsLang("Index") & "' " & LangSel & ">" &  rsLang("language") & "</option>" & vbCrlf
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
If Request.ServerVariables("REQUEST_METHOD") = "POST" Or Request("reload") = 1 Then
	'get open appt
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT r.[index] AS myID, [appdate], [timestamp], r.[instID], [deptID], [LangID], [apptimeFROM], [apptimeTo], [city], [state]" & _
			", [spec_cir], [clname], [cfname], [lbcomment], [comment], [cliadd], [ccity], [cstate], [HPID], [emergency], [emerfee], [class] " & _
			"FROM request_T AS r INNER JOIN language_T AS t ON r.[langID]=t.[index] INNER JOIN dept_T AS d ON r.DeptId = d.[index] " & _
			"WHERE [IntrID] <= 0 " & _
			"AND [status] <> 2 " & _
			"AND [status] <> 3 " & _
			"AND [status] <> 4 "
	If Z_Czero(Request("selLang")) > 0 Then sqlApp = sqlApp & "AND [LangID] = " & Request("selLang") & " "
	If Z_Czero(Request("selClass")) > 0 Then
		sqlApp = sqlApp & "AND [class] = " & Request("selClass") & " "
		selClass( Z_Czero(Request("selClass")) ) = "selected"
	End If
	If (Z_FixNull(Request("frmdte")) = "") AND (Z_FixNull(Request("todte")) = "") Then
		sqlApp = sqlApp & "AND [appdate] >= '" & date & "' "
	Else 
		If Z_FixNull(Request("frmdte")) <> "" Then sqlApp = sqlApp & "AND [appdate] >= '" & Request("frmdte") & "' "
		If Z_FixNull(Request("todte")) <> "" Then
			sqlApp = sqlApp & "AND [appdate] <= '" & Request("todte") & "' "
			If Z_FixNull(Request("frmdte")) = "" Then
				sqlApp = sqlApp & "AND [appdate] >= '" & date & "' "
			End If
		End If
	End If
	sqlApp = sqlApp & "ORDER BY [language], [appdate] "
	'response.write "<code>" & sqlApp & "</code>"
	rsApp.Open sqlApp, g_strCONN, 3, 1
	xctr = 0
	Set rsYes = Server.CreateObject("ADODB.RecordSet")
	Do Until rsApp.EOF
		If Z_IntrYesNo(rsApp("myID")) Then
			kulay = "#FFFFFF"
			If Z_IsOdd(xctr) Then kulay = "#FBEEB7"
			xctr = xctr + 1
			'yes intr
			
			rsYes.Open "SELECT IntrID, appID, ansTS FROM appt_T WHERE accept = 1 AND AppID = " & rsApp("myID"), g_strCONN, 3, 1
			intrCtr = 0
			strYes = ""
			Do Until rsYes.EOF
				intrName = GetIntr(rsYes("intrID"))
				intrCity = "(" & Z_GetIntrCity(rsYes("intrID")) & ")"
				appdate = Z_DateNull(Z_GetAppDate(rsYes("appID")))
				intrPastSked = "Timestamp: " & rsYes("ansTS") & vbCrLf & Z_GetPastSked(rsYes("intrID"), appdate)
				warnAssign = 0
				tmpAppDate = Z_GetInfoFROMAppID(rsApp("myID"), "appdate")
				tmpAppTFrom2 =  Z_GetInfoFROMAppID(rsApp("myID"), "appTimeFrom")
				tmpAppTTo2 =  Z_GetInfoFROMAppID(rsApp("myID"), "appTimeTo")
				'response.write rsYes("intrID") & " - " &  rsApp("index") & " - " & tmpAppDate & " - " & tmpAppTFrom2 & " - " & tmpAppTTo2
				If SkedCheck(rsYes("intrID"), rsApp("myID"), tmpAppDate, tmpAppTFrom2, tmpAppTTo2) = 1 Then warnAssign = 1
				vaca = 0
				If OnVacation(rsYes("intrID"), tmpAppDate) Then vaca = 1
				If intrCtr = 0 Then
					strYes = strYes & "<td class='tblgrnUnder' style='text-align: left;' title=""" & intrPastSked & """><nobr>" & intrName & " " & intrCity & _
						"</td>" & vbCrLf & "<td class='tblgrnUnder'><input class='btntbl' type='button' value='ASSIGN' style='width: 50px;' onmouseover=""this.className='hovbtntbl'"" onmouseout=""this.className='btntbl'"" onclick='Assign(" & warnAssign & "," & vaca & "," & rsYes("intrID") & "," & rsApp("myID") & ",""" & Request("frmdte") & """,""" & Request("todte") & """,""" & intrName & """," & myLang & ");'></td></tr>" & vbCrLf
				Else
					strYes = strYes & "<tr bgcolor='" & kulay & "'><td class='tblgrnUnder' style='text-align: left;' title=""" & intrPastSked & """><nobr>" & intrName & " " & intrCity & _
						"</td>" & vbCrLf & "<td class='tblgrnUnder'><input class='btntbl' type='button' value='ASSIGN' style='width: 50px;' onmouseover=""this.className='hovbtntbl'"" onmouseout=""this.className='btntbl'"" onclick='Assign(" & warnAssign & "," & vaca & "," & rsYes("intrID") & "," & rsApp("myID") & ",""" & Request("frmdte") & """,""" & Request("todte") & """,""" & intrName & """," & myLang & ");'></td></tr>" & vbCrLf
				End If
				rsYes.MoveNext
				intrCtr = intrCtr + 1
			Loop
			rsYes.Close

			'No intr
			Set rsNo = Server.CreateObject("ADODB.RecordSet")
			rsNo.Open "SELECT IntrID, UID, ansTS FROM appt_T WHERE accept = 2 AND AppID = " & rsApp("myID"), g_strCONN, 3, 1
			strNo = ""
			Do Until rsNo.EOF
				strNo = strNo & "<tr><td class='tblgrn2' aling='left'><a class='SmallLink' title='Timestamp: " & rsNO("ansTS") & "' href=""JavaScript: ResetMe(" & rsNo("UID") & ",'" & frmdte & "','" & todte & "', '" & GetIntrFN(rsNo("intrID")) & "', " & myLang & ");"">NO-" & GetIntrFN(rsNo("intrID")) & "</a></td></tr>"
				'strNo = strNo & NointrName
				rsNo.MoveNext
			Loop
			rsNo.Close
			Set rsNo = Nothing	
			ts = rsApp("timestamp")
			tmpIname = GetInst(rsApp("instID"))
			myDept = GetDept(rsApp("deptID"))
			tmpSalita = GetLang(rsApp("langID"))
			timeframe = Z_FormatTime(rsApp("appTimeFrom"), 4) & " - " & Z_FormatTime(rsApp("appTimeTo"), 4)
			clntname = rsApp("clname") & ", " & rsApp("cfname")
			
			If rsApp("cliAdd") Then
				tmpcity = Trim(rsApp("ccity"))
				If Z_FixNull(rsApp("cstate")) <> "" Then tmpCity = tmpCity & ", " & rsApp("cstate")
			Else
				tmpcity = Trim(rsApp("city"))
				If Z_FixNull(rsApp("state")) <> "" Then tmpCity = tmpCity & ", " & rsApp("state")
			End If
			tmpHPID = Z_CZero(rsApp("HPID"))
			IntrCommentTitle = ""
			IntrComment = ""
			emer = ""
			If rsApp("Emergency") Then emer = "checked"
			emerfee = ""
			If rsApp("Emerfee") Then emerfee = "checked"
			LBcomment = rsApp("LBcomment")
			AppReason = ""
			IntrComment = ""
			IntrCommentTitle = ""
			tmpsc = rsApp("spec_cir")
			If Len(tmpsc) > 25 Then 
				tmpsc = Left(tmpsc, 22) & "..."
				tmpsctitle = rsApp("spec_cir")
			End If
			If tmpHPID <> 0  THen
				Set rsHP = Server.CreateObject("ADODB.RecordSet")
				sqlHP = "SELECT * FROM Appointment_T WHERE [index] = " & tmpHPID
				rsHP.Open sqlHP, g_StrCONNHP, 3, 1
				If Not rsHP.EOF Then
					IntrComment =  rsHP("lbcom") & "<br>" & rsHP("comment")
					IntrCommentTitle = rsHP("lbcom") & "<br>" & rsHP("comment")
					If Len(IntrComment) > 25 Then 
						IntrComment = Left(IntrComment, 22) & "..."
						IntrCommentTitle = rsHP("lbcom") & vbCrLf & rsHP("comment")
					End If
					AppReason = GetReas(Z_Replace(rsHP("reason"),", ", "|"))
				End If
				rsHp.Close
				Set rsHp = Nothing
			End If
			tmpCom = rsApp("Comment")
			tmpComTitle = rsApp("Comment")
			If Len(tmpCom) > 25 Then 
				tmpCom = Left(tmpCom, 22) & "..."
				tmpComTitle = rsApp("Comment")
			End If
			strtbl = strtbl & "<tr bgcolor='" & kulay & "' style='vertical-align: top;'>" & vbCrLf & _ 
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'><input type='hidden' name='hid" & x & "' value='" & rsApp("myID") & "' ><a class='link2' href='reqconfirm.asp?ID=" & rsApp("myID") & "'><b>" & rsApp("myID") & "</b></a></td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & ts & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & tmpIname & " - " &  myDept & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & tmpCity & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & tmpSalita & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & rsApp("appdate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & timeframe & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & clntname & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'>" & AppReason & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "' title='" & tmpsctitle & "'>" & tmpsc & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "' title='" & tmpComTitle & "'>" & tmpCom & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "' title='" & IntrCommentTitle & "'>" & IntrComment & "</td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'><input type='checkbox' name='chkEmer" & xctr & "' " & emer & " value='1'></td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'><input type='checkbox' name='chkEmerfee" & xctr & "' " & emerfee & " value='1'></td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'><textarea name='txtLBcom" & xctr & "' class='main' onkeyup='bawal(this);' style='width: 200px;' rows='2'>" & LBcomment & "</textarea></td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'><input class='btntbl' type='button' value='SAVE' style='width: 50px;' onmouseover=""this.className='hovbtntbl'"" onmouseout=""this.className='btntbl'"" onclick=""SaveMe(" & rsApp("myID") & "," & xctr & ",'" & frmdte & "','" & todte & "'," & myLang & ");""></td>" & vbCrLf & _
				"<td class='tblgrn2' rowspan='" & IntrCtr & "'><table border='0'>" & strNo & "</table></td>" & vbCrLf & _
				strYes & vbCrLf
		End If
		rsApp.MoveNext
	Loop
	Set rsYes = Nothing	
	rsApp.Close
	Set rsApp = Nothing
End If
<!-- #include file="_closeSQL.asp" -->
%>
<html>
	<head>
		<title>Language Bank - Open Appointments</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function ResetMe(myid, frmdte, todte, intrname, myLang) {
			var ans = window.confirm("Reset answer for " + intrname + "?\nClick Cancel to stop.");
			if (ans) {
				document.frmTbl.action = "resetme.asp?appID=" + myid + "&frmdte=" + frmdte + "&todte=" + todte + "&selLang=" + myLang;
				document.frmTbl.submit();
			}
			else {
				return;
			}
		}
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
		}
			function Assign(warnIntr, vaca, intrID, appID, frmdte, todte, IntrName, myLang) {
				var txt = "Assign " + IntrName + " for Appointment " + appID + "? \nClick Cancel to stop.";
				var ans = window.confirm(txt);
				if (ans) {
					if (warnIntr == 1) {
						var ans = window.confirm("WARNING: Interpreter already has an appointment for this date and time range.\nPlease check the calendar. \nClick OK to override. \nClick Cancel to stop."); 
						if (ans) {
							document.frmTbl.action = "apptassign.asp?intrID=" + intrID + "&appID=" + appID + "&frmdte=" + frmdte + "&todte=" + todte + "&selLang=" + myLang;
							document.frmTbl.submit();
						}
						else {
							return;
						}
					}
					else if (vaca == 1) {
						alert("Interpreter on vacation.\n Assigning not allowed.");
						return;
					}
					else {
						document.frmTbl.action = "apptassign.asp?intrID=" + intrID + "&appID=" + appID + "&frmdte=" + frmdte + "&todte=" + todte + "&selLang=" + myLang;
						document.frmTbl.submit();
					}
				}
			}
			function CalendarView(strDate)
		{
			document.frmTbl.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmTbl.submit();
		}
		function SaveMe(xxx, ctr, frmdte, todte, myLang) {
			document.frmTbl.action = "SaveOpen.asp?ReqID=" + xxx + "&ctr=" + ctr + "&frmdte=" + frmdte + "&todte=" + todte + "&selLang=" + myLang;
			document.frmTbl.submit();
		}
		function FindOpen(frmdte, todte, myLang) {
			document.frmTbl.action = "Openappts.asp?frmdte=" + frmdte + "&todte=" + todte + "&selLang=" + myLang;
			document.frmTbl.submit();
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
	          top:expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
		<body >
			<form method='post' name='frmTbl'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top'>
							<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
								<tr>
									<td>
										<table cellpadding='0' cellspacing='0' width='100%' border='0'>
											<tr>
												<td align='left'>
													
												</td>
												
														<td align='right'>
															&nbsp;
															<input type='hidden' name='Hctr' value='<%=x%>'>
														</td>
											</tr>
										</table>
									</td>
								</tr>
								<% If Session("MSG") <> "" Then %>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='14' align='left'>
											<div name="dErr" style="width:300px; height:40px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
								<% End If %>
								<tr>
									<td colspan='10' align='left'>
										<table class="reqtble">	
											<tr>
												<td align='right'>Date Range:</td>
												<td>
													<input class='main' size='10' maxlength='10' name='txtFromDate'  readonly value='<%=frmdte%>'>
													<input type="button" value="..." title='Calendar' name="calFrom" style="width: 19px;"
													onclick="showCalendarControl(document.frmTbl.txtFromDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
													&nbsp;&nbsp;-&nbsp;&nbsp;
													<input class='main' size='10' maxlength='10' name='txtToDate'  readonly value='<%=todte%>'>
													<input type="button" value="..." title='Calendar' name="calTo" style="width: 19px;"
													onclick="showCalendarControl(document.frmTbl.txtToDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
													&nbsp;&nbsp;
													Language:
													<select class="seltxt" style='width: 150px;' name="selLang">
														<option value="0">&nbsp;</option>
														<%=strLang%>
													</select>
													&nbsp;&nbsp;
Classification:
<select class='seltxt' style='width: 100px;' name='selClass' id='selClass'>
	<option value='-1'>&nbsp;</option>
	<option value='1' <%=selClass(1)%>>Social Services</option>
	<option value='2' <%=selClass(2)%>>Private</option>
	<option value='3' <%=selClass(3)%>>Court</option>
	<option value='4' <%=selClass(4)%>>Medical</option>
	<option value='5' <%=selClass(5)%>>Legal</option>
	<option value='6' <%=selClass(6)%>>Mental Health</option>
</select>
&nbsp;
													<input type='button' value='Search' name='btnSearch' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='FindOpen(document.frmTbl.txtFromDate.value, document.frmTbl.txtToDate.value, document.frmTbl.selLang.value);'>
												</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr>
									<td colspan='10' align='left'>
										<div class='container' style='height: 500px; width:95%; position: relative;'>
											<table class="reqtble" width='100%' >	
												<thead>
													<tr class="noscroll">	
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Request ID</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Date Requested</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Institution - Department</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">City</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Language</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Appointment Date</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Start and End Time</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Client</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Reason</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Special Circumstances/Precautions</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Comment</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">LB Comment (Preferred/Unpreferred Interpreter)</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Emergency</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Emergency Fee</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">LB Notes</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">&nbsp;</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">NO interpreters</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Available Interpreter</td>
														<td class='tblgrn' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">&nbsp;</td>
													</tr>
												</thead>
												<tbody style="OVERFLOW: auto;">
													<%=strtbl%>
												</tbody>
											</table>
										</div>	
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table width='100%'  border='0'>
								<tr>
									<td align='left'>
										&nbsp;
									</td>
									<td align='right'>
										<% If xctr <> 0 Then %>
											<b><u><%=xctr%></u></b> record/s &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<% End If %>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td height='50px' valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
	</head>
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
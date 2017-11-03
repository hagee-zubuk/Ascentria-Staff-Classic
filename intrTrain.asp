<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
tmpPage = "document.frmAdmin."
server.scripttimeout = 360000
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then
	Session("MSG") = "Error: Please sign-in as user with admin rights."
	Response.Redirect "default.asp"
End If
'create train List
Set rsTrain = Server.CreateObject("ADODB.RecordSet")
sqlTrain = "SELECT * FROM Training_T ORDER BY Training"
rsTrain.Open sqlTrain, g_strCONN, 1, 3
Do Until rsTrain.EOF
	strTrain1 = strTrain1 & "<option value=" & rsTrain("index") & ">" & rsTrain("Training") & "</option>"
	rsTrain.MoveNext
Loop
rsTrain.Close
Set rsTrain = Nothing
ScriptName = Request.ServerVariables("SCRIPT_NAME")
''''''Site
NumPerPage = 10

If Request.QueryString("LogPage") = "" Then
	CurrPage = 1
Else
	CurrPage = CInt(Request.QueryString("LogPage"))
End if
tmpName = GetIntr(Request("IntrID"))
Set rsLog = Server.CreateObject("ADODB.RecordSet")
sqlLog = "SELECT * FROM [IntrTraining_T] WHERE [intrID] = " & Request("IntrID") & " ORDER BY [date] DESC"
rsLog.Open sqlLog, g_strCONN, 1, 3
TotalPages = 10
If Not rsLog.EOF Then
	rsLog.MoveFirst
	rsLog.PageSize = NumPerPage
	TotalPages = rsLog.PageCount
	rsLog.AbsolutePage = CurrPage
Else
	strLog = "<tr><td colspan='5' align='center'><font size='1'><i><-------- N/A -------></i></font></td></tr>"
End If
ctr = 0
Do While Not rsLog.EOF And ctr < rsLog.PageSize
	If rsLog("date") <> "" Then
		if Z_IsOdd(ctr) = true then 
			kulay = "#FFFAF0" 
		else 
			kulay = "#FFFFFF"
		end if
		tmpTrain = rsLog("type")
		CertMe = ""
		If tmpTrain = 3 Then CertMe = "disabled"
		strLog = strLog & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkeval" & ctr & _
			"' value='" & rsLog("index") & "'></td><td align='center'><input type='text' " & _
			"style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtDate" & ctr & "' value='" & _
			rsLog("date") & "' class='main'><td align='center'><input type='text' " & _
			"style='font-size: 10px; height: 20px;' size='6' maxlength='5' name='txthrs" & ctr & "' value='" & _
			rsLog("hours") & "' class='main'><td align='center'><select " & CertMe & " class='seltxt' name='selTrain" & ctr & "' style='width: 100px;'><option value='0'>&nbsp;</option>"
		

		'create train List
		strTrain = ""
		Set rsTrain = Server.CreateObject("ADODB.RecordSet")
		sqlTrain = "SELECT * FROM Training_T ORDER BY Training"
		rsTrain.Open sqlTrain, g_strCONN, 1, 3
		Do Until rsTrain.EOF
			trainselect = ""
			certme = ""
			If tmpTrain = rsTrain("index") Then trainselect = "Selected"
			strTrain = strTrain & "<option value=" & rsTrain("index") & " " & trainselect & ">" & rsTrain("Training") & "</option>"
			rsTrain.MoveNext
		Loop
		rsTrain.Close
		Set rsTrain = Nothing
		strLog = strLog & strTrain	
		
		strLog = strLog & "</select></td>" & vbCrLf
		
		If tmpTrain = 3 Then
			strLog = strLog & "<td valign='top' align='center'>" & _
				"<input class='main' style='font-size: 10px; height: 20px;' size='51'  maxlength='50' name='txtCert" & ctr & "' " & _
				"value='" & rsLog("cert") & "'>" & _
				"<input type='hidden' name='hidcert" & ctr & "' value='3'></td>" & vbCrLf
		End If
		strLog = strLog & "</tr>" & vbCrLf
		'rsLog.MoveNext
		ctr = ctr + 1
	End If
	rsLog.MoveNext
	
Loop
rsLog.Close

Set rsLog = Nothing

%>
<html>
	<head>
		<title>Language Bank - Interpreter Page - Training</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
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
			function CalendarView(strDate)
			{
				document.frmAdmin.action = 'calendarview2.asp?appDate=' + strDate;
				document.frmAdmin.submit();
			}
			function SaveMe(xxx)
			{
				document.frmAdmin.action = "intraction.asp?action=5&intrID=" + xxx;
				document.frmAdmin.submit();
			}
			function KillMe(xxx)
			{
				document.frmAdmin.action = "intraction.asp?action=6&intrID=" + xxx;
				document.frmAdmin.submit();
			}
			function certme(xxx)
			{
				document.frmAdmin.txtCert.disabled = true;
				if (xxx == 3)
				{
					document.frmAdmin.txtCert.disabled = false;
				}
			}
		-->
		</script>	
	</head>
	<body>
		<form method='post' name='frmAdmin'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top' >
						<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
							<!-- #include file="_greetme.asp" -->
							<% If Request("IntrID") = 0 Then %>
								<tr>
									<td class='header' colspan='10'><nobr>Interpreter</td>
								</tr>
							<% Else %>
								<tr>
									<td class='header' colspan='10'>
										<a href='adminIntr.asp?intrID=<%=Request("intrID")%>&type=<%=Request("type")%>'' class='intrlink'>[Interpreter Info]</a> 
										<a href='intreval.asp?intrID=<%=Request("intrID")%>&type=<%=Request("type")%>'' class='intrlink'>[Evalution/Feedback]</a> 
										<nobr>Training
										<a href='intrsched.asp?intrID=<%=Request("intrID")%>&type=<%=Request("type")%>'' class='intrlink'>[Schedule]</a>
									</td>
								</tr>
							<% End If %>
							<tr><td>
							<div style='width:100%; height:100%;OVERFLOW: AUTO; ' align='center'>
								<table cellspacing='0' cellpadding='0' border='0' align='center'>
									<tr>
										<td align='left' colspan='5'><u>INTERPRETER:</u> <%=tmpName%></td>
									</tr>
									<tr>
										<td valign='top'>
										<table border='0'>
											<tr bgcolor='#336601'>
												<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#ffffffff, endColorstr=#336601);">
													<font size='2' face='trebuchet MS' color='white'>Training</font>
													</td></tr>
												<tr><td colspan='3' width='175px' align='center'>
													<%  'Site
														If Not CurrPage = 1 Then 
															Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage - 1 &  "&LogPage2=" & CurrPage2 & "&LogPage3=" & CurrPage3 & "'><font size='1' face='trebuchet MS'>Prev</a> | "
															Session("page") = CurrPage 
														Else
															Response.Write "<font size='1' face='trebuchet MS'>Prev | "
														End If
														If Not CurrPage = TotalPages Then
															Response.Write "<a href='" & ScriptName & "?LogPage=" & CurrPage + 1 & "'>Next</font></a>"
															Session("page") = CurrPage 
														Else
															Response.Write "Next</font>"
														End If
														
													%>
												</td>
												<td colspan='3' align='right'><font size='1' face='trebuchet MS'><%=CurrPage%> of <%=TotalPages%></font></td>
											</tr>
											<tr><td>&nbsp;</td>
												<td align='center'><font size='1' face='trebuchet MS'>Date</font></td>
												<td align='center'><font size='1' face='trebuchet MS'>Hours</font></td>
												<td align='center'><font size='1' face='trebuchet MS'>Training</font></td>
												<td align='center'><font size='1' face='trebuchet MS'>Certificate</font></td>
												</tr>
											<%=strLog%>
											<tr bgcolor='#336601'>
												<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#ffffffff, endColorstr=#336601);">
													<font size='2' face='trebuchet MS' color='white'>New Entries</font>
											</td></tr>
											<tr>
												<td>&nbsp;</td>
												<td valign='top' align='center'>
													<input class='main' style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txtdate'>
												</td>
												<td valign='top' align='center'>
													<input class='main' style='font-size: 10px; height: 20px;' size='9'  maxlength='10' name='txthrs'>
												</td>
												<td valign='top' align='center'>
													<select class='seltxt' name='selTrain' style='width: 100px;' onchange='certme(this.value);'>
														<option value='0'>&nbsp;</option>
														<%=strTrain1%>
													</select>
												</td>
												<td valign='top' align='center'>
													<input class='main' style='font-size: 10px; height: 20px;' size='51'  maxlength='50' name='txtCert'>
												</td>
											</tr>
											<tr>
												<td colspan='5' align='center'>
													<input type='button' style='width: 130px;' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"  value='Save Entry' onclick='SaveMe(<%=Request("Intrid")%>);'>
													<input type='button' style='width: 130px;' value='Delete Checked Entry' class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe(<%=Request("Intrid")%>);'>
												</td>
												<input type='hidden' name='ctr' value='<%=ctr%>'>
												<input type='hidden' name='tmpType' value='<%=Request("type")%>'>
											</tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
											<tr><td>&nbsp;</td></tr>
										</table>
									</table>
								</div>
							</td></tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign='bottom'>
						<!-- #include file="_footer.asp" -->
					</td>
				</tr>
			</table>
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
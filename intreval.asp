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
sqlLog = "SELECT * FROM [InterpreterEval_T] WHERE [intrID] = " & Request("IntrID") & " ORDER BY [date] DESC"
response.write "<!--SQL: " & sqlLog & "-->"
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
		strLog = strLog & "<tr bgcolor='" & kulay & "'><td align='center'><input type='checkbox' name='chkeval" & ctr & _
			"' value='" & rsLog("index") & "'></td><td align='center'><input type='text' " & _
			"style='font-size: 10px; height: 20px;' size='9' maxlength='10' name='txtDate" & ctr & "' value='" & _
			rsLog("date") & "' class='main'></td><td colspan='2'><textarea class='main' cols='18' name='txtcom" & ctr & "'>" & rsLog("comment") & "</textarea></td></tr>"  & vbCRlf
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
		<title>Language Bank - Interpreter Page - Evaluation/Feedback</title>
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
				document.frmAdmin.action = "intraction.asp?action=3&intrID=" + xxx;
				document.frmAdmin.submit();
			}
			function KillMe(xxx)
			{
				document.frmAdmin.action = "intraction.asp?action=4&intrID=" + xxx;
				document.frmAdmin.submit();
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
										<nobr>Evalution/Feedback
										<a href='intrtrain.asp?intrID=<%=Request("intrID")%>&type=<%=Request("type")%>'' class='intrlink'>[Training]</a>
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
													<font size='2' face='trebuchet MS' color='white'>Evalution/Feedback</font>
													</td></tr>
												<tr><td colspan='2' width='175px' align='center'>
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
											<tr><td colspan='2' align='center'><font size='1' face='trebuchet MS'>Date</font></td>
												<td colspan='3' align='center'><font size='1' face='trebuchet MS'>Comment</font></td>
												</tr>
											<%=strLog%>
											<tr bgcolor='#336601'>
												<td colspan='5' align='center' style="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0, startColorstr=#ffffffff, endColorstr=#336601);">
													<font size='2' face='trebuchet MS' color='white'>New Entries</font>
											</td></tr>
											<tr>
												<td>&nbsp;</td>
												<td valign='top' align='center'>
													<input style='font-size: 10px; height: 20px;' size='9'  maxlength='10' class='main' name='txtdate'>
												</td>
												<td>
													<textarea class='main' cols='18' name='txtcom'></textarea>
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
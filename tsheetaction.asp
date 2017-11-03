<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Session("UIntr") = "" Then 
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.Redirect "default.asp"
	End If
MyNow = Now
If Request("action") = 1 Then
	If Request("confirm") <> 1 Then		
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T]"
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("ctr" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					tmpAppDate = GetAppDate(tmpctr)
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						If Z_FixNull(Request("sunstart" & i)) <> "" And Z_FixNull(Request("sunend" & i)) <> "" Then
							date1st = Date & " " & Z_dates(Request("sunstart" & i))
							date2nd = Date & " " & Z_dates(Request("sunend" & i))
							If datediff("n", date1st, date2nd) >= 0 Then
								minTime = DateDiff("n", date1st, date2nd)
							Else
								minTime = DateDiff("n", date1st, dateadd("d", 1, date2nd))
							End If
							rsTS("totalhrs") = MakeTime(Z_CZero(minTime))
						End If
						If Z_FixNull(Request("sunstart" & i)) <> "" Then 
							rsTS("AStarttime") = tmpAppDate & " " & Z_dates(Request("sunstart" & i))
						Else
							rsTS("AStarttime") = Empty
						End If
						If Z_FixNull(Request("sunend" & i)) <> "" Then 
							rsTS("AEndtime") = tmpAppDate & " " & Z_dates(Request("sunend" & i))
						Else
							rsTS("AEndtime") = Empty
						End If
						rsTS("payhrs") = Request("hidpayhrs" & i)
						rsTS.Update
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	Else
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T]"
		'On Error Resume Next
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("chkcon" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						rsTS("confirmed") = MyNow
						rsTS.Update
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	End If
	Session("MSG") = "Timesheet saved."
	If Request("confirm") = 1 Then Session("MSG") = "Timesheet confirmed."
	Response.Redirect "tsheet.asp?tmpdate=" & Request("tmpDate")
ElseIf Request("action") = 2 Then
	If Request("confirm") <> 1 Then
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T]"
		'On Error Resume Next
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("ctr" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						rsTS("toll") = Z_CZero(Request("suntoll" & i))
						rsTS.Update
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	Else
		Set rsTS = Server.CreateObject("ADODB.Recordset")
		sqlTS = "SELECT * FROM [request_T]"
		'On Error Resume Next
		rsTS.Open sqlTS, g_strCONN, 1, 3
		ctrI = Request("myctr")
			For i = 0 to ctrI + 1 
				tmpctr = Request("chkcon" & i)
				If tmpctr <> "" Then
					strTmp = "index=" & tmpctr 
					rsTS.Find(strTmp)
					If Not rsTS.EOF Then
						rsTS("confirmedtoll") = MyNow
						rsTS.Update
					End If
				End If
				rsTS.MoveFirst
			Next
		rsTS.Close
		Set rsTS = Nothing
	End If
	Session("MSG") = "Mileage saved."
	If Request("confirm") = 1 Then Session("MSG") = "Mileage confirmed."
	Response.Redirect "mileage.asp?tmpMonth=" & Request("tmpMonth") & "&tmpYear=" & Request("tmpYear")
End If
	
%>
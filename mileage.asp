<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
	Function AmtRate(xxx)
		AmtRate = 0
		If Z_Czero(xxx) = 0 Then
			AmtRate = 0
			Exit Function
		End If
		Set rsRate = Server.CreateObject("ADODB.RecordSet")
		sqlRate = "SELECT * FROM MileageRate_T"
		rsRate.Open sqlRate, g_strCONN, 1, 3
		If Not rsRate.EOF Then
			AmtRate = rsRate("mileageRate") * xxx
		End If
		rsRate.Close
		Set rsRate = Nothing
	End Function
	tmpPage = "document.frmTS."
	'Get week range
	If Request("action") = 1 Then
		tmpYear = Request("tmpYear")
		If Request("tmpMonth") = 1 Then 
			tmpMonth = 12
			tmpYear = tmpYear - 1
		Else
			tmpMonth = Request("tmpMonth") - 1
		End If
	ElseIf Request("action") = 2 Then
		tmpYear = Request("tmpYear")
		If Request("tmpMonth") = 12 Then 
			tmpMonth = 1
			tmpYear = tmpYear + 1
		Else
			tmpMonth = Request("tmpMonth") + 1
		End If
	Else
		tmpMonth = Request("tmpMonth")
		tmpYear = Request("tmpyear")
	End If
	myMileage = MonthName(tmpMonth) & " - " & tmpYear
	Set rsTS = Server.CreateObject("ADODB.RecordSet")
	sqlTS = "SELECT * FROM request_T WHERE Month(appDate) = " & tmpMonth & " AND Year(appDate) = " & tmpYear & " AND IntrID = " & Session("UIntr") & " " & _
		"AND showintr = 1 AND status <> 2 AND Status <> 3 ORDER BY appDate, appTimeFrom"
	rsTS.Open sqlTS, g_strCONN, 1, 3
	'response.write sqlTS
	ctr = 0
	ctr2 = 0
	TottmpAMT = 0
	TotToll = 0
	If Not rsTS.EOF Then
		Do Until rsTS.EOF
			myStat = ""
			'If rsTS("Status") = 2 Or rsTS("Status") = 3 Or rsTS("Status") = 4 Or rsTS("confirmedtoll") <> "" Then myStat = "DISABLED"
			'If rsTS("Status") = 2 Or rsTS("Status") = 3 Or rsTS("Status") = 4 Then myStat = "DISABLED"
			tmpAMT = Z_FormatNumber(rsTS("actMil"), 2)
			TottmpAMT = TottmpAMT + Z_Czero(tmpAMT)
			TotToll = TotToll + Z_CZero(rsTS("toll"))
			tmpCon = ""
			LBcon = ""
			If rsTS("LbconfirmToll") = True Then 
				tmpCon = "checked" '"<b>*</b>"
				LBcon = "readonly"
				ctrCon = ctrCon + 1
			End If
			IntrCon = ""
			If rsTS("confirmedtoll") <> "" Then IntrCon = "disabled checked"
			myAct = rsTS("index") & " - " & GetInst(rsTS("InstID")) & " - " & left(rsTS("Cfname"), 1) & ". " & left(rsTS("Clname"), 1) & "."
			strMile = strMile & "<tr bgcolor='#F5F5F5'><td align='center'><input type='hidden' name='ctr" & ctr & "' value='" & rsTS("index") & "'>" & rsTS("appdate") & "</td>" & _
				"<td align='center'><nobr>" & myAct & "</td>" & _
			 	"<td align='center'>" & tmpAMT & "</td>" & _
				"<td align='center'>$<input class='main' size='6' " & myStat & " maxlength='5' name='suntoll" & ctr & "' value='" & rsTS("toll") & "' " & LBcon & "></td>" & _
				"<td align='center'><input type='checkbox' disabled " &  tmpCon & "></td>"
			If tmpCon = "checked" Then
				strMile = strMile & "<td align='center'><input type='checkbox' name='chkcon" & ctr & "' value='" & rsTS("index") & "' " & IntrCon & "></td>"
				ctr2 = ctr2 + 1
			End If
			strMile = strMile &	"</tr>"
			
			ctr = ctr + 1
			rsTS.MoveNext
		Loop
	End If
	rsTS.Close
	Set rsClose = Nothing
%>
<html>
	<head>
		<title>Language Bank - Interpreter Mileage</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function SaveTS()
		{
			document.frmTS.action = "tsheetaction.asp?action=2";
			document.frmTS.submit();
	
		}
		function SendTS()
		{
			if (document.frmTS.myCTR.value != document.frmTS.myCTR2.value)
			{
				alert("All entries must be approved first before you can confirm.")
				return;
			}
			var ans = window.confirm("Confirm CHECKED Appointments? \nCancel to stop.");
			if (ans)
			{
				document.frmTS.action = "tsheetaction.asp?action=2&confirm=1";
				document.frmTS.submit();
			}
		}
		function CalendarView(strDate)
		{
			document.frmTS.action = 'tsheet.asp?tmpdate=' + strDate;
			document.frmTS.submit();
		}
		function PrevMonth(xxx, yyy)
		{
			document.frmTS.action = "mileage.asp?action=1&tmpMonth=" + xxx + '&tmpYear=' + yyy;
			document.frmTS.submit();	
		}
		function NextMonth(xxx, yyy)
		{
			document.frmTS.action = "mileage.asp?action=2&tmpMonth=" + xxx + '&tmpYear=' + yyy;
			document.frmTS.submit();	
		}
		//-->
		</script>
	</head>
	<body>
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
								<tr>
								<td class='title' colspan='10' align='center'><nobr> Interpreter Mileage</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td  align='center' colspan='10'>
										<div name="dErr" style="width: 250px; height:55px;OVERFLOW: auto;">
											<table border='0' cellspacing='1'>		
												<tr>
													<td><span class='error'><%=Session("MSG")%></span></td>
												</tr>
											</table>
										</div>
									</td>
								</tr>
								<tr>
									<td align='right' width='200px'>Name:</td>
									<td class='confirm'><%=GetIntr(Session("UIntr"))%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Date:</td>
									<td class='confirm'><%=myMileage%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<form name='frmTS' method='POST'>
											<table border='0' cellpadding='1' cellspacing='2' width='75%'>
												<tr>
													<td align='center' class='tblgrn'>Date</td>
													<td align='center' class='tblgrn'>Activity</td>
													<td align='center' class='tblgrn'>Mileage</td>
													<td align='center' class='tblgrn'>Tolls & parking<br>with receipts</td>
													<td align='center' class='tblgrn'>Approved</td>
													<td align='center' class='tblgrn'>Confirmed</td>
												</tr>	
												<%=strMile%>
												<tr>
													<td colspan='2'>&nbsp;</td>
													<td align='center' class='main' bgcolor='#FFFFCE'><%=TottmpAMT%></td>
													<td align='center' class='main' bgcolor='#FFFFCE'>$<%=Z_formatnumber(TotToll,2)%></td>
												</tr>
											</table>
										
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' align='center'>
										*To <b>SAVE</b> your Appointment Tolls & parking, input Tolls & parking amount then click on 'Save' button.<br>
										Once approved, you can no longer edit Tolls & parking amount.
										<br><br>
										*To <b>CONFIRM</b> your appointment Tolls & parking, check the correspoding checkbox then click on 'Confirm' button.<br>
										Appointment Tolls & parking needs to be approved by Languagebank before you can confirm it.
									</td>
								</tr>
								<tr>
									<td colspan='10' align='center' height='100px' valign='bottom'>
										<input type='hidden' name='tmpMonth' value="<%=tmpMonth%>">
										<input type='hidden' name='tmpYear' value="<%=tmpYear%>">
										<input type='hidden' name='myCTR' value="<%=ctr%>">
										<input type='hidden' name='myCTR2' value="<%=ctr2%>">
										<input class='btn' type='button' value='<Prev Month'  onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PrevMonth(<%=tmpMonth%>, <%=tmpYear%>);'>
										<input class='btn' type='button' value='Save' <%=billedna%> <%=myStat%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveTS();'>
										<input class='btn' type='button' value='Confirm' <%=billedna%> <%=myStat%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SendTS();'>
										<input class='btn' type='button' value='Next Month>'  onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='NextMonth(<%=tmpMonth%>, <%=tmpYear%>);'>
										
									</td>
								</tr>
								</form>
							</table>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
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
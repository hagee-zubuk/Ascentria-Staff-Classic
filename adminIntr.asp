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
'GET INTERPRETER LIST
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
'change sql for active /inactive
If Request("type") = 0 Then sqlIntr = "SELECT * FROM interpreter_T WHERE Active = true ORDER BY [last name], [first name]"
If Request("type") = 1 Then sqlIntr = "SELECT * FROM interpreter_T WHERE Active = false ORDER BY [last name], [first name]"
If Request("type") = 2 Then sqlIntr = "SELECT * FROM interpreter_T ORDER BY [last name], [first name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	tmpIntr = Request("IntrID")
	IntrSel = ""
	If tmpIntr = "" Then tmpIntr = "-1"
	If z_Czero(tmpIntr) = rsIntr("index") Then IntrSel = "selected"
	strIntr = strIntr	& "<option " & IntrSel & " value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbCrlf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT * FROM language_T ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	strLang2 = strLang2	& "<option  value='" & Trim(rsLang("language")) & "'>" &  rsLang("language") & "</option>" & vbCrlf
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
stat2 = ""
active2 = ""
stat1 = "checked"
active1 = "checked"
If Request.ServerVariables("REQUEST_METHOD") = "POST" OR z_Czero(Request("IntrID")) <> 0 Then
	If Request("IntrID") <> 0 Then
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SeLECT * FROM interpreter_T WHERE index = " & Request("IntrID")
		rsIntr.Open sqlIntr, g_strCONN, 1, 3
		If Not rsIntr.EOF Then
			fname = rsIntr("first name")
			lname = rsIntr("last name")
			email = rsIntr("e-mail")
			p1 = rsIntr("Phone1")
			p1ext = rsIntr("p1ext")
			p2 = rsIntr("phone2")
			fax = rsIntr("fax")
			addrI = rsIntr("IntrAdrI")
			addr = rsIntr("Address1")
			city = rsIntr("City")
			state = rsIntr("State")
			zip = rsIntr("zip code")
			inHouse = ""
			If rsIntr("inhouse") = true Then inHouse = "checked"
			crimerec = rsIntr("crimeDate")
			driverec = rsIntr("driveDate")
			hiredate = rsIntr("datehired")
			dateterm = rsIntr("dateterm")
			stat1 = ""
			stat2 = ""
			If rsIntr("stat") = 0 Then stat1 = "checked"
			If rsIntr("stat") = 1 Then stat2 = "checked"
			drivedate = rsIntr("drivedate")
			crimedate = rsIntr("crimedate")
			ssnum = rsIntr("ssnum")
			passnum = rsIntr("passnum")
			passexp = rsIntr("passexp")
			drivenum = rsIntr("drivenum")
			driveexp = rsIntr("driveexp")
			greennum = rsIntr("greennum")
			greenexp = rsIntr("greenexp")
			employnum = rsIntr("employnum")
			employexp = rsIntr("employexp")
			carnum = rsIntr("carnum")
			carexp = rsIntr("carexp")
			active1 = ""
			active2 = ""
			If rsIntr("active") = true Then active1 = "checked"
			If rsIntr("active") = false Then active2 = "checked"
			myrate = rsIntr("rate")
			tmpLang1 = rsIntr("Language1")
			tmpLang2 = rsIntr("Language2")
			tmpLang3 = rsIntr("Language3")
			tmpLang4 = rsIntr("Language4")
			tmpLang5 = rsIntr("Language5")
			vacto = rsIntr("vacto")
			vacfrom = rsIntr("vacfrom")
			filenum = rsIntr("filenum")
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	End If
End If
'GET INTERPRETER RATES
Set rsRates = Server.CreateObject("ADODB.RecordSet")
sqlRates = "SELECT * FROM rate2_T ORDER BY Rate2"
rsRates.Open sqlRates, g_strCONN, 3, 1
Do Until rsRates.EOF
	RateKo2 = ""
	If myrate = rsRates("Rate2") Then RateKo2= "selected"
	strRateIntr = strRateIntr & "<option " & RateKo2 & " value='" & rsRates("Rate2") & "'>$" & Z_FormatNumber(rsRates("Rate2"), 2) & "</option>" & vbCrLf
	rsRates.MoveNext
Loop
rsRates.Close
Set rsRates = Nothing
%>
<html>
	<head>
		<title>Language Bank - Interpreter Page</title>
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
			function FindMe(xxx, yyy)
			{
				document.frmAdmin.action = "adminIntr.asp?IntrID=" + xxx + "&type=" + yyy;
				document.frmAdmin.submit();
			}
			function SaveMe()
			{
				if (document.frmAdmin.selIntr.value == "0")
				{
					if ((document.frmAdmin.txtIntrP1.value == "" && document.frmAdmin.txtIntrP2.value == "" && document.frmAdmin.txtIntrEmail.value == "" && document.frmAdmin.txtIntrFax.value == "") && document.frmAdmin.txtIntrLname.value != "" & document.frmAdmin.txtIntrFname.value != "")
					{
						alert("ERROR: Interpreter should at least have 1 contact information.")
						return;
					}
				}
				if (document.frmAdmin.txtvacto.value != ""  && document.frmAdmin.txtvacfrom.value == "")
				{
					alert("Please enter a 'from' vaction date.")
					return;
				}
				if (document.frmAdmin.txtvacto.value == ""  && document.frmAdmin.txtvacfrom.value != "")
				{
					alert("Please enter a 'to' vaction date.")
					return;
				}
				document.frmAdmin.action = "intraction.asp?action=1";
				document.frmAdmin.submit();
			}
			function KillMe()
			{
				var ans = window.confirm("Delete Interpreter?\nClick Cancel to stop.");
				if (ans)
				{
					document.frmAdmin.action = "interaction.asp?action=2";
					document.frmAdmin.submit();
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
							<tr>
									<td colspan='2' align='center'>
										<span class='error'><%=Session("MSG")%></span>
									</td>
								</tr>
							<% If z_Czero(Request("IntrID")) = 0 Then %>
								<tr>
									<td class='header' colspan='10'><nobr>Interpreter Info</td>
								</tr>
							<% Else %>
								<tr>
									<td class='header' colspan='10'>
										<nobr>Interpreter Info
										<a href='intreval.asp?intrID=<%=Request("intrID")%>&type=<%=Request("type")%>' class='intrlink'>[Evalution/Feedback]</a>
										<a href='intrtrain.asp?intrID=<%=Request("intrID")%>&type=<%=Request("type")%>' class='intrlink'>[Training]</a>
										<a href='intrsched.asp?intrID=<%=Request("intrID")%>&type=<%=Request("type")%>' class='intrlink'>[Schedule]</a>
									</td>
								</tr>
							<% End If %>
								<tr>
									<td align='right'>&nbsp;</td>
									<td>
										<select class='seltxt' name='selIntr' onchange='FindMe(document.frmAdmin.selIntr.value, <%=Request("type")%>);'>
											<option value='0'>&nbsp;</option>
											<%=strIntr%>
										</select>
										&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Name:</td>
									<td>
										<input class='main' size='20' maxlength='20' name='txtIntrLname' value='<%=lname%>' onkeyup='bawal(this);'>
										,&nbsp;
										<input class='main' size='20' maxlength='20' name='txtIntrFname' value='<%=fname%>' onkeyup='bawal(this);'>
										<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat' readonly value='last name, first name'>
										<input type='hidden' name='IntrName'>
										<input type='hidden' name='tmpType' value='<%=Request("type")%>'>
									</td>
								</tr>
								<tr>
									<td align='right' width='15%'>E-Mail:</td>
									<td><input class='main' size='50' maxlength='50' name='txtIntrEmail' value='<%=email%>' onkeyup='bawal(this);'></td>
								</tr>
								<tr>	
									<td align='right' width='15%'>Home Phone:</td>
									<td>
										<input class='main' size='12' maxlength='12' name='txtIntrP1' value='<%=p1%>' onkeyup='bawal(this);'>
										&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtIntrExt' value='<%=p1ext%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<tr>
									<td align='right' width='15%'>Mobile Phone:</td>
									<td><input class='main' size='12' maxlength='12' name='txtIntrP2' value='<%=p2%>' onkeyup='bawal(this);'></td>
								</tr>
								<tr>
									<td align='right' width='15%'>Fax:</td>
									<td><input class='main' size='12' maxlength='12' name='txtIntrFax' value='<%=fax%>' onkeyup='bawal(this);'></td>
								</tr>
								<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtIntrAddrI' value='<%=addrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
								<tr>
									<td align='right'>Address:</td>
									<td>
										<input class='main' size='50' maxlength='50' name='txtIntrAddr' value='<%=addr%>' onkeyup='bawal(this);'>
										<br>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
									</td>
								</tr>
								<tr>
									<td align='right'>City:</td>
									<td colspan='5'>
										<input class='main' size='25' maxlength='25' name='txtIntrCity' value='<%=city%>' onkeyup='bawal(this);'>&nbsp;State:
										<input class='main' size='2' maxlength='2' name='txtIntrState' value='<%=state%>' onkeyup='bawal(this);'>&nbsp;Zip:
										<input class='main' size='10' maxlength='10' name='txtIntrZip' value='<%=zip%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<tr>
									<td align='right'><nobr>In-House:</td>
									<td><input type='checkbox' name='chkInHouse' value='1' <%=inHouse%>></td>
								</tr>
								<tr>
									<td align='right' valign='top'>File Number:</td>
									<td>
										<input class='main' size='7' maxlength='6' name='txtfilenum' value='<%=filenum%>'>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Driving Record:</td>
									<td>
										<input class='main' size='11' maxlength='10' name='txtdrivedate' value='<%=drivedate%>' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Criminal Record:</td>
									<td>
										<input class='main' size='11' maxlength='10' name='txtCrimedate' value='<%=crimedate%>' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Training:</td>
									<td>
										<input class='main' size='50' maxlength='50' name='txtTrain' value='<%=tmpTrain%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Date of Hire:</td>
									<td>
										<input class='main' size='11' maxlength='10' name='txtHire' value='<%=hiredate%>' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<table cellspacing='1' cellpadding='1' border='0'>
											<tr>
												<td>&nbsp;</td>
												<td align='center'><u>Number</u></td>
												<td align='center'><u>Expiration Date</u></td>
											</tr>
											<tr>
												<td align='right' valign='top'>Social Security:</td>
												<td align='center'><input class='main' size='26' maxlength='25' name='txtss' value='<%=ssnum%>' onkeyup='bawal(this);'></td>
											</tr>
											<tr>
												<td align='right' valign='top'>Passport:</td>
												<td align='center'><input class='main' size='26' maxlength='25' name='txtpass' value='<%=passnum%>' onkeyup='bawal(this);'></td>
												<td align='center'><input class='main' size='11' maxlength='10' name='txtpassexp' value='<%=passexp%>' onkeyup='bawal(this);'></td>
											</tr>
											<tr>
												<td align='right' valign='top'>Driver's License:</td>
												<td align='center'><input class='main' size='26' maxlength='25' name='txtdrive' value='<%=drivenum%>' onkeyup='bawal(this);'></td>
												<td align='center'><input class='main' size='11' maxlength='10' name='txtdriveexp' value='<%=driveexp%>' onkeyup='bawal(this);'></td>
											</tr>
											<tr>
												<td align='right' valign='top'>Green Card:</td>
												<td align='center'><input class='main' size='26' maxlength='25' name='txtgreen' value='<%=greennum%>' onkeyup='bawal(this);'></td>
												<td align='center'><input class='main' size='11' maxlength='10' name='txtgreenexp' value='<%=greenexp%>' onkeyup='bawal(this);'></td>
											</tr>
											<tr>
												<td align='right' valign='top'>Employment Authorization:</td>
												<td align='center'><input class='main' size='26' maxlength='25' name='txtemploy' value='<%=employnum%>' onkeyup='bawal(this);'></td>
												<td align='center'><input class='main' size='11' maxlength='10' name='txtemployexp' value='<%=employexp%>' onkeyup='bawal(this);'></td>
											</tr>
											<tr>
												<td align='right' valign='top'>Car Insurance:</td>
												<td align='center'><input class='main' size='26' maxlength='25' name='txtcar' value='<%=carnum%>' onkeyup='bawal(this);'></td>
												<td align='center'><input class='main' size='11' maxlength='10' name='txtcarexp' value='<%=carexp%>' onkeyup='bawal(this);'></td>
											</tr>
										</table>
									</td>
								</tr>
                				<tr>
									<td align='right' valign='top'>Date of Termination:</td>
									<td>
										<input class='main' size='11' maxlength='10' name='txtTerm' value='<%=dateterm%>' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
									</td>
								</tr>
								<tr><td>&nbsp;</td</tr>
								<tr>
									<td align='right'><nobr>Status:</td>
									<td>
										<input type='radio' name='radioStatIntr' value='0' <%=stat1%>>Employee
										&nbsp;&nbsp;
										<input type='radio' name='radioStatIntr' value='1' <%=stat2%>>Outside Consultant
									</td>
								</tr>
								<tr>
									<td align='right'>&nbsp;</td>
									<td>
										<input type='radio' name='radioStatIntr1' value='0' <%=active1%>>Active
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<input type='radio' name='radioStatIntr1' value='1' <%=active2%>>Inactive
									</td>
								</tr>
								<tr>
									<td align='right'>Default Rate:</td>
									<td>
										<select class='seltxt' name='selIntrRate'  style='width:75px;'>
											<option value='0'>&nbsp;</option>
											<%=strRateIntr%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>Language:</td>
									<td>
										<select class='seltxt' name='selIntrLang'  style='width:150px;'>
											<option value='0'>&nbsp;</option>
											<%=strLang2%>
										</select>
										<input class='transmall' style='width: 400px;' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat' readonly value='* To delete a language, check the corresponding checkbox then save.'>
									</td>
								</tr>
									<td>&nbsp;</td>
									<td>
										<input type='checkbox' name='chkLang1' value='1'>
										<input style='width:150px;' class='main' readonly  name='txtLang1' value='<%=tmpLang1%>'>
										&nbsp;
										<input type='checkbox' name='chkLang2' value='1'>
										<input style='width:150px;' class='main' readonly  name='txtLang2' value='<%=tmpLang2%>'>
									</td>
								<tr>
									<td>&nbsp;</td>
									<td>
										<input type='checkbox' name='chkLang3' value='1'>
										<input style='width:150px;' class='main' readonly  name='txtLang3' value='<%=tmpLang3%>'>
										&nbsp;
										<input type='checkbox' name='chkLang4' value='1'>
										<input style='width:150px;' class='main' readonly  name='txtLang4' value='<%=tmpLang4%>'>
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<input type='checkbox' name='chkLang5' value='1'>
										<input style='width:150px;' class='main' readonly  name='txtLang5' value='<%=tmpLang5%>'>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Comments:</td>
									<td>
										<input class='main' size='50' maxlength='50' name='txtIntrCom' value='<%=tmpIntrCom%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>Vacation:</td>
									<td>
										From:<input class='main' size='11' maxlength='10' name='txtvacfrom' value='<%=vacfrom%>' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										&nbsp;&nbsp;-&nbsp;&nbsp;
										To:<input class='main' size='11' maxlength='10' name='txtvacto' value='<%=vacto%>' onkeyup='bawal(this);'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
									</td>
								</tr>
								<tr>
									<td colspan='10' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' style='width: 125px;' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveMe();">
										<input class='btn' type='button' style='width: 125px;' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe();'>
									</td>
								</tr>
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
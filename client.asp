<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
''	Session("MSG") = "Invalid account."
''	Response.Redirect "default.asp"
'maybe just log it this time?
End If
tmpPage = "document.frmClient."
intCli = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST" Or Request("cliid") <> "" Then
	
	If Request("cliid") <> "" Then
		intCli = Request("cliid")
	Else
		intCli = Request("selcli")
	End If
	Set rsCli = Server.CreateObject("ADODB.RecordSet")
	sqlCli = "SELECT * FROM c_need_T WHERE UID = " & intCli
	rsCli.Open sqlCli, g_strCONN, 3, 1
	If Not rsCli.EOF Then
		tmplname = rsCli("clname")
		tmpfname = rsCli("cfname")
		tmpDOB = Z_DateNull(rsCli("dob"))
		tmpemail = rsCli("email")
		tmpcomment = rsCli("comment")
		If rsCli("asl") Then chk1 = "checked"
		If rsCli("senglish") Then chk2 = "checked"
		If rsCli("cdeaf") Then chk3 = "checked"
		If rsCli("dblind") Then chk4 = "checked"
		If rsCli("cart") Then chk5 = "checked"
		If rsCli("cspeech") Then chk6 = "checked"
		If rsCli("dlow") Then chk7 = "checked"
		If rsCli("lprint") Then chk8 = "checked"
		If rsCli("cd") Then chk9 = "checked"
		If rsCli("alist") Then chk10 = "checked"
		If rsCli("braille") Then chk11 = "checked"
		If rsCli("laptop") Then chk12 = "checked"
		If rsCli("other") Then chk13 = "checked"
	End If
	rsCli.Close
	Set rsCli = Nothing
	'get pref intr
	Set rsCli = Server.CreateObject("ADODB.RecordSet")
	sqlCli = "SELECT * FROM c_need_intr_T WHERE CID = " & intCli
	rsCli.Open sqlCli, g_strCONN, 3, 1
	ctr = 0
	Do Until rsCli.EOF
		strPref = strPref & "<tr><td><input type='checkbox' name='chkpref" & ctr & "' value='" & rsCli("UID") & "'>" & GetIntr(rsCli("intrID")) & "</td></tr>"
		ctr = ctr + 1
		rsCli.MoveNext
	Loop
	rsCli.Close
	Set rsCli = Nothing
	'get not pref intr
	Set rsCli = Server.CreateObject("ADODB.RecordSet")
	sqlCli = "SELECT * FROM c_need_intr_no_T WHERE CID = " & intCli
	rsCli.Open sqlCli, g_strCONN, 3, 1
	ctr2 = 0
	Do Until rsCli.EOF
		strNoPref = strNoPref & "<tr><td><input type='checkbox' name='chkprefex" & ctr2 & "' value='" & rsCli("UID") & "'>" & GetIntr(rsCli("intrID")) & "</td></tr>"
		ctr2 = ctr2 + 1
		rsCli.MoveNext
	Loop
	rsCli.Close
	Set rsCli = Nothing
End If
'get client list
Set rsCli = Server.CreateObject("ADODB.RecordSet")
sqlCli = "SELECT * FROM c_need_T ORDER BY clname, cfname"
rsCli.Open sqlCli, g_strCONN, 3, 1
Do Until rsCli.EOF
	selCli = ""
	If Z_CZero(intCli) = rsCli("UID") Then selCli = "selected"
	strName = rsCli("clname") & ", " & rsCli("cfname")
	strCli = strCli & "<option value='" & rsCli("UID") & "' " & selCli & " >" & strName & "</option>"
	rsCli.MoveNext
Loop
rsCli.Close
Set rsCli = Nothing
'get interpreter list
Set rsCli = Server.CreateObject("ADODB.RecordSet")
sqlCli = "SELECT [last name], [first name],  [index] as intrID FROM interpreter_T WHERE Active = 1 ORDER BY [last name], [first name]"
rsCli.Open sqlCli, g_strCONN, 3, 1
Do Until rsCli.EOF
	strName = rsCli("last name") & ", " & rsCli("first name")
	strIntr = strIntr & "<option value='" & rsCli("intrID") & "'>" & strName & "</option>"
	rsCli.MoveNext
Loop
rsCli.Close
Set rsCli = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Client-Interpreter Preferred List</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function CalendarView(strDate)
		{
			document.frmClient.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmClient.submit();
		}
		function SaveInfo()
		{
			if (document.frmClient.txtClilname.value == "") {
				alert("Client's last name is required.")
				return;
			}
			if (document.frmClient.txtClifname.value == "") {
				alert("Client's first name is required.")
				return;
			}
			document.frmClient.action = 'clientaction.asp';
			document.frmClient.submit();
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
		function bawal2(tmpform)
		{
			var iChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz0123456789-,.\'"; //",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		tmp = tmp + tmpform.value.charAt(i);
		  		}
			  	else
		  		{
		  			alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  			
		  		}
		  	}
		}
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
		//-->
		</script>
	</head>
	<body>
		<form method='post' name='frmClient' method='POST' action='client.asp'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table cellSpacing='2' cellPadding='0' width="100%" border='0' align='center' >
							<!-- #include file="_greetme.asp" -->
							<tr>
								<td class='title' colspan='10' align='center'><nobr>Client-Interpreter Preferred List</td>
							</tr>
							<tr>
								<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right'>Client:</td>
								<td>
									<select class='seltxt' name='selcli' onchange='document.frmClient.submit();' onblur='document.frmClient.submit();'>
										<option value='0'>&nbsp;</option>
										<%=strCli%>
									</select>
									<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Leave blank to add new</span>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right'>Client Last Name:</td>
								<td>
									<input class='main' size='20' maxlength='20' name='txtClilname' value="<%=tmplname%>" onkeyup='bawal2(this);'>&nbsp;First Name:
									<input class='main' size='20' maxlength='20' name='txtClifname' value="<%=tmpfname%>" onkeyup='bawal2(this);'>
								</td>
							</tr>
							<tr>
								<td align='right'>DOB:</td>
								<td>
									<input class='main' size='11' maxlength='10' name='txtDOB' value='<%=tmpDOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
									<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
								</td>
							</tr>
							<tr>
								<td align='right'>Email:</td>
								<td>
									<input class='main' size='50' maxlength='50' name='txtemail' value="<%=tmpemail%>" onkeyup='bawal(this);'>
								</td>
							</tr>
							<tr>
								<td align='right' valign='top'>Type of Communication Access:</td>
								<td>
									<table>
										<tr>
											<td>
												<input type='checkbox' name='chk1' value='1' <%=chk1%> >ASL Sign Language Interpreter
											</td>
											<td>
												<input type='checkbox' name='chk2' value='1' <%=chk2%> >Signed English Interpreter
											</td>
										</tr>
										<tr>
											<td>
												<input type='checkbox' name='chk3' value='1' <%=chk3%> >Certified Deaf Interpreter
											</td>
											<td>
												<input type='checkbox' name='chk4' value='1' <%=chk4%> >Deaf-Blind Interpreter
											</td>
										</tr>
										<tr>
											<td>
												<input type='checkbox' name='chk5' value='1' <%=chk5%> >CART Services (Real Time Captioning)
											</td>
											<td>
												<input type='checkbox' name='chk6' value='1' <%=chk6%> >Cued Speech Interpreter
											</td>
										</tr>
										<tr>
											<td>
												<input type='checkbox' name='chk7' value='1' <%=chk7%> >Deaf-Low Vision Interpreter
											</td>
											<td>
												<input type='checkbox' name='chk8' value='1' <%=chk8%> >Large Print Version
											</td>
										</tr>
										<tr>
											<td>
												<input type='checkbox' name='chk9' value='1' <%=chk9%> >CD Version of Printed Information
											</td>
											<td>
												<input type='checkbox' name='chk10' value='1' <%=chk10%> >Assisted Listening Device
											</td>
										</tr>
										<tr>
											<td>
												<input type='checkbox' name='chk11' value='1' <%=chk11%> >Braille Version
											</td>
											<td>
												<input type='checkbox' name='chk12' value='1' <%=chk12%> >Laptop to Type On (for client with speech impairments)
											</td>
										</tr>
										<tr>
											<td>
												<input type='checkbox' name='chk13' value='1' <%=chk13%> >Other
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td align='right' valign='top'>Preferred Interpreter:</td>
								<td>
									<select class='seltxt' name='selintr'>
										<option value='0'>&nbsp;</option>
										<%=strIntr%>
									</select>
									<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*To remove interpreter, check checkbox then save</span>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>
									<table>
										<%=strPref%>
									</table>
								</td>
							</tr>
							<tr>
								<td align='right' valign='top'>Exclude Interpreter:</td>
								<td>
									<select class='seltxt' name='selintrex'>
										<option value='0'>&nbsp;</option>
										<%=strIntr%>
									</select>
									<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*To remove interpreter, check checkbox then save</span>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>
									<table>
										<%=strNoPref%>
									</table>
								</td>
							</tr>
							<tr>
								<td align='right' valign='top'>Comment:</td>

								<td>
									<textarea name='txtcom' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpcomment%></textarea>
								</td>

							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan='10' align='center' height='100px' valign='bottom'>
							<input type='hidden' name='ctr' value='<%=ctr%>'>
							<input type='hidden' name='ctr2' value='<%=ctr2%>'>
							<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveInfo();'>
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
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
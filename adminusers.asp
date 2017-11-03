<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
server.scripttimeout = 360000
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then
	Session("MSG") = "Error: Please sign-in as user with admin rights."
	Response.Redirect "default.asp"
End If
tmpPage = "document.frmAdmin."
'GET USER LIST
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT * FROM user_T ORDER BY [lname], [fname]"
rsUser.Open sqlUser, g_strCONN, 3, 1
Do Until rsUser.EOF
	tmpUser = Request("UserID")
	UserSel = ""
	If tmpUser = "" Then tmpIntr = "-1"
	If Z_Czero(tmpUser) = rsUser("index") Then UserSel = "selected"
	strUser = strUser	& "<option " & UserSel & " value='" & rsUser("Index") & "'>" & rsUser("lname") & ", " & rsUser("fname") & "</option>" & vbCrlf
	'strChkUser = strChkUser & "if (Ucase(document.frmAdmin.txtUserUname.value) == """ & UCase(rsUser("username")) & """ && document.frmAdmin.selUser.value == -1) " & vbCrLf & _
	'	"{alert(""ERROR: Username already exists.""); return;}" & vbCrLf
	'strChkUser2 = strChkUser2 & "if (Ucase(document.frmAdmin.txtUserUname.value) == """ & UCase(rsUser("username")) & """ && document.frmAdmin.selUser.value != -1) " & vbCrLf & _
	'	"{alert(""ERROR: Username already exists.""); return;}" & vbCrLf
	rsUser.MoveNext
Loop
rsUser.Close
Set rsUser = Nothing
If Request.ServerVariables("REQUEST_METHOD") = "POST" OR z_Czero(Request("userID")) <> 0 Then
	If Request("userID") <> 0 Then
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SeLECT * FROM user_T WHERE index = " & Request("userID")
		rsIntr.Open sqlIntr, g_strCONN, 1, 3
		If Not rsIntr.EOF Then
			tmpUserFname = rsIntr("fname")
			tmpUserLname = rsIntr("lname")
			tmpUserUname = rsIntr("username")
			pw = z_doDecrypt(rsIntr("password"))
			mytype = rsIntr("type")
			instID = rsIntr("instID")
			intrID = rsIntr("intrID")
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	End If
End If
uadmin0 = ""
uadmin1 = ""
uadmin2 = ""
uadmin3 = ""
uadmin4 = ""
If mytype = 0 Then 
	uadmin0 = "selected"
elseif mytype = 1 then
	uadmin1 = "selected"
elseif mytype = 2 then
	uadmin2 = "selected"
elseif mytype = 3 then
	uadmin3 = "selected"
elseif mytype = 4 then
	uadmin4 = "selected"
end if
'GET INTERPRETER
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE Active = True ORDER BY [last name], [first name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	IntrSel = ""
	If Z_Czero(intrID) = rsIntr("index") Then IntrSel = "selected"
	strIntr = strIntr	& "<option " & IntrSel & " value='" & rsIntr("Index") & "'>" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbCrlf
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsIntr = Nothing
%>
<html>
	<head>
		<title>Language Bank - Administrator - User Page</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script type='text/javascript' language='JavaScript'>
		<!--
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
		function FindMe(xxx)
		{
			document.frmAdmin.action = "adminusers.asp?userID=" + xxx;
			document.frmAdmin.submit();
		}
		function ChkPriv()
		{
			if (document.frmAdmin.selType.value == 2)
			{
				document.frmAdmin.selIntr2.disabled = false;
			}
			else
			{
				document.frmAdmin.selIntr2.value = 0;
				document.frmAdmin.selIntr2.disabled = true;
			}	
		}
		function KillMe()
		{
			var ans = window.confirm("DELETE user?");
			if (ans)
			{
				document.frmAdmin.action = "adminuseraction.asp?ctrl=2";
				document.frmAdmin.submit();
			}
		}
		function SaveMe()
		{
			document.frmAdmin.action = "adminuseraction.asp?ctrl=1";
			document.frmAdmin.submit();
		}
		-->
		</script>
	</head>
	<body onload='ChkPriv();'>
		<form method='post' name='frmAdmin'>
			<table cellSpacing='0' cellPadding='0'  width="100%" border='0' class='bgstyle2'>
				<tr>
					<td height='100px' valign='top' colspan='2'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<!-- #include file="_greetme.asp" -->
				<tr>
				<tr>
					<td class='header' colspan='10'><nobr>User</td>
				</tr>
				<tr>
					<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
				</tr>
				<tr>
					<td align='right'>&nbsp;</td>
					<td>
						<select class='seltxt' name='selUser' onchange='FindMe(document.frmAdmin.selUser.value);'>
							<option value='0'>&nbsp;</option>
							<%=strUser%>
						</select>
						&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
					</td>
				</tr>
				<tr>
					<td align='right'>Name:</td>
					<td>
						<input class='main' size='20' maxlength='20' name='txtUserLname' value='<%=tmpUserLname%>' onkeyup='bawal(this);'>
						,&nbsp;
						<input class='main' size='20' maxlength='20' name='txtUserFname' value='<%=tmpUserFname%>' onkeyup='bawal(this);'>
						<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat' readonly value='last name, first name'>
						<input type='hidden' name='UserName'>
						<input type='hidden' name='xUserName'>
					</td>
				</tr>
				<tr>
					<td align='right'>Username:</td>
					<td><input size='50' class='main' maxlength='50' name='txtUserUname' value='<%=tmpUserUname%>' onkeyup='bawal(this);'></td>
				</tr>
				<tr>
					<td align='right' width='15%'>Password:</td>
					<td><input type='password' class='main' size='50' maxlength='50' name='txtUserPword' value='<%=pw%>' onkeyup='bawal(this);'></td>
				</tr>
				<tr>
					<td align='right' width='15%'><nobr>Confirm Password:</td>
					<td><input type='password' class='main' size='50' maxlength='50' name='txtUserPword2' value='<%=pw%>' onkeyup='bawal(this);'></td>
				</tr>
				<tr>
					<td align='right'>Privilege:</td>
					<td>
						<select class='seltxt' name='selType' style='width:150px;' onchange='ChkPriv();'>
							<option value='0' <%=Uadmin0%>>Level 1</option>
							<option value='3' <%=Uadmin3%>>Level 2</option>
							<option value='4' <%=Uadmin4%>>Observer</option>
							<option value='1' <%=Uadmin1%>>Administrator</option>
							<option value='2' <%=Uadmin2%>>Interpreter</option>
						</select>
					</td>
				</tr>
				<tr>
					<td align='right'>Interpreter:</td>
					<td width='350px'>
						<select class='seltxt' name='selIntr2'  style='width:250px;'>
							<option value='0'>&nbsp;</option>
							<%=strIntr%>
						</select>
						<input type='hidden' name='hidIntr'>
					</td>
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
							
				<tr>
									<td colspan='10' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' style='width: 125px;' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveMe();">
										<input class='btn' type='button' style='width: 125px;' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe();'>
									</td>
								</tr>
			
				<tr>
					<td height='50px' valign='bottom' colspan='2'>
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
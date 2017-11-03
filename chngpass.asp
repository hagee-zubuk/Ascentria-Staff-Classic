<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
If Session("UIntr") = "" Then 
	Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
	Response.Redirect "default.asp"
End If
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Set rsUsr = Server.CreateObject("ADODB.RecordSet")
	sqlUsr = "SELECT * FROM user_T WHERE [index] = " & Request.Cookies("UID") 'Session("UIntr")
	rsUsr.Open sqlUsr, g_strCONN, 1, 3
	If Not rsUsr.EOF Then
		If Z_DoDecrypt(rsUsr("Password")) = Request("txtOPW") Then
			rsUsr("Password") = Z_DoEncrypt(Request("txtNPW"))
			rsUsr("reset") = False
			rsUsr.Update
			Session("MSG") = "New Password saved."
		Else
			Session("MSG") = "ERROR: Invalid password."
			err = 1
		End If
	End If
	rsUsr.CLose
	Set rsUsr = Nothing
	If Not err = 1 Then response.redirect "calendarview2.asp"
		
End If
%>
<html>
	<head>
		<title>Language Bank - Interpreter - Change Password</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function savePass()
		{
			if (document.frmTS.txtNPW.value != document.frmTS.txtCNPW.value)
			{
				alert("Password is not the same.")
				return;
			}
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
								<td class='title' colspan='10' align='center'><nobr> Interpreter - Change Password</td>
								</tr>
								<tr>
									<td  align='center' colspan='12'>
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
									<td>&nbsp;</td>
									<td>
										<form name='frmTS' method='POST'>
											<table border='0' cellpadding='1' cellspacing='2' width='75%'>
												<tr>
													<td align='right'>Old Password:</td>
													<td align='left'><input type='password' class='main' style='width: 130px;' maxlength='20' name='txtOPW'></td>
												</tr>
												<tr>
													<td align='right'>New Password:</td>
													<td align='left'><input type='password' class='main' style='width: 130px;' maxlength='20' name='txtNPW'></td>
												</tr>
												<tr>
													<td align='right'>Confirm New Password:</td>
													<td align='left'><input type='password' class='main' style='width: 130px;' maxlength='20' name='txtCNPW'></td>
												</tr>
											</table>
										
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
								<tr>
									<td colspan='12' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Save' <%=billedna%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='savePass();'>
										<input class='btn' type='reset' value='Clear' <%=billedna%> onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">

										
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
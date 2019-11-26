<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_utilsMedicaid.asp" -->
<%
	updatemedicad = false
	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		updatemedicad = true
		tmpmedicaid = Z_getmedicaid(request("reqid"))
		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		rsTBL.Open "SELECT * FROM medapprove_T WHERE medicaid = '" & tmpmedicaid & "'", g_strCONN, 3, 1
		if Not rstbl.EOF Then
			newlname = rsTBL("lname") 
			newfname = rsTBL("fname") 
			newdob = rsTBL("dob")
			'newgender = rsTBL("gender")
		end if
		rstbl.close
		set rstbl = nothing
		'If request("chkall") = 1 then
		'	set rsmed = server.createobject("adodb.recordset")
		'	rsmed.open "select clname, cfname, dob, hpid from request_t where medicaid = '" & tmpmedicaid & "' AND (status = 4 OR status = 1) AND (vermed = 0 OR vermed IS NULL)" , g_strconn, 1, 3
		'	Do until rsmed.eof
		'		rsmed("clname") = newlname
		'		rsmed("cfname") = newfname
		'		rsmed("dob") = newdob
		'		'rsmed("gender") = newgender
		'		rsmed.update
		'		hpid = rsmed("hpid")
		'	end if
		'	rsmed.close
		'	set rsmed = nothing
		'	if Z_czero(hpid) > 0 Then
		'		set rsmed = server.createobject("adodb.recordset")
		'		rsmed.open "select clname, cfname, dob, gender from appointment_T where [index] = " & hpid, g_strconnhp, 1, 3
		'		if not rsmed.eof then
		'			rsmed("clname") = z_doencrypt(newlname)
		'			rsmed("cfname") = z_doencrypt(newfname)
		'			rsmed("dob") = newdob
		'			'rsmed("gender") = newgender
		'			rsmed.update
		'		end if
		'		rsmed.close
		'		set rsmed = nothing
		'	end if
		'else
			set rsmed = server.createobject("adodb.recordset")
			rsmed.open "select clname, cfname, dob, hpid from request_t where [index] = " & request("reqid"), g_strconn, 1, 3
			if not rsmed.eof then
				rsmed("clname") = newlname
				rsmed("cfname") = newfname
				rsmed("dob") = newdob
				'rsmed("gender") = newgender
				rsmed.update
				hpid = rsmed("hpid")
			end if
			rsmed.close
			set rsmed = nothing
			if Z_czero(hpid) > 0 Then
				set rsmed = server.createobject("adodb.recordset")
				rsmed.open "select clname, cfname, dob, gender from appointment_T where [index] = " & hpid, g_strconnhp, 1, 3
				if not rsmed.eof then
					rsmed("clname") = z_doencrypt(newlname)
					rsmed("cfname") = z_doencrypt(newfname)
					rsmed("dob") = newdob
					'rsmed("gender") = newgender
					rsmed.update
				end if
				rsmed.close
				set rsmed = nothing
			end if
		'end if
		session("msg") = "Medicaid info updated. Please refresh the table to reflect the changes."
	End If
	Set rsMed = Server.CreateObject("ADODB.RecordSet")
	rsMed.Open "SELECT clname, cfname, dob, amerihealth, medicaid, meridian, nhhealth, wellsense, gender FROM request_T WHERE [index] = " & Request("ReqID"), g_strCONN, 3, 1
	If Not rsMed.EOF Then
		lname = Trim(Ucase(rsMed("clname")))
		fname = Trim(Ucase(rsMed("cfname")))
		dob = rsMed("dob")
		'hmo = Trim(Ucase(Z_FixNull(rsMed("medicaid")))) 
		'If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("meridian"))))
		'If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("nhhealth"))))
		'If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("wellsense")))) 
		hmo = Z_FixNull(Ucase(Trim(rsMed("medicaid")))) 
		If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("amerihealth"))))
		If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("meridian"))))
		If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("nhhealth"))))
		If hmo = "" Then hmo = Trim(Ucase(Z_FixNull(rsMed("wellsense")))) 

		gender = "Unknown"
		If Not IsNull(rsmed("gender"))  Then
			If rsmed("gender") = 1 Then
				gender = "Female"
			ElseIf rsmed("gender") = 0 Then
				gender = "Male"
			End If
		End If
	End If
	rsmed.close
	set rsmed = nothing
	Set rsTBL = Server.CreateObject("ADODB.RecordSet")
	rsTBL.Open "SELECT * FROM medapprove_T WHERE medicaid = '" & hmo & "'", g_strCONN, 3, 1
	if Not rstbl.EOF Then
		mlname = rsTBL("lname") 
		mfname = rsTBL("fname") 
		mdob = rsTBL("dob")
		mgender = "Unknown"
		If Not IsNull( rsTBL("gender") ) Then
			If rsTBL("gender") = 1 Then
				mgender = "Female"
			ElseIf rsTBL("gender") = 0 Then
				mgender = "Male"
			End If
		End If

		mmedicaid = rsTBL("medicaid")
	end if
	rstbl.close
	set rstbl = nothing
%>
<html>
	<head>
		<title>Language Bank - Medicaid Checker</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		<% if updatemedicad then %>
			
		<% end If %>
		function updatedmed(reqid) {
			document.frmMed.action = "medicaidcheck.asp?reqid=" + reqid;
			document.frmMed.submit();
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmMed' action=''>
			<table cellpadding='0' cellspacing='0' border='0' align='left' width='100%'> 
				<tr>
					<td height='25px'>&nbsp;</td>
					<td class='header' colspan='6'>
						<nobr>Medicaid Checker --&gt&gt
					</td>
				</tr>
				<tr>
					<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left' valign="top" width="150px;">
						Appointment Data:
					</td>
					<td align="left">
						<table>
							<tr>
								<td align="right">Last Name:</td>
								<td align="left"><%=lname%></td>
							</tr>
							<tr>
								<td align="right">First Name:</td>
								<td align="left"><%=fname%></td>
							</tr>
							<tr>
								<td align="right">DOB:</td>
								<td align="left"><%=dob%></td>
							</tr>
							<!--<tr>
								<td align="right">Gender:</td>
								<td align="left"><%=gender%></td>
							</tr>//-->
							<tr>
								<td align="right">Medicaid:</td>
								<td align="left"><%=hmo%></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr><td colspan='6'><hr align='center' width='75%'></td></tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left' valign="top">
						Approved Medicaid Data:
					</td>
					<td align="left">
						<table>
							<tr>
								<td align="right">Last Name:</td>
								<td align="left"><%=mlname%></td>
							</tr>
							<tr>
								<td align="right">First Name:</td>
								<td align="left"><%=mfname%></td>
							</tr>
							<tr>
								<td align="right">DOB:</td>
								<td align="left"><%=mdob%></td>
							</tr>
							<!--<tr>
								<td align="right">Gender:</td>
								<td align="left"><%=mgender%></td>
							</tr>//-->
							<tr>
								<td align="right">Medicaid:</td>
								<td align="left"><%=mmedicaid%></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr><td height='35px'>&nbsp;</td></tr>
				<tr>
					<td align='center' colspan='10' class='RemME'>
						<!--<input type='checkbox' name='chkall' value='1'>Include ALL appointments with this medicaid number which are for review<br>//-->
						<input class='btn' type='button' style='width: 253px;' value='Update Data' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="updatedmed(<%=Request("ReqID")%>);">
						
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